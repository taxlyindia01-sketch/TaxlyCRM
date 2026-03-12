from fastapi import APIRouter, HTTPException, Response, Request, Depends
from fastapi.responses import HTMLResponse, RedirectResponse
from pydantic import BaseModel, EmailStr
from typing import Optional
from datetime import datetime, timezone, timedelta
import uuid
import bcrypt
import httpx
import logging
import os
import secrets
import json

from sqlalchemy import select, update, delete
from sqlalchemy.ext.asyncio import AsyncSession

from database import (
    User, UserSession, SeriesConfig,
    serialize_doc, get_db
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/auth", tags=["Authentication"])

# ── Google OAuth config  ──────────────────────────────────────────────────
# Read lazily inside functions so load_dotenv() has already run by the time
# these are accessed (top-level reads happen before server.py calls load_dotenv)
def _google_client_id()     -> str: return os.environ.get("GOOGLE_CLIENT_ID", "")
def _google_client_secret() -> str: return os.environ.get("GOOGLE_CLIENT_SECRET", "")
def _google_redirect_uri(request: Request) -> str:
    """Return configured GOOGLE_REDIRECT_URI or auto-derive from the request."""
    configured = os.environ.get("GOOGLE_REDIRECT_URI", "")
    if configured:
        return configured
    # Auto-derive: same scheme+host as backend, at /api/auth/google/callback
    base = str(request.base_url).rstrip("/")
    return f"{base}/api/auth/google/callback"

# In-memory store for short-lived OAuth state -> frontend redirect_uri mapping
_oauth_state: dict[str, str] = {}
# In-memory store for completed OAuth sessions keyed by session_id
_oauth_sessions: dict[str, dict] = {}

DEMO_PERIOD_DAYS = 10


class UserRegister(BaseModel):
    email: EmailStr
    password: str
    name: str


class UserLogin(BaseModel):
    email: EmailStr
    password: str


def check_user_access(user: dict) -> tuple[bool, str]:
    is_approved = user.get('is_approved', False)
    if is_approved:
        return True, "approved"

    demo_expires_at = user.get('demo_expires_at')
    if demo_expires_at:
        if isinstance(demo_expires_at, str):
            demo_expires_at = datetime.fromisoformat(demo_expires_at.replace('Z', '+00:00'))
        if demo_expires_at.tzinfo is None:
            demo_expires_at = demo_expires_at.replace(tzinfo=timezone.utc)
        if datetime.now(timezone.utc) < demo_expires_at:
            days_remaining = (demo_expires_at - datetime.now(timezone.utc)).days
            return True, f"demo:{days_remaining}"

    return False, "expired"


async def get_current_user(request: Request, db: AsyncSession = Depends(get_db)) -> dict:
    session_token = request.cookies.get("session_token")
    if not session_token:
        auth_header = request.headers.get("Authorization")
        if auth_header and auth_header.startswith("Bearer "):
            session_token = auth_header.split(" ")[1]
    if not session_token:
        raise HTTPException(status_code=401, detail="Not authenticated")

    result = await db.execute(select(UserSession).where(UserSession.session_token == session_token))
    session_doc = result.scalar_one_or_none()
    if not session_doc:
        raise HTTPException(status_code=401, detail="Invalid session")

    expires_at = session_doc.expires_at
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
    if expires_at < datetime.now(timezone.utc):
        raise HTTPException(status_code=401, detail="Session expired")

    user_result = await db.execute(select(User).where(User.user_id == session_doc.user_id))
    user = user_result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=401, detail="User not found")

    if not user.is_active:
        raise HTTPException(status_code=403, detail="Account disabled by admin")

    user_dict = serialize_doc(user)
    has_access, access_status = check_user_access(user_dict)
    if not has_access:
        raise HTTPException(
            status_code=403,
            detail="Your 10-day demo has expired. Please contact admin for account approval."
        )

    user_dict['access_status'] = access_status
    return user_dict


@router.post("/register")
async def register(user_data: UserRegister, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(User).where(User.email == user_data.email))
    if result.scalar_one_or_none():
        raise HTTPException(status_code=400, detail="Email already registered")

    password_hash = bcrypt.hashpw(user_data.password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    user_id = f"user_{uuid.uuid4().hex[:12]}"
    tenant_id = f"tenant_{uuid.uuid4().hex[:12]}"
    demo_expires_at = datetime.now(timezone.utc) + timedelta(days=DEMO_PERIOD_DAYS)

    new_user = User(
        user_id=user_id,
        email=user_data.email,
        name=user_data.name,
        password_hash=password_hash,
        tenant_id=tenant_id,
        is_active=True,
        is_approved=False,
        approval_status="demo",
        demo_expires_at=demo_expires_at,
        signup_method="email",
        created_at=datetime.now(timezone.utc)
    )
    db.add(new_user)

    series = SeriesConfig(
        tenant_id=tenant_id,
        invoice_prefix="INV", invoice_counter=0,
        estimate_prefix="EST", estimate_counter=0,
        credit_note_prefix="CN", credit_note_counter=0
    )
    db.add(series)

    session_token = f"session_{uuid.uuid4().hex}"
    new_session = UserSession(
        session_token=session_token,
        user_id=user_id,
        expires_at=datetime.now(timezone.utc) + timedelta(days=7),
        created_at=datetime.now(timezone.utc)
    )
    db.add(new_session)
    await db.commit()

    return {
        "message": f"Registration successful! You have a {DEMO_PERIOD_DAYS}-day free demo period.",
        "user_id": user_id,
        "email": user_data.email,
        "name": user_data.name,
        "tenant_id": tenant_id,
        "session_token": session_token,
        "is_approved": False,
        "demo_expires_at": demo_expires_at.isoformat(),
        "demo_days_remaining": DEMO_PERIOD_DAYS,
        "access_status": f"demo:{DEMO_PERIOD_DAYS}"
    }


@router.post("/login")
async def login(request: Request, user_data: UserLogin, response: Response, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(User).where(User.email == user_data.email))
    user = result.scalar_one_or_none()

    if not user:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    if not user.password_hash:
        raise HTTPException(status_code=401, detail="This account was created with Google. Please use the 'Sign in with Google' button.")
    if not bcrypt.checkpw(user_data.password.encode('utf-8'), user.password_hash.encode('utf-8')):
        raise HTTPException(status_code=401, detail="Invalid credentials")
    if not user.is_active:
        raise HTTPException(status_code=403, detail="Account disabled by admin")

    user_dict = serialize_doc(user)
    has_access, access_status = check_user_access(user_dict)
    if not has_access:
        raise HTTPException(status_code=403, detail="Your 10-day demo has expired. Please contact admin for account approval.")

    session_token = f"session_{uuid.uuid4().hex}"
    new_session = UserSession(
        session_token=session_token,
        user_id=user.user_id,
        expires_at=datetime.now(timezone.utc) + timedelta(days=7),
        created_at=datetime.now(timezone.utc)
    )
    db.add(new_session)
    await db.commit()

    _is_https = str(request.url).startswith("https")
    response.set_cookie(key="session_token", value=session_token, httponly=True,
                        secure=_is_https, samesite="none" if _is_https else "lax",
                        max_age=604800, path="/")

    demo_info = {}
    if not user.is_approved and user.demo_expires_at:
        expires_at = user.demo_expires_at
        if expires_at.tzinfo is None:
            expires_at = expires_at.replace(tzinfo=timezone.utc)
        days_remaining = (expires_at - datetime.now(timezone.utc)).days
        demo_info = {"demo_expires_at": expires_at.isoformat(), "demo_days_remaining": max(0, days_remaining)}

    return {
        "user_id": user.user_id,
        "email": user.email,
        "name": user.name,
        "tenant_id": user.tenant_id,
        "session_token": session_token,
        "is_approved": user.is_approved,
        "access_status": access_status,
        **demo_info
    }


# ── Helper: upsert Google user and return session response data ────────────
async def _upsert_google_user_and_create_session(
    google_user: dict,
    response: Response,
    db: AsyncSession,
) -> dict:
    """Given a verified google_user dict (email, name, picture, sub),
    create or update the local User record and mint a session token."""
    result = await db.execute(select(User).where(User.email == google_user["email"]))
    user = result.scalar_one_or_none()
    is_new_user = False

    if user:
        user.name = google_user["name"]
        user.picture = google_user.get("picture")
        await db.commit()
        await db.refresh(user)
    else:
        is_new_user = True
        user_id = f"user_{uuid.uuid4().hex[:12]}"
        tenant_id = f"tenant_{uuid.uuid4().hex[:12]}"
        demo_expires_at = datetime.now(timezone.utc) + timedelta(days=DEMO_PERIOD_DAYS)

        user = User(
            user_id=user_id,
            email=google_user["email"],
            name=google_user["name"],
            tenant_id=tenant_id,
            picture=google_user.get("picture"),
            is_active=True,
            is_approved=False,
            approval_status="demo",
            demo_expires_at=demo_expires_at,
            signup_method="google",
            created_at=datetime.now(timezone.utc),
        )
        db.add(user)
        db.add(SeriesConfig(
            tenant_id=tenant_id,
            invoice_prefix="INV", invoice_counter=0,
            estimate_prefix="EST", estimate_counter=0,
            credit_note_prefix="CN", credit_note_counter=0,
        ))
        await db.commit()
        await db.refresh(user)

    if not user.is_active:
        raise HTTPException(status_code=403, detail="Account disabled by admin")

    user_dict = serialize_doc(user)
    has_access, access_status = check_user_access(user_dict)
    if not has_access:
        raise HTTPException(status_code=403, detail="Your 10-day demo has expired. Please contact admin for account approval.")

    session_token = secrets.token_hex(32)
    db.add(UserSession(
        session_token=session_token,
        user_id=user.user_id,
        expires_at=datetime.now(timezone.utc) + timedelta(days=7),
        created_at=datetime.now(timezone.utc),
    ))
    await db.commit()

    # Cookie security adapts to http vs https automatically
    response.set_cookie(
        key="session_token", value=session_token, httponly=True,
        secure=False, samesite="lax", max_age=604800, path="/",
    )

    demo_info: dict = {}
    if not user.is_approved and user.demo_expires_at:
        exp = user.demo_expires_at
        if exp.tzinfo is None:
            exp = exp.replace(tzinfo=timezone.utc)
        days_remaining = (exp - datetime.now(timezone.utc)).days
        demo_info = {"demo_expires_at": exp.isoformat(), "demo_days_remaining": max(0, days_remaining)}

    return {
        "user_id": user.user_id,
        "email": user.email,
        "name": user.name,
        "tenant_id": user.tenant_id,
        "picture": user.picture,
        "session_token": session_token,
        "is_approved": user.is_approved,
        "access_status": access_status,
        "is_new_user": is_new_user,
        **demo_info,
    }



# ── GET /auth/config-check  ─  verify env vars are loaded (dev helper) ────
@router.get("/config-check")
async def config_check():
    """Quick sanity check — shows which OAuth env vars are present."""
    cid = _google_client_id()
    csec = _google_client_secret()
    return {
        "google_client_id_set": bool(cid),
        "google_client_id_prefix": cid[:16] + "..." if cid else None,
        "google_client_secret_set": bool(csec),
        "google_redirect_uri_env": os.environ.get("GOOGLE_REDIRECT_URI", "(auto-derive from request)"),
        "env_hint": "If google_client_id_set is false, make sure .env exists in backend/ and restart uvicorn",
    }


# ── GET /auth/google/init  ─  redirect browser to Google consent screen ───
@router.get("/google/init")
async def google_oauth_init(request: Request, redirect_uri: str = ""):
    """
    Frontend opens this URL in a popup. We build the Google consent URL,
    save the state -> redirect_uri mapping, then redirect the popup there.
    """
    # Re-load env in case server started without .env (belt-and-suspenders)
    from pathlib import Path
    from dotenv import load_dotenv
    _env = Path(__file__).parent.parent / ".env"
    _env_ex = Path(__file__).parent.parent / ".env.example"
    if _env.exists():
        load_dotenv(_env, override=True)
    elif _env_ex.exists():
        load_dotenv(_env_ex, override=True)

    client_id = _google_client_id()
    if not client_id:
        return HTMLResponse("""
<html><head><style>
  body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;
       height:100vh;margin:0;background:#fef2f2;}
  .box{background:#fff;border:1px solid #fca5a5;border-radius:12px;padding:2rem;
       max-width:420px;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.1);}
  h2{color:#dc2626;margin:0 0 .5rem}
  p{color:#6b7280;font-size:.9rem;margin:.5rem 0}
  code{background:#f1f5f9;padding:2px 6px;border-radius:4px;font-size:.85rem}
</style></head><body><div class="box">
  <h2>⚙️ Google OAuth Not Configured</h2>
  <p>Add these to your <code>.env</code> file:</p>
  <p><code>GOOGLE_CLIENT_ID=your_client_id</code></p>
  <p><code>GOOGLE_CLIENT_SECRET=your_secret</code></p>
  <p>Then restart the backend server.</p>
  <br><button onclick="window.close()" style="background:#2563eb;color:#fff;border:none;
    padding:.5rem 1.2rem;border-radius:6px;cursor:pointer;font-size:.9rem">Close</button>
</div></body></html>""", status_code=501)

    state = secrets.token_urlsafe(24)
    _oauth_state[state] = redirect_uri

    backend_callback = _google_redirect_uri(request)
    from urllib.parse import quote
    params = (
        "https://accounts.google.com/o/oauth2/v2/auth"
        f"?client_id={quote(client_id)}"
        f"&redirect_uri={quote(backend_callback)}"
        "&response_type=code"
        "&scope=openid%20email%20profile"
        f"&state={state}"
        "&access_type=offline"
        "&prompt=select_account"
    )
    return RedirectResponse(url=params)


# ── GET /auth/google/callback  ─  Google posts code here ─────────────────
@router.get("/google/callback")
async def google_oauth_callback(
    code: str = "",
    state: str = "",
    error: str = "",
    request: Request = None,
    response: Response = None,
    db: AsyncSession = Depends(get_db),
):
    """
    Google redirects here after the user approves. We exchange the code for
    tokens, fetch the user profile, mint our own session, then redirect the
    popup back to the frontend with ?oauth=1&session_id=<sid>.
    """
    if error:
        return HTMLResponse(f"<script>window.close();</script>", status_code=200)

    frontend_redirect = _oauth_state.pop(state, "")
    if not frontend_redirect:
        raise HTTPException(status_code=400, detail="Invalid OAuth state")

    client_id     = _google_client_id()
    client_secret = _google_client_secret()
    if not client_id or not client_secret:
        raise HTTPException(status_code=501, detail="Google OAuth not configured")

    backend_callback = _google_redirect_uri(request)

    # Exchange auth code for tokens
    async with httpx.AsyncClient() as client:
        token_res = await client.post(
            "https://oauth2.googleapis.com/token",
            data={
                "code": code,
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uri": backend_callback,
                "grant_type": "authorization_code",
            },
            timeout=15.0,
        )
        if token_res.status_code != 200:
            logger.error(f"Token exchange failed: {token_res.text}")
            raise HTTPException(status_code=400, detail="Failed to exchange Google auth code")

        token_data = token_res.json()
        access_token = token_data.get("access_token")

        # Fetch user profile from Google
        profile_res = await client.get(
            "https://www.googleapis.com/oauth2/v2/userinfo",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=10.0,
        )
        if profile_res.status_code != 200:
            raise HTTPException(status_code=400, detail="Failed to fetch Google profile")

        google_profile = profile_res.json()

    google_user = {
        "email": google_profile.get("email"),
        "name": google_profile.get("name", google_profile.get("email", "").split("@")[0]),
        "picture": google_profile.get("picture"),
        "sub": google_profile.get("id"),
    }

    # Upsert user + create session
    session_data = await _upsert_google_user_and_create_session(google_user, response, db)

    # Store session data briefly so the frontend can retrieve it by session_id
    sid = secrets.token_urlsafe(24)
    _oauth_sessions[sid] = session_data

    # Build final redirect URL back to the frontend popup
    sep = "&" if "?" in frontend_redirect else "?"
    final_url = f"{frontend_redirect}{sep}session_id={sid}"
    return RedirectResponse(url=final_url)


# ── POST /auth/session  ─  frontend exchanges session_id for full user data ─
@router.post("/session")
async def create_session_from_google(request: Request, response: Response, db: AsyncSession = Depends(get_db)):
    """
    After the OAuth popup closes, the frontend posts the session_id it received
    in the URL. We look it up in our in-memory store and return the full user object.
    """
    session_id = request.headers.get("X-Session-ID")
    if not session_id:
        raise HTTPException(status_code=400, detail="Session ID required")

    session_data = _oauth_sessions.pop(session_id, None)
    if not session_data:
        raise HTTPException(status_code=400, detail="Invalid or expired OAuth session")

    # Re-set the cookie on this response so the browser gets it on the main window too
    session_token = session_data.get("session_token", "")
    # Use samesite=lax + secure=False for http (localhost dev); strict https in prod
    is_https = str(request.base_url).startswith("https")
    response.set_cookie(
        key="session_token", value=session_token, httponly=True,
        secure=is_https, samesite="none" if is_https else "lax",
        max_age=604800, path="/",
    )
    return session_data


@router.get("/me")
async def get_me(user: dict = Depends(get_current_user)):
    return user


@router.post("/logout")
async def logout(request: Request, response: Response, db: AsyncSession = Depends(get_db)):
    session_token = request.cookies.get("session_token")
    if session_token:
        await db.execute(delete(UserSession).where(UserSession.session_token == session_token))
        await db.commit()
    response.delete_cookie("session_token", path="/")
    return {"message": "Logged out successfully"}
