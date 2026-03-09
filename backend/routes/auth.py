from fastapi import APIRouter, HTTPException, Response, Request, Depends
from pydantic import BaseModel, EmailStr
from typing import Optional
from datetime import datetime, timezone, timedelta
import uuid
import bcrypt
import httpx
import logging

from sqlalchemy import select, update, delete
from sqlalchemy.ext.asyncio import AsyncSession

from database import (
    User, UserSession, SeriesConfig,
    serialize_doc, get_db
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/auth", tags=["Authentication"])

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
async def login(user_data: UserLogin, response: Response, db: AsyncSession = Depends(get_db)):
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

    response.set_cookie(key="session_token", value=session_token, httponly=True,
                        secure=True, samesite="none", max_age=604800, path="/")

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


@router.post("/session")
async def create_session_from_google(request: Request, response: Response, db: AsyncSession = Depends(get_db)):
    session_id = request.headers.get("X-Session-ID")
    if not session_id:
        raise HTTPException(status_code=400, detail="Session ID required")

    async with httpx.AsyncClient() as client:
        try:
            res = await client.get(
                "https://demobackend.emergentagent.com/auth/v1/env/oauth/session-data",
                headers={"X-Session-ID": session_id},
                timeout=10.0
            )
            if res.status_code != 200:
                raise HTTPException(status_code=400, detail="Invalid session")
            google_user = res.json()
        except Exception as e:
            logger.error(f"Error fetching Google session: {e}")
            raise HTTPException(status_code=400, detail="Failed to verify session")

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
            created_at=datetime.now(timezone.utc)
        )
        db.add(user)
        db.add(SeriesConfig(
            tenant_id=tenant_id,
            invoice_prefix="INV", invoice_counter=0,
            estimate_prefix="EST", estimate_counter=0,
            credit_note_prefix="CN", credit_note_counter=0
        ))
        await db.commit()
        await db.refresh(user)

    if not user.is_active:
        raise HTTPException(status_code=403, detail="Account disabled by admin")

    user_dict = serialize_doc(user)
    has_access, access_status = check_user_access(user_dict)
    if not has_access:
        raise HTTPException(status_code=403, detail="Your 10-day demo has expired. Please contact admin for account approval.")

    session_token = google_user["session_token"]
    db.add(UserSession(
        session_token=session_token,
        user_id=user.user_id,
        expires_at=datetime.now(timezone.utc) + timedelta(days=7),
        created_at=datetime.now(timezone.utc)
    ))
    await db.commit()

    response.set_cookie(key="session_token", value=session_token, httponly=True,
                        secure=True, samesite="none", max_age=604800, path="/")

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
        "picture": user.picture,
        "session_token": session_token,
        "is_approved": user.is_approved,
        "access_status": access_status,
        "is_new_user": is_new_user,
        **demo_info
    }


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
