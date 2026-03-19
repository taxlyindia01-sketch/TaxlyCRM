"""
Taxly Invoice Generator CRM - Main Application (PostgreSQL version)
"""
# ── load_dotenv MUST run before any route modules are imported ─────────────
from pathlib import Path
from dotenv import load_dotenv
import os as _os

_ROOT_DIR = Path(__file__).parent
_env_file = _ROOT_DIR / ".env"
_env_example = _ROOT_DIR / ".env.example"

if _env_file.exists():
    load_dotenv(_env_file, override=True)
    print(f"[ENV] Loaded: {_env_file}")
elif _env_example.exists():
    load_dotenv(_env_example, override=True)
    print(f"[ENV] WARNING: .env not found — loaded .env.example as fallback.")
    print(f"[ENV] Copy .env.example to .env and fill in production values!")
else:
    print("[ENV] WARNING: No .env or .env.example found. Using system environment only.")

# Confirm Google OAuth loaded
_gid = _os.environ.get("GOOGLE_CLIENT_ID", "")
print(f"[ENV] GOOGLE_CLIENT_ID = {'SET (' + _gid[:12] + '...)' if _gid else 'NOT SET — Google login will be disabled'}")
# ──────────────────────────────────────────────────────────────────────────

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from starlette.middleware.cors import CORSMiddleware
import os
import logging

from database import init_db

from routes.auth import router as auth_router
from routes.admin import router as admin_router
from routes.business import router as business_router
from routes.clients import router as clients_router
from routes.invoices import router as invoices_router
from routes.estimates import router as estimates_router
from routes.credit_notes import router as credit_notes_router
from routes.payments import router as payments_router
from routes.advances import router as advances_router
from routes.dashboard import router as dashboard_router
from routes.master_data import router as master_data_router
from routes.series_config import router as series_config_router
from routes.exports import router as exports_router
from routes.reports import router as reports_router
from routes.support import router as support_router

ROOT_DIR = _ROOT_DIR

app = FastAPI(
    title="Taxly Invoice Generator CRM",
    description="GST-compliant multi-tenant invoice management system (PostgreSQL)",
    version="3.0.0"
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

cors_origins = os.environ.get('CORS_ORIGINS', '*')

# Always include ALL local dev origins (covers Live Server on any port 5500-5510)
_local_dev_origins = [
    "http://localhost:5500", "http://127.0.0.1:5500",
    "http://localhost:5501", "http://127.0.0.1:5501",
    "http://localhost:5502", "http://127.0.0.1:5502",
    "http://localhost:5503", "http://127.0.0.1:5503",
    "http://localhost:8000", "http://127.0.0.1:8000",
    "http://localhost:3000", "http://127.0.0.1:3000",
    "http://localhost:8080", "http://127.0.0.1:8080",
    "http://localhost",      "http://127.0.0.1",
    "null",  # file:// protocol
]

if cors_origins.strip() == '*':
    # Wildcard + credentials is rejected by browsers; use explicit local list
    origins = _local_dev_origins
    allow_credentials = True
else:
    origins = list({o.strip() for o in cors_origins.split(',') if o.strip()} | set(_local_dev_origins))
    allow_credentials = True

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=allow_credentials,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS", "PATCH"],
    allow_headers=["*", "Authorization", "Content-Type", "X-Session-ID"],
    expose_headers=["*"],
)


@app.get("/health")
async def health_check_root():
    return {"status": "healthy", "database": "postgresql", "version": "3.0.0"}


@app.get("/api/health")
async def health_check_api():
    return {"status": "healthy", "database": "postgresql", "version": "3.0.0"}


app.include_router(auth_router, prefix="/api")
app.include_router(admin_router, prefix="/api")
app.include_router(business_router, prefix="/api")
app.include_router(clients_router, prefix="/api")
app.include_router(invoices_router, prefix="/api")
app.include_router(estimates_router, prefix="/api")
app.include_router(credit_notes_router, prefix="/api")
app.include_router(payments_router, prefix="/api")
app.include_router(advances_router, prefix="/api")
app.include_router(dashboard_router, prefix="/api")
app.include_router(master_data_router, prefix="/api")
app.include_router(series_config_router, prefix="/api")
app.include_router(exports_router, prefix="/api")
app.include_router(reports_router, prefix="/api")
app.include_router(support_router, prefix="/api")


# ── Serve frontend static files ───────────────────────────────────────────
FRONTEND_DIR = ROOT_DIR.parent / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

    @app.get("/", include_in_schema=False)
    @app.get("/{path:path}", include_in_schema=False)
    async def serve_frontend(path: str = ""):
        index = FRONTEND_DIR / "index.html"
        if index.exists():
            return FileResponse(str(index))
        return {"error": "Frontend not found"}
else:
    logger.warning(f"Frontend directory not found at {FRONTEND_DIR}")


@app.on_event("startup")
async def startup():
    await init_db()
    logger.info("PostgreSQL tables created/verified")
    logger.info("Taxly CRM (PostgreSQL) started successfully")


@app.on_event("shutdown")
async def shutdown():
    logger.info("Shutting down Taxly CRM")
