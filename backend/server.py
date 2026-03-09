"""
Taxly Invoice Generator CRM - Main Application (PostgreSQL version)
"""
from fastapi import FastAPI
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
import os
import logging
from pathlib import Path

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

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

app = FastAPI(
    title="Taxly Invoice Generator CRM",
    description="GST-compliant multi-tenant invoice management system (PostgreSQL)",
    version="3.0.0"
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

cors_origins = os.environ.get('CORS_ORIGINS', '*')
if cors_origins == '*':
    origins = ["*"]
    allow_credentials = False
else:
    origins = [o.strip() for o in cors_origins.split(',') if o.strip()]
    allow_credentials = True

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=allow_credentials,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS", "PATCH"],
    allow_headers=["*"],
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


@app.on_event("startup")
async def startup():
    await init_db()
    logger.info("PostgreSQL tables created/verified")
    logger.info("Taxly CRM (PostgreSQL) started successfully")


@app.on_event("shutdown")
async def shutdown():
    logger.info("Shutting down Taxly CRM")
