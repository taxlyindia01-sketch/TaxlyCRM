"""
PostgreSQL Database Configuration for Taxly Invoice Generator CRM
Replaces MongoDB (motor) with SQLAlchemy async + asyncpg
"""
import os
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Any

from dotenv import load_dotenv
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase
from sqlalchemy import (
    Column, String, Boolean, Float, DateTime, Text, JSON,
    UniqueConstraint, Index, text
)

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------
DATABASE_URL = os.environ.get(
    'DATABASE_URL',
    'postgresql+asyncpg://postgres:password@localhost:5432/taxly_db'
)

engine = create_async_engine(DATABASE_URL, echo=False, pool_pre_ping=True)
AsyncSessionLocal = async_sessionmaker(engine, expire_on_commit=False)


# ---------------------------------------------------------------------------
# Base & ORM Models
# ---------------------------------------------------------------------------
class Base(DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"
    user_id        = Column(String, primary_key=True)
    email          = Column(String, nullable=False, unique=True)
    name           = Column(String, nullable=False)
    password_hash  = Column(String)
    tenant_id      = Column(String, nullable=False, index=True)
    picture        = Column(String)
    is_active      = Column(Boolean, default=True)
    is_approved    = Column(Boolean, default=False)
    approval_status= Column(String, default="demo")
    demo_expires_at= Column(DateTime(timezone=True))
    signup_method  = Column(String, default="email")
    rejection_reason = Column(String)
    approved_at    = Column(DateTime(timezone=True))
    approved_by    = Column(String)
    created_at     = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    __table_args__ = (
        UniqueConstraint('tenant_id', name='uq_users_tenant_id'),
    )


class UserSession(Base):
    __tablename__ = "user_sessions"
    session_token = Column(String, primary_key=True)
    user_id       = Column(String, nullable=False, index=True)
    expires_at    = Column(DateTime(timezone=True), nullable=False)
    created_at    = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


class AdminSession(Base):
    __tablename__ = "admin_sessions"
    session_token = Column(String, primary_key=True)
    username      = Column(String, nullable=False)
    expires_at    = Column(DateTime(timezone=True), nullable=False)
    created_at    = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


class Business(Base):
    __tablename__ = "businesses"
    tenant_id             = Column(String, primary_key=True)
    business_name         = Column(String)
    gstin                 = Column(String)
    email                 = Column(String)
    phone                 = Column(String)
    address               = Column(String)
    city                  = Column(String)
    state                 = Column(String)
    pincode               = Column(String)
    bank_name             = Column(String)
    account_number        = Column(String)
    ifsc_code             = Column(String)
    swift_code            = Column(String)
    authorised_signatory  = Column(String)
    terms_and_conditions  = Column(Text)
    upi_id                = Column(String)
    logo_url              = Column(String)
    qr_code_url           = Column(String)


class Client(Base):
    __tablename__ = "clients"
    client_id         = Column(String, primary_key=True)
    tenant_id         = Column(String, nullable=False, index=True)
    name              = Column(String, nullable=False)
    type              = Column(String)
    gstin             = Column(String)
    email             = Column(String)
    phone             = Column(String)
    address           = Column(String)
    city              = Column(String)
    state             = Column(String)
    pincode           = Column(String)
    outstanding_balance = Column(Float, default=0.0)
    created_at        = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


class Invoice(Base):
    __tablename__ = "invoices"
    invoice_id        = Column(String, primary_key=True)
    tenant_id         = Column(String, nullable=False, index=True)
    client_id         = Column(String, nullable=False, index=True)
    client_name       = Column(String)
    client_gstin      = Column(String)
    client_state      = Column(String)
    invoice_number    = Column(String, unique=True)
    invoice_date      = Column(String)
    due_date          = Column(String)
    diary_no          = Column(String)
    ref_no            = Column(String)
    transporter_name  = Column(String)
    transporter_gstin = Column(String)
    vehicle_no        = Column(String)
    items             = Column(JSON, default=list)
    subtotal          = Column(Float, default=0.0)
    cgst              = Column(Float, default=0.0)
    sgst              = Column(Float, default=0.0)
    igst              = Column(Float, default=0.0)
    total             = Column(Float, default=0.0)
    currency          = Column(String, default="INR")
    is_export         = Column(Boolean, default=False)
    is_pure_agent     = Column(Boolean, default=False)
    status            = Column(String, default="unpaid")
    advance_paid      = Column(Float, default=0.0)
    outstanding       = Column(Float, default=0.0)
    notes             = Column(Text)
    bill_to_name      = Column(String)
    bill_to_address   = Column(String)
    bill_to_city      = Column(String)
    bill_to_state     = Column(String)
    bill_to_pincode   = Column(String)
    bill_to_gstin     = Column(String)
    ship_to_same      = Column(Boolean, default=True)
    ship_to_name      = Column(String)
    ship_to_address   = Column(String)
    ship_to_city      = Column(String)
    ship_to_state     = Column(String)
    ship_to_pincode   = Column(String)
    ship_to_gstin     = Column(String)
    pdf_template      = Column(String, default="default")
    created_at        = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    updated_at        = Column(DateTime(timezone=True))


class Estimate(Base):
    __tablename__ = "estimates"
    estimate_id       = Column(String, primary_key=True)
    tenant_id         = Column(String, nullable=False, index=True)
    client_id         = Column(String, nullable=False)
    client_state      = Column(String)
    estimate_number   = Column(String, unique=True)
    estimate_date     = Column(String)
    valid_until       = Column(String)
    diary_no          = Column(String)
    ref_no            = Column(String)
    transporter_name  = Column(String)
    transporter_gstin = Column(String)
    vehicle_no        = Column(String)
    items             = Column(JSON, default=list)
    subtotal          = Column(Float, default=0.0)
    cgst              = Column(Float, default=0.0)
    sgst              = Column(Float, default=0.0)
    igst              = Column(Float, default=0.0)
    total             = Column(Float, default=0.0)
    currency          = Column(String, default="INR")
    is_export         = Column(Boolean, default=False)
    status            = Column(String, default="pending")
    notes             = Column(Text)
    bill_to_name      = Column(String)
    bill_to_address   = Column(String)
    bill_to_city      = Column(String)
    bill_to_state     = Column(String)
    bill_to_pincode   = Column(String)
    bill_to_gstin     = Column(String)
    ship_to_same      = Column(Boolean, default=True)
    ship_to_name      = Column(String)
    ship_to_address   = Column(String)
    ship_to_city      = Column(String)
    ship_to_state     = Column(String)
    ship_to_pincode   = Column(String)
    ship_to_gstin     = Column(String)
    pdf_template      = Column(String, default="default")
    created_at        = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    updated_at        = Column(DateTime(timezone=True))


class CreditNote(Base):
    __tablename__ = "credit_notes"
    credit_note_id     = Column(String, primary_key=True)
    tenant_id          = Column(String, nullable=False, index=True)
    invoice_id         = Column(String, nullable=False)
    client_id          = Column(String, nullable=False)
    client_state       = Column(String)
    credit_note_number = Column(String, unique=True)
    credit_date        = Column(String)
    diary_no           = Column(String)
    ref_no             = Column(String)
    items              = Column(JSON, default=list)
    subtotal           = Column(Float, default=0.0)
    cgst               = Column(Float, default=0.0)
    sgst               = Column(Float, default=0.0)
    igst               = Column(Float, default=0.0)
    total              = Column(Float, default=0.0)
    currency           = Column(String, default="INR")
    reason             = Column(Text)
    bill_to_name       = Column(String)
    bill_to_address    = Column(String)
    bill_to_city       = Column(String)
    bill_to_state      = Column(String)
    bill_to_pincode    = Column(String)
    bill_to_gstin      = Column(String)
    ship_to_same       = Column(Boolean, default=True)
    ship_to_name       = Column(String)
    ship_to_address    = Column(String)
    ship_to_city       = Column(String)
    ship_to_state      = Column(String)
    ship_to_pincode    = Column(String)
    ship_to_gstin      = Column(String)
    created_at         = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    updated_at         = Column(DateTime(timezone=True))


class SeriesConfig(Base):
    __tablename__ = "series_config"
    tenant_id            = Column(String, primary_key=True)
    invoice_prefix       = Column(String, default="INV")
    invoice_counter      = Column(Float, default=0)
    estimate_prefix      = Column(String, default="EST")
    estimate_counter     = Column(Float, default=0)
    credit_note_prefix   = Column(String, default="CN")
    credit_note_counter  = Column(Float, default=0)


class MasterData(Base):
    __tablename__ = "master_data"
    master_id   = Column(String, primary_key=True)
    tenant_id   = Column(String, nullable=False, index=True)
    description = Column(String)
    hsn_sac_code= Column(String)
    gst_rate    = Column(Float)
    type        = Column(String, default="unified")
    value       = Column(String)
    created_at  = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


class Payment(Base):
    __tablename__ = "payments"
    payment_id       = Column(String, primary_key=True)
    tenant_id        = Column(String, nullable=False, index=True)
    invoice_id       = Column(String, nullable=False)
    client_id        = Column(String, nullable=False)
    amount           = Column(Float, default=0.0)
    payment_mode     = Column(String)
    payment_date     = Column(String)
    reference_number = Column(String)
    notes            = Column(Text)
    created_at       = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


class Advance(Base):
    __tablename__ = "advances"
    advance_id       = Column(String, primary_key=True)
    tenant_id        = Column(String, nullable=False, index=True)
    client_id        = Column(String, nullable=False)
    amount           = Column(Float, default=0.0)
    payment_mode     = Column(String)
    payment_date     = Column(String)
    reference_number = Column(String)
    status           = Column(String, default="available")
    created_at       = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------
def serialize_doc(obj) -> Optional[dict]:
    """Convert an ORM row (or plain dict) to a JSON-serialisable dict."""
    if obj is None:
        return None
    if isinstance(obj, dict):
        data = obj
    else:
        data = {c.name: getattr(obj, c.name) for c in obj.__table__.columns}

    result = {}
    for key, value in data.items():
        if isinstance(value, datetime):
            result[key] = value.isoformat()
        elif isinstance(value, (list, dict)):
            result[key] = value   # JSON columns already deserialized by SA
        else:
            result[key] = value
    return result


async def get_db() -> AsyncSession:
    """FastAPI dependency – yields an AsyncSession."""
    async with AsyncSessionLocal() as session:
        yield session


async def init_db():
    """Create all tables (safe to call on every startup)."""
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    print("PostgreSQL tables created / verified successfully")
