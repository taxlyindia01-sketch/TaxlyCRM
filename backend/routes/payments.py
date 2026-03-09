"""
payments.py  – PostgreSQL version
"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from typing import Optional
from datetime import datetime, timezone
import uuid

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Payment, Invoice, Client, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/payments", tags=["Payments"])


class PaymentCreate(BaseModel):
    invoice_id: str
    client_id: str
    amount: float
    payment_mode: str
    payment_date: str
    reference_number: Optional[str] = None
    notes: Optional[str] = None


@router.get("")
async def get_payments(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Payment).where(Payment.tenant_id == user["tenant_id"]).order_by(Payment.created_at.desc()))
    return [serialize_doc(p) for p in result.scalars().all()]


@router.post("")
async def create_payment(payment_data: PaymentCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == payment_data.invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = inv_res.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")

    payment = Payment(
        payment_id=f"payment_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        invoice_id=payment_data.invoice_id,
        client_id=payment_data.client_id,
        amount=payment_data.amount,
        payment_mode=payment_data.payment_mode,
        payment_date=payment_data.payment_date,
        reference_number=payment_data.reference_number,
        notes=payment_data.notes,
        created_at=datetime.now(timezone.utc)
    )
    db.add(payment)

    new_advance = (invoice.advance_paid or 0) + payment_data.amount
    new_outstanding = max(0, round((invoice.outstanding or 0) - payment_data.amount, 2))
    invoice.advance_paid = round(new_advance, 2)
    invoice.outstanding = new_outstanding
    invoice.status = "paid" if new_outstanding <= 0 else "partial"

    c_res = await db.execute(select(Client).where(Client.client_id == payment_data.client_id))
    client = c_res.scalar_one_or_none()
    if client:
        client.outstanding_balance = (client.outstanding_balance or 0) - payment_data.amount

    await db.commit()
    await db.refresh(payment)
    return serialize_doc(payment)
