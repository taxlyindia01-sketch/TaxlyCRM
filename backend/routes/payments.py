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
    payments = [serialize_doc(p) for p in result.scalars().all()]
    # Enrich with invoice_number for better reconciliation
    inv_ids = list({p["invoice_id"] for p in payments if p.get("invoice_id")})
    if inv_ids:
        inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id.in_(inv_ids)))
        inv_map = {inv.invoice_id: inv.invoice_number for inv in inv_res.scalars().all()}
        for p in payments:
            p["invoice_number"] = inv_map.get(p.get("invoice_id"), "")
    return payments


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

@router.get("/{payment_id}")
async def get_payment(payment_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    res = await db.execute(select(Payment).where(Payment.payment_id == payment_id, Payment.tenant_id == user["tenant_id"]))
    p = res.scalar_one_or_none()
    if not p:
        raise HTTPException(status_code=404, detail="Payment not found")
    data = serialize_doc(p)
    # Enrich with invoice_number
    if p.invoice_id:
        inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == p.invoice_id))
        inv = inv_res.scalar_one_or_none()
        if inv:
            data["invoice_number"] = inv.invoice_number
    return data


@router.put("/{payment_id}")
async def update_payment(payment_id: str, payment_data: PaymentCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    res = await db.execute(select(Payment).where(Payment.payment_id == payment_id, Payment.tenant_id == user["tenant_id"]))
    p = res.scalar_one_or_none()
    if not p:
        raise HTTPException(status_code=404, detail="Payment not found")
    old_amount = p.amount or 0
    # Restore old amounts on invoice
    if p.invoice_id:
        inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == p.invoice_id))
        inv = inv_res.scalar_one_or_none()
        if inv:
            inv.advance_paid = round((inv.advance_paid or 0) - old_amount, 2)
            inv.outstanding = round((inv.outstanding or 0) + old_amount, 2)
    # Update payment
    p.invoice_id = payment_data.invoice_id
    p.client_id = payment_data.client_id
    p.amount = payment_data.amount
    p.payment_mode = payment_data.payment_mode
    p.payment_date = payment_data.payment_date
    p.reference_number = payment_data.reference_number
    p.notes = payment_data.notes
    # Apply new amounts
    if payment_data.invoice_id:
        inv_res2 = await db.execute(select(Invoice).where(Invoice.invoice_id == payment_data.invoice_id))
        inv2 = inv_res2.scalar_one_or_none()
        if inv2:
            inv2.advance_paid = round((inv2.advance_paid or 0) + payment_data.amount, 2)
            inv2.outstanding = max(0, round((inv2.outstanding or 0) - payment_data.amount, 2))
            inv2.status = "paid" if inv2.outstanding <= 0 else "partial"
    await db.commit()
    return serialize_doc(p)
