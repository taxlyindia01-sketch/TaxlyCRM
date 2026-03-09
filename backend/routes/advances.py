from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from typing import Optional
from datetime import datetime, timezone
import uuid

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Advance, Invoice, Client, Payment, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/advances", tags=["Advances"])


class AdvancePaymentCreate(BaseModel):
    client_id: str
    amount: float
    payment_mode: str
    payment_date: str
    reference_number: Optional[str] = None


class AdvanceAdjustment(BaseModel):
    invoice_id: str
    amount: float


@router.get("")
async def get_advances(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Advance).where(Advance.tenant_id == user["tenant_id"]).order_by(Advance.created_at.desc()))
    return [serialize_doc(a) for a in result.scalars().all()]


@router.post("")
async def create_advance(advance_data: AdvancePaymentCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    advance = Advance(
        advance_id=f"advance_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        client_id=advance_data.client_id,
        amount=advance_data.amount,
        payment_mode=advance_data.payment_mode,
        payment_date=advance_data.payment_date,
        reference_number=advance_data.reference_number,
        status="available",
        created_at=datetime.now(timezone.utc)
    )
    db.add(advance)
    await db.commit()
    await db.refresh(advance)
    return serialize_doc(advance)


@router.post("/{advance_id}/adjust")
async def adjust_advance(advance_id: str, adjustment: AdvanceAdjustment, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    adv_res = await db.execute(select(Advance).where(Advance.advance_id == advance_id, Advance.tenant_id == user["tenant_id"]))
    advance = adv_res.scalar_one_or_none()
    if not advance:
        raise HTTPException(status_code=404, detail="Advance not found")
    if advance.status != "available":
        raise HTTPException(status_code=400, detail="Advance already used")
    if (advance.amount or 0) < adjustment.amount:
        raise HTTPException(status_code=400, detail="Adjustment amount exceeds advance amount")

    inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == adjustment.invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = inv_res.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")

    new_advance_paid = (invoice.advance_paid or 0) + adjustment.amount
    new_outstanding = (invoice.total or 0) - new_advance_paid
    invoice.advance_paid = round(new_advance_paid, 2)
    invoice.outstanding = round(new_outstanding, 2)
    invoice.status = "paid" if new_outstanding <= 0 else "partial"

    c_res = await db.execute(select(Client).where(Client.client_id == invoice.client_id))
    client = c_res.scalar_one_or_none()
    if client:
        client.outstanding_balance = (client.outstanding_balance or 0) - adjustment.amount

    remaining = (advance.amount or 0) - adjustment.amount
    if remaining <= 0:
        advance.status = "used"; advance.amount = 0
    else:
        advance.amount = round(remaining, 2)

    db.add(Payment(
        payment_id=f"payment_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        invoice_id=adjustment.invoice_id,
        client_id=invoice.client_id,
        amount=adjustment.amount,
        payment_mode="Advance Adjustment",
        payment_date=datetime.now(timezone.utc).strftime("%Y-%m-%d"),
        reference_number=f"ADV-{advance_id[:8]}",
        notes=f"Adjusted from advance {advance_id}",
        created_at=datetime.now(timezone.utc)
    ))

    await db.commit()
    return {"message": "Advance adjusted successfully", "remaining": round(remaining, 2)}
