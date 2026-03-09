from fastapi import APIRouter, Depends
from datetime import datetime, timezone

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Invoice, Client, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/dashboard", tags=["Dashboard"])


@router.get("/stats")
async def get_dashboard_stats(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_result = await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"]).limit(10000))
    invoices = inv_result.scalars().all()

    cli_result = await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]).limit(10000))
    clients = cli_result.scalars().all()
    client_map = {c.client_id: c for c in clients}

    total_collection = sum((i.advance_paid or 0) for i in invoices)
    total_revenue = sum((i.subtotal or 0) for i in invoices)
    total_outstanding = sum((i.outstanding or 0) for i in invoices)

    recent = sorted(invoices, key=lambda x: x.created_at or datetime.min.replace(tzinfo=timezone.utc), reverse=True)[:5]
    recent_list = []
    for inv in recent:
        c = client_map.get(inv.client_id)
        recent_list.append({
            "invoice_id": inv.invoice_id,
            "invoice_number": inv.invoice_number,
            "invoice_date": inv.invoice_date,
            "total": inv.total,
            "currency": inv.currency,
            "status": inv.status,
            "outstanding": inv.outstanding,
            "client_name": c.name if c else "Unknown"
        })

    return {
        "total_collection": round(total_collection, 2),
        "total_revenue": round(total_revenue, 2),
        "total_outstanding": round(total_outstanding, 2),
        "total_invoices": len(invoices),
        "total_clients": len(clients),
        "recent_invoices": recent_list
    }


@router.get("/stats-by-currency")
async def get_dashboard_stats_by_currency(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_result = await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"]).limit(10000))
    invoices = inv_result.scalars().all()

    currency_stats: dict = {}
    for inv in invoices:
        currency = inv.currency or 'INR'
        if currency not in currency_stats:
            currency_stats[currency] = {"total_collection": 0.0, "total_revenue": 0.0, "total_outstanding": 0.0, "invoice_count": 0}
        currency_stats[currency]["total_collection"] += inv.advance_paid or 0
        currency_stats[currency]["total_revenue"] += inv.subtotal or 0
        currency_stats[currency]["total_outstanding"] += inv.outstanding or 0
        currency_stats[currency]["invoice_count"] += 1

    for c in currency_stats:
        for k in ["total_collection", "total_revenue", "total_outstanding"]:
            currency_stats[c][k] = round(currency_stats[c][k], 2)

    return currency_stats
