from fastapi import APIRouter, Depends
from datetime import datetime, timezone

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Invoice, Client, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/dashboard", tags=["Dashboard"])


def _to_inr(amount: float, currency: str, exchange_rate: float | None) -> float:
    """Convert an amount to INR using the stored exchange rate.
    If currency is INR or exchange_rate is missing, return as-is."""
    if not currency or currency == "INR":
        return amount
    if exchange_rate and exchange_rate > 0:
        return round(amount * exchange_rate, 2)
    return amount  # fallback: treat as INR if no rate stored


@router.get("/stats")
async def get_dashboard_stats(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_result = await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"]).limit(10000))
    invoices = inv_result.scalars().all()

    cli_result = await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]).limit(10000))
    clients = cli_result.scalars().all()
    client_map = {c.client_id: c for c in clients}

    # Convert all amounts to INR using stored exchange rates (exclude cancelled)
    active_invoices = [i for i in invoices if i.status != "cancelled"]
    total_collection = sum(_to_inr(i.advance_paid or 0, i.currency or "INR", i.exchange_rate) for i in active_invoices)
    total_revenue    = sum(_to_inr(i.total or 0,        i.currency or "INR", i.exchange_rate) for i in active_invoices)
    total_outstanding= sum(_to_inr(i.outstanding or 0,  i.currency or "INR", i.exchange_rate) for i in active_invoices)

    recent = sorted(invoices, key=lambda x: x.created_at or datetime.min.replace(tzinfo=timezone.utc), reverse=True)[:5]
    recent_list = []
    for inv in recent:
        c = client_map.get(inv.client_id)
        recent_list.append({
            "invoice_id":     inv.invoice_id,
            "invoice_number": inv.invoice_number,
            "invoice_date":   inv.invoice_date,
            "total":          inv.total,
            "total_inr":      _to_inr(inv.total or 0, inv.currency or "INR", inv.exchange_rate),
            "currency":       inv.currency,
            "exchange_rate":  inv.exchange_rate,
            "status":         inv.status,
            "outstanding":    inv.outstanding,
            "client_name":    c.name if c else "Unknown"
        })

    return {
        "total_collection":  round(total_collection, 2),
        "total_revenue":     round(total_revenue, 2),
        "total_outstanding": round(total_outstanding, 2),
        "total_invoices":    len(invoices),
        "total_clients":     len(clients),
        "recent_invoices":   recent_list
    }


@router.get("/stats-by-currency")
async def get_dashboard_stats_by_currency(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_result = await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"]).limit(10000))
    invoices = inv_result.scalars().all()

    currency_stats: dict = {}
    for inv in invoices:
        if inv.status == "cancelled":
            continue  # Exclude cancelled invoices from all stats
        currency = inv.currency or "INR"
        if currency not in currency_stats:
            currency_stats[currency] = {
                "total_collection": 0.0,
                "total_revenue":    0.0,
                "total_outstanding":0.0,
                "invoice_count":    0,
                "exchange_rate":    inv.exchange_rate  # store last-seen rate
            }
        currency_stats[currency]["total_collection"]  += inv.advance_paid or 0
        currency_stats[currency]["total_revenue"]     += inv.total or 0  # total = subtotal + tax
        currency_stats[currency]["total_outstanding"] += inv.outstanding or 0
        currency_stats[currency]["invoice_count"]     += 1
        if inv.exchange_rate:
            currency_stats[currency]["exchange_rate"] = inv.exchange_rate

    for c in currency_stats:
        for k in ["total_collection", "total_revenue", "total_outstanding"]:
            currency_stats[c][k] = round(currency_stats[c][k], 2)

    return currency_stats
