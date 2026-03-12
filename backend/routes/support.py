import uuid
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from database import SupportTicket, get_db
from routes.auth import get_current_user
from routes.admin import check_admin

router = APIRouter(prefix="/support", tags=["Support"])

class TicketIn(BaseModel):
    name: str
    email: str
    category: str
    message: str

@router.post("/tickets")
async def create_ticket(data: TicketIn, db: AsyncSession = Depends(get_db),
                        user: dict = Depends(get_current_user)):
    ticket = SupportTicket(
        ticket_id=uuid.uuid4().hex,
        name=data.name,
        email=data.email,
        category=data.category,
        message=data.message,
        tenant_id=user.get("tenant_id", ""),
        status="open"
    )
    db.add(ticket)
    await db.commit()
    return {"message": "Ticket submitted successfully", "ticket_id": ticket.ticket_id}

@router.get("/tickets")
async def list_tickets(db: AsyncSession = Depends(get_db),
                       _: bool = Depends(check_admin)):
    result = await db.execute(select(SupportTicket).order_by(SupportTicket.created_at.desc()))
    tickets = result.scalars().all()
    return [
        {
            "ticket_id": t.ticket_id,
            "name": t.name,
            "email": t.email,
            "category": t.category,
            "message": t.message,
            "status": t.status,
            "created_at": t.created_at.isoformat() if t.created_at else None
        }
        for t in tickets
    ]

@router.put("/tickets/{ticket_id}/status")
async def update_ticket_status(ticket_id: str, data: dict,
                                db: AsyncSession = Depends(get_db),
                                _: bool = Depends(check_admin)):
    result = await db.execute(select(SupportTicket).where(SupportTicket.ticket_id == ticket_id))
    ticket = result.scalar_one_or_none()
    if not ticket:
        raise HTTPException(status_code=404, detail="Ticket not found")
    ticket.status = data.get("status", ticket.status)
    await db.commit()
    return {"message": "Updated"}
