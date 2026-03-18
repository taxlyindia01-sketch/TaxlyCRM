import uuid, smtplib, os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from typing import Optional
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

class StatusUpdate(BaseModel):
    status: str
    admin_note: Optional[str] = ""

def _send_resolve_email(to_email: str, name: str, category: str, message: str, admin_note: str = ""):
    """Send resolution email to ticket submitter. Uses SMTP env vars if set."""
    smtp_host = os.getenv("SMTP_HOST", "")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    from_email = os.getenv("SMTP_FROM", smtp_user)

    if not smtp_host or not smtp_user:
        # No SMTP configured — skip silently (frontend will handle mailto fallback)
        return False

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "Your Support Ticket Has Been Resolved — Taxly"
        msg["From"] = f"Taxly Support <{from_email}>"
        msg["To"] = to_email

        note_html = f"<p><strong>Admin Note:</strong> {admin_note}</p>" if admin_note else ""
        html = f"""
        <div style="font-family:sans-serif;max-width:520px;margin:0 auto">
          <div style="background:#0f172a;padding:1.2rem 1.5rem;border-radius:10px 10px 0 0">
            <h2 style="color:#fff;margin:0;font-size:1.1rem">Taxly Support</h2>
          </div>
          <div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:1.5rem;border-radius:0 0 10px 10px">
            <p>Dear <strong>{name}</strong>,</p>
            <p>Your support ticket has been <strong style="color:#16a34a">resolved</strong>. Here are the details:</p>
            <div style="background:#f8fafc;border-left:3px solid #b45309;padding:.75rem 1rem;margin:1rem 0;border-radius:0 6px 6px 0">
              <p style="margin:0 0 .3rem"><strong>Category:</strong> {category}</p>
              <p style="margin:0;color:#475569"><strong>Your message:</strong> {message}</p>
            </div>
            {note_html}
            <p>If you have further questions, feel free to raise a new ticket.</p>
            <p style="color:#64748b;font-size:.85rem;margin-top:1.5rem">Thank you for using Taxly!</p>
          </div>
        </div>
        """
        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(from_email, to_email, msg.as_string())
        return True
    except Exception as e:
        print(f"[Email error] {e}")
        return False


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


@router.get("/my-tickets")
async def list_my_tickets(db: AsyncSession = Depends(get_db),
                          user: dict = Depends(get_current_user)):
    """Return tickets submitted by the currently logged-in tenant."""
    result = await db.execute(
        select(SupportTicket)
        .where(SupportTicket.tenant_id == user.get("tenant_id", ""))
        .order_by(SupportTicket.created_at.desc())
    )
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
async def update_ticket_status(ticket_id: str, data: StatusUpdate,
                                db: AsyncSession = Depends(get_db),
                                _: bool = Depends(check_admin)):
    result = await db.execute(select(SupportTicket).where(SupportTicket.ticket_id == ticket_id))
    ticket = result.scalar_one_or_none()
    if not ticket:
        raise HTTPException(status_code=404, detail="Ticket not found")

    ticket.status = data.status
    await db.commit()

    email_sent = False
    if data.status == "resolved":
        email_sent = _send_resolve_email(
            to_email=ticket.email,
            name=ticket.name,
            category=ticket.category,
            message=ticket.message,
            admin_note=data.admin_note or ""
        )

    return {
        "message": "Updated",
        "email_sent": email_sent,
        "ticket_email": ticket.email,
        "ticket_name": ticket.name
    }
