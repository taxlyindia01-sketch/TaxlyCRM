from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime, timezone
import uuid
from io import BytesIO
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Invoice, Business, Client, SeriesConfig, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/invoices", tags=["Invoices"])

ROOT_DIR = Path(__file__).parent.parent
LOGO_UPLOAD_DIR = ROOT_DIR / "uploads" / "logos"

CURRENCY_SYMBOLS = {"INR": "Rs.", "USD": "$", "EUR": "EUR ", "GBP": "GBP "}


class InvoiceItem(BaseModel):
    description: str
    hsn_sac_code: Optional[str] = None
    quantity: float
    rate: float
    gst_rate: float
    amount: float
    item_notes: Optional[str] = None
    is_pure_agent: Optional[bool] = False


class InvoiceCreate(BaseModel):
    client_id: str
    invoice_date: str
    due_date: str
    items: List[InvoiceItem]
    currency: str = "INR"
    is_export: bool = False
    diary_no: Optional[str] = None
    ref_no: Optional[str] = None
    transporter_name: Optional[str] = None
    transporter_gstin: Optional[str] = None
    vehicle_no: Optional[str] = None
    notes: Optional[str] = None
    bill_to_name: Optional[str] = None
    bill_to_address: Optional[str] = None
    bill_to_city: Optional[str] = None
    bill_to_state: Optional[str] = None
    bill_to_pincode: Optional[str] = None
    bill_to_gstin: Optional[str] = None
    ship_to_same: Optional[bool] = True
    ship_to_name: Optional[str] = None
    ship_to_address: Optional[str] = None
    ship_to_city: Optional[str] = None
    ship_to_state: Optional[str] = None
    ship_to_pincode: Optional[str] = None
    ship_to_gstin: Optional[str] = None
    pdf_template: Optional[str] = "default"
    exchange_rate: Optional[float] = None


async def get_next_invoice_number(tenant_id: str, db: AsyncSession) -> str:
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == tenant_id))
    config = result.scalar_one_or_none()
    if not config:
        config = SeriesConfig(tenant_id=tenant_id, invoice_prefix="INV", invoice_counter=0,
                              estimate_prefix="EST", estimate_counter=0,
                              credit_note_prefix="CN", credit_note_counter=0)
        db.add(config)
        await db.flush()
    config.invoice_counter = (config.invoice_counter or 0) + 1
    await db.flush()
    return f"{config.invoice_prefix}-{int(config.invoice_counter):05d}"


@router.get("")
async def get_invoices(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(
        select(Invoice).where(Invoice.tenant_id == user["tenant_id"]).order_by(Invoice.created_at.desc())
    )
    invoices = result.scalars().all()
    out = []
    for inv in invoices:
        d = serialize_doc(inv)
        if not d.get('client_name') or not d.get('client_gstin'):
            c_res = await db.execute(select(Client).where(Client.client_id == inv.client_id))
            c = c_res.scalar_one_or_none()
            if c:
                d['client_name'] = c.name
                d['client_gstin'] = c.gstin
        out.append(d)
    return out


@router.post("")
async def create_invoice(invoice_data: InvoiceCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == invoice_data.client_id))
    client = c_res.scalar_one_or_none()
    if not business or not client:
        raise HTTPException(status_code=400, detail="Business or Client not found")

    invoice_number = await get_next_invoice_number(user["tenant_id"], db)
    subtotal = sum(item.amount for item in invoice_data.items)
    cgst = sgst = igst = 0.0

    if not invoice_data.is_export:
        if business.state == client.state:
            for item in invoice_data.items:
                if not item.is_pure_agent:
                    cgst += item.amount * (item.gst_rate / 2) / 100
                    sgst += item.amount * (item.gst_rate / 2) / 100
        else:
            for item in invoice_data.items:
                if not item.is_pure_agent:
                    igst += item.amount * item.gst_rate / 100

    total = subtotal + cgst + sgst + igst
    invoice_id = f"invoice_{uuid.uuid4().hex[:12]}"

    inv = Invoice(
        invoice_id=invoice_id,
        tenant_id=user["tenant_id"],
        client_id=invoice_data.client_id,
        client_name=client.name,
        client_gstin=client.gstin,
        client_state=client.state,
        invoice_number=invoice_number,
        invoice_date=invoice_data.invoice_date,
        due_date=invoice_data.due_date,
        diary_no=invoice_data.diary_no,
        ref_no=invoice_data.ref_no,
        transporter_name=invoice_data.transporter_name,
        transporter_gstin=invoice_data.transporter_gstin,
        vehicle_no=invoice_data.vehicle_no,
        items=[item.model_dump() for item in invoice_data.items],
        subtotal=round(subtotal, 2),
        cgst=round(cgst, 2), sgst=round(sgst, 2), igst=round(igst, 2),
        total=round(total, 2),
        currency=invoice_data.currency,
        is_export=invoice_data.is_export,
        status="unpaid",
        advance_paid=0.0,
        outstanding=round(total, 2),
        notes=invoice_data.notes,
        bill_to_name=invoice_data.bill_to_name or client.name,
        bill_to_address=invoice_data.bill_to_address or client.address,
        bill_to_city=invoice_data.bill_to_city or client.city,
        bill_to_state=invoice_data.bill_to_state or client.state,
        bill_to_pincode=invoice_data.bill_to_pincode or client.pincode,
        bill_to_gstin=invoice_data.bill_to_gstin or client.gstin,
        ship_to_same=invoice_data.ship_to_same,
        ship_to_name=invoice_data.ship_to_name if not invoice_data.ship_to_same else None,
        ship_to_address=invoice_data.ship_to_address if not invoice_data.ship_to_same else None,
        ship_to_city=invoice_data.ship_to_city if not invoice_data.ship_to_same else None,
        ship_to_state=invoice_data.ship_to_state if not invoice_data.ship_to_same else None,
        ship_to_pincode=invoice_data.ship_to_pincode if not invoice_data.ship_to_same else None,
        ship_to_gstin=invoice_data.ship_to_gstin if not invoice_data.ship_to_same else None,
        pdf_template=invoice_data.pdf_template,
        exchange_rate=invoice_data.exchange_rate,
        created_at=datetime.now(timezone.utc)
    )
    db.add(inv)
    client.outstanding_balance = (client.outstanding_balance or 0) + round(total, 2)
    await db.commit()
    await db.refresh(inv)
    return serialize_doc(inv)


@router.post("/{invoice_id}/cancel")
async def cancel_invoice(invoice_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Invoice).where(Invoice.invoice_id == invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if invoice.status == 'cancelled':
        raise HTTPException(status_code=400, detail="Invoice is already cancelled")
    if invoice.status == 'paid':
        raise HTTPException(status_code=400, detail="Cannot cancel a paid invoice")

    old_outstanding = invoice.outstanding or 0
    invoice.status = "cancelled"
    invoice.outstanding = 0

    c_res = await db.execute(select(Client).where(Client.client_id == invoice.client_id))
    client = c_res.scalar_one_or_none()
    if client:
        client.outstanding_balance = (client.outstanding_balance or 0) - old_outstanding

    await db.commit()
    return {"message": "Invoice cancelled successfully"}


@router.get("/{invoice_id}")
async def get_invoice(invoice_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Invoice).where(Invoice.invoice_id == invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    return serialize_doc(invoice)


@router.put("/{invoice_id}")
async def update_invoice(invoice_id: str, invoice_data: InvoiceCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Invoice).where(Invoice.invoice_id == invoice_id, Invoice.tenant_id == user["tenant_id"]))
    existing = result.scalar_one_or_none()
    if not existing:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if existing.status == 'cancelled':
        raise HTTPException(status_code=400, detail="Cannot edit a cancelled invoice")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == invoice_data.client_id))
    client = c_res.scalar_one_or_none()
    if not business or not client:
        raise HTTPException(status_code=400, detail="Business or Client not found")

    subtotal = sum(item.amount for item in invoice_data.items)
    cgst = sgst = igst = 0.0
    if not invoice_data.is_export:
        if business.state == client.state:
            for item in invoice_data.items:
                if not item.is_pure_agent:
                    cgst += item.amount * (item.gst_rate / 2) / 100
                    sgst += item.amount * (item.gst_rate / 2) / 100
        else:
            for item in invoice_data.items:
                if not item.is_pure_agent:
                    igst += item.amount * item.gst_rate / 100

    total = subtotal + cgst + sgst + igst
    old_outstanding = existing.outstanding or 0
    advance_paid = existing.advance_paid or 0
    new_outstanding = round(total - advance_paid, 2)
    new_status = "paid" if new_outstanding <= 0 else ("partial" if advance_paid > 0 else "unpaid")

    existing.client_id = invoice_data.client_id
    existing.client_name = client.name; existing.client_gstin = client.gstin; existing.client_state = client.state
    existing.invoice_date = invoice_data.invoice_date; existing.due_date = invoice_data.due_date
    existing.diary_no = invoice_data.diary_no; existing.ref_no = invoice_data.ref_no
    existing.transporter_name = invoice_data.transporter_name; existing.transporter_gstin = invoice_data.transporter_gstin
    existing.vehicle_no = invoice_data.vehicle_no
    existing.items = [item.model_dump() for item in invoice_data.items]
    existing.subtotal = round(subtotal, 2); existing.cgst = round(cgst, 2); existing.sgst = round(sgst, 2); existing.igst = round(igst, 2)
    existing.total = round(total, 2); existing.currency = invoice_data.currency; existing.is_export = invoice_data.is_export
    existing.outstanding = new_outstanding; existing.status = new_status
    existing.notes = invoice_data.notes
    existing.bill_to_name = invoice_data.bill_to_name or client.name
    existing.bill_to_address = invoice_data.bill_to_address or client.address
    existing.bill_to_city = invoice_data.bill_to_city or client.city
    existing.bill_to_state = invoice_data.bill_to_state or client.state
    existing.bill_to_pincode = invoice_data.bill_to_pincode or client.pincode
    existing.bill_to_gstin = invoice_data.bill_to_gstin or client.gstin
    existing.ship_to_same = invoice_data.ship_to_same
    existing.ship_to_name = invoice_data.ship_to_name if not invoice_data.ship_to_same else None
    existing.ship_to_address = invoice_data.ship_to_address if not invoice_data.ship_to_same else None
    existing.ship_to_city = invoice_data.ship_to_city if not invoice_data.ship_to_same else None
    existing.ship_to_state = invoice_data.ship_to_state if not invoice_data.ship_to_same else None
    existing.ship_to_pincode = invoice_data.ship_to_pincode if not invoice_data.ship_to_same else None
    existing.ship_to_gstin = invoice_data.ship_to_gstin if not invoice_data.ship_to_same else None
    existing.pdf_template = invoice_data.pdf_template
    existing.updated_at = datetime.now(timezone.utc)

    outstanding_diff = new_outstanding - old_outstanding
    if outstanding_diff != 0:
        client.outstanding_balance = (client.outstanding_balance or 0) + outstanding_diff

    await db.commit()
    await db.refresh(existing)
    return serialize_doc(existing)


@router.get("/{invoice_id}/pdf")
async def download_invoice_pdf(invoice_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Invoice).where(Invoice.invoice_id == invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == invoice.client_id))
    client = c_res.scalar_one_or_none()
    if not business or not client:
        raise HTTPException(status_code=400, detail="Business or client not found")

    inv = serialize_doc(invoice)
    biz = serialize_doc(business)
    cli = serialize_doc(client)

    use_compact = inv.get('pdf_template') == 'compact'
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []; styles = getSampleStyleSheet()

    if inv.get('is_export'):
        invoice_title = "EXPORT INVOICE"; theme_color = colors.HexColor('#7C3AED'); light_bg = colors.HexColor('#F5F3FF')
    elif inv.get('is_pure_agent'):
        invoice_title = "INVOICE (PURE AGENT)"; theme_color = colors.HexColor('#EA580C'); light_bg = colors.HexColor('#FFF7ED')
    else:
        invoice_title = "TAX INVOICE"; theme_color = colors.HexColor('#2563EB'); light_bg = colors.HexColor('#EFF6FF')

    currency_symbol = CURRENCY_SYMBOLS.get(inv.get('currency', 'INR'), "Rs.")
    logo_element = None
    if biz.get('logo_url'):
        try:
            logo_path = LOGO_UPLOAD_DIR / biz['logo_url'].split("/")[-1]
            if logo_path.exists():
                logo_element = RLImage(str(logo_path), width=2*inch, height=1*inch)
        except:
            pass

    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24 if not use_compact else 20,
                                  textColor=theme_color, alignment=TA_CENTER, fontName='Helvetica-Bold')
    if logo_element:
        logo_table = Table([[logo_element, Paragraph(invoice_title, title_style)]], colWidths=[2.5*inch, 4*inch])
        logo_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0),(0,0), 'LEFT'), ('ALIGN', (1,0),(1,0), 'RIGHT')]))
        elements.append(logo_table)
    else:
        elements.append(Paragraph(invoice_title, title_style))
    elements.append(Spacer(1, 0.2*inch))

    from_style = ParagraphStyle('FromStyle', parent=styles['Normal'], fontSize=8 if use_compact else 9, leading=11 if use_compact else 12)
    invoice_details = f"<b>Invoice No:</b> {inv.get('invoice_number','')}<br/><b>Date:</b> {inv.get('invoice_date','')}<br/><b>Due Date:</b> {inv.get('due_date','')}<br/><b>Currency:</b> {inv.get('currency','INR')}"
    if inv.get('diary_no'): invoice_details += f"<br/><b>Diary No:</b> {inv.get('diary_no')}"
    if inv.get('ref_no'):   invoice_details += f"<br/><b>Ref No:</b> {inv.get('ref_no')}"

    header_data = [[
        Paragraph(f"<b>From:</b><br/><b>{biz.get('business_name','')}</b><br/>{biz.get('address','')}<br/>{biz.get('city','')}, {biz.get('state','')} - {biz.get('pincode','')}<br/>GSTIN: {biz.get('gstin','')}<br/>Email: {biz.get('email','')}<br/>Phone: {biz.get('phone','')}", from_style),
        Paragraph(invoice_details, from_style)
    ]]
    header_table = Table(header_data, colWidths=[3.5*inch, 3*inch])
    header_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('ALIGN',(1,0),(1,0),'RIGHT'),
                                       ('BOX',(0,0),(-1,-1),1,theme_color), ('BACKGROUND',(0,0),(-1,-1),light_bg),
                                       ('LEFTPADDING',(0,0),(-1,-1),10), ('RIGHTPADDING',(0,0),(-1,-1),10),
                                       ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(header_table); elements.append(Spacer(1, 0.15*inch))

    bill_to_name = inv.get('bill_to_name') or cli.get('name','')
    bill_to_text = f"<b>Bill To:</b><br/><b>{bill_to_name}</b><br/>{inv.get('bill_to_address') or cli.get('address','')}<br/>{inv.get('bill_to_city') or cli.get('city','')}, {inv.get('bill_to_state') or cli.get('state','')} - {inv.get('bill_to_pincode') or cli.get('pincode','')}"
    if inv.get('bill_to_gstin') or cli.get('gstin'):
        bill_to_text += f"<br/>GSTIN: {inv.get('bill_to_gstin') or cli.get('gstin')}"

    if not inv.get('ship_to_same', True) and inv.get('ship_to_name'):
        ship_to_text = f"<b>Ship To:</b><br/><b>{inv.get('ship_to_name','')}</b><br/>{inv.get('ship_to_address','')}<br/>{inv.get('ship_to_city','')}, {inv.get('ship_to_state','')} - {inv.get('ship_to_pincode','')}"
        if inv.get('ship_to_gstin'): ship_to_text += f"<br/>GSTIN: {inv.get('ship_to_gstin')}"
        address_table = Table([[Paragraph(bill_to_text, from_style), Paragraph(ship_to_text, from_style)]], colWidths=[3.25*inch, 3.25*inch])
    else:
        address_table = Table([[Paragraph(bill_to_text, from_style)]], colWidths=[6.5*inch])
    address_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOX',(0,0),(-1,-1),1,theme_color),
                                        ('BACKGROUND',(0,0),(-1,-1),colors.white), ('LEFTPADDING',(0,0),(-1,-1),10),
                                        ('RIGHTPADDING',(0,0),(-1,-1),10), ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(address_table); elements.append(Spacer(1, 0.15*inch))

    # ── Transporter / E-Way Details ─────────────────────────────────────
    if inv.get('transporter_name') or inv.get('eway_bill_no'):
        trans_parts = []
        if inv.get('transporter_name'):   trans_parts.append(f"<b>Transporter:</b> {inv['transporter_name']}")
        if inv.get('transporter_gstin'):  trans_parts.append(f"<b>Transporter GSTIN:</b> {inv['transporter_gstin']}")
        if inv.get('eway_bill_no'):       trans_parts.append(f"<b>E-Way Bill No:</b> {inv['eway_bill_no']}")
        if inv.get('vehicle_no'):         trans_parts.append(f"<b>Vehicle No:</b> {inv['vehicle_no']}")
        trans_style = ParagraphStyle('TransStyle', parent=styles['Normal'], fontSize=8, leading=11)
        trans_table = Table([[Paragraph("  ".join(trans_parts), trans_style)]], colWidths=[6.5*inch])
        trans_table.setStyle(TableStyle([('BOX',(0,0),(-1,-1),1,theme_color),
                                          ('BACKGROUND',(0,0),(-1,-1),light_bg),
                                          ('LEFTPADDING',(0,0),(-1,-1),10),
                                          ('TOPPADDING',(0,0),(-1,-1),6),
                                          ('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        elements.append(trans_table); elements.append(Spacer(1, 0.1*inch))

    items_data = [['#', 'Description', 'HSN/SAC', 'Qty', 'Rate', 'GST%', 'Amount']]
    for idx, item in enumerate(inv.get('items', []), 1):
        desc = item.get('description', '')
        if item.get('item_notes') and not item.get('is_pure_agent'): desc = f"{desc}\n({item.get('item_notes')})"
        if item.get('is_pure_agent'): desc = f"{desc}\n[Pure Agent - No GST]"
        gst_display = "N/A" if item.get('is_pure_agent') else f"{item.get('gst_rate',0)}%"
        items_data.append([str(idx), Paragraph(desc.replace('\n','<br/>'), ParagraphStyle('ItemDesc', parent=styles['Normal'], fontSize=8 if use_compact else 9)),
                            item.get('hsn_sac_code','-'), str(item.get('quantity',0)),
                            f"{currency_symbol}{item.get('rate',0):.2f}", gst_display, f"{currency_symbol}{item.get('amount',0):.2f}"])

    items_table = Table(items_data, colWidths=[0.35*inch, 2.2*inch, 0.75*inch, 0.5*inch, 0.85*inch, 0.55*inch, 1.1*inch])
    items_table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),theme_color), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                                      ('ALIGN',(0,0),(-1,-1),'CENTER'), ('ALIGN',(4,1),(4,-1),'RIGHT'), ('ALIGN',(6,1),(6,-1),'RIGHT'),
                                      ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,-1),9),
                                      ('GRID',(0,0),(-1,-1),0.5,theme_color), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,light_bg]),
                                      ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    elements.append(items_table); elements.append(Spacer(1, 0.15*inch))

    summary_data = [['Subtotal:', f"{currency_symbol}{inv.get('subtotal',0):.2f}"]]
    if not inv.get('is_export') and not inv.get('is_pure_agent'):
        if inv.get('cgst',0) > 0:
            summary_data += [['CGST:', f"{currency_symbol}{inv.get('cgst',0):.2f}"], ['SGST:', f"{currency_symbol}{inv.get('sgst',0):.2f}"]]
        if inv.get('igst',0) > 0:
            summary_data.append(['IGST:', f"{currency_symbol}{inv.get('igst',0):.2f}"])
    else:
        summary_data.append(['GST:', 'Not Applicable'])
    summary_data.append(['Total:', f"{currency_symbol}{inv.get('total',0):.2f}"])
    if inv.get('advance_paid',0) > 0:
        summary_data += [['Advance Paid:', f"{currency_symbol}{inv.get('advance_paid',0):.2f}"],
                         ['Outstanding:', f"{currency_symbol}{inv.get('outstanding',0):.2f}"]]

    summary_table = Table(summary_data, colWidths=[4.8*inch, 1.5*inch])
    summary_table.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'RIGHT'), ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                                        ('FONTSIZE',(0,-1),(-1,-1),11), ('TEXTCOLOR',(0,-1),(-1,-1),theme_color),
                                        ('TOPPADDING',(0,-1),(-1,-1),6), ('LINEABOVE',(0,-1),(-1,-1),2,theme_color)]))
    elements.append(summary_table)

    if biz.get('bank_name'):
        elements.append(Spacer(1, 0.2*inch))
        bank_details = f"<b>Bank Details:</b><br/>Bank: {biz.get('bank_name')}<br/>Account: {biz.get('account_number') or 'N/A'}<br/>IFSC: {biz.get('ifsc_code') or 'N/A'}"
        if biz.get('swift_code'): bank_details += f"<br/>SWIFT: {biz.get('swift_code')}"
        if biz.get('upi_id'):     bank_details += f"<br/><b>UPI ID:</b> {biz.get('upi_id')}"
        elements.append(Paragraph(bank_details, ParagraphStyle('BankStyle', parent=styles['Normal'], fontSize=8)))

    if biz.get('terms_and_conditions'):
        elements.append(Spacer(1, 0.2*inch))
        tc_style = ParagraphStyle('TCStyle', parent=styles['Normal'], fontSize=7, textColor=colors.HexColor('#64748B'), leading=9)
        elements.append(Paragraph("<b>Terms & Conditions:</b>", tc_style))
        for line in biz.get('terms_and_conditions','').split('\n')[:5]:
            if line.strip(): elements.append(Paragraph(line.strip(), tc_style))

    if biz.get('authorised_signatory') or biz.get('signature_url'):
        elements.append(Spacer(1, 0.3*inch))
        from_s = ParagraphStyle('FromStyle', parent=styles['Normal'], fontSize=9, leading=12, alignment=1)
        sig_img_cell = ''
        if biz.get('signature_url'):
            try:
                sig_path = ROOT_DIR / biz.get('signature_url','').lstrip('/').replace('api/','').replace('business/signature/','uploads/signatures/')
                if not sig_path.exists():
                    import re as _re
                    fname = _re.split(r'[/\\]', biz.get('signature_url',''))[-1]
                    sig_path = ROOT_DIR / 'uploads' / 'signatures' / fname
                if sig_path.exists():
                    sig_img_cell = RLImage(str(sig_path), width=1.4*inch, height=0.7*inch)
            except Exception: pass
        sig_label = Paragraph(f"<b>Authorised Signatory</b><br/>{biz.get('authorised_signatory','')}", from_s)
        # Right-align: empty spacer left, signatory block right
        right_cell = [sig_img_cell, sig_label] if sig_img_cell else [sig_label]
        sig_table = Table([['', sig_img_cell, sig_label]], colWidths=[3.5*inch, 1.5*inch, 2.5*inch])
        sig_table.setStyle(TableStyle([
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('VALIGN',(0,0),(-1,-1),'BOTTOM'),
            ('LINEABOVE',(2,0),(2,0),1,colors.black),
        ]))
        elements.append(sig_table)

    doc.build(elements)
    buffer.seek(0)
    return StreamingResponse(buffer, media_type="application/pdf",
                              headers={"Content-Disposition": f"attachment; filename=invoice_{inv.get('invoice_number')}.pdf"})
