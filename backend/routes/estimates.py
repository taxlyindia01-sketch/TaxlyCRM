from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime, timezone, timedelta
import uuid
from io import BytesIO
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Estimate, Business, Client, SeriesConfig, Invoice, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/estimates", tags=["Estimates"])

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


class ServiceEstimateCreate(BaseModel):
    client_id: str
    estimate_date: str
    valid_until: str
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


async def get_next_estimate_number(tenant_id: str, db: AsyncSession) -> str:
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == tenant_id))
    config = result.scalar_one_or_none()
    if not config:
        config = SeriesConfig(tenant_id=tenant_id, invoice_prefix="INV", invoice_counter=0,
                              estimate_prefix="EST", estimate_counter=0,
                              credit_note_prefix="CN", credit_note_counter=0)
        db.add(config); await db.flush()
    config.estimate_counter = (config.estimate_counter or 0) + 1
    await db.flush()
    return f"{config.estimate_prefix}-{int(config.estimate_counter):05d}"


async def get_next_invoice_number(tenant_id: str, db: AsyncSession) -> str:
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == tenant_id))
    config = result.scalar_one_or_none()
    config.invoice_counter = (config.invoice_counter or 0) + 1
    await db.flush()
    return f"{config.invoice_prefix}-{int(config.invoice_counter):05d}"


def _calc_taxes(items, is_same_state, is_export):
    cgst = sgst = igst = 0.0
    if not is_export:
        for item in items:
            if item.is_pure_agent: continue
            if is_same_state:
                cgst += item.amount * (item.gst_rate / 2) / 100
                sgst += item.amount * (item.gst_rate / 2) / 100
            else:
                igst += item.amount * item.gst_rate / 100
    return cgst, sgst, igst


@router.get("")
async def get_estimates(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Estimate).where(Estimate.tenant_id == user["tenant_id"]).order_by(Estimate.created_at.desc()))
    return [serialize_doc(e) for e in result.scalars().all()]


@router.post("")
async def create_estimate(estimate_data: ServiceEstimateCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == estimate_data.client_id))
    client = c_res.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=400, detail="Client not found. Please add the client first.")
    if not business:
        import types
        business = types.SimpleNamespace(state=client.state or "")

    estimate_number = await get_next_estimate_number(user["tenant_id"], db)
    subtotal = sum(item.amount for item in estimate_data.items)
    cgst, sgst, igst = _calc_taxes(estimate_data.items, business.state == client.state, estimate_data.is_export)
    total = subtotal + cgst + sgst + igst

    est = Estimate(
        estimate_id=f"estimate_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        client_id=estimate_data.client_id,
        client_state=client.state,
        estimate_number=estimate_number,
        estimate_date=estimate_data.estimate_date,
        valid_until=estimate_data.valid_until,
        diary_no=estimate_data.diary_no,
        ref_no=estimate_data.ref_no,
        transporter_name=estimate_data.transporter_name,
        transporter_gstin=estimate_data.transporter_gstin,
        vehicle_no=estimate_data.vehicle_no,
        items=[item.model_dump() for item in estimate_data.items],
        subtotal=round(subtotal, 2),
        cgst=round(cgst, 2), sgst=round(sgst, 2), igst=round(igst, 2),
        total=round(total, 2),
        currency=estimate_data.currency,
        is_export=estimate_data.is_export,
        exchange_rate=estimate_data.exchange_rate,
        status="pending",
        notes=estimate_data.notes,
        bill_to_name=estimate_data.bill_to_name or client.name,
        bill_to_address=estimate_data.bill_to_address or client.address,
        bill_to_city=estimate_data.bill_to_city or client.city,
        bill_to_state=estimate_data.bill_to_state or client.state,
        bill_to_pincode=estimate_data.bill_to_pincode or client.pincode,
        bill_to_gstin=estimate_data.bill_to_gstin or client.gstin,
        ship_to_same=estimate_data.ship_to_same,
        ship_to_name=estimate_data.ship_to_name if not estimate_data.ship_to_same else None,
        ship_to_address=estimate_data.ship_to_address if not estimate_data.ship_to_same else None,
        ship_to_city=estimate_data.ship_to_city if not estimate_data.ship_to_same else None,
        ship_to_state=estimate_data.ship_to_state if not estimate_data.ship_to_same else None,
        ship_to_pincode=estimate_data.ship_to_pincode if not estimate_data.ship_to_same else None,
        ship_to_gstin=estimate_data.ship_to_gstin if not estimate_data.ship_to_same else None,
        created_at=datetime.now(timezone.utc)
    )
    db.add(est)
    await db.commit()
    await db.refresh(est)
    return serialize_doc(est)


@router.delete("/{estimate_id}")
async def delete_estimate(estimate_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    from sqlalchemy import delete as sql_delete
    result = await db.execute(sql_delete(Estimate).where(Estimate.estimate_id == estimate_id, Estimate.tenant_id == user["tenant_id"]))
    if result.rowcount == 0:
        raise HTTPException(status_code=404, detail="Estimate not found")
    await db.commit()
    return {"message": "Estimate deleted successfully"}


@router.get("/{estimate_id}")
async def get_estimate(estimate_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Estimate).where(Estimate.estimate_id == estimate_id, Estimate.tenant_id == user["tenant_id"]))
    est = result.scalar_one_or_none()
    if not est:
        raise HTTPException(status_code=404, detail="Estimate not found")
    return serialize_doc(est)


@router.put("/{estimate_id}")
async def update_estimate(estimate_id: str, estimate_data: ServiceEstimateCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    est_res = await db.execute(select(Estimate).where(Estimate.estimate_id == estimate_id, Estimate.tenant_id == user["tenant_id"]))
    est = est_res.scalar_one_or_none()
    if not est:
        raise HTTPException(status_code=404, detail="Estimate not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == estimate_data.client_id))
    client = c_res.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=400, detail="Client not found.")
    if not business:
        import types
        business = types.SimpleNamespace(state=client.state or "")

    subtotal = sum(item.amount for item in estimate_data.items)
    cgst, sgst, igst = _calc_taxes(estimate_data.items, business.state == client.state, estimate_data.is_export)
    total = subtotal + cgst + sgst + igst

    est.client_id = estimate_data.client_id; est.client_state = client.state
    est.estimate_date = estimate_data.estimate_date; est.valid_until = estimate_data.valid_until
    est.diary_no = estimate_data.diary_no; est.ref_no = estimate_data.ref_no
    est.items = [item.model_dump() for item in estimate_data.items]
    est.subtotal = round(subtotal, 2); est.cgst = round(cgst, 2); est.sgst = round(sgst, 2); est.igst = round(igst, 2)
    est.total = round(total, 2); est.currency = estimate_data.currency; est.is_export = estimate_data.is_export; est.exchange_rate = estimate_data.exchange_rate
    est.notes = estimate_data.notes; est.updated_at = datetime.now(timezone.utc)
    await db.commit()
    await db.refresh(est)
    return serialize_doc(est)


@router.post("/{estimate_id}/convert-to-invoice")
async def convert_estimate_to_invoice(estimate_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    est_res = await db.execute(select(Estimate).where(Estimate.estimate_id == estimate_id, Estimate.tenant_id == user["tenant_id"]))
    estimate = est_res.scalar_one_or_none()
    if not estimate:
        raise HTTPException(status_code=404, detail="Estimate not found")

    invoice_number = await get_next_invoice_number(user["tenant_id"], db)
    invoice_id = f"invoice_{uuid.uuid4().hex[:12]}"

    inv = Invoice(
        invoice_id=invoice_id,
        tenant_id=estimate.tenant_id,
        client_id=estimate.client_id,
        client_state=estimate.client_state,
        invoice_number=invoice_number,
        invoice_date=datetime.now(timezone.utc).strftime("%Y-%m-%d"),
        due_date=(datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%d"),
        items=estimate.items,
        subtotal=estimate.subtotal, cgst=estimate.cgst, sgst=estimate.sgst, igst=estimate.igst,
        total=estimate.total, currency=estimate.currency, is_export=estimate.is_export,
        status="unpaid", advance_paid=0.0, outstanding=estimate.total,
        notes=estimate.notes, created_at=datetime.now(timezone.utc)
    )
    db.add(inv)

    c_res = await db.execute(select(Client).where(Client.client_id == estimate.client_id))
    client = c_res.scalar_one_or_none()
    if client:
        client.outstanding_balance = (client.outstanding_balance or 0) + (estimate.total or 0)

    estimate.status = "converted"
    await db.commit()
    return {"invoice_id": invoice_id, "invoice_number": invoice_number, "total": estimate.total}


@router.get("/{estimate_id}/pdf")
async def download_estimate_pdf(estimate_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    est_res = await db.execute(select(Estimate).where(Estimate.estimate_id == estimate_id, Estimate.tenant_id == user["tenant_id"]))
    estimate = est_res.scalar_one_or_none()
    if not estimate:
        raise HTTPException(status_code=404, detail="Estimate not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == estimate.client_id))
    client = c_res.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=400, detail="Client not found")
    if not business:
        # Use placeholder business info if profile not set up yet
        from sqlalchemy.orm import MappedColumn
        class _DefaultBiz:
            business_name = "Your Business Name"
            gstin = ""; address = ""; city = ""; state = ""; pincode = ""
            phone = ""; email = ""; logo_url = None; qr_code_url = None
            signature_url = None; authorised_signatory = ""; upi_id = ""
            bank_name = ""; account_number = ""; ifsc_code = ""; swift_code = ""
            terms_and_conditions = "Please set up your Business Settings."
        business = _DefaultBiz()

    est = serialize_doc(estimate); biz = serialize_doc(business); cli = serialize_doc(client)
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []; styles = getSampleStyleSheet()
    theme_color = colors.HexColor('#2563EB'); light_bg = colors.HexColor('#EFF6FF')
    currency_symbol = CURRENCY_SYMBOLS.get(est.get('currency', 'INR'), "Rs.")

    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, textColor=theme_color, alignment=TA_CENTER, fontName='Helvetica-Bold')
    elements.append(Paragraph("PROFORMA INVOICE", title_style)); elements.append(Spacer(1, 0.2*inch))

    from_style = ParagraphStyle('FromStyle', parent=styles['Normal'], fontSize=9, leading=12)
    est_details = f"<b>Proforma No:</b> {est.get('estimate_number','')}<br/><b>Date:</b> {est.get('estimate_date','')}<br/><b>Valid Until:</b> {est.get('valid_until','')}<br/><b>Currency:</b> {est.get('currency','INR')}"
    header_data = [[
        Paragraph(f"<b>From:</b><br/><b>{biz.get('business_name','')}</b><br/>{biz.get('address','')}<br/>{biz.get('city','')}, {biz.get('state','')} - {biz.get('pincode','')}<br/>GSTIN: {biz.get('gstin','')}", from_style),
        Paragraph(est_details, from_style)
    ]]
    header_table = Table(header_data, colWidths=[3.5*inch, 3*inch])
    header_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOX',(0,0),(-1,-1),1,theme_color),
                                       ('BACKGROUND',(0,0),(-1,-1),light_bg), ('LEFTPADDING',(0,0),(-1,-1),10),
                                       ('RIGHTPADDING',(0,0),(-1,-1),10), ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(header_table); elements.append(Spacer(1, 0.15*inch))

    bill_to_text = f"<b>Bill To:</b><br/><b>{est.get('bill_to_name') or cli.get('name','')}</b><br/>{est.get('bill_to_address') or cli.get('address','')}<br/>{est.get('bill_to_city') or cli.get('city','')}, {est.get('bill_to_state') or cli.get('state','')} - {est.get('bill_to_pincode') or cli.get('pincode','')}"
    address_table = Table([[Paragraph(bill_to_text, from_style)]], colWidths=[6.5*inch])
    address_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOX',(0,0),(-1,-1),1,theme_color),
                                        ('BACKGROUND',(0,0),(-1,-1),colors.white), ('LEFTPADDING',(0,0),(-1,-1),10),
                                        ('RIGHTPADDING',(0,0),(-1,-1),10), ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(address_table); elements.append(Spacer(1, 0.15*inch))

    items_data = [['#', 'Description', 'HSN/SAC', 'Qty', 'Rate', 'GST%', 'Amount']]
    for idx, item in enumerate(est.get('items', []), 1):
        desc = item.get('description','')
        if item.get('item_notes') and not item.get('is_pure_agent'): desc = f"{desc}\n({item.get('item_notes')})"
        gst_display = "N/A" if item.get('is_pure_agent') else f"{item.get('gst_rate',0)}%"
        items_data.append([str(idx), Paragraph(desc.replace('\n','<br/>'), ParagraphStyle('ID', parent=styles['Normal'], fontSize=9)),
                            item.get('hsn_sac_code','-'), str(item.get('quantity',0)),
                            f"{currency_symbol}{item.get('rate',0):.2f}", gst_display, f"{currency_symbol}{item.get('amount',0):.2f}"])

    items_table = Table(items_data, colWidths=[0.35*inch, 2.2*inch, 0.75*inch, 0.5*inch, 0.85*inch, 0.55*inch, 1.1*inch])
    items_table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),theme_color), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                                      ('ALIGN',(0,0),(-1,-1),'CENTER'), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                                      ('FONTSIZE',(0,0),(-1,-1),9), ('GRID',(0,0),(-1,-1),0.5,theme_color),
                                      ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,light_bg]), ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    elements.append(items_table); elements.append(Spacer(1, 0.15*inch))

    summary_data = [['Subtotal:', f"{currency_symbol}{est.get('subtotal',0):.2f}"]]
    if not est.get('is_export'):
        if est.get('cgst',0)>0: summary_data += [['CGST:',f"{currency_symbol}{est.get('cgst',0):.2f}"],['SGST:',f"{currency_symbol}{est.get('sgst',0):.2f}"]]
        if est.get('igst',0)>0: summary_data.append(['IGST:',f"{currency_symbol}{est.get('igst',0):.2f}"])
    summary_data.append(['Total:', f"{currency_symbol}{est.get('total',0):.2f}"])
    summary_table = Table(summary_data, colWidths=[4.8*inch, 1.5*inch])
    summary_table.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'RIGHT'), ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                                        ('FONTSIZE',(0,-1),(-1,-1),11), ('TEXTCOLOR',(0,-1),(-1,-1),theme_color),
                                        ('LINEABOVE',(0,-1),(-1,-1),2,theme_color)]))
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

    doc.build(elements); buffer.seek(0)
    return StreamingResponse(buffer, media_type="application/pdf",
                              headers={"Content-Disposition": f"attachment; filename=proforma_{est.get('estimate_number')}.pdf"})
