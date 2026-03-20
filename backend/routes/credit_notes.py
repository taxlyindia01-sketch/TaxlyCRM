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
from reportlab.lib.enums import TA_CENTER

from sqlalchemy import select, delete as sql_delete
from sqlalchemy.ext.asyncio import AsyncSession

from database import CreditNote, Invoice, Business, Client, SeriesConfig, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/credit-notes", tags=["Credit Notes"])

ROOT_DIR = Path(__file__).parent.parent
LOGO_UPLOAD_DIR = ROOT_DIR / "uploads" / "logos"
CURRENCY_SYMBOLS = {"INR": "Rs.", "USD": "$", "EUR": "EUR ", "GBP": "GBP "}


class CreditNoteItem(BaseModel):
    description: str
    hsn_sac_code: Optional[str] = None
    quantity: float
    rate: float
    gst_rate: float
    amount: float
    item_notes: Optional[str] = None
    is_pure_agent: Optional[bool] = False


class CreditNoteCreate(BaseModel):
    invoice_id: str
    client_id: str
    credit_date: str
    items: List[CreditNoteItem]
    currency: str = "INR"
    diary_no: Optional[str] = None
    ref_no: Optional[str] = None
    reason: Optional[str] = None
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
    is_export: Optional[bool] = False
    exchange_rate: Optional[float] = None


async def get_next_credit_note_number(tenant_id: str, db: AsyncSession) -> str:
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == tenant_id))
    config = result.scalar_one_or_none()
    if not config:
        config = SeriesConfig(tenant_id=tenant_id, invoice_prefix="INV", invoice_counter=0,
                              estimate_prefix="EST", estimate_counter=0,
                              credit_note_prefix="CN", credit_note_counter=0)
        db.add(config); await db.flush()
    config.credit_note_counter = (config.credit_note_counter or 0) + 1
    await db.flush()
    return f"{config.credit_note_prefix}-{int(config.credit_note_counter):05d}"


def _calc_taxes(items, is_same_state, is_export=False):
    cgst = sgst = igst = 0.0
    if is_export: return 0.0, 0.0, 0.0
    for item in items:
        if item.is_pure_agent: continue
        if is_same_state:
            cgst += item.amount * (item.gst_rate / 2) / 100
            sgst += item.amount * (item.gst_rate / 2) / 100
        else:
            igst += item.amount * item.gst_rate / 100
    return cgst, sgst, igst


@router.get("")
async def get_credit_notes(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(CreditNote).where(CreditNote.tenant_id == user["tenant_id"]).order_by(CreditNote.created_at.desc()))
    return [serialize_doc(cn) for cn in result.scalars().all()]


@router.post("")
async def create_credit_note(cn_data: CreditNoteCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == cn_data.invoice_id, Invoice.tenant_id == user["tenant_id"]))
    invoice = inv_res.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == cn_data.client_id))
    client = c_res.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=400, detail="Client not found. Please add the client first.")
    if not business:
        import types
        business = types.SimpleNamespace(state=client.state or "")

    cn_number = await get_next_credit_note_number(user["tenant_id"], db)
    subtotal = sum(item.amount for item in cn_data.items)
    cgst, sgst, igst = _calc_taxes(cn_data.items, business.state == client.state, cn_data.is_export or False)
    total = subtotal + cgst + sgst + igst

    cn = CreditNote(
        credit_note_id=f"cn_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        invoice_id=cn_data.invoice_id,
        client_id=cn_data.client_id,
        client_state=client.state,
        credit_note_number=cn_number,
        credit_date=cn_data.credit_date,
        diary_no=cn_data.diary_no, ref_no=cn_data.ref_no,
        is_export=cn_data.is_export or False,
        exchange_rate=cn_data.exchange_rate,
        items=[item.model_dump() for item in cn_data.items],
        subtotal=round(subtotal, 2), cgst=round(cgst, 2), sgst=round(sgst, 2), igst=round(igst, 2),
        total=round(total, 2), currency=cn_data.currency, reason=cn_data.reason,
        created_at=datetime.now(timezone.utc)
    )
    db.add(cn)

    new_outstanding = round((invoice.outstanding or 0) - total, 2)
    invoice.outstanding = max(0, new_outstanding)  # Can't go below 0
    if invoice.outstanding <= 0:
        invoice.status = "paid"
    elif invoice.advance_paid and invoice.advance_paid > 0:
        invoice.status = "partial"
    client.outstanding_balance = (client.outstanding_balance or 0) - total

    await db.commit(); await db.refresh(cn)
    return serialize_doc(cn)


@router.delete("/{credit_note_id}")
async def delete_credit_note(credit_note_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    cn_res = await db.execute(select(CreditNote).where(CreditNote.credit_note_id == credit_note_id, CreditNote.tenant_id == user["tenant_id"]))
    cn = cn_res.scalar_one_or_none()
    if not cn:
        raise HTTPException(status_code=404, detail="Credit note not found")

    inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == cn.invoice_id))
    invoice = inv_res.scalar_one_or_none()
    if invoice:
        invoice.outstanding = (invoice.outstanding or 0) + cn.total

    c_res = await db.execute(select(Client).where(Client.client_id == cn.client_id))
    client = c_res.scalar_one_or_none()
    if client:
        client.outstanding_balance = (client.outstanding_balance or 0) + cn.total

    await db.execute(sql_delete(CreditNote).where(CreditNote.credit_note_id == credit_note_id))
    await db.commit()
    return {"message": "Credit note deleted successfully"}


@router.get("/{credit_note_id}")
async def get_credit_note(credit_note_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(CreditNote).where(CreditNote.credit_note_id == credit_note_id, CreditNote.tenant_id == user["tenant_id"]))
    cn = result.scalar_one_or_none()
    if not cn:
        raise HTTPException(status_code=404, detail="Credit note not found")
    return serialize_doc(cn)


@router.put("/{credit_note_id}")
async def update_credit_note(credit_note_id: str, cn_data: CreditNoteCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    cn_res = await db.execute(select(CreditNote).where(CreditNote.credit_note_id == credit_note_id, CreditNote.tenant_id == user["tenant_id"]))
    cn = cn_res.scalar_one_or_none()
    if not cn:
        raise HTTPException(status_code=404, detail="Credit note not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == cn_data.client_id))
    client = c_res.scalar_one_or_none()
    inv_res = await db.execute(select(Invoice).where(Invoice.invoice_id == cn_data.invoice_id))
    invoice = inv_res.scalar_one_or_none()
    if not business or not client or not invoice:
        raise HTTPException(status_code=400, detail="Business, Client, or Invoice not found")

    subtotal = sum(item.amount for item in cn_data.items)
    cgst, sgst, igst = _calc_taxes(cn_data.items, business.state == client.state, cn_data.is_export or False)
    total = subtotal + cgst + sgst + igst
    old_total = cn.total or 0
    diff = total - old_total

    cn.invoice_id = cn_data.invoice_id; cn.client_id = cn_data.client_id; cn.client_state = client.state
    cn.credit_date = cn_data.credit_date; cn.diary_no = cn_data.diary_no; cn.ref_no = cn_data.ref_no
    cn.items = [item.model_dump() for item in cn_data.items]
    cn.subtotal = round(subtotal, 2); cn.cgst = round(cgst, 2); cn.sgst = round(sgst, 2); cn.igst = round(igst, 2)
    cn.total = round(total, 2); cn.currency = cn_data.currency; cn.reason = cn_data.reason
    cn.is_export = cn_data.is_export or False; cn.exchange_rate = cn_data.exchange_rate
    cn.updated_at = datetime.now(timezone.utc)

    if diff != 0:
        invoice.outstanding = (invoice.outstanding or 0) - diff
        client.outstanding_balance = (client.outstanding_balance or 0) - diff

    await db.commit(); await db.refresh(cn)
    return serialize_doc(cn)


@router.get("/{credit_note_id}/pdf")
async def download_credit_note_pdf(credit_note_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    cn_res = await db.execute(select(CreditNote).where(CreditNote.credit_note_id == credit_note_id, CreditNote.tenant_id == user["tenant_id"]))
    credit_note = cn_res.scalar_one_or_none()
    if not credit_note:
        raise HTTPException(status_code=404, detail="Credit note not found")

    b_res = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = b_res.scalar_one_or_none()
    c_res = await db.execute(select(Client).where(Client.client_id == credit_note.client_id))
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

    cn = serialize_doc(credit_note); biz = serialize_doc(business); cli = serialize_doc(client)
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []; styles = getSampleStyleSheet()
    theme_color = colors.HexColor('#DC2626'); light_bg = colors.HexColor('#FEF2F2')
    currency_symbol = CURRENCY_SYMBOLS.get(cn.get('currency', 'INR'), "Rs.")

    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, textColor=theme_color, alignment=TA_CENTER, fontName='Helvetica-Bold')
    elements.append(Paragraph("CREDIT NOTE", title_style)); elements.append(Spacer(1, 0.2*inch))

    from_style = ParagraphStyle('FromStyle', parent=styles['Normal'], fontSize=9, leading=12)
    cn_details = f"<b>Credit Note No:</b> {cn.get('credit_note_number','')}<br/><b>Date:</b> {cn.get('credit_date','')}<br/><b>Currency:</b> {cn.get('currency','INR')}"

    header_data = [[
        Paragraph(f"<b>From:</b><br/><b>{biz.get('business_name','')}</b><br/>{biz.get('address','')}<br/>{biz.get('city','')}, {biz.get('state','')} - {biz.get('pincode','')}<br/>GSTIN: {biz.get('gstin','')}", from_style),
        Paragraph(cn_details, from_style)
    ]]
    header_table = Table(header_data, colWidths=[3.5*inch, 3*inch])
    header_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOX',(0,0),(-1,-1),1,theme_color),
                                       ('BACKGROUND',(0,0),(-1,-1),light_bg), ('LEFTPADDING',(0,0),(-1,-1),10),
                                       ('RIGHTPADDING',(0,0),(-1,-1),10), ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(header_table); elements.append(Spacer(1, 0.15*inch))

    bill_to_text = f"<b>Credit To:</b><br/><b>{cn.get('bill_to_name') or cli.get('name','')}</b><br/>{cn.get('bill_to_address') or cli.get('address','')}<br/>{cn.get('bill_to_city') or cli.get('city','')}, {cn.get('bill_to_state') or cli.get('state','')} - {cn.get('bill_to_pincode') or cli.get('pincode','')}"
    address_table = Table([[Paragraph(bill_to_text, from_style)]], colWidths=[6.5*inch])
    address_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOX',(0,0),(-1,-1),1,theme_color),
                                        ('BACKGROUND',(0,0),(-1,-1),colors.white), ('LEFTPADDING',(0,0),(-1,-1),10),
                                        ('RIGHTPADDING',(0,0),(-1,-1),10), ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    elements.append(address_table); elements.append(Spacer(1, 0.15*inch))

    items_data = [['#', 'Description', 'HSN/SAC', 'Qty', 'Rate', 'GST%', 'Amount']]
    for idx, item in enumerate(cn.get('items', []), 1):
        desc = item.get('description','')
        if item.get('item_notes') and not item.get('is_pure_agent'): desc += f"\n({item.get('item_notes')})"
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

    is_exp_cn = cn.get('is_export', False)
    summary_data = [['Subtotal:', f"{currency_symbol}{cn.get('subtotal',0):.2f}"]]
    if is_exp_cn:
        summary_data.append(['GST:', 'Zero Rated (Export)'])
    else:
        if cn.get('cgst',0)>0: summary_data += [['CGST:',f"{currency_symbol}{cn.get('cgst',0):.2f}"],['SGST:',f"{currency_symbol}{cn.get('sgst',0):.2f}"]]
        if cn.get('igst',0)>0: summary_data.append(['IGST:',f"{currency_symbol}{cn.get('igst',0):.2f}"])
    summary_data.append(['Total Credit:', f"{currency_symbol}{cn.get('total',0):.2f}"])
    summary_table = Table(summary_data, colWidths=[4.8*inch, 1.5*inch])
    summary_table.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'RIGHT'), ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                                        ('FONTSIZE',(0,-1),(-1,-1),11), ('TEXTCOLOR',(0,-1),(-1,-1),theme_color),
                                        ('LINEABOVE',(0,-1),(-1,-1),2,theme_color)]))
    elements.append(summary_table)

    if cn.get('reason'):
        elements.append(Spacer(1, 0.15*inch))
        elements.append(Paragraph(f"<b>Reason:</b> {cn.get('reason')}", from_style))

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
                              headers={"Content-Disposition": f"attachment; filename=credit_note_{cn.get('credit_note_number')}.pdf"})
