from fastapi import APIRouter, Depends, Query
from fastapi.responses import StreamingResponse
from typing import Optional
from io import BytesIO
from datetime import datetime, timezone

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Invoice, Payment, Client, Estimate, CreditNote, Advance, Business, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/reports", tags=["Reports"])


def style_header(ws, row=1):
    fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[row]:
        cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border


def auto_cols(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 50)


@router.get("/gstr1")
async def get_gstr1_report(start_date: str = Query(...), end_date: str = Query(...),
                             user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"],
                                                         Invoice.invoice_date >= start_date, Invoice.invoice_date <= end_date))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}

    b2b, b2c = [], []
    for inv in invoices:
        c = clients.get(inv.client_id, None)
        gstin = c.gstin if c else inv.client_gstin
        row = {"invoice_number": inv.invoice_number, "invoice_date": inv.invoice_date,
               "client_name": c.name if c else inv.client_name or "Unknown",
               "client_gstin": gstin or "", "place_of_supply": inv.client_state or (c.state if c else ""),
               "taxable_value": inv.subtotal or 0, "cgst": inv.cgst or 0, "sgst": inv.sgst or 0,
               "igst": inv.igst or 0, "total_tax": (inv.cgst or 0)+(inv.sgst or 0)+(inv.igst or 0),
               "invoice_value": inv.total or 0}
        (b2b if gstin else b2c).append(row)

    wb = Workbook()
    ws_b2b = wb.active; ws_b2b.title = "B2B Invoices"
    b2b_headers = ["Invoice No", "Invoice Date", "Client Name", "GSTIN", "Place of Supply",
                   "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Invoice Value"]
    for col, h in enumerate(b2b_headers, 1): ws_b2b.cell(row=1, column=col, value=h)
    for r, d in enumerate(b2b, 2):
        for col, k in enumerate(["invoice_number","invoice_date","client_name","client_gstin","place_of_supply","taxable_value","cgst","sgst","igst","total_tax","invoice_value"], 1):
            ws_b2b.cell(row=r, column=col, value=d[k])
    style_header(ws_b2b); auto_cols(ws_b2b)

    ws_b2c = wb.create_sheet("B2C Invoices")
    b2c_headers = ["Invoice No", "Invoice Date", "Client Name", "Place of Supply", "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Invoice Value"]
    for col, h in enumerate(b2c_headers, 1): ws_b2c.cell(row=1, column=col, value=h)
    for r, d in enumerate(b2c, 2):
        for col, k in enumerate(["invoice_number","invoice_date","client_name","place_of_supply","taxable_value","cgst","sgst","igst","total_tax","invoice_value"], 1):
            ws_b2c.cell(row=r, column=col, value=d[k])
    style_header(ws_b2c); auto_cols(ws_b2c)

    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=GSTR1_{start_date}_to_{end_date}.xlsx"})


@router.get("/payment-register")
async def get_payment_register(start_date: str = Query(...), end_date: str = Query(...),
                                 user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    payments = (await db.execute(select(Payment).where(Payment.tenant_id == user["tenant_id"],
                                                         Payment.payment_date >= start_date, Payment.payment_date <= end_date))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}
    invoices = {i.invoice_id: i for i in (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"]))).scalars().all()}

    wb = Workbook(); ws = wb.active; ws.title = "Payment Register"
    headers = ["Payment Date", "Invoice No", "Client Name", "GSTIN", "Payment Mode", "Reference No", "Amount", "Notes"]
    for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
    for r, p in enumerate(payments, 2):
        c = clients.get(p.client_id, {}); inv = invoices.get(p.invoice_id, {})
        ws.cell(row=r, column=1, value=p.payment_date)
        ws.cell(row=r, column=2, value=inv.invoice_number if hasattr(inv, 'invoice_number') else 'N/A')
        ws.cell(row=r, column=3, value=c.name if hasattr(c, 'name') else 'Unknown')
        ws.cell(row=r, column=4, value=c.gstin if hasattr(c, 'gstin') else '')
        ws.cell(row=r, column=5, value=p.payment_mode); ws.cell(row=r, column=6, value=p.reference_number or '')
        ws.cell(row=r, column=7, value=p.amount); ws.cell(row=r, column=8, value=p.notes or '')
    style_header(ws); auto_cols(ws)
    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=Payment_Register_{start_date}_{end_date}.xlsx"})


@router.get("/outstanding-register")
async def get_outstanding_register(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"], Invoice.outstanding > 0))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}

    wb = Workbook(); ws = wb.active; ws.title = "Outstanding Register"
    headers = ["Invoice No", "Invoice Date", "Due Date", "Client Name", "GSTIN", "Invoice Value", "Paid Amount", "Outstanding", "Overdue Days", "Status"]
    for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
    today = datetime.now().date()
    for r, inv in enumerate(invoices, 2):
        c = clients.get(inv.client_id)
        try:
            due = datetime.strptime(inv.due_date or '', '%Y-%m-%d').date()
            overdue = (today - due).days if today > due else 0
        except: overdue = 0
        ws.cell(row=r, column=1, value=inv.invoice_number); ws.cell(row=r, column=2, value=inv.invoice_date)
        ws.cell(row=r, column=3, value=inv.due_date)
        ws.cell(row=r, column=4, value=c.name if c else inv.client_name or 'Unknown')
        ws.cell(row=r, column=5, value=c.gstin if c else inv.client_gstin or '')
        ws.cell(row=r, column=6, value=inv.total or 0); ws.cell(row=r, column=7, value=inv.advance_paid or 0)
        ws.cell(row=r, column=8, value=inv.outstanding or 0); ws.cell(row=r, column=9, value=overdue)
        ws.cell(row=r, column=10, value="Overdue" if overdue > 0 else "Not Due")
    style_header(ws); auto_cols(ws)
    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": "attachment; filename=Outstanding_Register.xlsx"})


@router.get("/sale-register")
async def get_sale_register(start_date: str = Query(...), end_date: str = Query(...),
                              register_type: str = Query("summary"), user: dict = Depends(get_current_user),
                              db: AsyncSession = Depends(get_db)):
    invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"],
                                                         Invoice.invoice_date >= start_date, Invoice.invoice_date <= end_date))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}

    wb = Workbook(); ws = wb.active
    if register_type == "summary":
        ws.title = "Sale Register Summary"
        headers = ["Invoice No", "Date", "Client", "GSTIN", "State", "Diary No", "Ref No", "Subtotal", "CGST", "SGST", "IGST", "Total", "Status"]
        for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
        for r, inv in enumerate(invoices, 2):
            c = clients.get(inv.client_id)
            ws.cell(row=r,column=1,value=inv.invoice_number); ws.cell(row=r,column=2,value=inv.invoice_date)
            ws.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
            ws.cell(row=r,column=4,value=c.gstin if c else inv.client_gstin or '')
            ws.cell(row=r,column=5,value=inv.client_state or ''); ws.cell(row=r,column=6,value=inv.diary_no or '')
            ws.cell(row=r,column=7,value=inv.ref_no or ''); ws.cell(row=r,column=8,value=inv.subtotal or 0)
            ws.cell(row=r,column=9,value=inv.cgst or 0); ws.cell(row=r,column=10,value=inv.sgst or 0)
            ws.cell(row=r,column=11,value=inv.igst or 0); ws.cell(row=r,column=12,value=inv.total or 0)
            ws.cell(row=r,column=13,value=inv.status)
    else:
        ws.title = "Sale Register Items"
        headers = ["Invoice No", "Date", "Client", "GSTIN", "State", "Description", "Item Notes", "HSN/SAC", "Qty", "Rate", "GST%", "Amount", "Pure Agent"]
        for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
        r = 2
        for inv in invoices:
            c = clients.get(inv.client_id)
            for item in (inv.items or []):
                ws.cell(row=r,column=1,value=inv.invoice_number); ws.cell(row=r,column=2,value=inv.invoice_date)
                ws.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
                ws.cell(row=r,column=4,value=c.gstin if c else inv.client_gstin or '')
                ws.cell(row=r,column=5,value=inv.client_state or ''); ws.cell(row=r,column=6,value=item.get('description',''))
                ws.cell(row=r,column=7,value=item.get('item_notes','')); ws.cell(row=r,column=8,value=item.get('hsn_sac_code',''))
                ws.cell(row=r,column=9,value=item.get('quantity',0)); ws.cell(row=r,column=10,value=item.get('rate',0))
                ws.cell(row=r,column=11,value=item.get('gst_rate',0)); ws.cell(row=r,column=12,value=item.get('amount',0))
                ws.cell(row=r,column=13,value="Yes" if item.get('is_pure_agent') else "No"); r += 1

    style_header(ws); auto_cols(ws)
    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=Sale_Register_{register_type}_{start_date}_{end_date}.xlsx"})


@router.get("/tds-receivable")
async def get_tds_receivable_report(start_date: str = Query(...), end_date: str = Query(...),
                                     user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"],
                                                         Invoice.invoice_date >= start_date, Invoice.invoice_date <= end_date))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}

    wb = Workbook(); ws = wb.active; ws.title = "TDS Receivable"
    headers = ["Invoice No", "Invoice Date", "Client Name", "GSTIN", "Invoice Value", "TDS Rate (%)", "TDS Amount", "Net Receivable"]
    for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
    for r, inv in enumerate(invoices, 2):
        c = clients.get(inv.client_id)
        val = inv.total or 0; tds = round(val * 2 / 100, 2)
        ws.cell(row=r,column=1,value=inv.invoice_number); ws.cell(row=r,column=2,value=inv.invoice_date)
        ws.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
        ws.cell(row=r,column=4,value=c.gstin if c else inv.client_gstin or '')
        ws.cell(row=r,column=5,value=val); ws.cell(row=r,column=6,value=2)
        ws.cell(row=r,column=7,value=tds); ws.cell(row=r,column=8,value=round(val-tds,2))
    style_header(ws); auto_cols(ws)
    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=TDS_Receivable_{start_date}_{end_date}.xlsx"})


@router.get("/complete-report")
async def get_complete_report(start_date: str = Query(...), end_date: str = Query(...),
                               user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    business = (await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))).scalar_one_or_none()
    company_name = business.business_name if business else "Company"

    invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"],
                                                         Invoice.invoice_date >= start_date, Invoice.invoice_date <= end_date))).scalars().all()
    payments = (await db.execute(select(Payment).where(Payment.tenant_id == user["tenant_id"],
                                                         Payment.payment_date >= start_date, Payment.payment_date <= end_date))).scalars().all()
    credit_notes = (await db.execute(select(CreditNote).where(CreditNote.tenant_id == user["tenant_id"],
                                                               CreditNote.credit_date >= start_date, CreditNote.credit_date <= end_date))).scalars().all()
    estimates = (await db.execute(select(Estimate).where(Estimate.tenant_id == user["tenant_id"],
                                                          Estimate.estimate_date >= start_date, Estimate.estimate_date <= end_date))).scalars().all()
    outstanding_invoices = (await db.execute(select(Invoice).where(Invoice.tenant_id == user["tenant_id"], Invoice.outstanding > 0))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()}

    wb = Workbook()

    # Dashboard
    ws_d = wb.active; ws_d.title = "Dashboard"
    ws_d['A1'] = company_name; ws_d['A1'].font = Font(size=18, bold=True)
    ws_d['A2'] = f"Report: {start_date} to {end_date}"
    ws_d['A3'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_d['A5'] = "Metric"; ws_d['B5'] = "Count"; ws_d['C5'] = "Amount"
    summary = [
        ("Total Invoices", len(invoices), sum(i.total or 0 for i in invoices)),
        ("Total Estimates", len(estimates), sum(e.total or 0 for e in estimates)),
        ("Total Credit Notes", len(credit_notes), sum(cn.total or 0 for cn in credit_notes)),
        ("Total Payments", len(payments), sum(p.amount or 0 for p in payments)),
        ("Outstanding", "-", sum(i.outstanding or 0 for i in outstanding_invoices)),
    ]
    for r, (label, count, amount) in enumerate(summary, 6):
        ws_d.cell(row=r,column=1,value=label); ws_d.cell(row=r,column=2,value=count); ws_d.cell(row=r,column=3,value=amount)
    style_header(ws_d, 5); auto_cols(ws_d)

    # Sale Register
    ws_s = wb.create_sheet("Sale Register")
    hdrs = ["Invoice No","Date","Client","GSTIN","State","Subtotal","CGST","SGST","IGST","Total","Status"]
    for col, h in enumerate(hdrs, 1): ws_s.cell(row=1, column=col, value=h)
    for r, inv in enumerate(invoices, 2):
        c = clients.get(inv.client_id)
        ws_s.cell(row=r,column=1,value=inv.invoice_number); ws_s.cell(row=r,column=2,value=inv.invoice_date)
        ws_s.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
        ws_s.cell(row=r,column=4,value=c.gstin if c else inv.client_gstin or '')
        ws_s.cell(row=r,column=5,value=inv.client_state or ''); ws_s.cell(row=r,column=6,value=inv.subtotal or 0)
        ws_s.cell(row=r,column=7,value=inv.cgst or 0); ws_s.cell(row=r,column=8,value=inv.sgst or 0)
        ws_s.cell(row=r,column=9,value=inv.igst or 0); ws_s.cell(row=r,column=10,value=inv.total or 0)
        ws_s.cell(row=r,column=11,value=inv.status)
    style_header(ws_s); auto_cols(ws_s)

    # Payment Register
    ws_p = wb.create_sheet("Payment Register")
    inv_map = {i.invoice_id: i for i in invoices}
    for col, h in enumerate(["Date","Invoice No","Client","Mode","Reference","Amount"], 1): ws_p.cell(row=1, column=col, value=h)
    for r, p in enumerate(payments, 2):
        c = clients.get(p.client_id); inv = inv_map.get(p.invoice_id)
        ws_p.cell(row=r,column=1,value=p.payment_date)
        ws_p.cell(row=r,column=2,value=inv.invoice_number if inv else 'N/A')
        ws_p.cell(row=r,column=3,value=c.name if c else 'Unknown')
        ws_p.cell(row=r,column=4,value=p.payment_mode); ws_p.cell(row=r,column=5,value=p.reference_number or '')
        ws_p.cell(row=r,column=6,value=p.amount)
    style_header(ws_p); auto_cols(ws_p)

    # Outstanding
    ws_o = wb.create_sheet("Outstanding Register")
    for col, h in enumerate(["Invoice No","Date","Due Date","Client","GSTIN","Invoice Value","Paid","Outstanding","Overdue Days"], 1): ws_o.cell(row=1, column=col, value=h)
    today = datetime.now().date()
    for r, inv in enumerate(outstanding_invoices, 2):
        c = clients.get(inv.client_id)
        try: due = datetime.strptime(inv.due_date or '', '%Y-%m-%d').date(); overdue = max(0, (today-due).days)
        except: overdue = 0
        ws_o.cell(row=r,column=1,value=inv.invoice_number); ws_o.cell(row=r,column=2,value=inv.invoice_date)
        ws_o.cell(row=r,column=3,value=inv.due_date)
        ws_o.cell(row=r,column=4,value=c.name if c else inv.client_name or 'Unknown')
        ws_o.cell(row=r,column=5,value=c.gstin if c else inv.client_gstin or '')
        ws_o.cell(row=r,column=6,value=inv.total or 0); ws_o.cell(row=r,column=7,value=inv.advance_paid or 0)
        ws_o.cell(row=r,column=8,value=inv.outstanding or 0); ws_o.cell(row=r,column=9,value=overdue)
    style_header(ws_o); auto_cols(ws_o)

    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=Complete_Report_{start_date}_{end_date}.xlsx"})
