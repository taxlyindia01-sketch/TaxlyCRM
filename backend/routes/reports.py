from fastapi import APIRouter, Depends, Query, HTTPException
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

def add_totals_row(ws, numeric_cols, start_row=2, label_col=1, label="TOTAL"):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    if ws.max_row < start_row: return
    total_row = ws.max_row + 1
    total_fill = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    total_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='medium'), bottom=Side(style='medium'))
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=total_row, column=col_idx)
        if col_idx == label_col:
            cell.value = label
        elif col_idx in numeric_cols:
            col_vals = [ws.cell(row=r, column=col_idx).value or 0
                        for r in range(start_row, total_row)
                        if isinstance(ws.cell(row=r, column=col_idx).value, (int, float))]
            cell.value = round(sum(col_vals), 2)
        cell.fill = total_fill; cell.font = total_font
        cell.alignment = Alignment(horizontal='right' if col_idx in numeric_cols else 'center')
        cell.border = border

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
    add_totals_row(ws_b2b, {6,7,8,9,10,11})
    style_header(ws_b2b); auto_cols(ws_b2b)

    ws_b2c = wb.create_sheet("B2C Invoices")
    b2c_headers = ["Invoice No", "Invoice Date", "Client Name", "Place of Supply", "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Invoice Value"]
    for col, h in enumerate(b2c_headers, 1): ws_b2c.cell(row=1, column=col, value=h)
    for r, d in enumerate(b2c, 2):
        for col, k in enumerate(["invoice_number","invoice_date","client_name","place_of_supply","taxable_value","cgst","sgst","igst","total_tax","invoice_value"], 1):
            ws_b2c.cell(row=r, column=col, value=d[k])
    add_totals_row(ws_b2c, {5,6,7,8,9,10})
    style_header(ws_b2c); auto_cols(ws_b2c)

    # Credit Notes sheet (CDNR - Credit/Debit Note Register)
    credit_notes_list = (await db.execute(select(CreditNote).where(
        CreditNote.tenant_id == user["tenant_id"],
        CreditNote.credit_date >= start_date, CreditNote.credit_date <= end_date
    ))).scalars().all()

    ws_cdnr = wb.create_sheet("CDNR - Credit Notes")
    cdnr_headers = ["Credit Note No", "Credit Note Date", "Client Name", "Client GSTIN",
                    "Place of Supply", "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Credit Note Value"]
    for col, h in enumerate(cdnr_headers, 1): ws_cdnr.cell(row=1, column=col, value=h)
    for r, cn in enumerate(credit_notes_list, 2):
        cli = clients.get(cn.client_id)
        ws_cdnr.cell(row=r, column=1, value=cn.credit_note_number)
        ws_cdnr.cell(row=r, column=2, value=cn.credit_date)
        ws_cdnr.cell(row=r, column=3, value=cli.name if cli else "Unknown")
        ws_cdnr.cell(row=r, column=4, value=cli.gstin if cli else "")
        ws_cdnr.cell(row=r, column=5, value=cn.client_state or (cli.state if cli else ""))
        ws_cdnr.cell(row=r, column=6, value=cn.subtotal or 0)
        ws_cdnr.cell(row=r, column=7, value=cn.cgst or 0)
        ws_cdnr.cell(row=r, column=8, value=cn.sgst or 0)
        ws_cdnr.cell(row=r, column=9, value=cn.igst or 0)
        total_tax = (cn.cgst or 0) + (cn.sgst or 0) + (cn.igst or 0)
        ws_cdnr.cell(row=r, column=10, value=total_tax)
        ws_cdnr.cell(row=r, column=11, value=cn.total or 0)
    add_totals_row(ws_cdnr, {6,7,8,9,10,11})
    style_header(ws_cdnr); auto_cols(ws_cdnr)

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
    add_totals_row(ws, {7})
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
    add_totals_row(ws, {6,7,8,9})
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
    elif register_type in ('itemized', 'items', 'detail', 'item'):
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

    # Credit Notes sheet
    credit_notes_sr = (await db.execute(select(CreditNote).where(
        CreditNote.tenant_id == user["tenant_id"],
        CreditNote.credit_date >= start_date, CreditNote.credit_date <= end_date
    ))).scalars().all()
    if credit_notes_sr:
        ws_cn = wb.create_sheet("Credit Notes Issued")
        cn_hdrs = ["Credit Note No", "Date", "Client", "GSTIN", "State", "Reason", "Subtotal", "CGST", "SGST", "IGST", "Total"]
        for col, h in enumerate(cn_hdrs, 1): ws_cn.cell(row=1, column=col, value=h)
        for r, cn in enumerate(credit_notes_sr, 2):
            c = clients.get(cn.client_id)
            ws_cn.cell(row=r,column=1,value=cn.credit_note_number); ws_cn.cell(row=r,column=2,value=cn.credit_date)
            ws_cn.cell(row=r,column=3,value=c.name if c else "Unknown"); ws_cn.cell(row=r,column=4,value=c.gstin if c else "")
            ws_cn.cell(row=r,column=5,value=cn.client_state or ""); ws_cn.cell(row=r,column=6,value=cn.reason or "")
            ws_cn.cell(row=r,column=7,value=cn.subtotal or 0); ws_cn.cell(row=r,column=8,value=cn.cgst or 0)
            ws_cn.cell(row=r,column=9,value=cn.sgst or 0); ws_cn.cell(row=r,column=10,value=cn.igst or 0)
            ws_cn.cell(row=r,column=11,value=cn.total or 0)
        add_totals_row(ws_cn, {7,8,9,10,11})
        style_header(ws_cn); auto_cols(ws_cn)

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
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.series import DataPoint

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
    advances = (await db.execute(select(Advance).where(Advance.tenant_id == user["tenant_id"],
                                                        Advance.payment_date >= start_date, Advance.payment_date <= end_date))).scalars().all()
    clients_list = (await db.execute(select(Client).where(Client.tenant_id == user["tenant_id"]))).scalars().all()
    clients = {c.client_id: c for c in clients_list}

    today_dt = datetime.now().date()

    # ── Computed summaries ────────────────────────────────────────
    total_revenue = sum(i.total or 0 for i in invoices)
    total_collected = sum(p.amount or 0 for p in payments)
    total_outstanding = sum(i.outstanding or 0 for i in outstanding_invoices)
    total_cn = sum(cn.total or 0 for cn in credit_notes)
    total_est = sum(e.total or 0 for e in estimates)
    total_adv = sum(a.amount or 0 for a in advances)
    paid_inv = sum(1 for i in invoices if i.status == 'Paid')
    partial_inv = sum(1 for i in invoices if i.status == 'Partial')
    unpaid_inv = sum(1 for i in invoices if i.status not in ('Paid', 'Partial'))

    # ── COLOURS ───────────────────────────────────────────────────
    C_DARK   = "0F172A"   # header bg
    C_BLUE   = "1E40AF"   # totals row
    C_TEAL   = "0D9488"   # accent 1
    C_GREEN  = "16A34A"   # accent 2
    C_AMBER  = "D97706"   # accent 3
    C_RED    = "DC2626"   # accent 4
    C_INDIGO = "4338CA"   # accent 5
    C_WHITE  = "FFFFFF"
    C_LGRAY  = "F1F5F9"   # alt row bg
    C_LGRAY2 = "E2E8F0"

    def hdr_style(ws, row=1, color=C_DARK):
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        font = Font(color=C_WHITE, bold=True, size=10)
        bdr  = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ws[row]:
            cell.fill = fill; cell.font = font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = bdr
        ws.row_dimensions[row].height = 18

    def alt_row(ws, row, start_col=1):
        fill = PatternFill(start_color=C_LGRAY, end_color=C_LGRAY, fill_type="solid")
        for col in range(start_col, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    def totals(ws, num_cols):
        add_totals_row(ws, num_cols)

    wb = Workbook()

    # ══════════════════════════════════════════════════════════════
    # SHEET 1 — MIS DASHBOARD
    # ══════════════════════════════════════════════════════════════
    ws_d = wb.active; ws_d.title = "MIS Dashboard"
    ws_d.sheet_view.showGridLines = False
    ws_d.column_dimensions['A'].width = 28
    ws_d.column_dimensions['B'].width = 18
    ws_d.column_dimensions['C'].width = 18
    ws_d.column_dimensions['D'].width = 18
    ws_d.column_dimensions['E'].width = 18

    # Title block
    ws_d.merge_cells('A1:E1')
    t = ws_d['A1']; t.value = company_name
    t.font = Font(size=20, bold=True, color=C_DARK)
    t.alignment = Alignment(horizontal='center', vertical='center')
    t.fill = PatternFill(start_color=C_LGRAY2, end_color=C_LGRAY2, fill_type="solid")
    ws_d.row_dimensions[1].height = 32

    ws_d.merge_cells('A2:E2')
    t2 = ws_d['A2']; t2.value = f"MIS Report  |  Period: {start_date}  to  {end_date}  |  Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    t2.font = Font(size=10, color="64748B", italic=True)
    t2.alignment = Alignment(horizontal='center')
    ws_d.row_dimensions[2].height = 16

    # KPI tiles (row 4–9)
    kpis = [
        ("💰 Total Revenue",      f"₹{total_revenue:,.2f}",  C_INDIGO),
        ("✅ Total Collected",     f"₹{total_collected:,.2f}", C_GREEN),
        ("⏳ Outstanding",         f"₹{total_outstanding:,.2f}", C_AMBER),
        ("📋 Invoices (Period)",   f"{len(invoices)}",         C_BLUE),
        ("🔴 Credit Notes",        f"₹{total_cn:,.2f}",        C_RED),
        ("📊 Estimates (Period)",  f"₹{total_est:,.2f}",       C_TEAL),
        ("💳 Advance Received",    f"₹{total_adv:,.2f}",       C_INDIGO),
        ("🏢 Active Clients",      f"{len(clients_list)}",     C_GREEN),
    ]
    ws_d.row_dimensions[4].height = 14
    for i, (label, value, color) in enumerate(kpis):
        row = 5 + i
        ws_d.row_dimensions[row].height = 20
        lc = ws_d.cell(row=row, column=1, value=label)
        lc.font = Font(bold=True, size=10, color=color)
        lc.fill = PatternFill(start_color=C_LGRAY, end_color=C_LGRAY, fill_type="solid")
        vc = ws_d.cell(row=row, column=2, value=value)
        vc.font = Font(bold=True, size=11, color=C_DARK)
        vc.alignment = Alignment(horizontal='right')
        vc.fill = PatternFill(start_color=C_LGRAY, end_color=C_LGRAY, fill_type="solid")

    # Invoice status breakdown table (for chart)
    ws_d['A14'] = "Invoice Status"; ws_d['B14'] = "Count"
    for c in [ws_d['A14'], ws_d['B14']]:
        c.font = Font(bold=True, color=C_WHITE)
        c.fill = PatternFill(start_color=C_DARK, end_color=C_DARK, fill_type="solid")
        c.alignment = Alignment(horizontal='center')
    ws_d['A15'] = "Paid";    ws_d['B15'] = paid_inv
    ws_d['A16'] = "Partial"; ws_d['B16'] = partial_inv
    ws_d['A17'] = "Unpaid";  ws_d['B17'] = unpaid_inv

    # Revenue summary table (for bar chart)
    ws_d['D14'] = "Category";    ws_d['E14'] = "Amount (₹)"
    for c in [ws_d['D14'], ws_d['E14']]:
        c.font = Font(bold=True, color=C_WHITE)
        c.fill = PatternFill(start_color=C_DARK, end_color=C_DARK, fill_type="solid")
        c.alignment = Alignment(horizontal='center')
    ws_d['D15'] = "Revenue";       ws_d['E15'] = round(total_revenue, 2)
    ws_d['D16'] = "Collected";     ws_d['E16'] = round(total_collected, 2)
    ws_d['D17'] = "Outstanding";   ws_d['E17'] = round(total_outstanding, 2)
    ws_d['D18'] = "Credit Notes";  ws_d['E18'] = round(total_cn, 2)

    # Pie chart — invoice status
    pie = PieChart(); pie.title = "Invoice Status Breakdown"; pie.style = 10
    pie.width = 12; pie.height = 10
    data_ref = Reference(ws_d, min_col=2, min_row=14, max_row=17)
    cats_ref = Reference(ws_d, min_col=1, min_row=15, max_row=17)
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.dataLabels = None
    slices = [DataPoint(idx=0), DataPoint(idx=1), DataPoint(idx=2)]
    slices[0].graphicalProperties.solidFill = C_GREEN
    slices[1].graphicalProperties.solidFill = C_AMBER
    slices[2].graphicalProperties.solidFill = C_RED
    pie.series[0].data_points = slices
    ws_d.add_chart(pie, "A19")

    # Bar chart — revenue vs collection
    bar = BarChart(); bar.type = "col"; bar.title = "Revenue vs Collection vs Outstanding"
    bar.style = 10; bar.width = 16; bar.height = 10
    bar.grouping = "clustered"
    bdata = Reference(ws_d, min_col=5, min_row=14, max_row=18)
    bcats = Reference(ws_d, min_col=4, min_row=15, max_row=18)
    bar.add_data(bdata, titles_from_data=True)
    bar.set_categories(bcats)
    bar.series[0].graphicalProperties.solidFill = C_INDIGO
    ws_d.add_chart(bar, "D19")

    # ── HELPER: write invoice list sheet ──────────────────────────
    def write_invoice_sheet(ws, inv_list, color=C_DARK):
        hdrs = ["Invoice No","Date","Client","GSTIN","State","Diary No","Ref No","Subtotal","CGST","SGST","IGST","Total","Status"]
        for col, h in enumerate(hdrs, 1): ws.cell(row=1, column=col, value=h)
        for r, inv in enumerate(inv_list, 2):
            c = clients.get(inv.client_id)
            ws.cell(row=r,column=1,value=inv.invoice_number); ws.cell(row=r,column=2,value=inv.invoice_date)
            ws.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
            ws.cell(row=r,column=4,value=c.gstin if c else getattr(inv,'client_gstin','') or '')
            ws.cell(row=r,column=5,value=getattr(inv,'client_state','') or '')
            ws.cell(row=r,column=6,value=getattr(inv,'diary_no','') or '')
            ws.cell(row=r,column=7,value=getattr(inv,'ref_no','') or '')
            ws.cell(row=r,column=8,value=getattr(inv,'subtotal',0) or 0)
            ws.cell(row=r,column=9,value=getattr(inv,'cgst',0) or 0)
            ws.cell(row=r,column=10,value=getattr(inv,'sgst',0) or 0)
            ws.cell(row=r,column=11,value=getattr(inv,'igst',0) or 0)
            ws.cell(row=r,column=12,value=getattr(inv,'total',0) or 0)
            ws.cell(row=r,column=13,value=getattr(inv,'status',''))
            if r % 2 == 0: alt_row(ws, r)
        add_totals_row(ws, {8,9,10,11,12})
        hdr_style(ws, color=color); auto_cols(ws)

    def write_items_sheet(ws, doc_list, color=C_DARK):
        hdrs = ["Doc No","Date","Client","GSTIN","Description","Item Notes","HSN/SAC","Qty","Rate","GST%","Amount","Pure Agent"]
        for col, h in enumerate(hdrs, 1): ws.cell(row=1, column=col, value=h)
        r = 2
        for doc in doc_list:
            c = clients.get(doc.client_id)
            for item in (doc.items or []):
                ws.cell(row=r,column=1,value=getattr(doc,'invoice_number',None) or getattr(doc,'estimate_number',None) or getattr(doc,'credit_note_number',None))
                ws.cell(row=r,column=2,value=getattr(doc,'invoice_date',None) or getattr(doc,'estimate_date',None) or getattr(doc,'credit_date',None))
                ws.cell(row=r,column=3,value=c.name if c else getattr(doc,'client_name','') or 'Unknown')
                ws.cell(row=r,column=4,value=c.gstin if c else getattr(doc,'client_gstin','') or '')
                ws.cell(row=r,column=5,value=item.get('description',''))
                ws.cell(row=r,column=6,value=item.get('item_notes',''))
                ws.cell(row=r,column=7,value=item.get('hsn_sac_code',''))
                ws.cell(row=r,column=8,value=item.get('quantity',0))
                ws.cell(row=r,column=9,value=item.get('rate',0))
                ws.cell(row=r,column=10,value=item.get('gst_rate',0))
                ws.cell(row=r,column=11,value=item.get('amount',0))
                ws.cell(row=r,column=12,value="Yes" if item.get('is_pure_agent') else "No")
                if r % 2 == 0: alt_row(ws, r)
                r += 1
        add_totals_row(ws, {8,9,11})
        hdr_style(ws, color=color); auto_cols(ws)

    # ══════════════════════════════════════════════════════════════
    # SHEET 2 — GSTR-1 B2B
    # ══════════════════════════════════════════════════════════════
    ws_b2b = wb.create_sheet("GSTR-1 B2B")
    b2b_hdrs = ["Invoice No","Invoice Date","Client Name","GSTIN","Place of Supply","Taxable Value","CGST","SGST","IGST","Total Tax","Invoice Value"]
    for col, h in enumerate(b2b_hdrs, 1): ws_b2b.cell(row=1, column=col, value=h)
    r = 2
    for inv in invoices:
        c = clients.get(inv.client_id)
        gstin = c.gstin if c else inv.client_gstin
        if not gstin: continue
        ws_b2b.cell(row=r,column=1,value=inv.invoice_number); ws_b2b.cell(row=r,column=2,value=inv.invoice_date)
        ws_b2b.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
        ws_b2b.cell(row=r,column=4,value=gstin); ws_b2b.cell(row=r,column=5,value=inv.client_state or '')
        ws_b2b.cell(row=r,column=6,value=inv.subtotal or 0); ws_b2b.cell(row=r,column=7,value=inv.cgst or 0)
        ws_b2b.cell(row=r,column=8,value=inv.sgst or 0); ws_b2b.cell(row=r,column=9,value=inv.igst or 0)
        ws_b2b.cell(row=r,column=10,value=(inv.cgst or 0)+(inv.sgst or 0)+(inv.igst or 0))
        ws_b2b.cell(row=r,column=11,value=inv.total or 0)
        if r % 2 == 0: alt_row(ws_b2b, r)
        r += 1
    add_totals_row(ws_b2b, {6,7,8,9,10,11}); hdr_style(ws_b2b, color=C_INDIGO); auto_cols(ws_b2b)

    # ══════════════════════════════════════════════════════════════
    # SHEET 3 — GSTR-1 B2C
    # ══════════════════════════════════════════════════════════════
    ws_b2c = wb.create_sheet("GSTR-1 B2C")
    b2c_hdrs = ["Invoice No","Invoice Date","Client Name","Place of Supply","Taxable Value","CGST","SGST","IGST","Total Tax","Invoice Value"]
    for col, h in enumerate(b2c_hdrs, 1): ws_b2c.cell(row=1, column=col, value=h)
    r = 2
    for inv in invoices:
        c = clients.get(inv.client_id)
        gstin = c.gstin if c else inv.client_gstin
        if gstin: continue
        ws_b2c.cell(row=r,column=1,value=inv.invoice_number); ws_b2c.cell(row=r,column=2,value=inv.invoice_date)
        ws_b2c.cell(row=r,column=3,value=c.name if c else inv.client_name or 'Unknown')
        ws_b2c.cell(row=r,column=4,value=inv.client_state or '')
        ws_b2c.cell(row=r,column=5,value=inv.subtotal or 0); ws_b2c.cell(row=r,column=6,value=inv.cgst or 0)
        ws_b2c.cell(row=r,column=7,value=inv.sgst or 0); ws_b2c.cell(row=r,column=8,value=inv.igst or 0)
        ws_b2c.cell(row=r,column=9,value=(inv.cgst or 0)+(inv.sgst or 0)+(inv.igst or 0))
        ws_b2c.cell(row=r,column=10,value=inv.total or 0)
        if r % 2 == 0: alt_row(ws_b2c, r)
        r += 1
    add_totals_row(ws_b2c, {5,6,7,8,9,10}); hdr_style(ws_b2c, color=C_TEAL); auto_cols(ws_b2c)

    # ══════════════════════════════════════════════════════════════
    # SHEET 4 — Sale Summary
    # ══════════════════════════════════════════════════════════════
    ws_ss = wb.create_sheet("Sale Summary")
    write_invoice_sheet(ws_ss, invoices, color=C_BLUE)

    # ══════════════════════════════════════════════════════════════
    # SHEET 5 — Sale Item Wise
    # ══════════════════════════════════════════════════════════════
    ws_si = wb.create_sheet("Sale Item Wise")
    write_items_sheet(ws_si, invoices, color=C_DARK)

    # ══════════════════════════════════════════════════════════════
    # SHEET 6 — Payment Register
    # ══════════════════════════════════════════════════════════════
    ws_pr = wb.create_sheet("Payment Register")
    inv_map = {i.invoice_id: i for i in invoices}
    for col, h in enumerate(["Payment Date","Invoice No","Client Name","GSTIN","Mode","Reference","Amount","Notes"], 1):
        ws_pr.cell(row=1, column=col, value=h)
    for r, p in enumerate(payments, 2):
        c = clients.get(p.client_id); inv = inv_map.get(p.invoice_id)
        ws_pr.cell(row=r,column=1,value=p.payment_date)
        ws_pr.cell(row=r,column=2,value=inv.invoice_number if inv else 'N/A')
        ws_pr.cell(row=r,column=3,value=c.name if c else 'Unknown')
        ws_pr.cell(row=r,column=4,value=c.gstin if c else '')
        ws_pr.cell(row=r,column=5,value=p.payment_mode); ws_pr.cell(row=r,column=6,value=p.reference_number or '')
        ws_pr.cell(row=r,column=7,value=p.amount); ws_pr.cell(row=r,column=8,value=p.notes or '')
        if r % 2 == 0: alt_row(ws_pr, r)
    add_totals_row(ws_pr, {7}); hdr_style(ws_pr, color=C_GREEN); auto_cols(ws_pr)

    # ══════════════════════════════════════════════════════════════
    # SHEET 7 — Outstanding Register
    # ══════════════════════════════════════════════════════════════
    ws_or = wb.create_sheet("Outstanding Register")
    for col, h in enumerate(["Invoice No","Invoice Date","Due Date","Client","GSTIN","Invoice Value","Paid","Outstanding","Overdue Days","Status"], 1):
        ws_or.cell(row=1, column=col, value=h)
    for r, inv in enumerate(outstanding_invoices, 2):
        c = clients.get(inv.client_id)
        try: due = datetime.strptime(inv.due_date or '', '%Y-%m-%d').date(); overdue = max(0,(today_dt-due).days)
        except: overdue = 0
        ws_or.cell(row=r,column=1,value=inv.invoice_number); ws_or.cell(row=r,column=2,value=inv.invoice_date)
        ws_or.cell(row=r,column=3,value=inv.due_date)
        ws_or.cell(row=r,column=4,value=c.name if c else inv.client_name or 'Unknown')
        ws_or.cell(row=r,column=5,value=c.gstin if c else inv.client_gstin or '')
        ws_or.cell(row=r,column=6,value=inv.total or 0); ws_or.cell(row=r,column=7,value=inv.advance_paid or 0)
        ws_or.cell(row=r,column=8,value=inv.outstanding or 0); ws_or.cell(row=r,column=9,value=overdue)
        ws_or.cell(row=r,column=10,value="Overdue" if overdue>0 else "Not Due")
        if r % 2 == 0: alt_row(ws_or, r)
    add_totals_row(ws_or, {6,7,8,9}); hdr_style(ws_or, color=C_AMBER); auto_cols(ws_or)

    # ══════════════════════════════════════════════════════════════
    # SHEET 8 — Client Register
    # ══════════════════════════════════════════════════════════════
    ws_cr = wb.create_sheet("Client Register")
    for col, h in enumerate(["Client Name","GSTIN","PAN","Phone","Email","State","Type","Outstanding"], 1):
        ws_cr.cell(row=1, column=col, value=h)
    for r, cl in enumerate(clients_list, 2):
        cl_outstanding = sum(i.outstanding or 0 for i in outstanding_invoices if i.client_id == cl.client_id)
        ws_cr.cell(row=r,column=1,value=cl.name); ws_cr.cell(row=r,column=2,value=cl.gstin or '')
        ws_cr.cell(row=r,column=3,value=getattr(cl,'pan','') or ''); ws_cr.cell(row=r,column=4,value=cl.phone or '')
        ws_cr.cell(row=r,column=5,value=cl.email or ''); ws_cr.cell(row=r,column=6,value=cl.state or '')
        ws_cr.cell(row=r,column=7,value=cl.type or 'B2B'); ws_cr.cell(row=r,column=8,value=cl_outstanding)
        if r % 2 == 0: alt_row(ws_cr, r)
    add_totals_row(ws_cr, {8}); hdr_style(ws_cr, color=C_TEAL); auto_cols(ws_cr)

    # ══════════════════════════════════════════════════════════════
    # SHEET 9 — Advance Register
    # ══════════════════════════════════════════════════════════════
    ws_ar = wb.create_sheet("Advance Register")
    for col, h in enumerate(["Date","Client","Amount","Mode","Status","Reference"], 1):
        ws_ar.cell(row=1, column=col, value=h)
    for r, adv in enumerate(advances, 2):
        c = clients.get(adv.client_id)
        ws_ar.cell(row=r,column=1,value=adv.payment_date)
        ws_ar.cell(row=r,column=2,value=c.name if c else 'Unknown')
        ws_ar.cell(row=r,column=3,value=adv.amount or 0)
        ws_ar.cell(row=r,column=4,value=adv.payment_mode or '')
        ws_ar.cell(row=r,column=5,value=adv.status or '')
        ws_ar.cell(row=r,column=6,value=adv.reference_number or '')
        if r % 2 == 0: alt_row(ws_ar, r)
    add_totals_row(ws_ar, {3}); hdr_style(ws_ar, color=C_INDIGO); auto_cols(ws_ar)

    # ══════════════════════════════════════════════════════════════
    # SHEET 10 — CN Summary
    # ══════════════════════════════════════════════════════════════
    ws_cns = wb.create_sheet("CN Summary")
    for col, h in enumerate(["CN No","Date","Client","GSTIN","State","Subtotal","CGST","SGST","IGST","Total","Status"], 1):
        ws_cns.cell(row=1, column=col, value=h)
    for r, cn in enumerate(credit_notes, 2):
        c = clients.get(cn.client_id)
        ws_cns.cell(row=r,column=1,value=cn.credit_note_number); ws_cns.cell(row=r,column=2,value=cn.credit_date)
        ws_cns.cell(row=r,column=3,value=c.name if c else 'Unknown')
        ws_cns.cell(row=r,column=4,value=c.gstin if c else '')
        ws_cns.cell(row=r,column=5,value=cn.client_state or '')
        ws_cns.cell(row=r,column=6,value=cn.subtotal or 0); ws_cns.cell(row=r,column=7,value=cn.cgst or 0)
        ws_cns.cell(row=r,column=8,value=cn.sgst or 0); ws_cns.cell(row=r,column=9,value=cn.igst or 0)
        ws_cns.cell(row=r,column=10,value=cn.total or 0); ws_cns.cell(row=r,column=11,value='issued')
        if r % 2 == 0: alt_row(ws_cns, r)
    add_totals_row(ws_cns, {6,7,8,9,10}); hdr_style(ws_cns, color=C_RED); auto_cols(ws_cns)

    # ══════════════════════════════════════════════════════════════
    # SHEET 11 — CN Item Wise
    # ══════════════════════════════════════════════════════════════
    ws_cni = wb.create_sheet("CN Item Wise")
    write_items_sheet(ws_cni, credit_notes, color=C_RED)

    # ══════════════════════════════════════════════════════════════
    # SHEET 12 — Estimate Summary
    # ══════════════════════════════════════════════════════════════
    ws_es = wb.create_sheet("Estimate Summary")
    for col, h in enumerate(["Estimate No","Date","Client","GSTIN","State","Subtotal","CGST","SGST","IGST","Total","Status"], 1):
        ws_es.cell(row=1, column=col, value=h)
    for r, est in enumerate(estimates, 2):
        c = clients.get(est.client_id)
        ws_es.cell(row=r,column=1,value=est.estimate_number); ws_es.cell(row=r,column=2,value=est.estimate_date)
        ws_es.cell(row=r,column=3,value=c.name if c else est.client_name or 'Unknown')
        ws_es.cell(row=r,column=4,value=c.gstin if c else est.client_gstin or '')
        ws_es.cell(row=r,column=5,value=est.client_state or '')
        ws_es.cell(row=r,column=6,value=est.subtotal or 0); ws_es.cell(row=r,column=7,value=est.cgst or 0)
        ws_es.cell(row=r,column=8,value=est.sgst or 0); ws_es.cell(row=r,column=9,value=est.igst or 0)
        ws_es.cell(row=r,column=10,value=est.total or 0); ws_es.cell(row=r,column=11,value=est.status or '')
        if r % 2 == 0: alt_row(ws_es, r)
    add_totals_row(ws_es, {6,7,8,9,10}); hdr_style(ws_es, color=C_TEAL); auto_cols(ws_es)

    # ══════════════════════════════════════════════════════════════
    # SHEET 13 — Estimate Item Wise
    # ══════════════════════════════════════════════════════════════
    ws_ei = wb.create_sheet("Estimate Item Wise")
    write_items_sheet(ws_ei, estimates, color=C_TEAL)

    output = BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              headers={"Content-Disposition": f"attachment; filename=Complete_MIS_Report_{start_date}_{end_date}.xlsx"})


@router.get("/account-register")
async def get_account_register(start_date: str = Query(...), end_date: str = Query(...),
                                user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    """
    Account Register: one row per invoice, with each item in its own set of columns.
    Columns: Invoice No, Date, Client, GSTIN, State, Diary No, Ref No,
             [Item N Desc, Item N Notes, Item N Amount, Item N GST%] x N items,
             CGST, SGST, IGST, Total, Status
    The sheet auto-expands column count based on max items in any invoice.
    """
    invoices = (await db.execute(select(Invoice).where(
        Invoice.tenant_id == user["tenant_id"],
        Invoice.invoice_date >= start_date,
        Invoice.invoice_date <= end_date
    ))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(
        select(Client).where(Client.tenant_id == user["tenant_id"])
    )).scalars().all()}

    # Determine max items across all invoices
    max_items = max((len(inv.items or []) for inv in invoices), default=1)
    max_items = max(max_items, 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Account Register"

    # Build header row
    base_headers = ["Invoice No", "Date", "Client", "GSTIN", "State", "Diary No", "Ref No"]
    item_headers = []
    for n in range(1, max_items + 1):
        item_headers += [
            f"Item {n} Description",
            f"Item {n} Notes",
            f"Item {n} Amount",
            f"Item {n} GST Rate",
        ]
    tail_headers = ["CGST", "SGST", "IGST", "Total", "Status"]
    headers = base_headers + item_headers + tail_headers

    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    for row_idx, inv in enumerate(invoices, 2):
        c = clients.get(inv.client_id)
        items = inv.items or []

        # Base columns
        ws.cell(row=row_idx, column=1, value=inv.invoice_number)
        ws.cell(row=row_idx, column=2, value=inv.invoice_date)
        ws.cell(row=row_idx, column=3, value=c.name if c else inv.client_name or "Unknown")
        ws.cell(row=row_idx, column=4, value=c.gstin if c else inv.client_gstin or "")
        ws.cell(row=row_idx, column=5, value=inv.client_state or "")
        ws.cell(row=row_idx, column=6, value=inv.diary_no or "")
        ws.cell(row=row_idx, column=7, value=inv.ref_no or "")

        # Item columns (pad with blanks if fewer items than max)
        col_offset = len(base_headers) + 1
        for item_n in range(max_items):
            item = items[item_n] if item_n < len(items) else {}
            base_col = col_offset + item_n * 4
            ws.cell(row=row_idx, column=base_col,     value=item.get("description", ""))
            ws.cell(row=row_idx, column=base_col + 1, value=item.get("item_notes", "") or item.get("note", ""))
            ws.cell(row=row_idx, column=base_col + 2, value=item.get("amount", 0) or round((item.get("quantity", 0) or item.get("qty", 0)) * (item.get("rate", 0) or 0), 2))
            ws.cell(row=row_idx, column=base_col + 3, value=item.get("gst_rate", 0))

        # Tail columns
        tail_col = col_offset + max_items * 4
        ws.cell(row=row_idx, column=tail_col,     value=inv.cgst or 0)
        ws.cell(row=row_idx, column=tail_col + 1, value=inv.sgst or 0)
        ws.cell(row=row_idx, column=tail_col + 2, value=inv.igst or 0)
        ws.cell(row=row_idx, column=tail_col + 3, value=inv.total or 0)
        ws.cell(row=row_idx, column=tail_col + 4, value=inv.status or "")

    style_header(ws)
    auto_cols(ws)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Account_Register_{start_date}_{end_date}.xlsx"}
    )


@router.get("/credit-note-register")
async def get_credit_note_register(start_date: str = Query(...), end_date: str = Query(...),
                                    user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    """
    Credit Note Register: one row per credit note, with each item in its own set of columns.
    Same format as Account Register but for credit notes.
    """
    credit_notes = (await db.execute(select(CreditNote).where(
        CreditNote.tenant_id == user["tenant_id"],
        CreditNote.credit_date >= start_date,
        CreditNote.credit_date <= end_date
    ))).scalars().all()
    clients = {c.client_id: c for c in (await db.execute(
        select(Client).where(Client.tenant_id == user["tenant_id"])
    )).scalars().all()}

    max_items = max((len(cn.items or []) for cn in credit_notes), default=1)
    max_items = max(max_items, 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Credit Note Register"

    base_headers = ["Credit Note No", "Date", "Client", "GSTIN", "State", "Diary No", "Ref No", "Reason"]
    item_headers = []
    for n in range(1, max_items + 1):
        item_headers += [f"Item {n} Description", f"Item {n} Notes", f"Item {n} Amount", f"Item {n} GST Rate"]
    tail_headers = ["CGST", "SGST", "IGST", "Total"]
    headers = base_headers + item_headers + tail_headers

    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    for row_idx, cn in enumerate(credit_notes, 2):
        c = clients.get(cn.client_id)
        items = cn.items or []

        ws.cell(row=row_idx, column=1, value=cn.credit_note_number)
        ws.cell(row=row_idx, column=2, value=cn.credit_date)
        ws.cell(row=row_idx, column=3, value=c.name if c else "Unknown")
        ws.cell(row=row_idx, column=4, value=c.gstin if c else "")
        ws.cell(row=row_idx, column=5, value=cn.client_state or "")
        ws.cell(row=row_idx, column=6, value=cn.diary_no or "")
        ws.cell(row=row_idx, column=7, value=cn.ref_no or "")
        ws.cell(row=row_idx, column=8, value=cn.reason or "")

        col_offset = len(base_headers) + 1
        for item_n in range(max_items):
            item = items[item_n] if item_n < len(items) else {}
            base = col_offset + item_n * 4
            ws.cell(row=row_idx, column=base,     value=item.get("description", ""))
            ws.cell(row=row_idx, column=base + 1, value=item.get("item_notes", ""))
            ws.cell(row=row_idx, column=base + 2, value=item.get("amount", 0))
            ws.cell(row=row_idx, column=base + 3, value=item.get("gst_rate", 0))

        tail_col = col_offset + max_items * 4
        ws.cell(row=row_idx, column=tail_col,     value=cn.cgst or 0)
        ws.cell(row=row_idx, column=tail_col + 1, value=cn.sgst or 0)
        ws.cell(row=row_idx, column=tail_col + 2, value=cn.igst or 0)
        ws.cell(row=row_idx, column=tail_col + 3, value=cn.total or 0)

        if row_idx % 2 == 0:
            alt_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = alt_fill

    style_header(ws)
    auto_cols(ws)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Credit_Note_Register_{start_date}_{end_date}.xlsx"}
    )
