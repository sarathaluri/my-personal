"""
APCRDA PROJECT 2: Contractor Payment Aging & Liability Analysis
===============================================================
Purpose : Track all contractor bills — pending, cleared, overdue.
          Flags MSME interest exposure, aging buckets, and work-stall risk.

Run     : python project2_contractor_aging.py
Output  : Contractor_Payment_Aging.xlsx

OFFLINE — replace generate_contractor_data() with real register values.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from datetime import datetime, timedelta
import random
import warnings
warnings.filterwarnings("ignore")

random.seed(99)
np.random.seed(99)

TODAY = datetime.today()

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
MSME_INTEREST_RATE  = 0.18        # 18% p.a. per MSMED Act
STANDARD_PAY_DAYS   = 30          # days within which bill should be cleared
WORK_CATEGORIES     = [
    "Roads & Bridges", "Water Supply", "Sewerage & Drainage",
    "Buildings & Structures", "Electrical & Mechanical",
    "Landscaping & Green Infra", "ICT Infrastructure",
]
CONTRACTORS = [
    ("Srinivasa Infra Pvt Ltd",       "MSME"),
    ("Lanco Amaravati Projects",       "Large"),
    ("KNR Constructions Ltd",          "Large"),
    ("Patel Engineering Ltd",          "Large"),
    ("Gayatri Projects Ltd",           "Large"),
    ("Sri Balaji Contractors",         "MSME"),
    ("Nagarjuna Construction Co.",     "Large"),
    ("Rithwik Projects Pvt Ltd",       "MSME"),
    ("IL&FS Engineering",              "Large"),
    ("Venkateswara Infra Works",       "MSME"),
    ("Shapoorji Pallonji & Co.",       "Large"),
    ("Triveni Infracon Pvt Ltd",       "MSME"),
]
BILL_TYPES = ["Running Account (RA)", "Final Bill", "Secured Advance", "Mobilization Advance Recovery"]
PAYMENT_STATUS = ["Cleared", "Cleared", "Cleared", "Pending Treasury", "Pending Accounts", "Under Scrutiny"]

# ─────────────────────────────────────────────────────────────────────────────
# STYLING HELPERS (same palette as Project 1 for consistency)
# ─────────────────────────────────────────────────────────────────────────────
C_NAVY  = "1B2A4A"; C_GOLD  = "C9A84C"; C_RED   = "C0392B"
C_AMBER = "E67E22"; C_GREEN = "1E8449"; C_WHITE = "FFFFFF"
C_ALT   = "EAF0FB"; C_BORDER= "BDC3C7"

def hfill(h): return PatternFill("solid", fgColor=h)
def hfont(h="000000", bold=False, sz=10): return Font(color=h, bold=bold, size=sz, name="Calibri")
def bdr():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)
def ctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)

def hdr_row(ws, row, cols, col_start=1):
    for i, label in enumerate(cols):
        c = ws.cell(row=row, column=col_start+i, value=label)
        c.fill = hfill(C_NAVY); c.font = hfont(C_WHITE, True, 10)
        c.alignment = ctr(); c.border = bdr()

def style_row(ws, row, ncols, col_start=1, alt=False):
    bg = C_ALT if alt else C_WHITE
    for col in range(col_start, col_start+ncols):
        c = ws.cell(row=row, column=col)
        c.fill = hfill(bg); c.font = hfont(sz=9)
        c.alignment = lft(); c.border = bdr()

def risk_fill(cell, bucket):
    mapping = {
        "0-30d"  : ("D5F5E3", C_GREEN),
        "31-60d" : ("FDEBD0", C_AMBER),
        "61-90d" : ("FAD7A0", "784212"),
        "90d+"   : ("FADBD8", C_RED),
    }
    if bucket in mapping:
        bg, fg = mapping[bucket]
        cell.fill = hfill(bg)
        cell.font = Font(color=fg, bold=(bucket == "90d+"), size=9, name="Calibri")


# ─────────────────────────────────────────────────────────────────────────────
# DATA GENERATOR
# ─────────────────────────────────────────────────────────────────────────────
def rand_date(days_ago_min, days_ago_max):
    return TODAY - timedelta(days=random.randint(days_ago_min, days_ago_max))

def aging_bucket(days: int) -> str:
    if days <= 30:  return "0-30d"
    if days <= 60:  return "31-60d"
    if days <= 90:  return "61-90d"
    return "90d+"

def msme_interest(amount_cr: float, days_pending: int, is_msme: bool) -> float:
    if not is_msme or days_pending <= STANDARD_PAY_DAYS:
        return 0.0
    overdue_days = days_pending - STANDARD_PAY_DAYS
    return round(amount_cr * MSME_INTEREST_RATE * overdue_days / 365, 4)

def generate_contractor_data():
    bills = []
    work_orders = []

    # Work orders
    wo_pool = []
    for wo_idx in range(30):
        contractor = random.choice(CONTRACTORS)
        category   = random.choice(WORK_CATEGORIES)
        wo_value   = round(random.uniform(1.5, 85.0), 2)
        wo_date    = rand_date(800, 200)
        wo_pool.append({
            "WO_Number"        : f"APCRDA/WO/{2023 + wo_idx%2}/{str(wo_idx+1).zfill(4)}",
            "Contractor_Name"  : contractor[0],
            "Contractor_Type"  : contractor[1],
            "Work_Category"    : category,
            "WO_Value_Cr"      : wo_value,
            "WO_Date"          : wo_date,
            "Completion_Target": wo_date + timedelta(days=random.randint(180, 730)),
            "Physical_Progress": f"{random.randint(15, 95)}%",
        })
    work_orders = pd.DataFrame(wo_pool)

    # Bills against work orders
    bill_id = 1
    for _, wo in work_orders.iterrows():
        num_bills = random.randint(1, 5)
        for b in range(num_bills):
            bill_date   = wo["WO_Date"] + timedelta(days=(b+1)*random.randint(30, 90))
            if bill_date > TODAY:
                continue
            gross       = round(wo["WO_Value_Cr"] * random.uniform(0.05, 0.25), 4)
            tds         = round(gross * 0.02, 4)
            gst_tds     = round(gross * 0.02, 4)
            labour_cess = round(gross * 0.01, 4)
            security_dep= round(gross * 0.05, 4)
            adv_rec     = round(gross * random.uniform(0, 0.08), 4)
            net_pay     = round(gross - tds - gst_tds - labour_cess - security_dep - adv_rec, 4)
            status      = random.choice(PAYMENT_STATUS)
            bill_type   = random.choice(BILL_TYPES)
            sub_date    = bill_date
            clr_date    = sub_date + timedelta(days=random.randint(5, 120)) if "Cleared" in status else None
            days_pending= (TODAY - sub_date).days if status != "Cleared" else 0
            if clr_date and clr_date > TODAY:
                clr_date = None; status = "Pending Treasury"
                days_pending = (TODAY - sub_date).days

            interest = msme_interest(net_pay, days_pending, wo["Contractor_Type"] == "MSME")

            bills.append({
                "Bill_ID"           : f"BILL/{str(bill_id).zfill(5)}",
                "WO_Number"         : wo["WO_Number"],
                "Contractor_Name"   : wo["Contractor_Name"],
                "Contractor_Type"   : wo["Contractor_Type"],
                "Work_Category"     : wo["Work_Category"],
                "Bill_Type"         : bill_type,
                "Bill_Date"         : bill_date,
                "Gross_Amount_Cr"   : gross,
                "TDS_Cr"            : tds,
                "GST_TDS_Cr"        : gst_tds,
                "Labour_Cess_Cr"    : labour_cess,
                "Security_Deposit_Cr": security_dep,
                "Advance_Recovery_Cr": adv_rec,
                "Net_Payable_Cr"    : net_pay,
                "Payment_Status"    : status,
                "Submission_Date"   : sub_date,
                "Cleared_Date"      : clr_date,
                "Days_Pending"      : days_pending,
                "Aging_Bucket"      : aging_bucket(days_pending) if status != "Cleared" else "Cleared",
                "MSME_Interest_Cr"  : interest,
                "Stall_Risk"        : "YES" if days_pending > 60 and status != "Cleared" else "NO",
            })
            bill_id += 1

    bills_df = pd.DataFrame(bills)
    return work_orders, bills_df


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def compute_aging_kpis(bills: pd.DataFrame) -> dict:
    pending = bills[bills["Payment_Status"] != "Cleared"]
    cleared = bills[bills["Payment_Status"] == "Cleared"]

    total_gross    = bills["Gross_Amount_Cr"].sum()
    total_pending  = pending["Net_Payable_Cr"].sum()
    total_cleared  = cleared["Net_Payable_Cr"].sum()
    msme_exposure  = pending[pending["Contractor_Type"] == "MSME"]["MSME_Interest_Cr"].sum()
    stall_risk_amt = pending[pending["Stall_Risk"] == "YES"]["Net_Payable_Cr"].sum()

    bucket_summary = (
        pending.groupby("Aging_Bucket")["Net_Payable_Cr"]
        .agg(["count", "sum"]).reset_index()
        .rename(columns={"count": "Bills", "sum": "Amount_Cr"})
    )

    contractor_aging = (
        pending.groupby(["Contractor_Name", "Contractor_Type"])
        .agg(
            Pending_Bills=("Bill_ID", "count"),
            Pending_Amount_Cr=("Net_Payable_Cr", "sum"),
            Max_Days_Pending=("Days_Pending", "max"),
            MSME_Interest_Cr=("MSME_Interest_Cr", "sum"),
        ).reset_index().sort_values("Pending_Amount_Cr", ascending=False)
    )

    return {
        "total_gross"       : round(total_gross, 2),
        "total_pending"     : round(total_pending, 2),
        "total_cleared"     : round(total_cleared, 2),
        "pending_bills"     : len(pending),
        "cleared_bills"     : len(cleared),
        "msme_interest"     : round(msme_exposure, 4),
        "stall_risk_amount" : round(stall_risk_amt, 2),
        "stall_risk_count"  : len(pending[pending["Stall_Risk"] == "YES"]),
        "bucket_summary"    : bucket_summary,
        "contractor_aging"  : contractor_aging,
        "avg_days_pending"  : round(pending["Days_Pending"].mean(), 1),
        "max_days_pending"  : int(pending["Days_Pending"].max()),
    }


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(work_orders, bills, kpis):
    wb = Workbook()

    # ── SHEET 1: KPI ──────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "📊 KPI Summary"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 22

    ws.merge_cells("B2:C2")
    t = ws["B2"]
    t.value = "CONTRACTOR PAYMENT AGING — KPI DASHBOARD"
    t.fill = hfill(C_NAVY); t.font = Font(color=C_GOLD, bold=True, size=14, name="Calibri")
    t.alignment = ctr(); ws.row_dimensions[2].height = 30

    ws.merge_cells("B3:C3")
    ws["B3"].value = f"As of {TODAY.strftime('%d %B %Y')}  |  All figures in ₹ Crore"
    ws["B3"].font = hfont(C_NAVY, sz=10); ws["B3"].alignment = ctr()

    card_data = [
        ("Total Bills Raised",        f"₹ {kpis['total_gross']:,.2f} Cr",      C_NAVY,  f"{len(bills)} bills total"),
        ("Total Pending",             f"₹ {kpis['total_pending']:,.2f} Cr",    C_RED,   f"{kpis['pending_bills']} bills unpaid"),
        ("Total Cleared",             f"₹ {kpis['total_cleared']:,.2f} Cr",    C_GREEN, f"{kpis['cleared_bills']} bills cleared"),
        ("Avg Days Pending",          f"{kpis['avg_days_pending']} days",       C_AMBER, f"Max: {kpis['max_days_pending']} days"),
        ("MSME Interest Exposure",    f"₹ {kpis['msme_interest']:,.4f} Cr",    C_RED,   "Accrued under MSMED Act Sec 16"),
        ("Work-Stall Risk Bills",     f"{kpis['stall_risk_count']} bills",      C_RED,   f"₹ {kpis['stall_risk_amount']:,.2f} Cr — >60 days unpaid"),
    ]

    r = 5
    for label, value, color, note in card_data:
        for row in [r, r+1, r+2]:
            ws.row_dimensions[row].height = 14 if row != r+1 else 26
        ws.merge_cells(f"B{r}:C{r}")
        ws.merge_cells(f"B{r+1}:C{r+1}")
        ws.merge_cells(f"B{r+2}:C{r+2}")
        for row in [r, r+1, r+2]:
            ws[f"B{row}"].fill = hfill(color)
        ws[f"B{r}"].value = label.upper()
        ws[f"B{r}"].font = Font(color=C_WHITE, size=8, name="Calibri")
        ws[f"B{r}"].alignment = ctr()
        ws[f"B{r+1}"].value = value
        ws[f"B{r+1}"].font = Font(color=C_WHITE, bold=True, size=16, name="Calibri")
        ws[f"B{r+1}"].alignment = ctr()
        ws[f"B{r+2}"].value = note
        ws[f"B{r+2}"].font = Font(color="DDDDDD", size=8, name="Calibri")
        ws[f"B{r+2}"].alignment = ctr()
        r += 4

    # Aging bucket table
    r += 1
    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"].value = "PENDING BILLS BY AGING BUCKET"
    ws[f"B{r}"].fill = hfill(C_NAVY)
    ws[f"B{r}"].font = hfont(C_GOLD, True, 11); ws[f"B{r}"].alignment = ctr()
    ws.row_dimensions[r].height = 22; r += 1

    hdr_row(ws, r, ["Aging Bucket", "No. of Bills", "Amount (₹Cr)", "% of Total Pending"], col_start=2)
    r += 1
    bucket_order = ["0-30d", "31-60d", "61-90d", "90d+"]
    for bkt in bucket_order:
        brow = kpis["bucket_summary"][kpis["bucket_summary"]["Aging_Bucket"] == bkt]
        if brow.empty:
            continue
        ws.row_dimensions[r].height = 18
        style_row(ws, r, 4, col_start=2)
        count  = int(brow["Bills"].values[0])
        amount = round(float(brow["Amount_Cr"].values[0]), 2)
        pct    = round(amount / kpis["total_pending"] * 100, 1) if kpis["total_pending"] else 0
        for c_idx, val in enumerate([bkt, count, amount, f"{pct}%"], start=2):
            cell = ws.cell(row=r, column=c_idx, value=val)
            risk_fill(cell, bkt)
            if c_idx == 4:
                cell.number_format = '#,##0.00'
        r += 1

    # ── SHEET 2: BILL REGISTER ────────────────────────────────────────────────
    ws2 = wb.create_sheet("📋 Bill Register")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "A3"

    cols2 = [
        "Bill ID", "WO Number", "Contractor", "Type", "Work Category",
        "Bill Type", "Bill Date", "Gross (₹Cr)", "TDS (₹Cr)", "GST TDS (₹Cr)",
        "Labour Cess", "Security Dep", "Adv Recovery", "Net Payable (₹Cr)",
        "Status", "Submitted", "Cleared", "Days Pending", "Aging Bucket",
        "MSME Interest (₹Cr)", "Stall Risk",
    ]
    widths2 = [14,22,28,10,20,22,14,14,12,12,12,12,14,16,20,14,14,14,14,18,12]
    for i,w in enumerate(widths2,1): ws2.column_dimensions[get_column_letter(i)].width = w

    ws2.row_dimensions[1].height = 28
    ws2.merge_cells(f"A1:{get_column_letter(len(cols2))}1")
    t2 = ws2["A1"]
    t2.value = "CONTRACTOR BILL REGISTER — ALL BILLS"
    t2.fill = hfill(C_NAVY); t2.font = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t2.alignment = ctr()

    ws2.row_dimensions[2].height = 30
    hdr_row(ws2, 2, cols2)

    field_map2 = [
        "Bill_ID","WO_Number","Contractor_Name","Contractor_Type","Work_Category",
        "Bill_Type","Bill_Date","Gross_Amount_Cr","TDS_Cr","GST_TDS_Cr",
        "Labour_Cess_Cr","Security_Deposit_Cr","Advance_Recovery_Cr","Net_Payable_Cr",
        "Payment_Status","Submission_Date","Cleared_Date","Days_Pending","Aging_Bucket",
        "MSME_Interest_Cr","Stall_Risk",
    ]
    money_cols = {"Gross_Amount_Cr","TDS_Cr","GST_TDS_Cr","Labour_Cess_Cr",
                  "Security_Deposit_Cr","Advance_Recovery_Cr","Net_Payable_Cr","MSME_Interest_Cr"}

    for r_idx, (_, row) in enumerate(bills.iterrows(), start=3):
        ws2.row_dimensions[r_idx].height = 16
        style_row(ws2, r_idx, len(cols2), alt=r_idx%2==0)
        for c_idx, fld in enumerate(field_map2, start=1):
            val = row[fld]
            cell = ws2.cell(row=r_idx, column=c_idx)
            if pd.isna(val) or val is None:
                cell.value = "—"
            elif isinstance(val, datetime):
                cell.value = val.strftime("%d-%b-%Y")
            elif fld in money_cols:
                cell.value = round(float(val), 4)
                cell.number_format = '#,##0.0000'
            else:
                cell.value = val

            if fld == "Aging_Bucket" and val != "Cleared":
                risk_fill(cell, str(val))
            if fld == "Stall_Risk" and val == "YES":
                cell.fill = hfill("FADBD8")
                cell.font = Font(color=C_RED, bold=True, size=9, name="Calibri")
            if fld == "Payment_Status":
                if val == "Cleared":
                    cell.fill = hfill("D5F5E3")
                    cell.font = Font(color=C_GREEN, size=9, name="Calibri")
                elif "Pending" in str(val):
                    cell.fill = hfill("FDEBD0")
                    cell.font = Font(color=C_AMBER, size=9, name="Calibri")
                elif "Scrutiny" in str(val):
                    cell.fill = hfill("FADBD8")
                    cell.font = Font(color=C_RED, size=9, name="Calibri")

    # ── SHEET 3: CONTRACTOR AGING SUMMARY ─────────────────────────────────────
    ws3 = wb.create_sheet("🏗️ Contractor Summary")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A3"

    cols3 = ["Contractor Name", "Type", "Pending Bills", "Pending Amount (₹Cr)",
             "Max Days Pending", "MSME Interest (₹Cr)", "Risk Level"]
    widths3 = [35, 10, 16, 22, 18, 22, 14]
    for i,w in enumerate(widths3,1): ws3.column_dimensions[get_column_letter(i)].width = w

    ws3.row_dimensions[1].height = 28
    ws3.merge_cells(f"A1:{get_column_letter(len(cols3))}1")
    t3 = ws3["A1"]
    t3.value = "CONTRACTOR-WISE AGING SUMMARY"
    t3.fill = hfill(C_NAVY); t3.font = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t3.alignment = ctr()

    ws3.row_dimensions[2].height = 30
    hdr_row(ws3, 2, cols3)

    ca = kpis["contractor_aging"]
    for r_idx, (_, row) in enumerate(ca.iterrows(), start=3):
        ws3.row_dimensions[r_idx].height = 18
        style_row(ws3, r_idx, len(cols3), alt=r_idx%2==0)
        risk = "🔴 HIGH" if row["Max_Days_Pending"] > 90 else (
               "🟡 MEDIUM" if row["Max_Days_Pending"] > 60 else "🟢 LOW")
        vals = [
            row["Contractor_Name"], row["Contractor_Type"],
            int(row["Pending_Bills"]), round(row["Pending_Amount_Cr"], 2),
            int(row["Max_Days_Pending"]), round(row["MSME_Interest_Cr"], 4), risk,
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            if c_idx in (4, 6): cell.number_format = '#,##0.0000'
            if c_idx == 7:
                if "HIGH" in str(val):
                    cell.fill = hfill("FADBD8"); cell.font = Font(color=C_RED, bold=True, size=9, name="Calibri")
                elif "MEDIUM" in str(val):
                    cell.fill = hfill("FDEBD0"); cell.font = Font(color=C_AMBER, size=9, name="Calibri")
                else:
                    cell.fill = hfill("D5F5E3"); cell.font = Font(color=C_GREEN, size=9, name="Calibri")
            if c_idx == 2 and val == "MSME":
                cell.fill = hfill("D6EAF8"); cell.font = Font(color="1A5276", bold=True, size=9, name="Calibri")

    # ── SHEET 4: STALL RISK ALERTS ────────────────────────────────────────────
    ws4 = wb.create_sheet("🚨 Stall Risk Alerts")
    ws4.sheet_view.showGridLines = False

    stall = bills[(bills["Stall_Risk"] == "YES") & (bills["Payment_Status"] != "Cleared")].copy()
    stall = stall.sort_values("Days_Pending", ascending=False)

    ws4.row_dimensions[1].height = 30
    ws4.merge_cells(f"A1:{get_column_letter(len(cols2))}1")
    t4 = ws4["A1"]
    t4.value = f"⚠  WORK-STALL RISK — {len(stall)} BILLS >60 DAYS UNPAID  |  ₹ {stall['Net_Payable_Cr'].sum():,.2f} Cr outstanding"
    t4.fill = hfill(C_RED); t4.font = Font(color=C_WHITE, bold=True, size=13, name="Calibri")
    t4.alignment = ctr()

    cols4 = ["Bill ID", "Contractor", "Work Category", "Net Payable (₹Cr)",
             "Submitted", "Days Pending", "Aging Bucket", "MSME Interest (₹Cr)", "Status"]
    widths4 = [14, 30, 22, 18, 14, 14, 14, 20, 22]
    for i,w in enumerate(widths4,1): ws4.column_dimensions[get_column_letter(i)].width = w

    ws4.row_dimensions[2].height = 30
    hdr_row(ws4, 2, cols4)

    field_map4 = ["Bill_ID","Contractor_Name","Work_Category","Net_Payable_Cr",
                  "Submission_Date","Days_Pending","Aging_Bucket","MSME_Interest_Cr","Payment_Status"]
    for r_idx, (_, row) in enumerate(stall.iterrows(), start=3):
        ws4.row_dimensions[r_idx].height = 18
        for c_idx, fld in enumerate(field_map4, start=1):
            val = row[fld]
            cell = ws4.cell(row=r_idx, column=c_idx)
            cell.fill = hfill("FADBD8"); cell.font = Font(color=C_RED, size=9, name="Calibri")
            cell.border = bdr(); cell.alignment = lft()
            if pd.isna(val) or val is None: cell.value = "—"
            elif isinstance(val, datetime): cell.value = val.strftime("%d-%b-%Y")
            elif fld in ("Net_Payable_Cr","MSME_Interest_Cr"):
                cell.value = round(float(val), 4); cell.number_format = '#,##0.0000'
            elif fld == "Days_Pending": cell.value = int(val)
            else: cell.value = val

    # ── SHEET 5: WORK ORDER REGISTER ──────────────────────────────────────────
    ws5 = wb.create_sheet("📁 Work Order Register")
    ws5.sheet_view.showGridLines = False
    ws5.freeze_panes = "A3"

    cols5 = ["WO Number", "Contractor", "Type", "Category",
             "WO Value (₹Cr)", "WO Date", "Target Completion", "Physical Progress"]
    widths5 = [26, 30, 10, 22, 16, 14, 18, 18]
    for i,w in enumerate(widths5,1): ws5.column_dimensions[get_column_letter(i)].width = w

    ws5.row_dimensions[1].height = 28
    ws5.merge_cells(f"A1:{get_column_letter(len(cols5))}1")
    t5 = ws5["A1"]
    t5.value = "WORK ORDER MASTER REGISTER"
    t5.fill = hfill(C_NAVY); t5.font = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t5.alignment = ctr()
    hdr_row(ws5, 2, cols5)

    field_map5 = ["WO_Number","Contractor_Name","Contractor_Type","Work_Category",
                  "WO_Value_Cr","WO_Date","Completion_Target","Physical_Progress"]
    for r_idx, (_, row) in enumerate(work_orders.iterrows(), start=3):
        ws5.row_dimensions[r_idx].height = 18
        style_row(ws5, r_idx, len(cols5), alt=r_idx%2==0)
        for c_idx, fld in enumerate(field_map5, start=1):
            val = row[fld]
            cell = ws5.cell(row=r_idx, column=c_idx)
            if isinstance(val, datetime): cell.value = val.strftime("%d-%b-%Y")
            elif fld == "WO_Value_Cr":
                cell.value = round(float(val), 2); cell.number_format = '#,##0.00'
            else: cell.value = val
            if fld == "Contractor_Type" and val == "MSME":
                cell.fill = hfill("D6EAF8"); cell.font = Font(color="1A5276", bold=True, size=9, name="Calibri")

    return wb


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  APCRDA PROJECT 2 — CONTRACTOR PAYMENT AGING")
    print("=" * 60)

    print("\n[1/4] Generating work order and bill data...")
    work_orders, bills = generate_contractor_data()

    print("[2/4] Computing aging KPIs...")
    kpis = compute_aging_kpis(bills)

    print("\n── KPI SNAPSHOT ─────────────────────────────────────────")
    print(f"  Total Bills Raised    : ₹ {kpis['total_gross']:>10,.2f} Cr")
    print(f"  Total Pending         : ₹ {kpis['total_pending']:>10,.2f} Cr  ({kpis['pending_bills']} bills)")
    print(f"  Total Cleared         : ₹ {kpis['total_cleared']:>10,.2f} Cr  ({kpis['cleared_bills']} bills)")
    print(f"  Avg Days Pending      : {kpis['avg_days_pending']} days (Max: {kpis['max_days_pending']})")
    print(f"  MSME Interest Risk    : ₹ {kpis['msme_interest']:,.4f} Cr")
    print(f"  Work-Stall Risk Bills : {kpis['stall_risk_count']} bills  (₹ {kpis['stall_risk_amount']:,.2f} Cr)")
    print("─" * 60)
    print("\n  AGING BREAKDOWN (Pending Only):")
    for _, row in kpis["bucket_summary"].iterrows():
        print(f"  {row['Aging_Bucket']:>8} : {int(row['Bills']):>4} bills | ₹ {row['Amount_Cr']:>8,.2f} Cr")
    print("─" * 60)

    print("\n[3/4] Building Excel workbook...")
    wb = build_excel(work_orders, bills, kpis)

    out = "/mnt/user-data/outputs/Contractor_Payment_Aging.xlsx"
    wb.save(out)
    print(f"[4/4] Saved → {out}")

    import os
    print(f"\n  File size : {os.path.getsize(out)/1024:.1f} KB")
    print("\n  SHEETS:")
    print("  📊 KPI Summary       → aging buckets + MSME interest exposure")
    print("  📋 Bill Register     → every bill with full deduction breakdown")
    print("  🏗️ Contractor Summary → per-contractor outstanding + risk level")
    print("  🚨 Stall Risk Alerts → only >60-day unpaid bills")
    print("  📁 Work Order Register → all WOs with progress status")
    print("════════════════════════════════════════════════════════════\n")


if __name__ == "__main__":
    main()
