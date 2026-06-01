"""
APCRDA PROJECT 1: UC Compliance Dashboard
==========================================
Author  : APCRDA Internship Toolkit
Purpose : Track Utilization Certificate compliance across all active
          centrally-sponsored and state schemes. Flags overdue UCs,
          reconciliation gaps, and next-tranche risk.

Run     : python project1_uc_dashboard.py
Output  : UC_Compliance_Dashboard.xlsx  (fully formatted, ready to present)

OFFLINE — no internet required. Uses synthetic but structurally realistic data.
Replace generate_sample_data() inputs with real figures when available.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from datetime import datetime, timedelta
import random
import warnings
warnings.filterwarnings("ignore")

random.seed(42)
np.random.seed(42)

TODAY = datetime.today()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1: SYNTHETIC DATA GENERATOR
# Replace this section's values with actuals from your registers.
# Structure must remain identical.
# ─────────────────────────────────────────────────────────────────────────────

SCHEMES = [
    # (Scheme Name, Funding Type, Total Sanction ₹Cr, GoI%, GoAP%, APCRDA%)
    ("AMRUT 2.0 - Water Supply",          "CSS",   320.00, 0.50, 0.35, 0.15),
    ("AMRUT 2.0 - Sewerage",              "CSS",   180.00, 0.50, 0.35, 0.15),
    ("Smart City Mission - Core Infra",   "CSS",   450.00, 0.50, 0.40, 0.10),
    ("Smart City Mission - ICT",          "CSS",   120.00, 0.50, 0.40, 0.10),
    ("PMAY-Urban - Housing",              "CSS",   210.00, 0.60, 0.30, 0.10),
    ("AP State Capital Roads Ph-I",       "State", 560.00, 0.00, 0.80, 0.20),
    ("AP State Capital Roads Ph-II",      "State", 340.00, 0.00, 0.80, 0.20),
    ("Amaravati Green Grid",              "State", 95.00,  0.00, 0.70, 0.30),
    ("World Bank - Urban Infra",          "EAP",   780.00, 0.00, 0.75, 0.25),
    ("ADB - Resilient Infra",             "EAP",   620.00, 0.00, 0.75, 0.25),
    ("NCRMP-II Coastal Protection",       "CSS",   145.00, 0.75, 0.20, 0.05),
    ("HRIDAY - Heritage Areas",           "CSS",   48.00,  0.60, 0.30, 0.10),
]

TRANCHE_LABELS = ["1st Tranche", "2nd Tranche", "3rd Tranche", "4th Tranche"]


def random_date(start_days_ago: int, end_days_ago: int) -> datetime:
    delta = random.randint(end_days_ago, start_days_ago)
    return TODAY - timedelta(days=delta)


def generate_sample_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Returns three DataFrames:
      1. scheme_master   — one row per scheme
      2. tranche_detail  — one row per tranche per scheme
      3. uc_register     — one row per UC submitted/pending
    """

    # ── SCHEME MASTER ────────────────────────────────────────────────────────
    scheme_rows = []
    for idx, (name, ftype, sanction, goi_pct, goap_pct, apcrda_pct) in enumerate(SCHEMES):
        utilization_pct  = round(random.uniform(0.25, 0.88), 4)
        expenditure      = round(sanction * utilization_pct, 2)
        uc_submitted     = round(expenditure * random.uniform(0.55, 0.95), 2)
        uc_pending       = round(expenditure - uc_submitted, 2)
        start_date       = random_date(900, 400)
        end_date         = start_date + timedelta(days=random.randint(730, 1460))
        num_tranches     = random.randint(2, 4)

        scheme_rows.append({
            "Scheme_ID"         : f"APCRDA-{2023 + idx % 2}-{str(idx+1).zfill(3)}",
            "Scheme_Name"       : name,
            "Funding_Type"      : ftype,        # CSS / State / EAP
            "Total_Sanction_Cr" : sanction,
            "GoI_Share_Pct"     : goi_pct,
            "GoAP_Share_Pct"    : goap_pct,
            "APCRDA_Share_Pct"  : apcrda_pct,
            "Total_Released_Cr" : round(sanction * random.uniform(0.40, 0.75), 2),
            "Total_Expended_Cr" : expenditure,
            "UC_Submitted_Cr"   : uc_submitted,
            "UC_Pending_Cr"     : uc_pending,
            "Scheme_Start"      : start_date,
            "Scheme_End"        : end_date,
            "Num_Tranches"      : num_tranches,
            "Nodal_Ministry"    : (
                "MoHUA" if ftype == "CSS" else
                "GoAP Finance" if ftype == "State" else
                "DEA / Finance Ministry"
            ),
        })

    sm = pd.DataFrame(scheme_rows)

    # ── TRANCHE DETAIL ───────────────────────────────────────────────────────
    tranche_rows = []
    for _, s in sm.iterrows():
        released_so_far = s["Total_Released_Cr"]
        for t_idx in range(s["Num_Tranches"]):
            t_amount   = round(released_so_far / s["Num_Tranches"] * random.uniform(0.85, 1.15), 2)
            t_date     = s["Scheme_Start"] + timedelta(days=t_idx * random.randint(90, 180))
            uc_due     = t_date + timedelta(days=180)          # standard 6-month UC cycle
            uc_filed   = t_date + timedelta(days=random.randint(30, 240))
            filed      = uc_filed <= TODAY
            overdue    = (not filed) and (TODAY > uc_due)
            days_overdue = max(0, (TODAY - uc_due).days) if overdue else 0

            tranche_rows.append({
                "Scheme_ID"          : s["Scheme_ID"],
                "Scheme_Name"        : s["Scheme_Name"],
                "Funding_Type"       : s["Funding_Type"],
                "Tranche"            : TRANCHE_LABELS[t_idx],
                "Release_Date"       : t_date,
                "Amount_Released_Cr" : t_amount,
                "UC_Due_Date"        : uc_due,
                "UC_Filed_Date"      : uc_filed if filed else pd.NaT,
                "UC_Status"          : "Filed" if filed else ("OVERDUE" if overdue else "Pending"),
                "Days_Overdue"       : days_overdue,
                "Next_Tranche_Risk"  : "HIGH" if overdue else ("MEDIUM" if not filed else "LOW"),
            })

    td = pd.DataFrame(tranche_rows)

    # ── UC REGISTER ──────────────────────────────────────────────────────────
    uc_rows = []
    for _, s in sm.iterrows():
        num_ucs = random.randint(1, 5)
        cumulative = 0
        for u in range(num_ucs):
            amount = round(s["UC_Submitted_Cr"] / num_ucs * random.uniform(0.8, 1.2), 2)
            cumulative += amount
            sub_date  = s["Scheme_Start"] + timedelta(days=(u + 1) * random.randint(60, 120))
            ack_date  = sub_date + timedelta(days=random.randint(5, 45))
            approved  = ack_date + timedelta(days=random.randint(10, 60))
            status    = random.choice(["Approved", "Approved", "Pending Approval", "Returned for Correction"])

            uc_rows.append({
                "UC_Number"          : f"UC/{s['Scheme_ID']}/{2023}/{str(u+1).zfill(2)}",
                "Scheme_ID"          : s["Scheme_ID"],
                "Scheme_Name"        : s["Scheme_Name"],
                "Funding_Type"       : s["Funding_Type"],
                "Nodal_Ministry"     : s["Nodal_Ministry"],
                "Period_From"        : sub_date - timedelta(days=90),
                "Period_To"          : sub_date,
                "Amount_Cr"          : amount,
                "Submission_Date"    : sub_date,
                "Acknowledgement_Dt" : ack_date if status != "Pending Approval" else pd.NaT,
                "Approval_Date"      : approved if status == "Approved" else pd.NaT,
                "UC_Status"          : status,
                "Remarks"            : (
                    "" if status == "Approved" else
                    "Awaiting ministry sign-off" if status == "Pending Approval" else
                    "MB mismatch flagged — resubmit with corrected annexure"
                ),
            })

    ur = pd.DataFrame(uc_rows)
    return sm, td, ur


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2: ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def compute_kpis(sm: pd.DataFrame, td: pd.DataFrame, ur: pd.DataFrame) -> dict:
    total_sanction    = sm["Total_Sanction_Cr"].sum()
    total_released    = sm["Total_Released_Cr"].sum()
    total_expended    = sm["Total_Expended_Cr"].sum()
    total_uc_sub      = sm["UC_Submitted_Cr"].sum()
    total_uc_pending  = sm["UC_Pending_Cr"].sum()

    overdue_tranches  = td[td["UC_Status"] == "OVERDUE"]
    high_risk         = td[td["Next_Tranche_Risk"] == "HIGH"]
    uc_returned       = ur[ur["UC_Status"] == "Returned for Correction"]

    utilization_rate  = total_expended / total_released * 100 if total_released else 0
    uc_coverage_rate  = total_uc_sub / total_expended * 100 if total_expended else 0

    return {
        "total_sanction"      : total_sanction,
        "total_released"      : total_released,
        "total_expended"      : total_expended,
        "total_uc_submitted"  : total_uc_sub,
        "total_uc_pending"    : total_uc_pending,
        "utilization_rate"    : round(utilization_rate, 1),
        "uc_coverage_rate"    : round(uc_coverage_rate, 1),
        "num_overdue_tranches": len(overdue_tranches),
        "overdue_amount"      : round(overdue_tranches["Amount_Released_Cr"].sum(), 2),
        "high_risk_schemes"   : len(high_risk["Scheme_ID"].unique()),
        "uc_returned_count"   : len(uc_returned),
        "uc_returned_amount"  : round(uc_returned["Amount_Cr"].sum(), 2),
    }


def risk_band(row) -> str:
    if row["UC_Status"] == "OVERDUE":
        return "🔴 HIGH"
    if row["UC_Status"] == "Pending" and (TODAY - row["UC_Due_Date"]).days > -30:
        return "🟡 MEDIUM"
    return "🟢 LOW"


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3: EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────

# ── Color Palette ─────────────────────────────────────────────────────────────
C_NAVY      = "1B2A4A"
C_GOLD      = "C9A84C"
C_RED       = "C0392B"
C_AMBER     = "E67E22"
C_GREEN     = "1E8449"
C_LIGHT_BG  = "F4F6F9"
C_WHITE     = "FFFFFF"
C_HEADER_BG = "1B2A4A"
C_ALT_ROW   = "EAF0FB"
C_BORDER    = "BDC3C7"

def hfill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def hfont(hex_color: str = "000000", bold: bool = False, size: int = 10) -> Font:
    return Font(color=hex_color, bold=bold, size=size, name="Calibri")

def border_all() -> Border:
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def center() -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_al() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def write_header_row(ws, row: int, cols: list, col_start: int = 1):
    for i, label in enumerate(cols):
        c = ws.cell(row=row, column=col_start + i, value=label)
        c.fill      = hfill(C_HEADER_BG)
        c.font      = hfont(C_WHITE, bold=True, size=10)
        c.alignment = center()
        c.border    = border_all()


def style_data_row(ws, row: int, num_cols: int, col_start: int = 1, alt: bool = False):
    bg = C_ALT_ROW if alt else C_WHITE
    for col in range(col_start, col_start + num_cols):
        c = ws.cell(row=row, column=col)
        c.fill      = hfill(bg)
        c.font      = hfont(size=9)
        c.alignment = left_al()
        c.border    = border_all()


def apply_risk_color(cell, value: str):
    if "HIGH" in str(value) or "OVERDUE" in str(value):
        cell.fill = hfill("FADBD8")
        cell.font = Font(color=C_RED, bold=True, size=9, name="Calibri")
    elif "MEDIUM" in str(value) or "Pending" in str(value):
        cell.fill = hfill("FDEBD0")
        cell.font = Font(color="A04000", bold=False, size=9, name="Calibri")
    elif "LOW" in str(value) or "Filed" in str(value) or "Approved" in str(value):
        cell.fill = hfill("D5F5E3")
        cell.font = Font(color=C_GREEN, bold=False, size=9, name="Calibri")


# ── SHEET 1: TITLE / COVER ────────────────────────────────────────────────────
def build_cover(wb: Workbook):
    ws = wb.active
    ws.title = "📋 Cover"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 30

    # Title block
    ws.row_dimensions[3].height = 40
    ws.merge_cells("B3:C3")
    t = ws["B3"]
    t.value     = "APCRDA — UC COMPLIANCE DASHBOARD"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=20, name="Calibri")
    t.alignment = center()

    ws.merge_cells("B4:C4")
    s = ws["B4"]
    s.value     = "Andhra Pradesh Capital Region Development Authority | Accounts Directorate"
    s.fill      = hfill(C_NAVY)
    s.font      = Font(color=C_WHITE, size=11, name="Calibri")
    s.alignment = center()

    ws.row_dimensions[6].height = 20
    meta = [
        ("Report Generated"    , TODAY.strftime("%d %B %Y")),
        ("Prepared by"         , "Internship — Finance & Accounts Wing"),
        ("Reporting Period"    , "FY 2023-24 to FY 2025-26"),
        ("Scope"               , f"{len(SCHEMES)} Active Schemes | All Funding Types"),
        ("Classification"      , "INTERNAL — NOT FOR CIRCULATION"),
    ]
    for i, (label, value) in enumerate(meta, start=7):
        ws.row_dimensions[i].height = 18
        l = ws.cell(row=i, column=2, value=label)
        l.font      = hfont(C_NAVY, bold=True, size=10)
        l.alignment = left_al()
        v = ws.cell(row=i, column=3, value=value)
        v.font      = hfont(size=10)
        v.alignment = left_al()
        if label == "Classification":
            v.font = Font(color=C_RED, bold=True, size=10, name="Calibri")

    ws.row_dimensions[14].height = 20
    ws.merge_cells("B14:C14")
    n = ws["B14"]
    n.value     = "SHEETS IN THIS WORKBOOK"
    n.font      = hfont(C_NAVY, bold=True, size=11)
    n.alignment = left_al()

    sheet_index = [
        ("📋 Cover"             , "This page — report metadata and index"),
        ("📊 KPI Summary"       , "Top-level financial KPIs and health scores"),
        ("📁 Scheme Master"     , "One row per scheme — sanctions, releases, expenditure"),
        ("📅 Tranche Tracker"   , "Per-tranche UC due dates, filing status, risk flags"),
        ("📄 UC Register"       , "All UCs submitted — status, approval, remarks"),
        ("🚨 Overdue Alerts"    , "Filtered view — only OVERDUE and HIGH RISK items"),
        ("📈 Charts"            , "Visual summaries for presentation use"),
    ]
    for i, (sname, sdesc) in enumerate(sheet_index, start=15):
        ws.row_dimensions[i].height = 16
        ws.cell(row=i, column=2, value=sname).font = hfont(C_NAVY, bold=True, size=9)
        ws.cell(row=i, column=3, value=sdesc).font = hfont(size=9)


# ── SHEET 2: KPI SUMMARY ─────────────────────────────────────────────────────
def build_kpi_sheet(wb: Workbook, kpis: dict):
    ws = wb.create_sheet("📊 KPI Summary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22

    # Title
    ws.merge_cells("B2:E2")
    t = ws["B2"]
    t.value     = "KEY PERFORMANCE INDICATORS — UC COMPLIANCE"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=14, name="Calibri")
    t.alignment = center()
    ws.row_dimensions[2].height = 30

    ws.merge_cells("B3:E3")
    ws["B3"].value     = f"As of {TODAY.strftime('%d %B %Y')}  |  All figures in ₹ Crore"
    ws["B3"].font      = hfont(C_NAVY, size=10)
    ws["B3"].alignment = center()

    # KPI Cards — row 5 onward
    card_data = [
        ("Total Scheme Sanction",    f"₹ {kpis['total_sanction']:,.2f} Cr",    C_NAVY,  "Aggregate approved outlay across all schemes"),
        ("Total Funds Released",     f"₹ {kpis['total_released']:,.2f} Cr",    "1A5276", "Actual funds drawn from GoI/GoAP/EAP"),
        ("Total Expenditure",        f"₹ {kpis['total_expended']:,.2f} Cr",    "1E8449", "Certified expenditure booked in accounts"),
        ("Utilization Rate",         f"{kpis['utilization_rate']}%",            "1E8449" if kpis['utilization_rate'] > 70 else C_AMBER, "Expenditure ÷ Released Funds"),
        ("UC Submitted",             f"₹ {kpis['total_uc_submitted']:,.2f} Cr", "1E8449", "UCs filed and acknowledged by nodal ministry"),
        ("UC Pending",               f"₹ {kpis['total_uc_pending']:,.2f} Cr",   C_AMBER, "Expenditure incurred but UC not yet filed"),
        ("UC Coverage Rate",         f"{kpis['uc_coverage_rate']}%",            "1E8449" if kpis['uc_coverage_rate'] > 80 else C_AMBER, "UC Submitted ÷ Total Expenditure"),
        ("Overdue Tranches",         f"{kpis['num_overdue_tranches']} tranches", C_RED,  f"₹ {kpis['overdue_amount']:,.2f} Cr blocked"),
        ("High-Risk Schemes",        f"{kpis['high_risk_schemes']} schemes",     C_RED,  "Next tranche at risk of delay"),
        ("UCs Returned/Rejected",    f"{kpis['uc_returned_count']} UCs",         C_AMBER, f"₹ {kpis['uc_returned_amount']:,.2f} Cr to be resubmitted"),
    ]

    r = 5
    for i, (label, value, color, note) in enumerate(card_data):
        ws.row_dimensions[r].height = 14
        ws.row_dimensions[r+1].height = 28
        ws.row_dimensions[r+2].height = 14
        ws.row_dimensions[r+3].height = 8

        for col in range(2, 6):
            for row in [r, r+1, r+2]:
                ws.cell(row=row, column=col).fill = hfill(color)

        ws.merge_cells(f"B{r}:E{r}")
        lbl = ws[f"B{r}"]
        lbl.value     = label.upper()
        lbl.font      = Font(color=C_WHITE, bold=False, size=8, name="Calibri")
        lbl.alignment = center()

        ws.merge_cells(f"B{r+1}:E{r+1}")
        val = ws[f"B{r+1}"]
        val.value     = value
        val.font      = Font(color=C_WHITE, bold=True, size=16, name="Calibri")
        val.alignment = center()

        ws.merge_cells(f"B{r+2}:E{r+2}")
        nt = ws[f"B{r+2}"]
        nt.value     = note
        nt.font      = Font(color="DDDDDD", size=8, name="Calibri")
        nt.alignment = center()

        r += 4

    # Risk summary box
    r += 2
    ws.merge_cells(f"B{r}:E{r}")
    rb = ws[f"B{r}"]
    rb.value     = "AUDIT RISK ASSESSMENT"
    rb.fill      = hfill(C_RED)
    rb.font      = Font(color=C_WHITE, bold=True, size=11, name="Calibri")
    rb.alignment = center()
    ws.row_dimensions[r].height = 22

    risk_lines = [
        f"• {kpis['num_overdue_tranches']} tranches overdue → ₹ {kpis['overdue_amount']:,.2f} Cr next-tranche release blocked",
        f"• UC coverage at {kpis['uc_coverage_rate']}% — CAG threshold typically 85%+",
        f"• {kpis['uc_returned_count']} UCs returned for correction — resurface in next AG inspection",
        f"• {kpis['high_risk_schemes']} schemes with HIGH next-tranche risk — immediate action required",
    ]
    for line in risk_lines:
        r += 1
        ws.row_dimensions[r].height = 16
        ws.merge_cells(f"B{r}:E{r}")
        c = ws[f"B{r}"]
        c.value     = line
        c.fill      = hfill("FADBD8")
        c.font      = Font(color=C_RED, size=9, name="Calibri")
        c.alignment = left_al()


# ── SHEET 3: SCHEME MASTER ────────────────────────────────────────────────────
def build_scheme_master(wb: Workbook, sm: pd.DataFrame):
    ws = wb.create_sheet("📁 Scheme Master")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    cols = [
        "Scheme ID", "Scheme Name", "Funding Type", "Nodal Ministry",
        "Sanction (₹Cr)", "Released (₹Cr)", "Expended (₹Cr)",
        "UC Submitted (₹Cr)", "UC Pending (₹Cr)",
        "Utilization %", "UC Coverage %",
        "GoI %", "GoAP %", "APCRDA %",
        "Scheme Start", "Scheme End", "No. Tranches",
    ]

    widths = [18, 35, 14, 20, 16, 16, 16, 18, 16, 14, 14, 10, 10, 12, 16, 16, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    t = ws["A1"]
    t.value     = "SCHEME MASTER REGISTER — APCRDA ACCOUNTS DIRECTORATE"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t.alignment = center()

    ws.row_dimensions[2].height = 30
    write_header_row(ws, 2, cols)

    sm["Utilization_Pct"] = (sm["Total_Expended_Cr"] / sm["Total_Released_Cr"] * 100).round(1)
    sm["UC_Coverage_Pct"] = (sm["UC_Submitted_Cr"] / sm["Total_Expended_Cr"] * 100).round(1)

    display_map = {
        "Scheme_ID"          : "Scheme ID",
        "Scheme_Name"        : "Scheme Name",
        "Funding_Type"       : "Funding Type",
        "Nodal_Ministry"     : "Nodal Ministry",
        "Total_Sanction_Cr"  : "Sanction (₹Cr)",
        "Total_Released_Cr"  : "Released (₹Cr)",
        "Total_Expended_Cr"  : "Expended (₹Cr)",
        "UC_Submitted_Cr"    : "UC Submitted (₹Cr)",
        "UC_Pending_Cr"      : "UC Pending (₹Cr)",
        "Utilization_Pct"    : "Utilization %",
        "UC_Coverage_Pct"    : "UC Coverage %",
        "GoI_Share_Pct"      : "GoI %",
        "GoAP_Share_Pct"     : "GoAP %",
        "APCRDA_Share_Pct"   : "APCRDA %",
        "Scheme_Start"       : "Scheme Start",
        "Scheme_End"         : "Scheme End",
        "Num_Tranches"       : "No. Tranches",
    }

    for r_idx, (_, row) in enumerate(sm.iterrows(), start=3):
        ws.row_dimensions[r_idx].height = 18
        alt = r_idx % 2 == 0
        style_data_row(ws, r_idx, len(cols), alt=alt)
        for c_idx, src_col in enumerate(display_map.keys(), start=1):
            val = row[src_col]
            cell = ws.cell(row=r_idx, column=c_idx)
            if isinstance(val, datetime):
                cell.value = val.strftime("%d-%b-%Y")
            elif isinstance(val, float) and src_col.endswith("_Pct"):
                cell.value = f"{val:.1f}%"
                if val < 60:
                    apply_risk_color(cell, "HIGH")
                elif val < 80:
                    apply_risk_color(cell, "MEDIUM")
                else:
                    apply_risk_color(cell, "LOW")
            elif isinstance(val, float):
                cell.value = round(val, 2)
                cell.number_format = '#,##0.00'
            else:
                cell.value = val

            if src_col == "Funding_Type":
                if val == "CSS":
                    cell.fill = hfill("D6EAF8")
                elif val == "EAP":
                    cell.fill = hfill("D5F5E3")
                else:
                    cell.fill = hfill("FDEBD0")


# ── SHEET 4: TRANCHE TRACKER ──────────────────────────────────────────────────
def build_tranche_tracker(wb: Workbook, td: pd.DataFrame):
    ws = wb.create_sheet("📅 Tranche Tracker")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    cols = [
        "Scheme ID", "Scheme Name", "Funding Type", "Tranche",
        "Release Date", "Amount (₹Cr)", "UC Due Date",
        "UC Filed Date", "UC Status", "Days Overdue", "Next Tranche Risk",
    ]
    widths = [18, 35, 14, 14, 16, 14, 16, 16, 14, 14, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    t = ws["A1"]
    t.value     = "TRANCHE-WISE UC TRACKING REGISTER"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t.alignment = center()

    ws.row_dimensions[2].height = 30
    write_header_row(ws, 2, cols)

    field_map = [
        "Scheme_ID", "Scheme_Name", "Funding_Type", "Tranche",
        "Release_Date", "Amount_Released_Cr", "UC_Due_Date",
        "UC_Filed_Date", "UC_Status", "Days_Overdue", "Next_Tranche_Risk",
    ]

    for r_idx, (_, row) in enumerate(td.iterrows(), start=3):
        ws.row_dimensions[r_idx].height = 18
        alt = r_idx % 2 == 0
        style_data_row(ws, r_idx, len(cols), alt=alt)
        for c_idx, fld in enumerate(field_map, start=1):
            val = row[fld]
            cell = ws.cell(row=r_idx, column=c_idx)
            if pd.isna(val):
                cell.value = "—"
            elif isinstance(val, datetime):
                cell.value = val.strftime("%d-%b-%Y")
            elif fld == "Amount_Released_Cr":
                cell.value = round(float(val), 2)
                cell.number_format = '#,##0.00'
            else:
                cell.value = val

            if fld in ("UC_Status", "Next_Tranche_Risk"):
                apply_risk_color(cell, str(val))

            if fld == "Days_Overdue" and isinstance(val, (int, float)) and val > 0:
                cell.fill  = hfill("FADBD8")
                cell.font  = Font(color=C_RED, bold=True, size=9, name="Calibri")
                cell.value = int(val)


# ── SHEET 5: UC REGISTER ──────────────────────────────────────────────────────
def build_uc_register(wb: Workbook, ur: pd.DataFrame):
    ws = wb.create_sheet("📄 UC Register")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    cols = [
        "UC Number", "Scheme ID", "Scheme Name", "Funding Type",
        "Nodal Ministry", "Period From", "Period To", "Amount (₹Cr)",
        "Submission Date", "Acknowledgement Date", "Approval Date",
        "UC Status", "Remarks",
    ]
    widths = [28, 18, 35, 14, 20, 16, 16, 14, 16, 20, 16, 22, 45]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    t = ws["A1"]
    t.value     = "UTILIZATION CERTIFICATE (UC) SUBMISSION REGISTER"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t.alignment = center()

    ws.row_dimensions[2].height = 30
    write_header_row(ws, 2, cols)

    field_map = [
        "UC_Number", "Scheme_ID", "Scheme_Name", "Funding_Type",
        "Nodal_Ministry", "Period_From", "Period_To", "Amount_Cr",
        "Submission_Date", "Acknowledgement_Dt", "Approval_Date",
        "UC_Status", "Remarks",
    ]

    for r_idx, (_, row) in enumerate(ur.iterrows(), start=3):
        ws.row_dimensions[r_idx].height = 18
        alt = r_idx % 2 == 0
        style_data_row(ws, r_idx, len(cols), alt=alt)
        for c_idx, fld in enumerate(field_map, start=1):
            val = row[fld]
            cell = ws.cell(row=r_idx, column=c_idx)
            if pd.isna(val) or val == "":
                cell.value = "—"
            elif isinstance(val, datetime):
                cell.value = val.strftime("%d-%b-%Y")
            elif fld == "Amount_Cr":
                cell.value = round(float(val), 2)
                cell.number_format = '#,##0.00'
            else:
                cell.value = val

            if fld == "UC_Status":
                apply_risk_color(cell, str(val))


# ── SHEET 6: OVERDUE ALERTS ───────────────────────────────────────────────────
def build_overdue_alerts(wb: Workbook, td: pd.DataFrame, ur: pd.DataFrame):
    ws = wb.create_sheet("🚨 Overdue Alerts")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:K1")
    t = ws["A1"]
    t.value     = "⚠  OVERDUE UC ALERTS — IMMEDIATE ACTION REQUIRED"
    t.fill      = hfill(C_RED)
    t.font      = Font(color=C_WHITE, bold=True, size=14, name="Calibri")
    t.alignment = center()

    overdue_td = td[td["UC_Status"] == "OVERDUE"].copy().reset_index(drop=True)
    returned_ur = ur[ur["UC_Status"] == "Returned for Correction"].copy().reset_index(drop=True)

    # Section A: Overdue tranches
    ws.row_dimensions[3].height = 22
    ws.merge_cells("A3:K3")
    s = ws["A3"]
    s.value     = f"SECTION A: OVERDUE TRANCHES  ({len(overdue_td)} items  |  ₹ {overdue_td['Amount_Released_Cr'].sum():,.2f} Cr blocked)"
    s.fill      = hfill("922B21")
    s.font      = Font(color=C_WHITE, bold=True, size=11, name="Calibri")
    s.alignment = left_al()

    cols_a = ["Scheme ID", "Scheme Name", "Tranche", "Amount (₹Cr)",
              "UC Due Date", "Days Overdue", "Next Tranche Risk", "Action Required"]
    widths_a = [18, 38, 14, 14, 16, 14, 18, 45]
    for i, w in enumerate(widths_a, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    write_header_row(ws, 4, cols_a)
    for r_idx, (_, row) in enumerate(overdue_td.iterrows(), start=5):
        ws.row_dimensions[r_idx].height = 18
        style_data_row(ws, r_idx, len(cols_a), alt=r_idx % 2 == 0)
        action = f"File UC immediately — {int(row['Days_Overdue'])} days overdue. Next tranche blocked."
        vals = [
            row["Scheme_ID"], row["Scheme_Name"], row["Tranche"],
            round(row["Amount_Released_Cr"], 2),
            row["UC_Due_Date"].strftime("%d-%b-%Y"),
            int(row["Days_Overdue"]),
            row["Next_Tranche_Risk"],
            action,
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.fill   = hfill("FADBD8")
            cell.font   = Font(color=C_RED, size=9, name="Calibri")
            cell.border = border_all()
            cell.alignment = left_al()
            if c_idx == 4:
                cell.number_format = '#,##0.00'

    # Section B: Returned UCs
    row_b = 5 + len(overdue_td) + 2
    ws.row_dimensions[row_b].height = 22
    ws.merge_cells(f"A{row_b}:H{row_b}")
    s2 = ws[f"A{row_b}"]
    s2.value     = f"SECTION B: UCs RETURNED FOR CORRECTION  ({len(returned_ur)} items  |  ₹ {returned_ur['Amount_Cr'].sum():,.2f} Cr)"
    s2.fill      = hfill(C_AMBER)
    s2.font      = Font(color=C_WHITE, bold=True, size=11, name="Calibri")
    s2.alignment = left_al()

    cols_b = ["UC Number", "Scheme Name", "Amount (₹Cr)", "Submission Date", "Remarks", "Action"]
    write_header_row(ws, row_b + 1, cols_b)
    for r_idx, (_, row) in enumerate(returned_ur.iterrows(), start=row_b + 2):
        ws.row_dimensions[r_idx].height = 18
        style_data_row(ws, r_idx, len(cols_b), alt=r_idx % 2 == 0)
        vals = [
            row["UC_Number"], row["Scheme_Name"],
            round(row["Amount_Cr"], 2),
            row["Submission_Date"].strftime("%d-%b-%Y") if not pd.isna(row["Submission_Date"]) else "—",
            row["Remarks"],
            "Correct MB mismatch and resubmit within 15 days",
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.fill   = hfill("FDEBD0")
            cell.font   = Font(color="7D6608", size=9, name="Calibri")
            cell.border = border_all()
            cell.alignment = left_al()
            if c_idx == 3:
                cell.number_format = '#,##0.00'


# ── SHEET 7: CHARTS ───────────────────────────────────────────────────────────
def build_charts(wb: Workbook, sm: pd.DataFrame, td: pd.DataFrame):
    ws = wb.create_sheet("📈 Charts")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    ws.row_dimensions[1].height = 28
    ws.merge_cells("B1:Q1")
    t = ws["B1"]
    t.value     = "VISUAL SUMMARY — UC COMPLIANCE DASHBOARD"
    t.fill      = hfill(C_NAVY)
    t.font      = Font(color=C_GOLD, bold=True, size=13, name="Calibri")
    t.alignment = center()

    # ── Data table for charts (write at col 20+) ──────────────────────────────
    sm2 = sm.copy()
    sm2["UC_Cov"] = (sm2["UC_Submitted_Cr"] / sm2["Total_Expended_Cr"] * 100).round(1)
    sm2["Util"]   = (sm2["Total_Expended_Cr"] / sm2["Total_Released_Cr"] * 100).round(1)

    # Scheme names in T column (col 20)
    ws.cell(row=3, column=20, value="Scheme")
    ws.cell(row=3, column=21, value="UC Coverage %")
    ws.cell(row=3, column=22, value="Utilization %")
    ws.cell(row=3, column=23, value="Expended (₹Cr)")
    for i, (_, row) in enumerate(sm2.iterrows(), start=4):
        name = row["Scheme_Name"][:25]
        ws.cell(row=i, column=20, value=name)
        ws.cell(row=i, column=21, value=row["UC_Cov"])
        ws.cell(row=i, column=22, value=row["Util"])
        ws.cell(row=i, column=23, value=row["Total_Expended_Cr"])

    n = len(sm2) + 3

    # Chart 1: UC Coverage % by Scheme
    chart1 = BarChart()
    chart1.type     = "bar"
    chart1.title    = "UC Coverage % by Scheme"
    chart1.y_axis.title = "Scheme"
    chart1.x_axis.title = "Coverage %"
    chart1.width    = 18
    chart1.height   = 12
    data1  = Reference(ws, min_col=21, max_col=21, min_row=3, max_row=n)
    cats1  = Reference(ws, min_col=20, min_row=4, max_row=n)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.series[0].graphicalProperties.solidFill = "1B2A4A"
    ws.add_chart(chart1, "B3")

    # Chart 2: Utilization % by Scheme
    chart2 = BarChart()
    chart2.type     = "bar"
    chart2.title    = "Fund Utilization % by Scheme"
    chart2.y_axis.title = "Scheme"
    chart2.x_axis.title = "Utilization %"
    chart2.width    = 18
    chart2.height   = 12
    data2  = Reference(ws, min_col=22, max_col=22, min_row=3, max_row=n)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats1)
    chart2.series[0].graphicalProperties.solidFill = "C9A84C"
    ws.add_chart(chart2, "B24")

    # Chart 3: UC Status breakdown (Pie)
    status_counts = td["UC_Status"].value_counts()
    ws.cell(row=3,  column=25, value="UC Status")
    ws.cell(row=3,  column=26, value="Count")
    for i, (status, count) in enumerate(status_counts.items(), start=4):
        ws.cell(row=i, column=25, value=status)
        ws.cell(row=i, column=26, value=int(count))

    chart3 = PieChart()
    chart3.title  = "UC Status Distribution"
    chart3.width  = 14
    chart3.height = 10
    data3 = Reference(ws, min_col=26, max_col=26, min_row=3, max_row=3 + len(status_counts))
    cats3 = Reference(ws, min_col=25, min_row=4, max_row=3 + len(status_counts))
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    colors = ["1E8449", "E67E22", "C0392B"]
    for i, pt in enumerate(chart3.series[0].dPt if hasattr(chart3.series[0], 'dPt') else []):
        pass  # openpyxl pie coloring is limited; default is fine
    ws.add_chart(chart3, "Q3")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4: MAIN RUNNER
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  APCRDA PROJECT 1 — UC COMPLIANCE DASHBOARD")
    print("=" * 60)

    print("\n[1/6] Generating scheme and tranche data...")
    sm, td, ur = generate_sample_data()

    print("[2/6] Computing KPIs...")
    kpis = compute_kpis(sm, td, ur)

    print("\n── KPI SNAPSHOT ─────────────────────────────────────────")
    print(f"  Total Sanction      : ₹ {kpis['total_sanction']:>10,.2f} Cr")
    print(f"  Total Released      : ₹ {kpis['total_released']:>10,.2f} Cr")
    print(f"  Total Expended      : ₹ {kpis['total_expended']:>10,.2f} Cr")
    print(f"  Utilization Rate    : {kpis['utilization_rate']}%")
    print(f"  UC Submitted        : ₹ {kpis['total_uc_submitted']:>10,.2f} Cr")
    print(f"  UC Pending          : ₹ {kpis['total_uc_pending']:>10,.2f} Cr")
    print(f"  UC Coverage Rate    : {kpis['uc_coverage_rate']}%")
    print(f"  Overdue Tranches    : {kpis['num_overdue_tranches']}")
    print(f"  High-Risk Schemes   : {kpis['high_risk_schemes']}")
    print(f"  UCs Returned        : {kpis['uc_returned_count']}")
    print("─" * 60)

    print("\n[3/6] Building Excel workbook...")
    wb = Workbook()

    print("       → Cover sheet")
    build_cover(wb)

    print("       → KPI Summary")
    build_kpi_sheet(wb, kpis)

    print("       → Scheme Master")
    build_scheme_master(wb, sm)

    print("       → Tranche Tracker")
    build_tranche_tracker(wb, td)

    print("       → UC Register")
    build_uc_register(wb, ur)

    print("       → Overdue Alerts")
    build_overdue_alerts(wb, td, ur)

    print("       → Charts")
    build_charts(wb, sm, td)

    output_path = "/mnt/user-data/outputs/UC_Compliance_Dashboard.xlsx"
    print(f"\n[4/6] Saving to {output_path} ...")
    wb.save(output_path)

    print("\n[5/6] Verification:")
    import os
    size_kb = os.path.getsize(output_path) / 1024
    print(f"       File size    : {size_kb:.1f} KB")
    print(f"       Sheets built : 7")
    print(f"       Schemes      : {len(sm)}")
    print(f"       Tranches     : {len(td)}")
    print(f"       UC entries   : {len(ur)}")

    print("\n[6/6] HOW TO USE THIS FILE")
    print("────────────────────────────────────────────────────────────")
    print("  REPLACE SAMPLE DATA:")
    print("  → Open generate_sample_data() in this script")
    print("  → Replace SCHEMES list with your actual scheme names,")
    print("    sanctions, and funding splits")
    print("  → Replace tranche dates and amounts from your registers")
    print("  → Replace UC submission dates from your UC files")
    print("")
    print("  WHAT EACH SHEET DOES:")
    print("  📋 Cover          → Index + report metadata")
    print("  📊 KPI Summary    → 10 key metrics for your Director briefing")
    print("  📁 Scheme Master  → Full scheme-level financials")
    print("  📅 Tranche Tracker→ Per-tranche UC status + risk flags")
    print("  📄 UC Register    → Every UC filed, its status, remarks")
    print("  🚨 Overdue Alerts → Only the critical items — show this first")
    print("  📈 Charts         → Copy-paste into your final presentation")
    print("════════════════════════════════════════════════════════════")
    print("  OUTPUT: UC_Compliance_Dashboard.xlsx")
    print("════════════════════════════════════════════════════════════\n")


if __name__ == "__main__":
    main()
