"""
APCRDA PROJECT 3: Budget Utilization Velocity Model
====================================================
Purpose : Project end-of-year spend rate per scheme.
          Flags lapsing risk, underspend patterns, and March rush.
          Gives the Director a "will we utilize or lapse?" answer by scheme.

Run     : python project3_budget_velocity.py
Output  : Budget_Velocity_Model.xlsx

OFFLINE — replace generate_budget_data() with actuals.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from datetime import datetime, timedelta
import calendar
import random
import warnings
warnings.filterwarnings("ignore")

random.seed(7)
np.random.seed(7)

TODAY     = datetime.today()
FY_START  = datetime(TODAY.year if TODAY.month >= 4 else TODAY.year - 1, 4, 1)
FY_END    = datetime(FY_START.year + 1, 3, 31)
MONTHS_ELAPSED   = max(1, (TODAY.year - FY_START.year) * 12 + (TODAY.month - FY_START.month) + 1)
MONTHS_REMAINING = max(1, 12 - MONTHS_ELAPSED + 1)

# ─────────────────────────────────────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────────────────────────────────────
C_NAVY="1B2A4A"; C_GOLD="C9A84C"; C_RED="C0392B"; C_AMBER="E67E22"
C_GREEN="1E8449"; C_WHITE="FFFFFF"; C_ALT="EAF0FB"; C_BORDER="BDC3C7"

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
        c.fill=hfill(C_NAVY); c.font=hfont(C_WHITE,True,10)
        c.alignment=ctr(); c.border=bdr()
def style_row(ws, row, ncols, col_start=1, alt=False):
    bg = C_ALT if alt else C_WHITE
    for col in range(col_start, col_start+ncols):
        c=ws.cell(row=row, column=col)
        c.fill=hfill(bg); c.font=hfont(sz=9); c.alignment=lft(); c.border=bdr()


# ─────────────────────────────────────────────────────────────────────────────
# DATA GENERATOR
# ─────────────────────────────────────────────────────────────────────────────
SCHEMES = [
    ("AMRUT 2.0 - Water Supply",       280.0,  "CSS"),
    ("AMRUT 2.0 - Sewerage",           160.0,  "CSS"),
    ("Smart City Mission - Core Infra",400.0,  "CSS"),
    ("Smart City Mission - ICT",       100.0,  "CSS"),
    ("PMAY-Urban",                     185.0,  "CSS"),
    ("AP Capital Roads Ph-I",          520.0,  "State"),
    ("AP Capital Roads Ph-II",         310.0,  "State"),
    ("Amaravati Green Grid",           88.0,   "State"),
    ("World Bank Urban Infra",         650.0,  "EAP"),
    ("ADB Resilient Infra",            580.0,  "EAP"),
    ("NCRMP-II",                       130.0,  "CSS"),
    ("HRIDAY Heritage",                42.0,   "CSS"),
]

MONTH_NAMES = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]


def generate_budget_data():
    """
    Returns:
      scheme_df  : one row per scheme with annual budget + projected spend
      monthly_df : 12 monthly rows per scheme (actuals for elapsed, projected for remaining)
    """
    scheme_rows = []
    monthly_rows = []

    for s_idx, (name, annual_budget, ftype) in enumerate(SCHEMES):
        allotment  = round(annual_budget * random.uniform(0.75, 1.05), 2)  # released this FY
        # Simulate a realistic monthly spend profile
        # Government spending is typically back-loaded (slow Apr-Nov, heavy Dec-Mar)
        base_weights = [0.04, 0.05, 0.06, 0.07, 0.07, 0.08, 0.09, 0.10, 0.11, 0.12, 0.11, 0.10]
        # Add scheme-specific variance
        noise = [max(0.01, w + random.uniform(-0.02, 0.02)) for w in base_weights]
        total_weight = sum(noise)
        monthly_weights = [w / total_weight for w in noise]

        actual_monthly = []
        for m in range(12):
            if m < MONTHS_ELAPSED:
                target = allotment * monthly_weights[m]
                actual = round(target * random.uniform(0.50, 1.20), 2)
                actual_monthly.append(actual)
            else:
                actual_monthly.append(None)

        actual_spend_to_date = sum(x for x in actual_monthly if x is not None)
        elapsed_weight_sum   = sum(monthly_weights[:MONTHS_ELAPSED])
        expected_to_date     = allotment * elapsed_weight_sum

        # Projected remaining spend (linear extrapolation from current velocity)
        monthly_velocity = actual_spend_to_date / MONTHS_ELAPSED
        projected_remaining = round(monthly_velocity * MONTHS_REMAINING, 2)
        projected_total     = round(actual_spend_to_date + projected_remaining, 2)
        projected_utilization = round(projected_total / allotment * 100, 1) if allotment else 0
        lapsing_amount      = round(max(0, allotment - projected_total), 2)

        # Risk classification
        if projected_utilization < 70:
            lapse_risk = "🔴 HIGH LAPSE RISK"
        elif projected_utilization < 85:
            lapse_risk = "🟡 MODERATE RISK"
        else:
            lapse_risk = "🟢 ON TRACK"

        scheme_rows.append({
            "Scheme_ID"              : f"APCRDA-FY{FY_START.year%100}{str(s_idx+1).zfill(3)}",
            "Scheme_Name"            : name,
            "Funding_Type"           : ftype,
            "Annual_Budget_Cr"       : annual_budget,
            "FY_Allotment_Cr"        : allotment,
            "Actual_to_Date_Cr"      : round(actual_spend_to_date, 2),
            "Expected_to_Date_Cr"    : round(expected_to_date, 2),
            "Variance_Cr"            : round(actual_spend_to_date - expected_to_date, 2),
            "Monthly_Velocity_Cr"    : round(monthly_velocity, 2),
            "Months_Elapsed"         : MONTHS_ELAPSED,
            "Months_Remaining"       : MONTHS_REMAINING,
            "Projected_Remaining_Cr" : projected_remaining,
            "Projected_Total_Cr"     : projected_total,
            "Projected_Utilization_Pct": projected_utilization,
            "Lapsing_Amount_Cr"      : lapsing_amount,
            "Lapse_Risk"             : lapse_risk,
        })

        # Monthly detail
        for m in range(12):
            if m < MONTHS_ELAPSED:
                amt     = actual_monthly[m]
                is_proj = False
            else:
                amt     = round(monthly_velocity, 2)
                is_proj = True

            cumulative = 0
            for prev_m in range(m + 1):
                if prev_m < MONTHS_ELAPSED:
                    cumulative += actual_monthly[prev_m] if actual_monthly[prev_m] else 0
                else:
                    cumulative += monthly_velocity

            monthly_rows.append({
                "Scheme_ID"       : f"APCRDA-FY{FY_START.year%100}{str(s_idx+1).zfill(3)}",
                "Scheme_Name"     : name,
                "Funding_Type"    : ftype,
                "Month_Number"    : m + 1,
                "Month_Name"      : MONTH_NAMES[m],
                "FY_Month_Label"  : f"FY-M{str(m+1).zfill(2)} ({MONTH_NAMES[m]})",
                "Amount_Cr"       : round(amt, 2),
                "Is_Projected"    : is_proj,
                "Cumulative_Cr"   : round(cumulative, 2),
                "Budget_Allotment": allotment,
                "Cum_Utilization_Pct": round(cumulative / allotment * 100, 1) if allotment else 0,
            })

    return pd.DataFrame(scheme_rows), pd.DataFrame(monthly_rows)


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
def compute_portfolio_kpis(scheme_df: pd.DataFrame) -> dict:
    total_allotment    = scheme_df["FY_Allotment_Cr"].sum()
    total_actual       = scheme_df["Actual_to_Date_Cr"].sum()
    total_projected    = scheme_df["Projected_Total_Cr"].sum()
    total_lapsing      = scheme_df["Lapsing_Amount_Cr"].sum()
    high_risk          = scheme_df[scheme_df["Lapse_Risk"].str.contains("HIGH")]
    moderate_risk      = scheme_df[scheme_df["Lapse_Risk"].str.contains("MODERATE")]

    return {
        "total_allotment"   : round(total_allotment, 2),
        "total_actual"      : round(total_actual, 2),
        "actual_utilization": round(total_actual / total_allotment * 100, 1),
        "projected_total"   : round(total_projected, 2),
        "projected_util"    : round(total_projected / total_allotment * 100, 1),
        "total_lapsing"     : round(total_lapsing, 2),
        "high_risk_count"   : len(high_risk),
        "high_risk_lapsing" : round(high_risk["Lapsing_Amount_Cr"].sum(), 2),
        "moderate_risk_count": len(moderate_risk),
        "months_elapsed"    : MONTHS_ELAPSED,
        "months_remaining"  : MONTHS_REMAINING,
        "fy_label"          : f"FY {FY_START.year}-{str(FY_END.year)[-2:]}",
    }


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(scheme_df, monthly_df, kpis):
    wb = Workbook()

    # ── SHEET 1: PORTFOLIO KPI ────────────────────────────────────────────────
    ws = wb.active
    ws.title = "📊 Portfolio KPIs"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 24

    ws.merge_cells("B2:C2")
    t = ws["B2"]
    t.value = f"BUDGET VELOCITY MODEL — {kpis['fy_label']} PORTFOLIO VIEW"
    t.fill=hfill(C_NAVY); t.font=Font(color=C_GOLD,bold=True,size=14,name="Calibri")
    t.alignment=ctr(); ws.row_dimensions[2].height=30

    ws.merge_cells("B3:C3")
    ws["B3"].value = (f"As of {TODAY.strftime('%d %B %Y')}  |  "
                      f"Month {MONTHS_ELAPSED} of 12  |  {MONTHS_REMAINING} months remaining")
    ws["B3"].font=hfont(C_NAVY,sz=10); ws["B3"].alignment=ctr()

    card_data = [
        ("FY Allotment (Total)",    f"₹ {kpis['total_allotment']:,.2f} Cr",    C_NAVY,  f"{kpis['fy_label']} approved allotment"),
        ("Actual Spend to Date",    f"₹ {kpis['total_actual']:,.2f} Cr",       "1A5276", f"Utilization: {kpis['actual_utilization']}%"),
        ("Projected Year-End Spend",f"₹ {kpis['projected_total']:,.2f} Cr",    C_AMBER, f"At current burn rate: {kpis['projected_util']}%"),
        ("At-Risk Lapsing Amount",  f"₹ {kpis['total_lapsing']:,.2f} Cr",      C_RED,   "Will lapse if velocity doesn't improve"),
        ("High Lapse Risk Schemes", f"{kpis['high_risk_count']} schemes",       C_RED,   f"₹ {kpis['high_risk_lapsing']:,.2f} Cr combined exposure"),
        ("Moderate Risk Schemes",   f"{kpis['moderate_risk_count']} schemes",   C_AMBER, "Intervention needed within 60 days"),
    ]

    r = 5
    for label, value, color, note in card_data:
        for row in [r, r+1, r+2]:
            ws.row_dimensions[row].height = 14 if row!=r+1 else 26
        for row in [r, r+1, r+2]:
            ws.merge_cells(f"B{row}:C{row}")
            ws[f"B{row}"].fill = hfill(color)
        ws[f"B{r}"].value=label.upper()
        ws[f"B{r}"].font=Font(color=C_WHITE,size=8,name="Calibri"); ws[f"B{r}"].alignment=ctr()
        ws[f"B{r+1}"].value=value
        ws[f"B{r+1}"].font=Font(color=C_WHITE,bold=True,size=16,name="Calibri"); ws[f"B{r+1}"].alignment=ctr()
        ws[f"B{r+2}"].value=note
        ws[f"B{r+2}"].font=Font(color="DDDDDD",size=8,name="Calibri"); ws[f"B{r+2}"].alignment=ctr()
        r += 4

    # Remediation note
    r += 1
    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"].value = "WHAT THIS MEANS: CORRECTIVE ACTIONS"
    ws[f"B{r}"].fill=hfill(C_NAVY); ws[f"B{r}"].font=hfont(C_GOLD,True,11)
    ws[f"B{r}"].alignment=ctr(); ws.row_dimensions[r].height=22; r+=1

    actions = [
        f"• {kpis['high_risk_count']} schemes project <70% utilization — immediate acceleration required",
        f"• ₹ {kpis['total_lapsing']:,.2f} Cr at risk of lapsing on 31 March — requires re-appropriation or accelerated payments",
        f"• With {MONTHS_REMAINING} months remaining, required monthly spend = ₹ {kpis['total_lapsing']/max(MONTHS_REMAINING,1):,.2f} Cr extra per month",
        "• Recommend: monthly velocity review with EEs for all HIGH RISK schemes from next week",
    ]
    for action in actions:
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"].value=action; ws[f"B{r}"].fill=hfill("FDEBD0")
        ws[f"B{r}"].font=Font(color="784212",size=9,name="Calibri"); ws[f"B{r}"].alignment=lft()
        ws.row_dimensions[r].height=16; r+=1

    # ── SHEET 2: SCHEME VELOCITY TABLE ────────────────────────────────────────
    ws2 = wb.create_sheet("📋 Scheme Velocity")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "A3"

    cols2 = [
        "Scheme ID", "Scheme Name", "Funding Type",
        "Annual Budget (₹Cr)", "FY Allotment (₹Cr)", "Actual to Date (₹Cr)",
        "Expected to Date (₹Cr)", "Variance (₹Cr)",
        "Monthly Velocity (₹Cr)", "Months Remaining",
        "Projected Year-End (₹Cr)", "Projected Utilization %",
        "Lapsing Amount (₹Cr)", "Lapse Risk",
    ]
    widths2 = [18,35,14,18,18,20,20,16,20,16,22,22,20,24]
    for i,w in enumerate(widths2,1): ws2.column_dimensions[get_column_letter(i)].width=w

    ws2.row_dimensions[1].height=28
    ws2.merge_cells(f"A1:{get_column_letter(len(cols2))}1")
    t2=ws2["A1"]; t2.value="SCHEME-WISE BUDGET VELOCITY — DETAILED VIEW"
    t2.fill=hfill(C_NAVY); t2.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
    t2.alignment=ctr()
    hdr_row(ws2, 2, cols2)

    field_map2 = [
        "Scheme_ID","Scheme_Name","Funding_Type",
        "Annual_Budget_Cr","FY_Allotment_Cr","Actual_to_Date_Cr",
        "Expected_to_Date_Cr","Variance_Cr",
        "Monthly_Velocity_Cr","Months_Remaining",
        "Projected_Total_Cr","Projected_Utilization_Pct",
        "Lapsing_Amount_Cr","Lapse_Risk",
    ]
    money_cols2 = {"Annual_Budget_Cr","FY_Allotment_Cr","Actual_to_Date_Cr",
                   "Expected_to_Date_Cr","Variance_Cr","Monthly_Velocity_Cr",
                   "Projected_Total_Cr","Lapsing_Amount_Cr"}

    for r_idx, (_, row) in enumerate(scheme_df.iterrows(), start=3):
        ws2.row_dimensions[r_idx].height=18
        style_row(ws2, r_idx, len(cols2), alt=r_idx%2==0)
        for c_idx, fld in enumerate(field_map2, start=1):
            val = row[fld]
            cell = ws2.cell(row=r_idx, column=c_idx)
            if fld in money_cols2:
                cell.value=round(float(val),2); cell.number_format='#,##0.00'
                if fld == "Variance_Cr":
                    if float(val) < 0:
                        cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                    else:
                        cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
            elif fld == "Projected_Utilization_Pct":
                cell.value=f"{float(val):.1f}%"
                if float(val) < 70:
                    cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                elif float(val) < 85:
                    cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")
                else:
                    cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
            elif fld == "Lapse_Risk":
                cell.value=val
                if "HIGH" in str(val):
                    cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                elif "MODERATE" in str(val):
                    cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")
                else:
                    cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
            else:
                cell.value=val

    # ── SHEET 3: MONTHLY DRILL-DOWN (Pivot-style) ──────────────────────────────
    ws3 = wb.create_sheet("📅 Monthly Drill-Down")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "C3"

    ws3.row_dimensions[1].height=28
    month_cols = [f"M{str(m+1).zfill(2)}\n{MONTH_NAMES[m]}" for m in range(12)]
    all_cols3 = ["Scheme ID", "Scheme Name", "Type"] + month_cols + ["Total Projected"]
    widths3 = [18,35,10] + [12]*12 + [16]
    for i,w in enumerate(widths3,1): ws3.column_dimensions[get_column_letter(i)].width=w

    ws3.merge_cells(f"A1:{get_column_letter(len(all_cols3))}1")
    t3=ws3["A1"]; t3.value="MONTHLY EXPENDITURE DRILL-DOWN (ACTUALS + PROJECTIONS)"
    t3.fill=hfill(C_NAVY); t3.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
    t3.alignment=ctr()

    hdr_row(ws3, 2, all_cols3)

    # Add projected marker to header
    for m in range(12):
        if m >= MONTHS_ELAPSED:
            cell = ws3.cell(row=2, column=4+m)
            cell.fill = hfill("7D3C98")  # purple = projected
            cell.font = Font(color=C_WHITE, bold=True, size=10, name="Calibri")

    # Legend
    ws3.cell(row=2, column=4+MONTHS_ELAPSED).value = (
        ws3.cell(row=2, column=4+MONTHS_ELAPSED).value or "") + "\n[PROJECTED →]"

    pivot = monthly_df.pivot_table(
        index=["Scheme_ID","Scheme_Name","Funding_Type"],
        columns="Month_Number",
        values="Amount_Cr",
        aggfunc="sum"
    ).reset_index()

    proj_pivot = monthly_df.pivot_table(
        index=["Scheme_ID","Scheme_Name","Funding_Type"],
        columns="Month_Number",
        values="Is_Projected",
        aggfunc="max"
    ).reset_index()

    for r_idx, (_, row) in enumerate(pivot.iterrows(), start=3):
        ws3.row_dimensions[r_idx].height=18
        style_row(ws3, r_idx, len(all_cols3), alt=r_idx%2==0)
        ws3.cell(row=r_idx, column=1).value = row.get("Scheme_ID","")
        ws3.cell(row=r_idx, column=2).value = row.get("Scheme_Name","")
        ws3.cell(row=r_idx, column=3).value = row.get("Funding_Type","")

        row_total = 0
        for m in range(1, 13):
            val = row.get(m, 0) or 0
            row_total += val
            cell = ws3.cell(row=r_idx, column=3+m)
            cell.value = round(val, 2)
            cell.number_format = '#,##0.00'
            # projected months get a different shade
            is_proj = proj_pivot.iloc[r_idx-3].get(m, False) if r_idx-3 < len(proj_pivot) else False
            if is_proj:
                cell.fill = hfill("E8DAEF")
                cell.font = Font(color="7D3C98", size=9, name="Calibri", italic=True)

        total_cell = ws3.cell(row=r_idx, column=16)
        total_cell.value = round(row_total, 2)
        total_cell.number_format = '#,##0.00'
        total_cell.font = hfont(C_NAVY, bold=True, sz=9)

    # Totals row
    total_r = 3 + len(pivot)
    ws3.row_dimensions[total_r].height=20
    ws3.cell(row=total_r, column=1).value = "PORTFOLIO TOTAL"
    ws3.cell(row=total_r, column=1).font = hfont(C_NAVY, True, 10)
    ws3.cell(row=total_r, column=1).fill = hfill("D6EAF8")
    grand = 0
    for m in range(1, 13):
        col_sum = monthly_df[monthly_df["Month_Number"]==m]["Amount_Cr"].sum()
        grand += col_sum
        c = ws3.cell(row=total_r, column=3+m)
        c.value = round(col_sum, 2); c.number_format='#,##0.00'
        c.fill=hfill("D6EAF8"); c.font=hfont(C_NAVY, True, 9); c.border=bdr()
    gt = ws3.cell(row=total_r, column=16)
    gt.value=round(grand,2); gt.number_format='#,##0.00'
    gt.fill=hfill(C_NAVY); gt.font=hfont(C_GOLD, True, 10); gt.border=bdr()

    # ── SHEET 4: LAPSE RISK ALERTS ────────────────────────────────────────────
    ws4 = wb.create_sheet("🚨 Lapse Risk Alerts")
    ws4.sheet_view.showGridLines = False

    at_risk = scheme_df[scheme_df["Lapse_Risk"].str.contains("HIGH|MODERATE")].sort_values(
        "Lapsing_Amount_Cr", ascending=False)

    ws4.row_dimensions[1].height=30
    ws4.merge_cells(f"A1:{get_column_letter(8)}1")
    t4=ws4["A1"]
    t4.value = (f"⚠  FUND LAPSE RISK — {len(at_risk)} SCHEMES AT RISK  |  "
                f"₹ {at_risk['Lapsing_Amount_Cr'].sum():,.2f} Cr projected to lapse")
    t4.fill=hfill(C_RED); t4.font=Font(color=C_WHITE,bold=True,size=13,name="Calibri")
    t4.alignment=ctr()

    cols4=["Scheme Name","Type","FY Allotment","Actual to Date","Velocity/Month",
           "Projected Year-End","Projected Util %","Lapsing Amount","Risk Level",
           f"Extra Spend Needed/Month to Reach 90%"]
    widths4=[35,10,16,16,16,20,16,16,18,35]
    for i,w in enumerate(widths4,1): ws4.column_dimensions[get_column_letter(i)].width=w
    hdr_row(ws4, 2, cols4)

    for r_idx, (_, row) in enumerate(at_risk.iterrows(), start=3):
        ws4.row_dimensions[r_idx].height=18
        style_row(ws4, r_idx, len(cols4))
        target_spend   = row["FY_Allotment_Cr"] * 0.90
        needed_extra   = max(0, target_spend - row["Actual_to_Date_Cr"])
        extra_per_month= round(needed_extra / max(MONTHS_REMAINING, 1), 2)
        vals = [
            row["Scheme_Name"], row["Funding_Type"],
            round(row["FY_Allotment_Cr"],2), round(row["Actual_to_Date_Cr"],2),
            round(row["Monthly_Velocity_Cr"],2), round(row["Projected_Total_Cr"],2),
            f"{row['Projected_Utilization_Pct']:.1f}%",
            round(row["Lapsing_Amount_Cr"],2), row["Lapse_Risk"],
            f"₹ {extra_per_month:,.2f} Cr/month",
        ]
        risk_bg = "FADBD8" if "HIGH" in str(row["Lapse_Risk"]) else "FDEBD0"
        for c_idx, val in enumerate(vals, start=1):
            cell=ws4.cell(row=r_idx, column=c_idx, value=val)
            cell.fill=hfill(risk_bg); cell.border=bdr(); cell.alignment=lft()
            risk_color = C_RED if "HIGH" in str(row["Lapse_Risk"]) else C_AMBER
            cell.font=Font(color=risk_color, size=9, name="Calibri",
                           bold=(c_idx in (8,9)))
            if c_idx in (3,4,5,6,8): cell.number_format='#,##0.00'

    # ── SHEET 5: CHART DATA + CHARTS ──────────────────────────────────────────
    ws5 = wb.create_sheet("📈 Charts")
    ws5.sheet_view.showGridLines = False
    ws5.column_dimensions["A"].width=2

    ws5.row_dimensions[1].height=28
    ws5.merge_cells("B1:P1")
    t5=ws5["B1"]; t5.value="BUDGET VELOCITY — VISUAL SUMMARY"
    t5.fill=hfill(C_NAVY); t5.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
    t5.alignment=ctr()

    # Write data for chart 1: Projected utilization by scheme
    ws5.cell(row=3, column=18, value="Scheme"); ws5.cell(row=3, column=19, value="Proj Util %")
    ws5.cell(row=3, column=20, value="Lapsing Cr")
    for i, (_, row) in enumerate(scheme_df.iterrows(), start=4):
        ws5.cell(row=i, column=18, value=row["Scheme_Name"][:20])
        ws5.cell(row=i, column=19, value=row["Projected_Utilization_Pct"])
        ws5.cell(row=i, column=20, value=round(row["Lapsing_Amount_Cr"],2))

    n = len(scheme_df) + 3

    chart1 = BarChart()
    chart1.type="bar"; chart1.title="Projected Year-End Utilization % by Scheme"
    chart1.y_axis.title="Scheme"; chart1.x_axis.title="Utilization %"
    chart1.width=18; chart1.height=12
    d1=Reference(ws5, min_col=19, max_col=19, min_row=3, max_row=n)
    c1=Reference(ws5, min_col=18, min_row=4, max_row=n)
    chart1.add_data(d1, titles_from_data=True); chart1.set_categories(c1)
    chart1.series[0].graphicalProperties.solidFill = "C9A84C"
    ws5.add_chart(chart1, "B3")

    chart2 = BarChart()
    chart2.type="bar"; chart2.title="Projected Lapsing Amount (₹ Cr) by Scheme"
    chart2.y_axis.title="Scheme"; chart2.x_axis.title="₹ Crore"
    chart2.width=18; chart2.height=12
    d2=Reference(ws5, min_col=20, max_col=20, min_row=3, max_row=n)
    chart2.add_data(d2, titles_from_data=True); chart2.set_categories(c1)
    chart2.series[0].graphicalProperties.solidFill = "C0392B"
    ws5.add_chart(chart2, "B24")

    # Portfolio monthly trend
    ws5.cell(row=3, column=22, value="Month")
    ws5.cell(row=3, column=23, value="Portfolio Spend (₹Cr)")
    monthly_totals = monthly_df.groupby("Month_Number")["Amount_Cr"].sum().reset_index()
    for _, mrow in monthly_totals.iterrows():
        m = int(mrow["Month_Number"])
        ws5.cell(row=3+m, column=22, value=MONTH_NAMES[m-1])
        ws5.cell(row=3+m, column=23, value=round(mrow["Amount_Cr"],2))

    chart3 = LineChart()
    chart3.title="Portfolio Monthly Spend Trend (Actuals + Projections)"
    chart3.y_axis.title="₹ Crore"; chart3.x_axis.title="Month"
    chart3.width=22; chart3.height=12
    d3=Reference(ws5, min_col=23, max_col=23, min_row=3, max_row=15)
    c3=Reference(ws5, min_col=22, min_row=4, max_row=15)
    chart3.add_data(d3, titles_from_data=True); chart3.set_categories(c3)
    ws5.add_chart(chart3, "K3")

    return wb


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  APCRDA PROJECT 3 — BUDGET VELOCITY MODEL")
    print("=" * 60)

    print(f"\n  FY: {FY_START.strftime('%d %b %Y')} → {FY_END.strftime('%d %b %Y')}")
    print(f"  Months elapsed: {MONTHS_ELAPSED}  |  Remaining: {MONTHS_REMAINING}")

    print("\n[1/4] Generating budget data...")
    scheme_df, monthly_df = generate_budget_data()

    print("[2/4] Computing portfolio KPIs...")
    kpis = compute_portfolio_kpis(scheme_df)

    print("\n── KPI SNAPSHOT ─────────────────────────────────────────")
    print(f"  FY Allotment (Total)     : ₹ {kpis['total_allotment']:>10,.2f} Cr")
    print(f"  Actual Spend to Date     : ₹ {kpis['total_actual']:>10,.2f} Cr  ({kpis['actual_utilization']}%)")
    print(f"  Projected Year-End       : ₹ {kpis['projected_total']:>10,.2f} Cr  ({kpis['projected_util']}%)")
    print(f"  At-Risk Lapsing Amount   : ₹ {kpis['total_lapsing']:>10,.2f} Cr")
    print(f"  High Lapse Risk Schemes  : {kpis['high_risk_count']}")
    print(f"  Moderate Risk Schemes    : {kpis['moderate_risk_count']}")
    print("─" * 60)
    print("\n  SCHEME RISK SUMMARY:")
    for _, row in scheme_df[["Scheme_Name","Projected_Utilization_Pct","Lapsing_Amount_Cr","Lapse_Risk"]].iterrows():
        risk_flag = "🔴" if "HIGH" in row["Lapse_Risk"] else ("🟡" if "MODERATE" in row["Lapse_Risk"] else "🟢")
        print(f"  {risk_flag}  {row['Scheme_Name'][:38]:38}  {row['Projected_Utilization_Pct']:>5.1f}%  "
              f"| Lapse: ₹{row['Lapsing_Amount_Cr']:>7.2f} Cr")
    print("─" * 60)

    print("\n[3/4] Building Excel workbook...")
    wb = build_excel(scheme_df, monthly_df, kpis)

    out = "/mnt/user-data/outputs/Budget_Velocity_Model.xlsx"
    wb.save(out)
    print(f"[4/4] Saved → {out}")

    import os
    print(f"\n  File size : {os.path.getsize(out)/1024:.1f} KB")
    print("\n  SHEETS:")
    print("  📊 Portfolio KPIs      → commissioner-level summary + action items")
    print("  📋 Scheme Velocity     → every scheme with velocity + lapse projection")
    print("  📅 Monthly Drill-Down  → 12-month pivot (purple = projected)")
    print("  🚨 Lapse Risk Alerts   → only at-risk schemes + extra spend needed/month")
    print("  📈 Charts              → utilization bars + monthly spend trend line")
    print("════════════════════════════════════════════════════════════\n")


if __name__ == "__main__":
    main()
