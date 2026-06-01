"""
APCRDA PROJECT 4: PPP Contingent Liability Register
====================================================
Purpose : Map all PPP commitments — VGF tranches, government guarantees,
          annuity obligations, revenue share receivables, and contingent
          liabilities not yet on the financial statements.

Run     : python project4_ppp_liability.py
Output  : PPP_Contingent_Liability_Register.xlsx

OFFLINE — replace generate_ppp_data() with actual concession agreement data.
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

random.seed(21)
np.random.seed(21)

TODAY    = datetime.today()
FY_END   = datetime(TODAY.year if TODAY.month >= 4 else TODAY.year - 1, 3, 31) + timedelta(days=365)

# ─────────────────────────────────────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────────────────────────────────────
C_NAVY="1B2A4A"; C_GOLD="C9A84C"; C_RED="C0392B"; C_AMBER="E67E22"
C_GREEN="1E8449"; C_WHITE="FFFFFF"; C_ALT="EAF0FB"; C_BORDER="BDC3C7"
C_PURPLE="7D3C98"

def hfill(h): return PatternFill("solid", fgColor=h)
def hfont(h="000000", bold=False, sz=10): return Font(color=h, bold=bold, size=sz, name="Calibri")
def bdr():
    s=Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)
def ctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)
def hdr_row(ws, row, cols, col_start=1):
    for i,label in enumerate(cols):
        c=ws.cell(row=row, column=col_start+i, value=label)
        c.fill=hfill(C_NAVY); c.font=hfont(C_WHITE,True,10)
        c.alignment=ctr(); c.border=bdr()
def style_row(ws, row, ncols, col_start=1, alt=False):
    bg=C_ALT if alt else C_WHITE
    for col in range(col_start, col_start+ncols):
        c=ws.cell(row=row, column=col)
        c.fill=hfill(bg); c.font=hfont(sz=9); c.alignment=lft(); c.border=bdr()


# ─────────────────────────────────────────────────────────────────────────────
# DATA GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

PPP_PROJECTS = [
    # (Name, Sector, Model, Project Cost Cr, Private Equity Pct, Concession Yrs)
    ("Amaravati Convention Centre",       "Commercial",     "BOT-Annuity",       320.0, 0.35, 30),
    ("Capital City Bus Rapid Transit",    "Transport",      "BOT-Toll",          485.0, 0.40, 25),
    ("Amaravati Smart Parking System",    "Urban Mobility", "OMT",                58.0, 0.50, 15),
    ("Seed Access Road Package-I",        "Roads",          "HAM",               780.0, 0.40, 15),
    ("Seed Access Road Package-II",       "Roads",          "HAM",               640.0, 0.40, 15),
    ("Amaravati Solar Park",              "Energy",         "BOOT",              195.0, 0.60, 25),
    ("Capital Region Solid Waste Plant",  "Environment",    "BOO",               145.0, 0.55, 20),
    ("IT/ITES SEZ - Phase I",             "Commercial",     "DBFOT",             560.0, 0.70, 30),
    ("Amaravati Medical Hub",             "Healthcare",     "BOT-Annuity",       420.0, 0.45, 25),
    ("Greenfield Airport Connectivity",   "Transport",      "BOT-Toll",         1200.0, 0.50, 30),
]

PPP_MODELS = {
    "BOT-Annuity" : "Govt pays fixed annuity to private partner regardless of revenue",
    "BOT-Toll"    : "Private partner collects toll; govt provides VGF if needed",
    "HAM"         : "Hybrid Annuity — 40% govt upfront, 60% as annuity over concession",
    "OMT"         : "Operate-Maintain-Transfer — private operates existing asset",
    "BOOT"        : "Build-Own-Operate-Transfer — full private ownership during concession",
    "BOO"         : "Build-Own-Operate — no transfer at end",
    "DBFOT"       : "Design-Build-Finance-Operate-Transfer — comprehensive private role",
}


def generate_ppp_data():
    """
    Returns:
      project_df   : one row per PPP project
      vgf_df       : VGF tranche schedule per project
      annuity_df   : annual annuity obligation schedule
      guarantee_df : government guarantees and contingent liabilities
      revshare_df  : revenue share receivables from private partners
    """
    projects = []
    vgf_rows = []
    annuity_rows = []
    guarantee_rows = []
    revshare_rows = []

    for p_idx, (name, sector, model, cost, pvt_pct, concession_yrs) in enumerate(PPP_PROJECTS):
        pid = f"PPP-APCRDA-{str(p_idx+1).zfill(3)}"

        # Financial structure
        pvt_equity_cr    = round(cost * pvt_pct, 2)
        debt_cr          = round(cost * (1 - pvt_pct) * 0.70, 2)
        vgf_eligible     = model in ("BOT-Toll", "BOT-Annuity", "DBFOT", "HAM")
        vgf_amount       = round(cost * random.uniform(0.10, 0.35), 2) if vgf_eligible else 0
        govt_equity      = round(cost - pvt_equity_cr - debt_cr - vgf_amount, 2)
        govt_equity      = max(0, govt_equity)

        # HAM: 40% upfront from government
        if model == "HAM":
            ham_upfront  = round(cost * 0.40, 2)
            annuity_base = round(cost * 0.60 / concession_yrs, 2)
        else:
            ham_upfront  = 0
            annuity_base = round(vgf_amount / concession_yrs, 2) if model in ("BOT-Annuity",) else 0

        # Dates
        cod_date  = TODAY - timedelta(days=random.randint(30, 900))  # Commercial Operation Date
        conc_end  = cod_date + timedelta(days=concession_yrs * 365)
        sign_date = cod_date - timedelta(days=random.randint(180, 730))

        # Disclosed on FS?
        disclosed_on_fs = random.choice([True, True, False])
        if not disclosed_on_fs:
            disclosure_gap = f"Contingent liability of ₹{vgf_amount+annuity_base*concession_yrs:,.2f} Cr NOT in notes to accounts"
        else:
            disclosure_gap = "Disclosed in Schedule III notes"

        projects.append({
            "Project_ID"             : pid,
            "Project_Name"           : name,
            "Sector"                 : sector,
            "PPP_Model"              : model,
            "Model_Description"      : PPP_MODELS[model],
            "Total_Project_Cost_Cr"  : cost,
            "Private_Equity_Cr"      : pvt_equity_cr,
            "Debt_Financing_Cr"      : debt_cr,
            "Govt_Equity_Cr"         : govt_equity,
            "VGF_Amount_Cr"          : vgf_amount,
            "HAM_Upfront_Cr"         : ham_upfront,
            "Annual_Annuity_Cr"      : annuity_base,
            "Concession_Period_Yrs"  : concession_yrs,
            "Agreement_Sign_Date"    : sign_date,
            "COD_Date"               : cod_date,
            "Concession_End_Date"    : conc_end,
            "Private_Partner"        : random.choice([
                "GMR Amaravati Infra Ltd", "L&T Infrastructure Dev Projects",
                "GVK Power & Infra Ltd",   "Adani Road Transport Ltd",
                "IL&FS Transportation",    "Sterlite Power Ltd",
                "Cube Highways Fund",      "IRB Infrastructure",
            ]),
            "Govt_Guarantee_Issued"  : random.choice([True, False]),
            "Guarantee_Amount_Cr"    : round(debt_cr * random.uniform(0.5, 1.0), 2),
            "Disclosed_on_FS"        : disclosed_on_fs,
            "Disclosure_Status"      : disclosure_gap,
            "Total_Govt_Obligation_Cr": round(vgf_amount + ham_upfront + annuity_base * concession_yrs + govt_equity, 2),
        })

        # VGF Tranche Schedule
        if vgf_amount > 0:
            num_vgf_tranches = random.randint(2, 4)
            per_tranche = round(vgf_amount / num_vgf_tranches, 2)
            for vt in range(num_vgf_tranches):
                due_date  = sign_date + timedelta(days=(vt+1) * random.randint(90, 180))
                paid_date = due_date + timedelta(days=random.randint(0, 60)) if due_date < TODAY else None
                status    = "Paid" if paid_date else ("Overdue" if due_date < TODAY else "Upcoming")
                vgf_rows.append({
                    "Project_ID"   : pid,
                    "Project_Name" : name,
                    "Tranche"      : f"VGF-T{vt+1}",
                    "Amount_Cr"    : per_tranche,
                    "Due_Date"     : due_date,
                    "Paid_Date"    : paid_date,
                    "Status"       : status,
                    "Condition"    : f"Private equity ₹{round(pvt_equity_cr*(vt+1)/num_vgf_tranches,2)}Cr injected and certified",
                    "VGF_Balance_Cr": round(vgf_amount - (vt * per_tranche), 2),
                })

        # Annuity Schedule (next 5 years only for brevity)
        if annuity_base > 0:
            for yr in range(1, 6):
                ann_date  = cod_date + timedelta(days=yr * 365)
                ann_paid  = ann_date < TODAY
                annuity_rows.append({
                    "Project_ID"   : pid,
                    "Project_Name" : name,
                    "Sector"       : sector,
                    "PPP_Model"    : model,
                    "Year"         : yr,
                    "Due_Date"     : ann_date,
                    "Amount_Cr"    : annuity_base,
                    "Status"       : "Paid" if ann_paid else ("Overdue" if ann_date < TODAY else "Upcoming"),
                    "Cumulative_Obligation_Cr": round(annuity_base * concession_yrs, 2),
                })

        # Government Guarantees
        if projects[-1]["Govt_Guarantee_Issued"]:
            guarantee_rows.append({
                "Project_ID"          : pid,
                "Project_Name"        : name,
                "Sector"              : sector,
                "Guarantee_Type"      : random.choice(["Debt Service Guarantee", "Revenue Guarantee", "Equity Support Letter"]),
                "Guarantee_Amount_Cr" : projects[-1]["Guarantee_Amount_Cr"],
                "Issued_To"           : random.choice(["SBI Infrastructure Finance", "IIFCL", "Axis Bank", "HDFC Infrastructure"]),
                "Issue_Date"          : sign_date,
                "Expiry_Date"         : conc_end,
                "Invocation_Risk"     : random.choice(["LOW", "LOW", "MEDIUM", "HIGH"]),
                "Disclosed_on_FS"     : projects[-1]["Disclosed_on_FS"],
                "CAG_Flag"            : "YES — undisclosed guarantee" if not projects[-1]["Disclosed_on_FS"] else "NO",
            })

        # Revenue Share (only for revenue-generating models)
        if model in ("BOT-Toll", "OMT", "BOOT", "DBFOT", "BOO"):
            rev_share_pct  = random.uniform(0.05, 0.20)
            est_revenue_cr = round(cost * random.uniform(0.03, 0.08), 2)
            revshare_rows.append({
                "Project_ID"           : pid,
                "Project_Name"         : name,
                "Revenue_Share_Pct"    : round(rev_share_pct * 100, 1),
                "FY_Revenue_Estimated_Cr": est_revenue_cr,
                "APCRDA_Share_Due_Cr"  : round(est_revenue_cr * rev_share_pct, 2),
                "Amount_Received_Cr"   : round(est_revenue_cr * rev_share_pct * random.uniform(0, 1), 2),
                "Receivable_Cr"        : 0,  # computed below
                "Last_Audit_Date"      : TODAY - timedelta(days=random.randint(30, 365)),
                "Audit_Status"         : random.choice(["Verified", "Verified", "Pending Verification", "Disputed"]),
            })
            revshare_rows[-1]["Receivable_Cr"] = round(
                revshare_rows[-1]["APCRDA_Share_Due_Cr"] - revshare_rows[-1]["Amount_Received_Cr"], 2)

    return (
        pd.DataFrame(projects),
        pd.DataFrame(vgf_rows),
        pd.DataFrame(annuity_rows) if annuity_rows else pd.DataFrame(),
        pd.DataFrame(guarantee_rows) if guarantee_rows else pd.DataFrame(),
        pd.DataFrame(revshare_rows) if revshare_rows else pd.DataFrame(),
    )


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
def compute_ppp_kpis(projects, vgf, annuity, guarantee, revshare):
    total_project_cost = projects["Total_Project_Cost_Cr"].sum()
    total_govt_obligation = projects["Total_Govt_Obligation_Cr"].sum()
    total_vgf = projects["VGF_Amount_Cr"].sum()
    total_annuity_5yr = annuity["Amount_Cr"].sum() if not annuity.empty else 0
    total_guarantee = guarantee["Guarantee_Amount_Cr"].sum() if not guarantee.empty else 0
    undisclosed = projects[~projects["Disclosed_on_FS"]]
    undisclosed_liability = undisclosed["Total_Govt_Obligation_Cr"].sum()
    total_receivable = revshare["Receivable_Cr"].sum() if not revshare.empty else 0
    vgf_overdue = vgf[vgf["Status"] == "Overdue"] if not vgf.empty else pd.DataFrame()

    return {
        "total_project_cost"    : round(total_project_cost, 2),
        "total_govt_obligation" : round(total_govt_obligation, 2),
        "total_vgf"             : round(total_vgf, 2),
        "total_annuity_5yr"     : round(total_annuity_5yr, 2),
        "total_guarantee"       : round(total_guarantee, 2),
        "undisclosed_count"     : len(undisclosed),
        "undisclosed_liability" : round(undisclosed_liability, 2),
        "total_receivable"      : round(total_receivable, 2),
        "vgf_overdue_count"     : len(vgf_overdue),
        "vgf_overdue_amount"    : round(vgf_overdue["Amount_Cr"].sum(), 2) if not vgf_overdue.empty else 0,
        "high_inv_risk"         : len(guarantee[guarantee["Invocation_Risk"]=="HIGH"]) if not guarantee.empty else 0,
    }


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(projects, vgf, annuity, guarantee, revshare, kpis):
    wb = Workbook()

    # ── SHEET 1: KPI SUMMARY ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "📊 KPI Summary"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width=4
    ws.column_dimensions["B"].width=42
    ws.column_dimensions["C"].width=24

    ws.merge_cells("B2:C2")
    t=ws["B2"]; t.value="PPP CONTINGENT LIABILITY REGISTER — APCRDA"
    t.fill=hfill(C_NAVY); t.font=Font(color=C_GOLD,bold=True,size=14,name="Calibri")
    t.alignment=ctr(); ws.row_dimensions[2].height=30

    ws.merge_cells("B3:C3")
    ws["B3"].value=f"As of {TODAY.strftime('%d %B %Y')}  |  {len(projects)} Active PPP Projects"
    ws["B3"].font=hfont(C_NAVY,sz=10); ws["B3"].alignment=ctr()

    card_data = [
        ("Total PPP Portfolio Value",     f"₹ {kpis['total_project_cost']:,.2f} Cr",     C_NAVY,  "Aggregate cost across all PPP projects"),
        ("Total Govt Financial Obligation",f"₹ {kpis['total_govt_obligation']:,.2f} Cr", C_RED,   "VGF + Annuity + Guarantees + Equity"),
        ("Total VGF Committed",           f"₹ {kpis['total_vgf']:,.2f} Cr",              "1A5276", "Viability Gap Funding sanctioned"),
        ("5-Year Annuity Obligation",     f"₹ {kpis['total_annuity_5yr']:,.2f} Cr",      C_AMBER, "Fixed govt payments to private partners"),
        ("Contingent Guarantees",         f"₹ {kpis['total_guarantee']:,.2f} Cr",         C_RED,   f"{kpis['high_inv_risk']} guarantees at HIGH invocation risk"),
        ("UNDISCLOSED Liabilities",       f"₹ {kpis['undisclosed_liability']:,.2f} Cr",   C_RED,   f"{kpis['undisclosed_count']} projects not in FS notes — CAG risk"),
        ("Revenue Share Receivable",      f"₹ {kpis['total_receivable']:,.2f} Cr",        C_AMBER, "Uncollected revenue share from private partners"),
        ("Overdue VGF Tranches",          f"{kpis['vgf_overdue_count']} tranches",         C_RED,   f"₹ {kpis['vgf_overdue_amount']:,.2f} Cr — private partner equity not yet certified"),
    ]

    r=5
    for label, value, color, note in card_data:
        for row in [r, r+1, r+2]: ws.row_dimensions[row].height=14 if row!=r+1 else 26
        for row in [r, r+1, r+2]:
            ws.merge_cells(f"B{row}:C{row}"); ws[f"B{row}"].fill=hfill(color)
        ws[f"B{r}"].value=label.upper()
        ws[f"B{r}"].font=Font(color=C_WHITE,size=8,name="Calibri"); ws[f"B{r}"].alignment=ctr()
        ws[f"B{r+1}"].value=value
        ws[f"B{r+1}"].font=Font(color=C_WHITE,bold=True,size=15,name="Calibri"); ws[f"B{r+1}"].alignment=ctr()
        ws[f"B{r+2}"].value=note
        ws[f"B{r+2}"].font=Font(color="DDDDDD",size=8,name="Calibri"); ws[f"B{r+2}"].alignment=ctr()
        r+=4

    # Critical audit note
    r+=1
    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"].value="⚠  CRITICAL AUDIT OBSERVATIONS"
    ws[f"B{r}"].fill=hfill(C_RED); ws[f"B{r}"].font=hfont(C_WHITE,True,11)
    ws[f"B{r}"].alignment=ctr(); ws.row_dimensions[r].height=22; r+=1
    obs=[
        f"• {kpis['undisclosed_count']} PPP projects have undisclosed contingent liabilities — ₹{kpis['undisclosed_liability']:,.2f} Cr not in notes to accounts",
        f"• {kpis['high_inv_risk']} govt guarantees at HIGH invocation risk — not provisioned in financial statements",
        f"• ₹{kpis['total_receivable']:,.2f} Cr revenue share from private partners is outstanding — potential write-off risk",
        f"• {kpis['vgf_overdue_count']} VGF tranches overdue — private partner equity certification pending; releases blocked",
        "• Recommended: CAG-compliant contingent liability disclosure in next annual accounts",
    ]
    for ob in obs:
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"].value=ob; ws[f"B{r}"].fill=hfill("FADBD8")
        ws[f"B{r}"].font=Font(color=C_RED,size=9,name="Calibri"); ws[f"B{r}"].alignment=lft()
        ws.row_dimensions[r].height=16; r+=1

    # ── SHEET 2: PROJECT MASTER ────────────────────────────────────────────────
    ws2=wb.create_sheet("📁 PPP Project Master")
    ws2.sheet_view.showGridLines=False; ws2.freeze_panes="A3"

    cols2=["Project ID","Project Name","Sector","PPP Model","Model Description",
           "Project Cost (₹Cr)","Private Equity (₹Cr)","Debt (₹Cr)",
           "Govt Equity (₹Cr)","VGF Amount (₹Cr)","HAM Upfront (₹Cr)",
           "Annual Annuity (₹Cr)","Concession Yrs","Private Partner",
           "Agreement Date","COD Date","Concession End",
           "Govt Guarantee?","Total Govt Obligation (₹Cr)","Disclosed on FS?","Disclosure Status"]
    widths2=[14,35,16,16,45,18,18,14,14,16,14,18,14,30,16,16,16,16,24,16,40]
    for i,w in enumerate(widths2,1): ws2.column_dimensions[get_column_letter(i)].width=w

    ws2.row_dimensions[1].height=28
    ws2.merge_cells(f"A1:{get_column_letter(len(cols2))}1")
    t2=ws2["A1"]; t2.value="PPP PROJECT MASTER REGISTER"
    t2.fill=hfill(C_NAVY); t2.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
    t2.alignment=ctr()
    hdr_row(ws2,2,cols2)

    field_map2=["Project_ID","Project_Name","Sector","PPP_Model","Model_Description",
                "Total_Project_Cost_Cr","Private_Equity_Cr","Debt_Financing_Cr",
                "Govt_Equity_Cr","VGF_Amount_Cr","HAM_Upfront_Cr",
                "Annual_Annuity_Cr","Concession_Period_Yrs","Private_Partner",
                "Agreement_Sign_Date","COD_Date","Concession_End_Date",
                "Govt_Guarantee_Issued","Total_Govt_Obligation_Cr","Disclosed_on_FS","Disclosure_Status"]
    money_cols2={"Total_Project_Cost_Cr","Private_Equity_Cr","Debt_Financing_Cr",
                 "Govt_Equity_Cr","VGF_Amount_Cr","HAM_Upfront_Cr",
                 "Annual_Annuity_Cr","Total_Govt_Obligation_Cr"}

    for r_idx,(_, row) in enumerate(projects.iterrows(), start=3):
        ws2.row_dimensions[r_idx].height=18
        style_row(ws2, r_idx, len(cols2), alt=r_idx%2==0)
        for c_idx, fld in enumerate(field_map2, start=1):
            val=row[fld]; cell=ws2.cell(row=r_idx, column=c_idx)
            if isinstance(val, datetime): cell.value=val.strftime("%d-%b-%Y")
            elif fld in money_cols2:
                cell.value=round(float(val),2); cell.number_format='#,##0.00'
            elif fld=="Disclosed_on_FS":
                cell.value="YES" if val else "NO — RISK"
                if not val:
                    cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                else:
                    cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
            elif fld=="Govt_Guarantee_Issued":
                cell.value="YES" if val else "NO"
                if val: cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,bold=True,size=9,name="Calibri")
            else: cell.value=val

    # ── SHEET 3: VGF TRANCHE SCHEDULE ─────────────────────────────────────────
    ws3=wb.create_sheet("💰 VGF Schedule")
    ws3.sheet_view.showGridLines=False; ws3.freeze_panes="A3"

    cols3=["Project ID","Project Name","Tranche","Amount (₹Cr)","Due Date",
           "Paid Date","Status","VGF Balance (₹Cr)","Condition for Release"]
    widths3=[14,35,10,14,14,14,12,16,55]
    for i,w in enumerate(widths3,1): ws3.column_dimensions[get_column_letter(i)].width=w

    ws3.row_dimensions[1].height=28
    ws3.merge_cells(f"A1:{get_column_letter(len(cols3))}1")
    t3=ws3["A1"]; t3.value="VGF TRANCHE DISBURSEMENT SCHEDULE"
    t3.fill=hfill(C_NAVY); t3.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
    t3.alignment=ctr()
    hdr_row(ws3,2,cols3)

    field_map3=["Project_ID","Project_Name","Tranche","Amount_Cr","Due_Date",
                "Paid_Date","Status","VGF_Balance_Cr","Condition"]
    for r_idx,(_, row) in enumerate(vgf.iterrows(), start=3):
        ws3.row_dimensions[r_idx].height=18
        style_row(ws3, r_idx, len(cols3), alt=r_idx%2==0)
        for c_idx, fld in enumerate(field_map3, start=1):
            val=row[fld]; cell=ws3.cell(row=r_idx, column=c_idx)
            if pd.isna(val) or val is None: cell.value="—"
            elif isinstance(val, datetime): cell.value=val.strftime("%d-%b-%Y")
            elif fld in ("Amount_Cr","VGF_Balance_Cr"):
                cell.value=round(float(val),2); cell.number_format='#,##0.00'
            else: cell.value=val
            if fld=="Status":
                if val=="Paid": cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
                elif val=="Overdue": cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                elif val=="Upcoming": cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")

    # ── SHEET 4: ANNUITY SCHEDULE ──────────────────────────────────────────────
    if not annuity.empty:
        ws4=wb.create_sheet("📅 Annuity Schedule")
        ws4.sheet_view.showGridLines=False; ws4.freeze_panes="A3"

        cols4=["Project ID","Project Name","Sector","Model","Year","Due Date",
               "Annual Amount (₹Cr)","Status","Cumulative Obligation (₹Cr)"]
        widths4=[14,35,16,16,8,14,20,12,24]
        for i,w in enumerate(widths4,1): ws4.column_dimensions[get_column_letter(i)].width=w

        ws4.row_dimensions[1].height=28
        ws4.merge_cells(f"A1:{get_column_letter(len(cols4))}1")
        t4=ws4["A1"]; t4.value="ANNUITY PAYMENT OBLIGATION SCHEDULE (5-YEAR)"
        t4.fill=hfill(C_NAVY); t4.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
        t4.alignment=ctr()
        hdr_row(ws4,2,cols4)

        field_map4=["Project_ID","Project_Name","Sector","PPP_Model","Year","Due_Date",
                    "Amount_Cr","Status","Cumulative_Obligation_Cr"]
        for r_idx,(_, row) in enumerate(annuity.iterrows(), start=3):
            ws4.row_dimensions[r_idx].height=18
            style_row(ws4, r_idx, len(cols4), alt=r_idx%2==0)
            for c_idx, fld in enumerate(field_map4, start=1):
                val=row[fld]; cell=ws4.cell(row=r_idx, column=c_idx)
                if isinstance(val, datetime): cell.value=val.strftime("%d-%b-%Y")
                elif fld in ("Amount_Cr","Cumulative_Obligation_Cr"):
                    cell.value=round(float(val),2); cell.number_format='#,##0.00'
                else: cell.value=val
                if fld=="Status":
                    if val=="Paid": cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
                    elif val=="Overdue": cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                    else: cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")

    # ── SHEET 5: GUARANTEE REGISTER ───────────────────────────────────────────
    if not guarantee.empty:
        ws5=wb.create_sheet("🔒 Guarantee Register")
        ws5.sheet_view.showGridLines=False; ws5.freeze_panes="A3"

        cols5=["Project ID","Project Name","Sector","Guarantee Type","Amount (₹Cr)",
               "Issued To","Issue Date","Expiry Date","Invocation Risk","Disclosed?","CAG Flag"]
        widths5=[14,35,16,28,16,28,14,14,18,12,35]
        for i,w in enumerate(widths5,1): ws5.column_dimensions[get_column_letter(i)].width=w

        ws5.row_dimensions[1].height=28
        ws5.merge_cells(f"A1:{get_column_letter(len(cols5))}1")
        t5=ws5["A1"]; t5.value="GOVERNMENT GUARANTEE REGISTER — CONTINGENT LIABILITIES"
        t5.fill=hfill(C_NAVY); t5.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
        t5.alignment=ctr()
        hdr_row(ws5,2,cols5)

        field_map5=["Project_ID","Project_Name","Sector","Guarantee_Type","Guarantee_Amount_Cr",
                    "Issued_To","Issue_Date","Expiry_Date","Invocation_Risk","Disclosed_on_FS","CAG_Flag"]
        for r_idx,(_, row) in enumerate(guarantee.iterrows(), start=3):
            ws5.row_dimensions[r_idx].height=18
            style_row(ws5, r_idx, len(cols5), alt=r_idx%2==0)
            for c_idx, fld in enumerate(field_map5, start=1):
                val=row[fld]; cell=ws5.cell(row=r_idx, column=c_idx)
                if isinstance(val, datetime): cell.value=val.strftime("%d-%b-%Y")
                elif fld=="Guarantee_Amount_Cr":
                    cell.value=round(float(val),2); cell.number_format='#,##0.00'
                elif fld=="Disclosed_on_FS":
                    cell.value="YES" if val else "NO"
                    if not val: cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                    else: cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
                else: cell.value=val
                if fld=="Invocation_Risk":
                    if val=="HIGH": cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                    elif val=="MEDIUM": cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")
                    else: cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
                if fld=="CAG_Flag" and "YES" in str(val):
                    cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")

    # ── SHEET 6: REVENUE SHARE REGISTER ───────────────────────────────────────
    if not revshare.empty:
        ws6=wb.create_sheet("💹 Revenue Share")
        ws6.sheet_view.showGridLines=False; ws6.freeze_panes="A3"

        cols6=["Project ID","Project Name","Rev Share %","FY Revenue Est (₹Cr)",
               "APCRDA Share Due (₹Cr)","Amount Received (₹Cr)","Receivable (₹Cr)",
               "Last Audit Date","Audit Status"]
        widths6=[14,35,14,20,22,22,16,16,22]
        for i,w in enumerate(widths6,1): ws6.column_dimensions[get_column_letter(i)].width=w

        ws6.row_dimensions[1].height=28
        ws6.merge_cells(f"A1:{get_column_letter(len(cols6))}1")
        t6=ws6["A1"]; t6.value="REVENUE SHARE RECEIVABLES FROM PRIVATE PARTNERS"
        t6.fill=hfill(C_NAVY); t6.font=Font(color=C_GOLD,bold=True,size=13,name="Calibri")
        t6.alignment=ctr()
        hdr_row(ws6,2,cols6)

        field_map6=["Project_ID","Project_Name","Revenue_Share_Pct","FY_Revenue_Estimated_Cr",
                    "APCRDA_Share_Due_Cr","Amount_Received_Cr","Receivable_Cr",
                    "Last_Audit_Date","Audit_Status"]
        money6={"FY_Revenue_Estimated_Cr","APCRDA_Share_Due_Cr","Amount_Received_Cr","Receivable_Cr"}

        for r_idx,(_, row) in enumerate(revshare.iterrows(), start=3):
            ws6.row_dimensions[r_idx].height=18
            style_row(ws6, r_idx, len(cols6), alt=r_idx%2==0)
            for c_idx, fld in enumerate(field_map6, start=1):
                val=row[fld]; cell=ws6.cell(row=r_idx, column=c_idx)
                if isinstance(val, datetime): cell.value=val.strftime("%d-%b-%Y")
                elif fld=="Revenue_Share_Pct": cell.value=f"{float(val):.1f}%"
                elif fld in money6:
                    cell.value=round(float(val),2); cell.number_format='#,##0.00'
                    if fld=="Receivable_Cr" and float(val)>0:
                        cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,bold=True,size=9,name="Calibri")
                else: cell.value=val
                if fld=="Audit_Status":
                    if val=="Verified": cell.fill=hfill("D5F5E3"); cell.font=Font(color=C_GREEN,size=9,name="Calibri")
                    elif val=="Disputed": cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,bold=True,size=9,name="Calibri")
                    else: cell.fill=hfill("FDEBD0"); cell.font=Font(color=C_AMBER,size=9,name="Calibri")

    # ── SHEET 7: UNDISCLOSED LIABILITY ALERT ──────────────────────────────────
    ws7=wb.create_sheet("🚨 FS Disclosure Gaps")
    ws7.sheet_view.showGridLines=False

    undisclosed=projects[~projects["Disclosed_on_FS"]].copy()
    ws7.row_dimensions[1].height=30
    ws7.merge_cells(f"A1:{get_column_letter(6)}1")
    t7=ws7["A1"]
    t7.value=(f"⚠  FINANCIAL STATEMENT DISCLOSURE GAPS — "
              f"{len(undisclosed)} PROJECTS | ₹ {undisclosed['Total_Govt_Obligation_Cr'].sum():,.2f} Cr UNDISCLOSED")
    t7.fill=hfill(C_RED); t7.font=Font(color=C_WHITE,bold=True,size=13,name="Calibri")
    t7.alignment=ctr()

    cols7=["Project ID","Project Name","PPP Model","Total Govt Obligation (₹Cr)","Disclosure Status","Recommended Action"]
    widths7=[14,38,16,24,40,50]
    for i,w in enumerate(widths7,1): ws7.column_dimensions[get_column_letter(i)].width=w
    hdr_row(ws7,2,cols7)

    for r_idx,(_, row) in enumerate(undisclosed.iterrows(), start=3):
        ws7.row_dimensions[r_idx].height=22
        for c_idx, val in enumerate([
            row["Project_ID"], row["Project_Name"], row["PPP_Model"],
            round(row["Total_Govt_Obligation_Cr"],2),
            row["Disclosure_Status"],
            "Add Schedule III Note disclosing government obligation and contingent guarantee. Obtain legal sign-off before next AG inspection."
        ], start=1):
            cell=ws7.cell(row=r_idx, column=c_idx, value=val)
            cell.fill=hfill("FADBD8"); cell.font=Font(color=C_RED,size=9,name="Calibri")
            cell.border=bdr(); cell.alignment=lft()
            if c_idx==4: cell.number_format='#,##0.00'

    return wb


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  APCRDA PROJECT 4 — PPP CONTINGENT LIABILITY REGISTER")
    print("=" * 60)

    print("\n[1/4] Generating PPP project data...")
    projects, vgf, annuity, guarantee, revshare = generate_ppp_data()

    print("[2/4] Computing liability KPIs...")
    kpis = compute_ppp_kpis(projects, vgf, annuity, guarantee, revshare)

    print("\n── KPI SNAPSHOT ─────────────────────────────────────────")
    print(f"  Total PPP Portfolio       : ₹ {kpis['total_project_cost']:>10,.2f} Cr")
    print(f"  Total Govt Obligation     : ₹ {kpis['total_govt_obligation']:>10,.2f} Cr")
    print(f"  Total VGF Committed       : ₹ {kpis['total_vgf']:>10,.2f} Cr")
    print(f"  5-Year Annuity Obligation : ₹ {kpis['total_annuity_5yr']:>10,.2f} Cr")
    print(f"  Contingent Guarantees     : ₹ {kpis['total_guarantee']:>10,.2f} Cr")
    print(f"  UNDISCLOSED Liabilities   : ₹ {kpis['undisclosed_liability']:>10,.2f} Cr  ← CAG RISK")
    print(f"  Revenue Share Receivable  : ₹ {kpis['total_receivable']:>10,.2f} Cr")
    print(f"  VGF Tranches Overdue      : {kpis['vgf_overdue_count']}")
    print(f"  High Invocation Risk Guar : {kpis['high_inv_risk']}")
    print("─" * 60)

    print("\n[3/4] Building Excel workbook...")
    wb = build_excel(projects, vgf, annuity, guarantee, revshare, kpis)

    out="/mnt/user-data/outputs/PPP_Contingent_Liability_Register.xlsx"
    wb.save(out)
    print(f"[4/4] Saved → {out}")

    import os
    print(f"\n  File size : {os.path.getsize(out)/1024:.1f} KB")
    print("\n  SHEETS:")
    print("  📊 KPI Summary           → top-level liability exposure for Director")
    print("  📁 PPP Project Master    → all 10 projects, full financial structure")
    print("  💰 VGF Schedule          → tranche-wise VGF disbursement + status")
    print("  📅 Annuity Schedule      → 5-year annuity obligation calendar")
    print("  🔒 Guarantee Register    → all govt guarantees + invocation risk")
    print("  💹 Revenue Share         → receivables from private partners")
    print("  🚨 FS Disclosure Gaps    → undisclosed liabilities — present this first")
    print("════════════════════════════════════════════════════════════\n")


if __name__ == "__main__":
    main()
