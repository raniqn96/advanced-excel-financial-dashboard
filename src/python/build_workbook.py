"""
Automated workbook generator for the Excel Financial Dashboard & Forecasting Model.

Creates workbook/Financial_Dashboard_Model.xlsx with:
  - Data_Sales / Data_GL        : raw data from two source systems (ERP sales + General Ledger)
  - Lookups                     : product & region master tables (VLOOKUP / HLOOKUP sources)
  - Assumptions                 : forecast drivers (blue inputs)
  - Pivot_Summary               : pivot-style SUMIFS summary by region x month
  - Reconciliation              : ERP vs GL monthly reconciliation with tolerance checks
  - Budget_vs_Actual            : variance analysis with conditional formatting
  - Forecast                    : 12-month dynamic forecast driven by assumptions
  - Dashboard                   : KPI cards, interactive region selector, charts

Usage:
    pip install openpyxl
    python src/python/build_workbook.py [--rows 1200] [--seed 42] [--out workbook/Financial_Dashboard_Model.xlsx]
"""
import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment

FONT = "Arial"
NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "DDEBF7"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
REGIONS = [("R01", "North"), ("R02", "South"), ("R03", "East"), ("R04", "West")]
PRODUCTS = [
    ("P100", "Analytics Suite", "Software", 1200, 0.78),
    ("P200", "Cloud Storage", "Services", 450, 0.62),
    ("P300", "Consulting Hours", "Services", 180, 0.45),
    ("P400", "Hardware Kit", "Hardware", 950, 0.32),
    ("P500", "Support Plan", "Services", 300, 0.70),
    ("P600", "Training Pack", "Software", 250, 0.66),
]
YEAR = 2025

CUR = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_IN = Font(name=FONT, color="0000FF")
F_CALC = Font(name=FONT, color="000000")
F_LINK = Font(name=FONT, color="008000")
F_HDR = Font(name=FONT, bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", start_color=NAVY)
FILL_SUB = PatternFill("solid", start_color=LIGHT)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=FONT, italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def header(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def generate_data(rows, seed):
    rnd = random.Random(seed)
    sales = []
    start = date(YEAR, 1, 1)
    for i in range(rows):
        d = start + timedelta(days=rnd.randint(0, 364))
        p = rnd.choice(PRODUCTS)
        r = rnd.choice(REGIONS)
        qty = rnd.randint(1, 25)
        disc = rnd.choice([0, 0, 0.05, 0.1])
        sales.append([f"INV-{10000 + i}", d, r[0], p[0], qty, disc])
    sales.sort(key=lambda x: x[1])
    return sales


def build_gl(sales, seed):
    """GL postings per region/month; mostly matches ERP with a few deliberate breaks."""
    rnd = random.Random(seed + 1)
    price = {p[0]: p[3] for p in PRODUCTS}
    agg = {}
    for inv, d, reg, prod, qty, disc in sales:
        k = (reg, d.month)
        agg[k] = agg.get(k, 0) + qty * price[prod] * (1 - disc)
    gl = []
    for (reg, m), amt in sorted(agg.items()):
        noise = rnd.choice([0, 0, 0, 0, rnd.uniform(-0.04, 0.04)])
        gl.append([f"JE-{reg}-{m:02d}", date(YEAR, m, 28), reg, "4000-Revenue", round(amt * (1 + noise), 2)])
    return gl


def sheet_lookups(wb):
    ws = wb.create_sheet("Lookups")
    title(ws, "Master Data (Lookup Tables)", "Product master for VLOOKUP; region master for HLOOKUP")
    header(ws, 4, ["Product ID", "Product Name", "Category", "Unit Price ($)", "Gross Margin %"])
    for i, p in enumerate(PRODUCTS, start=5):
        for j, v in enumerate(p, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.border = F_IN, BORDER
        ws.cell(row=i, column=4).number_format = CUR
        ws.cell(row=i, column=5).number_format = PCT
    tab = Table(displayName="tblProducts", ref=f"A4:E{4 + len(PRODUCTS)}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tab)

    # Horizontal region table for HLOOKUP
    ws["G4"] = "Region ID"
    ws["G5"] = "Region Name"
    ws["G6"] = "Regional Manager"
    for c in ("G4", "G5", "G6"):
        ws[c].font, ws[c].fill, ws[c].border = F_HDR, FILL_HDR, BORDER
    mgrs = ["A. Sharma", "B. Patel", "C. Singh", "D. Rao"]
    for j, (rid, rname) in enumerate(REGIONS):
        col = 8 + j
        for r, v in zip((4, 5, 6), (rid, rname, mgrs[j])):
            c = ws.cell(row=r, column=col, value=v)
            c.font, c.border = F_IN, BORDER
    ws["G9"] = "Source: Synthetic demo master data generated by build_workbook.py"
    ws["G9"].font = Font(name=FONT, italic=True, size=8, color="808080")
    widths(ws, {"A": 12, "B": 20, "C": 12, "D": 14, "E": 14, "G": 18, "H": 12, "I": 12, "J": 12, "K": 12})
    return ws


def sheet_sales(wb, sales):
    ws = wb.create_sheet("Data_Sales")
    title(ws, "Source 1: ERP Sales Transactions", "Blue = raw imported data; black = calculated enrichment columns")
    cols = ["Invoice", "Date", "Region ID", "Product ID", "Qty", "Discount %",
            "Month", "Region", "Product", "Category", "Unit Price", "Net Revenue", "Gross Profit"]
    header(ws, 4, cols)
    for i, row in enumerate(sales, start=5):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = F_IN
        ws.cell(row=i, column=2).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=6).number_format = PCT
        ws.cell(row=i, column=7, value=f'=TEXT(B{i},"mmm")')
        ws.cell(row=i, column=8, value=f"=HLOOKUP(C{i},Lookups!$H$4:$K$5,2,FALSE)")
        ws.cell(row=i, column=9, value=f"=VLOOKUP(D{i},Lookups!$A$5:$E$10,2,FALSE)")
        ws.cell(row=i, column=10, value=f"=VLOOKUP(D{i},Lookups!$A$5:$E$10,3,FALSE)")
        ws.cell(row=i, column=11, value=f"=VLOOKUP(D{i},Lookups!$A$5:$E$10,4,FALSE)")
        ws.cell(row=i, column=12, value=f"=E{i}*K{i}*(1-F{i})")
        ws.cell(row=i, column=13, value=f"=L{i}*VLOOKUP(D{i},Lookups!$A$5:$E$10,5,FALSE)")
        for col in (11, 12, 13):
            ws.cell(row=i, column=col).number_format = CUR
        for col in range(7, 14):
            ws.cell(row=i, column=col).font = F_CALC
    last = 4 + len(sales)
    tab = Table(displayName="tblSales", ref=f"A4:M{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "A5"
    widths(ws, dict(zip("ABCDEFGHIJKLM", [12, 12, 10, 10, 7, 10, 8, 10, 18, 11, 11, 13, 13])))
    return last


def sheet_gl(wb, gl):
    ws = wb.create_sheet("Data_GL")
    title(ws, "Source 2: General Ledger Revenue Postings", "Month-end journal entries from finance system")
    header(ws, 4, ["Journal ID", "Posting Date", "Region ID", "Account", "Amount ($)", "Month"])
    for i, row in enumerate(gl, start=5):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=v).font = F_IN
        ws.cell(row=i, column=2).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=5).number_format = CUR
        ws.cell(row=i, column=6, value=f'=TEXT(B{i},"mmm")').font = F_CALC
    last = 4 + len(gl)
    tab = Table(displayName="tblGL", ref=f"A4:F{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws.add_table(tab)
    widths(ws, {"A": 14, "B": 13, "C": 10, "D": 15, "E": 14, "F": 8})
    return last


def sheet_assumptions(wb):
    ws = wb.create_sheet("Assumptions")
    title(ws, "Forecast & Control Assumptions", "Blue cells are inputs - change them to run scenarios")
    header(ws, 4, ["Driver", "Value", "Notes"])
    items = [
        ("Monthly revenue growth", 0.015, PCT, "Applied month-over-month to forecast base"),
        ("Seasonality uplift (Q4)", 0.08, PCT, "Extra uplift for Oct-Dec"),
        ("Gross margin target", 0.60, PCT, "Target used for forecast gross profit"),
        ("Budget growth vs prior-year run-rate", 0.05, PCT, "Budget = avg actual x (1+growth)"),
        ("Reconciliation tolerance", 0.01, PCT, "Max allowed ERP vs GL difference"),
        ("Variance alert threshold", 0.05, PCT, "Flags variances beyond +/- this %"),
    ]
    for i, (n, v, fmt, note) in enumerate(items, start=5):
        ws.cell(row=i, column=1, value=n).font = F_CALC
        c = ws.cell(row=i, column=2, value=v)
        c.font, c.number_format = F_IN, fmt
        c.fill = PatternFill("solid", start_color="FFFF00")
        ws.cell(row=i, column=3, value=note).font = Font(name=FONT, italic=True, color="595959")
        for col in range(1, 4):
            ws.cell(row=i, column=col).border = BORDER
    dv = DataValidation(type="decimal", operator="between", formula1="-0.5", formula2="1",
                        showErrorMessage=True, errorTitle="Invalid input", error="Enter a % between -50% and 100%")
    ws.add_data_validation(dv)
    dv.add("B5:B10")
    widths(ws, {"A": 36, "B": 12, "C": 46})


def sheet_pivot(wb, last):
    ws = wb.create_sheet("Pivot_Summary")
    title(ws, "Revenue Pivot: Region x Month ($)", "Pivot-style summary via SUMIFS (refresh-free); see VBA module for native PivotTable")
    header(ws, 4, ["Region"] + MONTHS + ["Total"])
    rng = lambda c: f"Data_Sales!${c}$5:${c}${last}"
    for i, (_, rname) in enumerate(REGIONS, start=5):
        ws.cell(row=i, column=1, value=rname).font = Font(name=FONT, bold=True)
        for j, m in enumerate(MONTHS, start=2):
            c = ws.cell(row=i, column=j, value=f'=SUMIFS({rng("L")},{rng("H")},$A{i},{rng("G")},"{m}")')
            c.number_format, c.font, c.border = CUR, F_LINK, BORDER
        t = ws.cell(row=i, column=14, value=f"=SUM(B{i}:M{i})")
        t.number_format, t.font, t.border = CUR, Font(name=FONT, bold=True), BORDER
    tr = 5 + len(REGIONS)
    ws.cell(row=tr, column=1, value="Total").font = Font(name=FONT, bold=True)
    for j in range(2, 15):
        L = get_column_letter(j)
        c = ws.cell(row=tr, column=j, value=f"=SUM({L}5:{L}{tr - 1})")
        c.number_format, c.font, c.fill, c.border = CUR, Font(name=FONT, bold=True), FILL_SUB, BORDER
    ws.conditional_formatting.add(f"B5:M{tr - 1}", ColorScaleRule(start_type="min", start_color="F8696B",
                                  mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))

    # Category x Region
    r0 = tr + 3
    ws.cell(row=r0 - 1, column=1, value="Revenue & Gross Profit by Category").font = Font(name=FONT, bold=True, size=12, color=NAVY)
    header(ws, r0, ["Category", "Net Revenue ($)", "Gross Profit ($)", "GP Margin %", "Share of Revenue"])
    cats = sorted({p[2] for p in PRODUCTS})
    for i, cat in enumerate(cats, start=r0 + 1):
        ws.cell(row=i, column=1, value=cat).font = Font(name=FONT, bold=True)
        ws.cell(row=i, column=2, value=f'=SUMIF({rng("J")},A{i},{rng("L")})').number_format = CUR
        ws.cell(row=i, column=3, value=f'=SUMIF({rng("J")},A{i},{rng("M")})').number_format = CUR
        ws.cell(row=i, column=4, value=f"=IF(B{i}=0,0,C{i}/B{i})").number_format = PCT
        ws.cell(row=i, column=5, value=f"=IF($N${tr}=0,0,B{i}/$N${tr})").number_format = PCT
        for col in range(2, 6):
            ws.cell(row=i, column=col).border = BORDER
    ce = r0 + len(cats)
    ws.conditional_formatting.add(f"B{r0 + 1}:B{ce}", DataBarRule(start_type="min", end_type="max", color=TEAL))
    widths(ws, {"A": 14, **{get_column_letter(c): 11 for c in range(2, 15)}, "N": 13})
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["E"].width = 15
    return tr, r0, ce


def sheet_recon(wb, last_sales, last_gl):
    ws = wb.create_sheet("Reconciliation")
    title(ws, "Monthly Reconciliation: ERP Sales vs General Ledger", "Automated tie-out with tolerance from Assumptions")
    header(ws, 4, ["Month", "ERP Revenue ($)", "GL Revenue ($)", "Difference ($)", "Diff %", "Status"])
    for i, m in enumerate(MONTHS, start=5):
        ws.cell(row=i, column=1, value=m).font = Font(name=FONT, bold=True)
        ws.cell(row=i, column=2, value=f'=SUMIF(Data_Sales!$G$5:$G${last_sales},A{i},Data_Sales!$L$5:$L${last_sales})')
        ws.cell(row=i, column=3, value=f'=SUMIF(Data_GL!$F$5:$F${last_gl},A{i},Data_GL!$E$5:$E${last_gl})')
        ws.cell(row=i, column=4, value=f"=C{i}-B{i}")
        ws.cell(row=i, column=5, value=f"=IF(B{i}=0,0,D{i}/B{i})")
        ws.cell(row=i, column=6, value=f'=IF(ABS(E{i})<=Assumptions!$B$9,"Reconciled","Investigate")')
        for col, fmt in zip(range(2, 6), [CUR, CUR, CUR, PCT]):
            ws.cell(row=i, column=col).number_format = fmt
        for col in range(1, 7):
            ws.cell(row=i, column=col).border = BORDER
        ws.cell(row=i, column=2).font = F_LINK
        ws.cell(row=i, column=3).font = F_LINK
    ws.cell(row=17, column=1, value="Total").font = Font(name=FONT, bold=True)
    for col in "BCD":
        ws[f"{col}17"] = f"=SUM({col}5:{col}16)"
        ws[f"{col}17"].number_format = CUR
    ws["E17"] = "=IF(B17=0,0,D17/B17)"
    ws["E17"].number_format = PCT
    ws["F17"] = '=COUNTIF(F5:F16,"Investigate")&" exception(s)"'
    for col in "ABCDEF":
        ws[f"{col}17"].fill, ws[f"{col}17"].font, ws[f"{col}17"].border = FILL_SUB, Font(name=FONT, bold=True), BORDER
    ws.conditional_formatting.add("F5:F16", CellIsRule(operator="equal", formula=['"Reconciled"'],
                                  fill=PatternFill("solid", start_color="C6EFCE"), font=Font(color="006100")))
    ws.conditional_formatting.add("F5:F16", CellIsRule(operator="equal", formula=['"Investigate"'],
                                  fill=PatternFill("solid", start_color="FFC7CE"), font=Font(color="9C0006")))
    widths(ws, {"A": 10, "B": 16, "C": 16, "D": 15, "E": 10, "F": 16})


def sheet_bva(wb, tr):
    ws = wb.create_sheet("Budget_vs_Actual")
    title(ws, "Budget vs Actual - Variance Analysis", "Budget derived from Assumptions; variances flagged vs alert threshold")
    header(ws, 4, ["Month", "Actual ($)", "Budget ($)", "Variance ($)", "Variance %", "Flag", "YTD Actual ($)", "YTD Budget ($)", "YTD Var %"])
    for i, m in enumerate(MONTHS, start=5):
        col = get_column_letter(i - 3)
        ws.cell(row=i, column=1, value=m).font = Font(name=FONT, bold=True)
        ws.cell(row=i, column=2, value=f"=Pivot_Summary!{col}{tr}").font = F_LINK
        ws.cell(row=i, column=3, value=f"=ROUND(AVERAGE(Pivot_Summary!$B${tr}:$M${tr})*(1+Assumptions!$B$8),-2)")
        ws.cell(row=i, column=4, value=f"=B{i}-C{i}")
        ws.cell(row=i, column=5, value=f"=IF(C{i}=0,0,D{i}/C{i})")
        ws.cell(row=i, column=6, value=f'=IF(E{i}>Assumptions!$B$10,"Above",IF(E{i}<-Assumptions!$B$10,"Below","On Track"))')
        ws.cell(row=i, column=7, value=f"=SUM($B$5:B{i})")
        ws.cell(row=i, column=8, value=f"=SUM($C$5:C{i})")
        ws.cell(row=i, column=9, value=f"=IF(H{i}=0,0,G{i}/H{i}-1)")
        for c, fmt in zip(range(2, 10), [CUR, CUR, CUR, PCT, None, CUR, CUR, PCT]):
            if fmt:
                ws.cell(row=i, column=c).number_format = fmt
        for c in range(1, 10):
            ws.cell(row=i, column=c).border = BORDER
    ws["A17"] = "Total"
    for col in "BCD":
        ws[f"{col}17"] = f"=SUM({col}5:{col}16)"
        ws[f"{col}17"].number_format = CUR
    ws["E17"] = "=IF(C17=0,0,D17/C17)"
    ws["E17"].number_format = PCT
    for col in "ABCDEFGHI":
        ws[f"{col}17"].fill, ws[f"{col}17"].font, ws[f"{col}17"].border = FILL_SUB, Font(name=FONT, bold=True), BORDER
    ws.conditional_formatting.add("D5:D16", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))
    ws.conditional_formatting.add("F5:F16", FormulaRule(formula=['F5="Below"'], fill=PatternFill("solid", start_color="FFC7CE")))
    ws.conditional_formatting.add("F5:F16", FormulaRule(formula=['F5="Above"'], fill=PatternFill("solid", start_color="C6EFCE")))
    ws.conditional_formatting.add("E5:E16", DataBarRule(start_type="min", end_type="max", color="5B9BD5"))
    widths(ws, {"A": 10, "B": 14, "C": 14, "D": 14, "E": 12, "F": 11, "G": 15, "H": 15, "I": 11})

    ch = BarChart()
    ch.title, ch.y_axis.title = "Actual vs Budget", "$"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=4, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=16))
    ch.height, ch.width = 8, 18
    ws.add_chart(ch, "K4")


def sheet_forecast(wb):
    ws = wb.create_sheet("Forecast")
    title(ws, f"12-Month Rolling Forecast ({YEAR + 1})", "Base = last-3-month average actual; driven by Assumptions")
    ws["A4"] = "Forecast base (avg last 3 months actual)"
    ws["D4"] = "=AVERAGE(Budget_vs_Actual!B14:B16)"
    ws["D4"].number_format, ws["D4"].font = CUR, F_LINK
    header(ws, 6, ["Month", "Period #", "Forecast Revenue ($)", "Forecast GP ($)", "Cumulative Revenue ($)", "MoM Growth %"])
    for i, m in enumerate(MONTHS, start=7):
        n = i - 6
        ws.cell(row=i, column=1, value=f"{m}-{str(YEAR + 1)[2:]}").font = Font(name=FONT, bold=True)
        ws.cell(row=i, column=2, value=n).font = F_IN
        ws.cell(row=i, column=3, value=f"=$D$4*(1+Assumptions!$B$5)^B{i}*IF(B{i}>=10,1+Assumptions!$B$6,1)")
        ws.cell(row=i, column=4, value=f"=C{i}*Assumptions!$B$7")
        ws.cell(row=i, column=5, value=f"=SUM($C$7:C{i})")
        ws.cell(row=i, column=6, value="=0" if n == 1 else f"=IF(C{i-1}=0,0,C{i}/C{i-1}-1)")
        for c, fmt in zip(range(3, 7), [CUR, CUR, CUR, PCT]):
            ws.cell(row=i, column=c).number_format = fmt
        for c in range(1, 7):
            ws.cell(row=i, column=c).border = BORDER
    ws["A19"] = "Total"
    ws["C19"] = "=SUM(C7:C18)"
    ws["D19"] = "=SUM(D7:D18)"
    for col in "ABCDEF":
        ws[f"{col}19"].fill, ws[f"{col}19"].font, ws[f"{col}19"].border = FILL_SUB, Font(name=FONT, bold=True), BORDER
        if col in "CD":
            ws[f"{col}19"].number_format = CUR
    widths(ws, {"A": 12, "B": 10, "C": 20, "D": 17, "E": 22, "F": 14})
    lc = LineChart()
    lc.title, lc.y_axis.title = "Forecast Revenue & Gross Profit", "$"
    lc.add_data(Reference(ws, min_col=3, max_col=4, min_row=6, max_row=18), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=1, min_row=7, max_row=18))
    lc.height, lc.width = 8, 18
    ws.add_chart(lc, "H4")


def sheet_dashboard(wb, last, tr, cr0, ce):
    ws = wb["Dashboard"]
    title(ws, "Executive Financial Dashboard", f"FY{YEAR} | Consolidated from ERP Sales + General Ledger | Select a region in C4")
    ws["B4"] = "Region filter:"
    ws["B4"].font = Font(name=FONT, bold=True)
    ws["C4"] = "All"
    ws["C4"].font, ws["C4"].fill, ws["C4"].border = F_IN, PatternFill("solid", start_color="FFFF00"), BORDER
    ws["C4"].comment = Comment("Pick All or a region - KPIs recalculate instantly.", "Model")
    dv = DataValidation(type="list", formula1='"All,North,South,East,West"', allow_blank=False, showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add("C4")

    s = lambda c: f"Data_Sales!${c}$5:${c}${last}"
    crit = f'IF($C$4="All","*",$C$4)'
    kpis = [
        ("Net Revenue", f"=SUMIFS({s('L')},{s('H')},{crit})", CUR),
        ("Gross Profit", f"=SUMIFS({s('M')},{s('H')},{crit})", CUR),
        ("GP Margin", "=IF(B8=0,0,C8/B8)", PCT),
        ("Transactions", f"=COUNTIFS({s('H')},{crit})", "#,##0"),
        ("Avg Deal Size", "=IF(E8=0,0,B8/E8)", CUR),
        ("Budget Var %", "=Budget_vs_Actual!E17", PCT),
        ("Recon Exceptions", '=COUNTIF(Reconciliation!F5:F16,"Investigate")', "0"),
    ]
    for j, (lbl, f, fmt) in enumerate(kpis, start=2):
        h = ws.cell(row=7, column=j, value=lbl)
        h.font, h.fill, h.alignment, h.border = F_HDR, PatternFill("solid", start_color=TEAL), Alignment(horizontal="center"), BORDER
        v = ws.cell(row=8, column=j, value=f)
        v.font = Font(name=FONT, size=14, bold=True, color=NAVY)
        v.number_format, v.alignment, v.border, v.fill = fmt, Alignment(horizontal="center"), BORDER, FILL_SUB
        ws.column_dimensions[get_column_letter(j)].width = 17
    ws.row_dimensions[8].height = 30
    ws.conditional_formatting.add("G8", CellIsRule(operator="lessThan", formula=["-Assumptions!$B$10"], font=Font(color="C00000", bold=True, size=14)))
    ws.conditional_formatting.add("H8", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", start_color="FFC7CE")))
    ws.column_dimensions["A"].width = 3

    # Monthly trend for selected region (chart source)
    ws["J6"] = "Monthly trend (selected region)"
    ws["J6"].font = Font(name=FONT, bold=True, color=NAVY)
    header(ws, 7, ["Month", "Revenue ($)", "Gross Profit ($)"], col=10)
    for i, m in enumerate(MONTHS, start=8):
        ws.cell(row=i, column=10, value=m)
        ws.cell(row=i, column=11, value=f'=SUMIFS({s("L")},{s("G")},J{i},{s("H")},{crit})').number_format = CUR
        ws.cell(row=i, column=12, value=f'=SUMIFS({s("M")},{s("G")},J{i},{s("H")},{crit})').number_format = CUR
    widths(ws, {"J": 9, "K": 14, "L": 15})

    lc = LineChart()
    lc.title = "Monthly Revenue & Gross Profit"
    lc.add_data(Reference(ws, min_col=11, max_col=12, min_row=7, max_row=19), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=10, min_row=8, max_row=19))
    lc.height, lc.width = 8, 16
    ws.add_chart(lc, "B11")

    pv = wb["Pivot_Summary"]
    bc = BarChart()
    bc.title = "Revenue by Region"
    bc.add_data(Reference(pv, min_col=14, min_row=4, max_row=4 + len(REGIONS)), titles_from_data=True)
    bc.set_categories(Reference(pv, min_col=1, min_row=5, max_row=4 + len(REGIONS)))
    bc.legend = None
    bc.height, bc.width = 8, 12
    ws.add_chart(bc, "B28")

    pc = PieChart()
    pc.title = "Revenue Mix by Category"
    pc.add_data(Reference(pv, min_col=2, min_row=cr0, max_row=ce), titles_from_data=True)
    pc.set_categories(Reference(pv, min_col=1, min_row=cr0 + 1, max_row=ce))
    pc.height, pc.width = 8, 12
    ws.add_chart(pc, "F28")


def export_csv(sales, gl, outdir):
    outdir.mkdir(parents=True, exist_ok=True)
    with open(outdir / "erp_sales.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Invoice", "Date", "RegionID", "ProductID", "Qty", "Discount"])
        for r in sales:
            w.writerow([r[0], r[1].isoformat(), *r[2:]])
    with open(outdir / "gl_revenue.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["JournalID", "PostingDate", "RegionID", "Account", "Amount"])
        for r in gl:
            w.writerow([r[0], r[1].isoformat(), *r[2:]])


def main():
    ap = argparse.ArgumentParser()
    root = Path(__file__).resolve().parents[2]
    ap.add_argument("--rows", type=int, default=1200)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", default=str(root / "workbook" / "Financial_Dashboard_Model.xlsx"))
    a = ap.parse_args()

    sales = generate_data(a.rows, a.seed)
    gl = build_gl(sales, a.seed)
    export_csv(sales, gl, root / "data")

    wb = Workbook()
    wb.active.title = "Dashboard"
    sheet_lookups(wb)
    last_s = sheet_sales(wb, sales)
    last_g = sheet_gl(wb, gl)
    sheet_assumptions(wb)
    tr, cr0, ce = sheet_pivot(wb, last_s)
    sheet_recon(wb, last_s, last_g)
    sheet_bva(wb, tr)
    sheet_forecast(wb)
    sheet_dashboard(wb, last_s, tr, cr0, ce)
    wb._sheets.insert(1, wb._sheets.pop(wb.sheetnames.index("Assumptions")))
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = NAVY if ws.title == "Dashboard" else ("FFC000" if ws.title == "Assumptions" else TEAL)
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    print(f"Workbook written: {a.out}")


if __name__ == "__main__":
    main()
