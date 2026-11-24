import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUMMARY_DIR = os.path.join(BASE_DIR, "outputs", "summary_tables")
DELIV_DIR = os.path.join(BASE_DIR, "deliverables")
os.makedirs(DELIV_DIR, exist_ok=True)


# Styling Palette 
C_NAVY = "1B365D"
C_BLUE = "005EB8"
C_SUBHDR = "D9E2EC"
C_ZEBRA = "F7F9FC"
C_INPUT_BG = "E8F1FA"
C_INPUT_FONT = "00338D"
C_TOTAL_BG = "EBF3FA"
C_GREEN_BG = "DEF7EC"
C_GREEN_FG = "03543F"
C_RED_BG = "FDE8E8"
C_RED_FG = "9B1C1C"
C_AMBER_BG = "FEF3C7"
C_AMBER_FG = "92400E"
C_BORDER = "CBD5E1"
C_DARK_BORDER = "1B365D"

THIN_SIDE = Side(border_style="thin", color=C_BORDER)
MED_NAVY = Side(border_style="medium", color=C_DARK_BORDER)
DOUBLE_BOT = Side(border_style="double", color=C_DARK_BORDER)

BORDER_DATA = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_TOTAL = Border(top=THIN_SIDE, bottom=DOUBLE_BOT, left=THIN_SIDE, right=THIN_SIDE)
BORDER_HDR = Border(top=MED_NAVY, bottom=MED_NAVY, left=THIN_SIDE, right=THIN_SIDE)

FONT_TITLE = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
FONT_SEC = Font(name="Calibri", size=11, bold=True, color=C_NAVY)
FONT_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_DATA = Font(name="Calibri", size=10, bold=False, color="1F2937")
FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="1F2937")
FONT_INPUT = Font(name="Calibri", size=10, bold=True, color=C_INPUT_FONT)
FONT_GREEN = Font(name="Calibri", size=10, bold=True, color=C_GREEN_FG)
FONT_NOTE = Font(name="Calibri", size=9, italic=True, color="4B5563")

FILL_NAVY = PatternFill("solid", fgColor=C_NAVY)
FILL_BLUE = PatternFill("solid", fgColor=C_BLUE)
FILL_SUBHDR = PatternFill("solid", fgColor=C_SUBHDR)
FILL_ZEBRA = PatternFill("solid", fgColor=C_ZEBRA)
FILL_INPUT = PatternFill("solid", fgColor=C_INPUT_BG)
FILL_TOTAL = PatternFill("solid", fgColor=C_TOTAL_BG)
FILL_GREEN = PatternFill("solid", fgColor=C_GREEN_BG)
FILL_RED = PatternFill("solid", fgColor=C_RED_BG)
FILL_AMBER = PatternFill("solid", fgColor=C_AMBER_BG)
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")

FMT_CURR = "$#,##0;($#,##0);\"-\""
FMT_CURR_DEC = "$#,##0.00;($#,##0.00);\"-\""
FMT_PCT = "0.0%"
FMT_PCT2 = "0.00%"
FMT_INT = "#,##0"
FMT_DEC = "0.00"
FMT_BPS = "+#,##0\" bps\";-#,##0\" bps\";\"0 bps\""


def style_banner(ws, max_col: int, title: str, subtitle: str):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c1 = ws.cell(1, 1, title)
    c1.font = FONT_TITLE
    c1.fill = FILL_NAVY
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2 = ws.cell(2, 1, subtitle)
    c2.font = FONT_SUBTITLE
    c2.fill = FILL_NAVY
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 19
    for col in range(1, max_col + 1):
        ws.cell(1, col).fill = FILL_NAVY
        ws.cell(2, col).fill = FILL_NAVY


def style_section_header(ws, row: int, start_col: int, end_col: int, text: str):
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    c = ws.cell(row, start_col, text)
    c.font = FONT_SEC
    c.fill = FILL_SUBHDR
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = FILL_SUBHDR
        cell.border = Border(bottom=MED_NAVY)
    ws.row_dimensions[row].height = 22


def style_table_header(ws, row: int, headers: list, start_col: int = 1):
    ws.row_dimensions[row].height = 24
    for i, h in enumerate(headers):
        cell = ws.cell(row, start_col + i, h)
        cell.font = FONT_HDR
        cell.fill = FILL_BLUE
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_HDR


def set_col_widths(ws, widths: dict):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


def build_workbook():
    # Load aggregated datasets produced by script 01
    with open(os.path.join(SUMMARY_DIR, "baseline_metrics.json")) as f:
        m = json.load(f)
    shipping_df = pd.read_csv(os.path.join(SUMMARY_DIR, "shipping_sla.csv"))
    payment_df = pd.read_csv(os.path.join(SUMMARY_DIR, "payment_risk.csv"))
    abc_xyz_df = pd.read_csv(os.path.join(SUMMARY_DIR, "abc_xyz_summary.csv"))
    disc_df = pd.read_csv(os.path.join(SUMMARY_DIR, "discount_tier_leakage.csv"))
    cat_df = pd.read_csv(os.path.join(SUMMARY_DIR, "category_abc_xyz.csv")).head(15)
    market_df = pd.read_csv(os.path.join(SUMMARY_DIR, "market_summary.csv"))

    wb = Workbook()
    ws_exec = wb.active
    ws_exec.title = "Executive_Summary"
    ws_diag = wb.create_sheet("Operational_Diagnostics")
    ws_ebitda = wb.create_sheet("EBITDA_Bridge_Model")
    ws_sens = wb.create_sheet("Sensitivity_Analysis")

    for ws in [ws_exec, ws_diag, ws_ebitda, ws_sens]:
        ws.views.sheetView[0].showGridLines = True


    # TAB 2: OPERATIONAL_DIAGNOSTICS
  
    style_banner(
        ws_diag, 9,
        "OPERATIONAL & SUPPLY CHAIN DIAGNOSTICS (DATACO GLOBAL OPERATIONS)",
        f"Granular Audit Across {m['total_rows']:,} Order Items | {m['unique_orders']:,} Unique Orders | {m['unique_categories']} Categories | 5 Markets (Jan 2015 - Jan 2018)"
    )

    # Table A: Shipping Mode SLA Performance (Rows 4-10)
    style_section_header(ws_diag, 4, 1, 9, "1. SHIPPING MODE SLA PERFORMANCE & EXPEDITED DELIVERY BOTTLENECK")
    style_table_header(ws_diag, 5, [
        "Shipping Mode", "Order Items", "Net Sales ($)", "Operating Profit ($)",
        "Margin %", "Scheduled Days", "Actual Days", "Avg Delay (Days)", "Late Delivery Risk %"
    ])
    r = 6
    for idx, row in shipping_df.iterrows():
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        vals = [
            row["shipping_mode"], int(row["order_items"]), float(row["net_sales"]),
            float(row["operating_profit"]), f"=D{r}/C{r}",
            float(row["avg_scheduled_days"]), float(row["avg_actual_days"]),
            f"=G{r}-F{r}", float(row["late_delivery_pct"]) / 100.0
        ]
        fmts = [None, FMT_INT, FMT_CURR, FMT_CURR, FMT_PCT, FMT_DEC, FMT_DEC, "+0.00;-0.00;0.00", FMT_PCT]
        for c_i, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_diag.cell(r, c_i, val)
            cell.font = FONT_DATA
            cell.fill = fill
            cell.border = BORDER_DATA
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="left" if c_i == 1 else "right", vertical="center")
        ws_diag.row_dimensions[r].height = 20
        r += 1

    # Total row for Shipping Mode (Row 10)
    ws_diag.cell(r, 1, "Total / Weighted Avg").font = FONT_BOLD
    ws_diag.cell(r, 2, f"=SUM(B6:B{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 3, f"=SUM(C6:C{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 4, f"=SUM(D6:D{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 5, f"=D{r}/C{r}").number_format = FMT_PCT
    ws_diag.cell(r, 6, f"=SUMPRODUCT(B6:B{r-1},F6:F{r-1})/B{r}").number_format = FMT_DEC
    ws_diag.cell(r, 7, f"=SUMPRODUCT(B6:B{r-1},G6:G{r-1})/B{r}").number_format = FMT_DEC
    ws_diag.cell(r, 8, f"=G{r}-F{r}").number_format = "+0.00;-0.00;0.00"
    ws_diag.cell(r, 9, f"=SUMPRODUCT(B6:B{r-1},I6:I{r-1})/B{r}").number_format = FMT_PCT
    for c_i in range(1, 10):
        cell = ws_diag.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="left" if c_i == 1 else "right", vertical="center")
    ws_diag.row_dimensions[r].height = 21

    # Conditional formatting on Late Delivery Risk % (I6:I9)
    ws_diag.conditional_formatting.add(
        "I6:I9",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.70"], fill=FILL_RED, font=Font(color=C_RED_FG, bold=True))
    )
    ws_diag.conditional_formatting.add(
        "I6:I9",
        CellIsRule(operator="between", formula=["0.45", "0.699"], fill=FILL_AMBER, font=Font(color=C_AMBER_FG, bold=True))
    )

    # Table B: Payment Channel & Fraud/Cancellation Risk Audit (Rows 12-18)
    style_section_header(ws_diag, 12, 1, 9, "2. PAYMENT CHANNEL RISK AUDIT (100% OF FRAUD & CANCELLATIONS ORIGINATE IN TRANSFER)")
    style_table_header(ws_diag, 13, [
        "Payment Channel", "Order Items", "Net Sales ($)", "Operating Profit ($)",
        "Suspected Fraud Items", "Canceled Items", "Fraud Rate %", "Cancel Rate %", "Failed Net Sales ($)"
    ])
    r = 14
    for idx, row in payment_df.iterrows():
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        vals = [
            row["payment_type"], int(row["order_items"]), float(row["net_sales"]),
            float(row["operating_profit"]), int(row["suspected_fraud_items"]),
            int(row["canceled_items"]), f"=E{r}/B{r}", f"=F{r}/B{r}", float(row["failed_net_sales"])
        ]
        fmts = [None, FMT_INT, FMT_CURR, FMT_CURR, FMT_INT, FMT_INT, FMT_PCT2, FMT_PCT2, FMT_CURR]
        for c_i, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_diag.cell(r, c_i, val)
            cell.font = FONT_BOLD if row["payment_type"] == "TRANSFER" else FONT_DATA
            cell.fill = FILL_RED if (row["payment_type"] == "TRANSFER" and c_i >= 5) else fill
            if row["payment_type"] == "TRANSFER" and c_i >= 5:
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_RED_FG)
            cell.border = BORDER_DATA
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="left" if c_i == 1 else "right", vertical="center")
        ws_diag.row_dimensions[r].height = 20
        r += 1

    ws_diag.cell(r, 1, "Total")
    ws_diag.cell(r, 2, f"=SUM(B14:B{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 3, f"=SUM(C14:C{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 4, f"=SUM(D14:D{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 5, f"=SUM(E14:E{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 6, f"=SUM(F14:F{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 7, f"=E{r}/B{r}").number_format = FMT_PCT2
    ws_diag.cell(r, 8, f"=F{r}/B{r}").number_format = FMT_PCT2
    ws_diag.cell(r, 9, f"=SUM(I14:I{r-1})").number_format = FMT_CURR
    for c_i in range(1, 10):
        cell = ws_diag.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="left" if c_i == 1 else "right", vertical="center")
    ws_diag.row_dimensions[r].height = 21

    # Table C: ABC-XYZ Inventory & Revenue Segmentation (Rows 20-29)
    style_section_header(ws_diag, 20, 1, 9, "3. ABC-XYZ INVENTORY SEGMENTATION & WORKING CAPITAL OPPORTUNITY")
    style_table_header(ws_diag, 21, [
        "ABC-XYZ Segment", "# Categories", "Order Items", "Net Sales ($)",
        "% Revenue Share", "Operating Profit ($)", "Margin %", "Avg Demand CV", "Recommended Supply Policy"
    ])
    policy_map = {
        "A-X": "Core Auto-Replenish (JIT / Tight Safety Stock)",
        "B-X": "Standard Automated Reorder (Min-Max)",
        "B-Y": "Dynamic Safety Stock + Promo Planning",
        "B-Z": "Shift to Make-to-Order / Vendor Consignment",
        "C-X": "Low-Touch Bulk Reorder (Quarterly)",
        "C-Y": "Selective Rationalization / Higher MOQ",
        "C-Z": "Rationalize SKU Tail / Drop-Ship Only",
    }
    r = 22
    abc_start = r
    for idx, row in abc_xyz_df.iterrows():
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        seg = row["abc_xyz"]
        vals = [
            seg, int(row["categories"]), int(row["order_items"]), float(row["net_sales"]),
            f"=D{r}/$D${abc_start + len(abc_xyz_df)}", float(row["operating_profit"]),
            f"=F{r}/D{r}", float(row["avg_cv"]), policy_map.get(seg, "Review")
        ]
        fmts = [None, FMT_INT, FMT_INT, FMT_CURR, FMT_PCT, FMT_CURR, FMT_PCT, FMT_DEC, None]
        for c_i, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_diag.cell(r, c_i, val)
            cell.font = FONT_BOLD if seg in ("A-X", "C-Z") else FONT_DATA
            cell.fill = FILL_GREEN if seg == "A-X" else (FILL_AMBER if seg in ("B-Z", "C-Z") else fill)
            cell.border = BORDER_DATA
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
        ws_diag.row_dimensions[r].height = 20
        r += 1

    ws_diag.cell(r, 1, "Total")
    ws_diag.cell(r, 2, f"=SUM(B{abc_start}:B{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 3, f"=SUM(C{abc_start}:C{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 4, f"=SUM(D{abc_start}:D{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 5, f"=SUM(E{abc_start}:E{r-1})").number_format = FMT_PCT
    ws_diag.cell(r, 6, f"=SUM(F{abc_start}:F{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 7, f"=F{r}/D{r}").number_format = FMT_PCT
    ws_diag.cell(r, 8, "-")
    ws_diag.cell(r, 9, "26 B-Z/C-Z Categories = 52% of Portfolio, Only 7.6% Rev")
    for c_i in range(1, 10):
        cell = ws_diag.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
    ws_diag.row_dimensions[r].height = 21

    # Table D: Discount Tier Margin Leakage (Rows 31-39)
    d_hdr = r + 2
    style_section_header(ws_diag, d_hdr, 1, 9, "4. COMMERCIAL DISCOUNTING & MARGIN LEAKAGE BY DISCOUNT TIER")
    style_table_header(ws_diag, d_hdr + 1, [
        "Discount Tier", "Order Items", "Gross Sales ($)", "Discounts Granted ($)",
        "Net Sales ($)", "Operating Profit ($)", "Margin %", "Loss-Making Order %", "Governance Action"
    ])
    gov_map = {
        "0%": "Benchmark Margin Tier",
        "1-5%": "Standard Volume Incentive",
        "6-10%": "Approved Tiered Promo",
        "11-15%": "Manager Sign-Off Threshold",
        "16-20%": "CAP & REMEDIATE (>15% Floor)",
        ">20%": "ELIMINATE / VP Exception Only",
    }
    r = d_hdr + 2
    disc_start = r
    for idx, row in disc_df.iterrows():
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        tier = str(row["discount_tier"])
        vals = [
            tier, int(row["order_items"]), float(row["gross_sales"]),
            float(row["total_discounts"]), float(row["net_sales"]),
            float(row["operating_profit"]), f"=F{r}/E{r}",
            float(row["loss_order_pct"]) / 100.0, gov_map.get(tier, "")
        ]
        fmts = [None, FMT_INT, FMT_CURR, FMT_CURR, FMT_CURR, FMT_CURR, FMT_PCT2, FMT_PCT, None]
        for c_i, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_diag.cell(r, c_i, val)
            cell.font = FONT_BOLD if tier in ("16-20%", ">20%") else FONT_DATA
            cell.fill = FILL_RED if (tier in ("16-20%", ">20%") and c_i in (4, 7, 9)) else fill
            if tier in ("16-20%", ">20%") and c_i in (4, 7, 9):
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_RED_FG)
            cell.border = BORDER_DATA
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
        ws_diag.row_dimensions[r].height = 20
        r += 1

    ws_diag.cell(r, 1, "Total")
    ws_diag.cell(r, 2, f"=SUM(B{disc_start}:B{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 3, f"=SUM(C{disc_start}:C{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 4, f"=SUM(D{disc_start}:D{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 5, f"=SUM(E{disc_start}:E{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 6, f"=SUM(F{disc_start}:F{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 7, f"=F{r}/E{r}").number_format = FMT_PCT2
    ws_diag.cell(r, 8, f"=SUMPRODUCT(B{disc_start}:B{r-1},H{disc_start}:H{r-1})/B{r}").number_format = FMT_PCT
    ws_diag.cell(r, 9, "$2.27M (60.8% of Discounts) in >15% Tiers")
    for c_i in range(1, 10):
        cell = ws_diag.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
    ws_diag.row_dimensions[r].height = 21

    # Table E: Market Performance Summary
    m_hdr = r + 2
    style_section_header(ws_diag, m_hdr, 1, 9, "5. GLOBAL MARKET PROFITABILITY & FULFILLMENT RISK SUMMARY")
    style_table_header(ws_diag, m_hdr + 1, [
        "Global Market", "Order Items", "Net Sales ($)", "% Revenue Share",
        "Operating Profit ($)", "Margin %", "Late Delivery Risk %", "Loss-Making Order %", "Priority Tier"
    ])
    r = m_hdr + 2
    mkt_start = r
    for idx, row in market_df.iterrows():
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        prio = "Tier 1 Focus (High Scale + Delay)" if row["market"] in ("Europe", "Pacific Asia") else "Tier 2 Optimize"
        vals = [
            row["market"], int(row["order_items"]), float(row["net_sales"]),
            f"=C{r}/$C${mkt_start + len(market_df)}", float(row["operating_profit"]),
            f"=E{r}/C{r}", float(row["late_delivery_pct"]) / 100.0,
            float(row["loss_order_pct"]) / 100.0, prio
        ]
        fmts = [None, FMT_INT, FMT_CURR, FMT_PCT, FMT_CURR, FMT_PCT, FMT_PCT, FMT_PCT, None]
        for c_i, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_diag.cell(r, c_i, val)
            cell.font = FONT_DATA
            cell.fill = fill
            cell.border = BORDER_DATA
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
        ws_diag.row_dimensions[r].height = 20
        r += 1

    ws_diag.cell(r, 1, "Total")
    ws_diag.cell(r, 2, f"=SUM(B{mkt_start}:B{r-1})").number_format = FMT_INT
    ws_diag.cell(r, 3, f"=SUM(C{mkt_start}:C{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 4, f"=SUM(D{mkt_start}:D{r-1})").number_format = FMT_PCT
    ws_diag.cell(r, 5, f"=SUM(E{mkt_start}:E{r-1})").number_format = FMT_CURR
    ws_diag.cell(r, 6, f"=E{r}/C{r}").number_format = FMT_PCT
    ws_diag.cell(r, 7, f"=SUMPRODUCT(B{mkt_start}:B{r-1},G{mkt_start}:G{r-1})/B{r}").number_format = FMT_PCT
    ws_diag.cell(r, 8, f"=SUMPRODUCT(B{mkt_start}:B{r-1},H{mkt_start}:H{r-1})/B{r}").number_format = FMT_PCT
    ws_diag.cell(r, 9, "Europe & Pacific Asia = 52% of Global Net Sales")
    for c_i in range(1, 10):
        cell = ws_diag.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="left" if c_i in (1, 9) else "right", vertical="center")
    ws_diag.row_dimensions[r].height = 21

    ws_diag.freeze_panes = "A6"
    set_col_widths(ws_diag, {
        "A": 22, "B": 14, "C": 17, "D": 18, "E": 17, "F": 17, "G": 15, "H": 18, "I": 38
    })

    # TAB 3: EBITDA_BRIDGE_MODEL (Live Formulas, Blue Inputs, 3 Scenarios)
  
    style_banner(
        ws_ebitda, 6,
        "3-SCENARIO EBITDA & WORKING CAPITAL TRANSFORMATION MODEL",
        "Annualized Baseline (FY2015-FY2017 Average Run-Rate) | Blue Shaded Cells = Editable Consulting Assumptions"
    )

    # Section 1: Annualized Baseline P&L & Operational Pools (Rows 4-17)
    style_section_header(ws_ebitda, 4, 1, 6, "1. ANNUALIZED BASELINE P&L & ADDRESSABLE LEAKAGE POOLS (FY2015-FY2017 AVG)")
    style_table_header(ws_ebitda, 5, [
        "Baseline Metric / Addressable Pool", "Annualized Value", "Unit",
        "3-Yr Cumulative (Dataset)", "Source / Audit Note", "Model Key"
    ])
    baseline_rows = [
        ("Annualized Gross Sales", m["ann_gross_sales"], "USD / yr", m["cum_gross_sales"], "FY15-17 average; avoids Jan-18 partial-yr skew", "B_GROSS", FMT_CURR),
        ("Less: Commercial Discounts", -m["ann_discounts"], "USD / yr", -m["cum_discounts"], "10.1% average discount rate on Gross Sales", "B_DISC", FMT_CURR),
        ("Annualized Net Sales", "=B6+B7", "USD / yr", m["cum_net_sales"], "Net realized top-line revenue (Order Item Total)", "B_NET", FMT_CURR),
        ("Less: Implied COGS & Fulfillment OpEx", "=B10-B8", "USD / yr", m["cum_operating_profit"] - m["cum_net_sales"], "Implied total operating cost structure", "B_OPEX", FMT_CURR),
        ("Baseline Annualized EBITDA (Op. Profit Proxy)", m["ann_operating_profit"], "USD / yr", m["cum_operating_profit"], "Baseline Benefit per Order annualized run-rate", "B_EBITDA", FMT_CURR),
        ("Baseline EBITDA Margin (%)", "=B10/B8", "% of Net Sales", m["cum_margin_pct"]/100.0, "12.01% baseline operating margin", "B_MGN", FMT_PCT2),
        ("Addressable Pool 1: Excess Discounts (>15% Cap)", m["ann_excess_disc_over_15"], "USD / yr", m["ann_excess_disc_over_15"]*3.0, "Pure discount dollars granted above 15% threshold", "P_DISC15", FMT_CURR),
        ("Addressable Pool 2: Loss-Making Order Leakage", m["ann_neg_profit_loss"], "USD / yr", m["cum_loss_pool"], f"18.7% of orders negative profit (~{m['ann_neg_items']:,.0f} items/yr)", "P_LOSS", FMT_CURR),
        ("Addressable Pool 3: Late Delivery Line Items", m["ann_late_items"], "Items / yr", m["ann_late_items"]*3.0, "54.8% late rate (First Class 95.3%, Second Class 76.6%)", "P_LATE", FMT_INT),
        ("Addressable Pool 4a: TRANSFER Failed Order Profit", m["ann_tf_fail_profit"], "USD / yr", m["ann_tf_fail_profit"]*3.0, "100% of fraud/cancellations in TRANSFER channel", "P_TF_PROFIT", FMT_CURR),
        ("Addressable Pool 4b: TRANSFER Failed Line Items", m["ann_tf_fail_items"], "Items / yr", m["cum_transfer_fail_items"], "Wasted pick/pack/reverse-logistics volume", "P_TF_ITEMS", FMT_INT),
    ]
    r = 6
    for idx, (label, val, unit, cum_v, note, key, fmt) in enumerate(baseline_rows):
        is_key_row = key in ("B_NET", "B_EBITDA", "B_MGN")
        fill = FILL_TOTAL if is_key_row else (FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE)
        ws_ebitda.cell(r, 1, label).font = FONT_BOLD if is_key_row else FONT_DATA
        c_val = ws_ebitda.cell(r, 2, val)
        c_val.number_format = fmt
        c_val.font = FONT_BOLD if is_key_row else FONT_DATA
        ws_ebitda.cell(r, 3, unit).font = FONT_DATA
        c_cum = ws_ebitda.cell(r, 4, cum_v)
        c_cum.number_format = fmt
        c_cum.font = FONT_DATA
        ws_ebitda.cell(r, 5, note).font = FONT_NOTE
        ws_ebitda.cell(r, 6, key).font = FONT_NOTE
        for c_i in range(1, 7):
            cell = ws_ebitda.cell(r, c_i)
            cell.fill = fill
            cell.border = BORDER_TOTAL if key == "B_EBITDA" else BORDER_DATA
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 5) else ("center" if c_i in (3, 6) else "right"), vertical="center")
        ws_ebitda.row_dimensions[r].height = 20
        r += 1

    # Section 2: Structural & Working Capital Modeling Assumptions (Rows 18-29)
    # Row 18 blank, Row 19 section header
    style_section_header(ws_ebitda, 19, 1, 6, "2. TRANSFORMATION LEVER ASSUMPTIONS (BLUE = EDITABLE MODEL INPUTS)")
    style_table_header(ws_ebitda, 20, [
        "Transformation Lever / Assumption", "Conservative", "Base Case", "Target Case", "Unit / Basis", "Operational Rationale"
    ])
    # Structural constants in rows 21-24 (shared across scenarios in col B, referenced by C/D)
    assump_rows = [
        ("Lever 1: Excess (>15%) Discount Cap Compliance", 0.30, 0.55, 0.80, "% Captured", "Enforce automated 15% discount guardrail in ERP", FMT_PCT),
        ("Lever 1b: Net Pricing Flow-Through (Post-Elasticity)", 0.85, 0.85, 0.85, "% Flow-Thru", "Assumes 15% volume elasticity offset on repriced orders", FMT_PCT),
        ("Lever 2: Loss-Making Order Tail Remediation", 0.10, 0.18, 0.25, "% Reduced", "Minimum margin floor + freight pass-through on loss tail", FMT_PCT),
        ("Lever 3: Expedited SLA Rework & Churn Savings / Late Item", 2.50, 4.00, 5.50, "$ / Late Item", "Carrier re-routing & realistic SLA tiers on 1st/2nd Class", FMT_CURR_DEC),
        ("Lever 4: TRANSFER Channel Fraud/Cancel Recovery Rate", 0.35, 0.55, 0.75, "% Mitigated", "Pre-shipment payment verification gate on TRANSFER only", FMT_PCT),
        ("Lever 4b: Pick/Pack & Reverse Logistics Cost / Failed Item", 12.00, 12.00, 12.00, "$ / Item", "Direct warehouse handling & chargeback overhead avoided", FMT_CURR_DEC),
        ("Lever 5: ABC-XYZ Inventory DIO Reduction (from 60d Base)", 6.0, 10.0, 15.0, "Days (DIO)", "Shift 26 B-Z/C-Z cats to MTO; tighten A-X safety stock", "0.0"),
        ("Working Capital Param: COGS Share of Total OpEx", 0.75, 0.75, 0.75, "% of Cost", "Implied COGS base for inventory valuation (75% of OpEx)", FMT_PCT),
        ("Working Capital Param: Annual Inventory Carrying Cost", 0.18, 0.18, 0.18, "% WACC+Whse", "Standard 18% holding, storage, insurance & obsolescence", FMT_PCT),
    ]
    r = 21
    for idx, (lbl, c_v, b_v, t_v, unit, rat, fmt) in enumerate(assump_rows):
        ws_ebitda.cell(r, 1, lbl).font = FONT_DATA
        ws_ebitda.cell(r, 1).fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        ws_ebitda.cell(r, 1).border = BORDER_DATA
        for col_i, val in [(2, c_v), (3, b_v), (4, t_v)]:
            cell = ws_ebitda.cell(r, col_i, val)
            cell.font = FONT_INPUT
            cell.fill = FILL_INPUT
            cell.border = BORDER_DATA
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal="right", vertical="center")
        c5 = ws_ebitda.cell(r, 5, unit)
        c5.font = FONT_DATA
        c5.fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        c5.border = BORDER_DATA
        c5.alignment = Alignment(horizontal="center", vertical="center")
        c6 = ws_ebitda.cell(r, 6, rat)
        c6.font = FONT_NOTE
        c6.fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        c6.border = BORDER_DATA
        c6.alignment = Alignment(horizontal="left", vertical="center")
        ws_ebitda.row_dimensions[r].height = 20
        r += 1

    # Section 3: Dynamic 3-Scenario EBITDA Bridge & Working Capital Output (Rows 31-47)
    style_section_header(ws_ebitda, 31, 1, 6, "3. DYNAMIC 3-SCENARIO EBITDA BRIDGE & WORKING CAPITAL IMPACT (100% LIVE FORMULAS)")
    style_table_header(ws_ebitda, 32, [
        "EBITDA Transformation Bridge Component", "Conservative", "Base Case", "Target Case", "Base % of Uplift", "Formula / Consulting Driver"
    ])
    # Row 33: Baseline Annualized EBITDA
    # Row 34: + Lever 1: Pricing & Discount Governance (>15% Cap)
    # Row 35: + Lever 2: Loss-Making Order Floor Remediation
    # Row 36: + Lever 3: Expedited SLA & Carrier Re-Engineering
    # Row 37: + Lever 4: TRANSFER Fraud & Cancellation Controls
    # Row 38: + Lever 5: ABC-XYZ Inventory Carrying Cost Savings
    # Row 39: Total Annualized EBITDA Uplift ($)
    # Row 40: Post-Transformation Annualized EBITDA ($)
    # Row 41: Post-Transformation Net Sales ($) [Baseline + Pricing Recovery]
    # Row 42: Baseline EBITDA Margin (%)
    # Row 43: Post-Transformation EBITDA Margin (%)
    # Row 44: EBITDA Margin Expansion (bps)
    # Row 45: EBITDA Growth vs. Baseline (%)
    # Row 46: One-Time Working Capital Cash Released ($)
    bridge_spec = [
        (33, "Baseline Annualized EBITDA ($)", "=$B$10", "=$B$10", "=$B$10", "-", "FY15-17 average Operating Profit proxy", FMT_CURR, True, FILL_TOTAL),
        (34, "+ Lever 1: Pricing & Discount Cap (>15% Guardrail)", "=$B$12*B21*B22", "=$B$12*C21*C22", "=$B$12*D21*D22", "=C34/$C$39", "Excess >15% discounts x compliance x flow-through", FMT_CURR, False, FILL_WHITE),
        (35, "+ Lever 2: Loss-Making Order Floor Remediation", "=$B$13*B23", "=$B$13*C23", "=$B$13*D23", "=C35/$C$39", "Negative-profit pool ($1.28M/yr) x reduction %", FMT_CURR, False, FILL_ZEBRA),
        (36, "+ Lever 3: Expedited SLA & 3PL Carrier Re-Routing", "=$B$14*B24", "=$B$14*C24", "=$B$14*D24", "=C36/$C$39", "Annualized late items (32.6K) x rework/churn savings", FMT_CURR, False, FILL_WHITE),
        (37, "+ Lever 4: TRANSFER Fraud & Cancellation Gate", "=($B$15+$B$16*B26)*B25", "=($B$15+$B$16*C26)*C25", "=($B$15+$B$16*D26)*D25", "=C37/$C$39", "Recovered margin + $12/item handling waste avoided", FMT_CURR, False, FILL_ZEBRA),
        (38, "+ Lever 5: ABC-XYZ Inventory Carrying Cost Lift", "=B46*B29", "=C46*C29", "=D46*D29", "=C38/$C$39", "Working capital cash freed x 18% carrying cost", FMT_CURR, False, FILL_WHITE),
        (39, "Total Annualized EBITDA Uplift ($)", "=SUM(B34:B38)", "=SUM(C34:C38)", "=SUM(D34:D38)", "=SUM(E34:E38)", "Sum of 5 operational & commercial levers", FMT_CURR, True, FILL_TOTAL),
        (40, "Post-Transformation Annualized EBITDA ($)", "=B33+B39", "=C33+C39", "=D33+D39", "-", "Target steady-state annualized EBITDA", FMT_CURR, True, FILL_GREEN),
        (41, "Post-Transformation Net Sales ($)", "=$B$8+B34", "=$B$8+C34", "=$B$8+D34", "-", "Baseline Net Sales + realized discount recovery", FMT_CURR, False, FILL_WHITE),
        (42, "Baseline EBITDA Margin (%)", "=$B$11", "=$B$11", "=$B$11", "-", "Baseline EBITDA / Baseline Net Sales", FMT_PCT2, False, FILL_ZEBRA),
        (43, "Post-Transformation EBITDA Margin (%)", "=B40/B41", "=C40/C41", "=D40/D41", "-", "Post-Transformation EBITDA / Net Sales", FMT_PCT2, True, FILL_GREEN),
        (44, "EBITDA Margin Expansion (bps)", "=(B43-B42)*10000", "=(C43-C42)*10000", "=(D43-D42)*10000", "-", "Basis points improvement over baseline", FMT_BPS, True, FILL_TOTAL),
        (45, "EBITDA Growth vs. Baseline (%)", "=B39/B33", "=C39/C33", "=D39/D33", "-", "% increase in annualized operating profit", FMT_PCT, True, FILL_WHITE),
        (46, "One-Time Working Capital Cash Released ($)", "=(ABS($B$9)*B28/365)*B27", "=(ABS($B$9)*C28/365)*C27", "=(ABS($B$9)*D28/365)*D27", "-", "Implied Daily COGS x DIO reduction days", FMT_CURR, True, FILL_AMBER),
    ]
    for row_num, lbl, f_c, f_b, f_t, f_pct, drv, fmt, is_b, fill in bridge_spec:
        ws_ebitda.cell(row_num, 1, lbl).font = FONT_GREEN if fill == FILL_GREEN else (FONT_BOLD if is_b else FONT_DATA)
        for col_i, form in [(2, f_c), (3, f_b), (4, f_t)]:
            c = ws_ebitda.cell(row_num, col_i, form)
            c.number_format = fmt
            c.font = FONT_GREEN if fill == FILL_GREEN else (FONT_BOLD if is_b else FONT_DATA)
        c5 = ws_ebitda.cell(row_num, 5, f_pct)
        if f_pct != "-":
            c5.number_format = FMT_PCT
        c5.font = FONT_BOLD if is_b else FONT_DATA
        c6 = ws_ebitda.cell(row_num, 6, drv)
        c6.font = FONT_NOTE
        for c_i in range(1, 7):
            cell = ws_ebitda.cell(row_num, c_i)
            cell.fill = fill
            cell.border = BORDER_TOTAL if row_num in (39, 40, 43, 46) else BORDER_DATA
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 6) else ("center" if c_i == 5 and f_pct == "-" else "right"), vertical="center")
        ws_ebitda.row_dimensions[row_num].height = 21 if is_b else 20

    # Embedded Native Excel Chart on EBITDA_Bridge_Model: Lever Contribution Comparison
    ch_lev = BarChart()
    ch_lev.type = "col"
    ch_lev.style = 10
    ch_lev.title = "Annualized EBITDA Uplift by Lever ($): Conservative vs. Base vs. Target"
    ch_lev.y_axis.title = "Annualized Uplift ($)"
    ch_lev.y_axis.number_format = "$#,##0"
    ch_lev.x_axis.title = "Transformation Lever"
    data_ref = Reference(ws_ebitda, min_col=2, max_col=4, min_row=32, max_row=38)
    cats_ref = Reference(ws_ebitda, min_col=1, min_row=34, max_row=38)
    # Note: row 33 is baseline, so let's chart rows 34-38 cleanly using series references
    from openpyxl.chart import Series
    s_cons = Series(Reference(ws_ebitda, min_col=2, min_row=34, max_row=38), title="Conservative")
    s_base = Series(Reference(ws_ebitda, min_col=3, min_row=34, max_row=38), title="Base Case")
    s_targ = Series(Reference(ws_ebitda, min_col=4, min_row=34, max_row=38), title="Target Case")
    s_cons.graphicalProperties.solidFill = "93C5FD"
    s_base.graphicalProperties.solidFill = "005EB8"
    s_targ.graphicalProperties.solidFill = "1B365D"
    ch_lev.append(s_cons)
    ch_lev.append(s_base)
    ch_lev.append(s_targ)
    ch_lev.set_categories(cats_ref)
    ch_lev.width = 22
    ch_lev.height = 10.5
    ch_lev.legend.position = "b"
    ws_ebitda.add_chart(ch_lev, "A48")

    ws_ebitda.freeze_panes = "B6"
    set_col_widths(ws_ebitda, {"A": 46, "B": 17, "C": 17, "D": 17, "E": 17, "F": 44})

    # TAB 4: SENSITIVITY_ANALYSIS (Live 2-Way Data Tables)
   
    style_banner(
        ws_sens, 8,
        "2-WAY SENSITIVITY & STRESS-TEST ANALYSIS",
        "Dynamic Formula-Driven Sensitivity Matrices Linked Directly to EBITDA_Bridge_Model"
    )

    # Matrix 1: Post-Transformation EBITDA Margin (%) — Loss-Tail Remediation % vs. Discount Cap Compliance %
    style_section_header(ws_sens, 4, 1, 8, "1. POST-TRANSFORMATION EBITDA MARGIN (%) — LOSS REMEDIATION vs. DISCOUNT CAP COMPLIANCE (HOLDING OTHER BASE LEVERS CONSTANT)")
    ws_sens.cell(5, 1, "Loss-Tail Remediation \\ Discount Cap %").font = FONT_HDR
    ws_sens.cell(5, 1).fill = FILL_NAVY
    ws_sens.cell(5, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_sens.cell(5, 1).border = BORDER_HDR
    ws_sens.row_dimensions[5].height = 26

    disc_caps = [0.20, 0.30, 0.45, 0.55, 0.70, 0.85, 0.95]
    loss_rems = [0.05, 0.10, 0.15, 0.18, 0.22, 0.25, 0.30]

    for j, dc in enumerate(disc_caps, 2):
        c = ws_sens.cell(5, j, dc)
        c.font = FONT_HDR
        c.fill = FILL_BLUE
        c.number_format = FMT_PCT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_HDR

    for i, lr in enumerate(loss_rems):
        row_i = 6 + i
        ws_sens.row_dimensions[row_i].height = 21
        rh = ws_sens.cell(row_i, 1, lr)
        rh.font = FONT_BOLD
        rh.fill = FILL_SUBHDR
        rh.number_format = FMT_PCT
        rh.alignment = Alignment(horizontal="center", vertical="center")
        rh.border = BORDER_DATA
        for j in range(2, 2 + len(disc_caps)):
            col_let = get_column_letter(j)
            # Live Excel formula computing Post-Transformation EBITDA Margin dynamically:
            # New EBITDA = Baseline + (Pool1 * ColDisc * FlowThru) + (Pool2 * RowLoss) + BaseL3 + BaseL4 + BaseL5
            # New Net Sales = BaseNet + (Pool1 * ColDisc * FlowThru)
            formula = (
                f"=(EBITDA_Bridge_Model!$B$10"
                f"+EBITDA_Bridge_Model!$B$12*{col_let}$5*EBITDA_Bridge_Model!$C$22"
                f"+EBITDA_Bridge_Model!$B$13*$A{row_i}"
                f"+EBITDA_Bridge_Model!$C$36+EBITDA_Bridge_Model!$C$37+EBITDA_Bridge_Model!$C$38)"
                f"/(EBITDA_Bridge_Model!$B$8+EBITDA_Bridge_Model!$B$12*{col_let}$5*EBITDA_Bridge_Model!$C$22)"
            )
            cell = ws_sens.cell(row_i, j, formula)
            cell.number_format = FMT_PCT2
            cell.font = FONT_BOLD if (lr == 0.18 and disc_caps[j-2] == 0.55) else FONT_DATA
            cell.border = BORDER_TOTAL if (lr == 0.18 and disc_caps[j-2] == 0.55) else BORDER_DATA
            cell.alignment = Alignment(horizontal="right", vertical="center")

    ws_sens.conditional_formatting.add(
        "B6:H12",
        ColorScaleRule(start_type="min", start_color="FDE8E8", mid_type="percentile", mid_value=50, mid_color="FEF3C7", end_type="max", end_color="DEF7EC")
    )
    ws_sens.cell(13, 1, "Note: Bold bordered cell (18.0% Loss Remediation, 55.0% Discount Compliance) represents the Base Case (16.57% EBITDA Margin vs. 12.01% Baseline).").font = FONT_NOTE

    # Matrix 2: Working Capital Cash Released ($) — COGS Share % vs. DIO Reduction (Days)
    style_section_header(ws_sens, 15, 1, 8, "2. WORKING CAPITAL CASH RELEASED ($) — DIO REDUCTION (DAYS) vs. IMPLIED COGS SHARE OF OPEX")
    ws_sens.cell(16, 1, "COGS Share \\ DIO Reduction (Days)").font = FONT_HDR
    ws_sens.cell(16, 1).fill = FILL_NAVY
    ws_sens.cell(16, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_sens.cell(16, 1).border = BORDER_HDR
    ws_sens.row_dimensions[16].height = 26

    dio_days = [4, 6, 8, 10, 12, 15, 20]
    cogs_shares = [0.60, 0.65, 0.70, 0.75, 0.80, 0.85]

    for j, d in enumerate(dio_days, 2):
        c = ws_sens.cell(16, j, d)
        c.font = FONT_HDR
        c.fill = FILL_BLUE
        c.number_format = "0\" days\""
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_HDR

    for i, cs in enumerate(cogs_shares):
        row_i = 17 + i
        ws_sens.row_dimensions[row_i].height = 21
        rh = ws_sens.cell(row_i, 1, cs)
        rh.font = FONT_BOLD
        rh.fill = FILL_SUBHDR
        rh.number_format = FMT_PCT
        rh.alignment = Alignment(horizontal="center", vertical="center")
        rh.border = BORDER_DATA
        for j in range(2, 2 + len(dio_days)):
            col_let = get_column_letter(j)
            formula = f"=(ABS(EBITDA_Bridge_Model!$B$9)*$A{row_i}/365)*{col_let}$16"
            cell = ws_sens.cell(row_i, j, formula)
            cell.number_format = FMT_CURR
            cell.font = FONT_BOLD if (cs == 0.75 and dio_days[j-2] == 10) else FONT_DATA
            cell.border = BORDER_TOTAL if (cs == 0.75 and dio_days[j-2] == 10) else BORDER_DATA
            cell.alignment = Alignment(horizontal="right", vertical="center")

    ws_sens.conditional_formatting.add(
        "B17:H22",
        ColorScaleRule(start_type="min", start_color="F7F9FC", mid_type="percentile", mid_value=50, mid_color="BFDBFE", end_type="max", end_color="60A5FA")
    )
    ws_sens.cell(23, 1, "Note: Bold bordered cell (75.0% COGS Share, 10-Day DIO Reduction) represents Base Case Working Capital Release ($197,420 one-time cash + $35,536/yr EBITDA lift).").font = FONT_NOTE

    # Matrix 3: Expedited SLA Savings ($) — Annualized Late Items Remediation % vs. Savings per Order ($)
    style_section_header(ws_sens, 25, 1, 8, "3. ANNUALIZED SLA REWORK SAVINGS ($) — SAVINGS PER LATE ORDER ($) vs. SLA REMEDIATION SCOPE (%)")
    ws_sens.cell(26, 1, "Scope % \\ Savings per Late Item").font = FONT_HDR
    ws_sens.cell(26, 1).fill = FILL_NAVY
    ws_sens.cell(26, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_sens.cell(26, 1).border = BORDER_HDR
    ws_sens.row_dimensions[26].height = 26

    unit_sav = [1.50, 2.50, 3.50, 4.00, 5.00, 5.50, 7.00]
    scopes = [0.50, 0.70, 0.85, 1.00]
    for j, u in enumerate(unit_sav, 2):
        c = ws_sens.cell(26, j, u)
        c.font = FONT_HDR
        c.fill = FILL_BLUE
        c.number_format = FMT_CURR_DEC
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_HDR
    for i, sc in enumerate(scopes):
        row_i = 27 + i
        ws_sens.row_dimensions[row_i].height = 21
        rh = ws_sens.cell(row_i, 1, sc)
        rh.font = FONT_BOLD
        rh.fill = FILL_SUBHDR
        rh.number_format = FMT_PCT
        rh.alignment = Alignment(horizontal="center", vertical="center")
        rh.border = BORDER_DATA
        for j in range(2, 2 + len(unit_sav)):
            col_let = get_column_letter(j)
            formula = f"=EBITDA_Bridge_Model!$B$14*$A{row_i}*{col_let}$26"
            cell = ws_sens.cell(row_i, j, formula)
            cell.number_format = FMT_CURR
            cell.font = FONT_BOLD if (sc == 1.00 and unit_sav[j-2] == 4.00) else FONT_DATA
            cell.border = BORDER_TOTAL if (sc == 1.00 and unit_sav[j-2] == 4.00) else BORDER_DATA
            cell.alignment = Alignment(horizontal="right", vertical="center")

    ws_sens.freeze_panes = "B6"
    set_col_widths(ws_sens, {"A": 34, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})

   
    # TAB 1: EXECUTIVE_SUMMARY (Linked Dynamically to Tabs 2 & 3)
    
    style_banner(
        ws_exec, 6,
        "SUPPLY CHAIN & WORKING CAPITAL EBITDA TRANSFORMATION — EXECUTIVE SUMMARY",
        "Strategy & Operations Consulting Deliverable | DataCo Global Supply Chain (180,519 Line Items | $33.05M Cumulative Net Sales)"
    )

    # KPI Scorecard (Rows 4-7)
    style_section_header(ws_exec, 4, 1, 6, "1. EXECUTIVE KPI SNAPSHOT (BASE-CASE ANNUALIZED TRANSFORMATION)")
    kpi_headers = [
        "Baseline Annual Net Sales",
        "Baseline EBITDA (12.0%)",
        "Base Target EBITDA",
        "EBITDA Margin Uplift",
        "Working Capital Freed",
        "Late Delivery Exposure"
    ]
    kpi_formulas = [
        ("=EBITDA_Bridge_Model!B8", FMT_CURR),
        ("=EBITDA_Bridge_Model!B10", FMT_CURR),
        ("=EBITDA_Bridge_Model!C40", FMT_CURR),
        ("=EBITDA_Bridge_Model!C44", FMT_BPS),
        ("=EBITDA_Bridge_Model!C46", FMT_CURR),
        ("=Operational_Diagnostics!I10", FMT_PCT),
    ]
    ws_exec.row_dimensions[5].height = 22
    ws_exec.row_dimensions[6].height = 28
    for idx, (h, (f_str, fmt)) in enumerate(zip(kpi_headers, kpi_formulas), 1):
        ch = ws_exec.cell(5, idx, h)
        ch.font = FONT_HDR
        ch.fill = FILL_BLUE
        ch.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ch.border = BORDER_HDR
        cv = ws_exec.cell(6, idx, f_str)
        cv.number_format = fmt
        cv.font = Font(name="Calibri", size=13, bold=True, color=C_GREEN_FG if idx in (3, 4, 5) else (C_RED_FG if idx == 6 else C_NAVY))
        cv.fill = FILL_GREEN if idx in (3, 4, 5) else (FILL_RED if idx == 6 else FILL_TOTAL)
        cv.alignment = Alignment(horizontal="center", vertical="center")
        cv.border = BORDER_TOTAL

    # 3-Scenario Summary Table (Rows 8-15)
    style_section_header(ws_exec, 8, 1, 6, "2. 3-SCENARIO FINANCIAL IMPACT SUMMARY (LINKED TO EBITDA BRIDGE MODEL)")
    style_table_header(ws_exec, 9, [
        "Financial Metric (Annualized FY15-17 Run-Rate)", "Baseline", "Conservative", "Base Case", "Target Case", "Key Value Driver"
    ])
    exec_rows = [
        ("Net Sales ($)", "=EBITDA_Bridge_Model!B8", "=EBITDA_Bridge_Model!B41", "=EBITDA_Bridge_Model!C41", "=EBITDA_Bridge_Model!D41", "Discount guardrail flows to Net Sales", FMT_CURR, False),
        ("Annualized EBITDA ($)", "=EBITDA_Bridge_Model!B10", "=EBITDA_Bridge_Model!B40", "=EBITDA_Bridge_Model!C40", "=EBITDA_Bridge_Model!D40", "5 operational & commercial levers", FMT_CURR, True),
        ("EBITDA Uplift vs. Baseline ($)", 0, "=EBITDA_Bridge_Model!B39", "=EBITDA_Bridge_Model!C39", "=EBITDA_Bridge_Model!D39", "+38.8% EBITDA growth in Base Case", FMT_CURR, True),
        ("EBITDA Margin (%)", "=EBITDA_Bridge_Model!B11", "=EBITDA_Bridge_Model!B43", "=EBITDA_Bridge_Model!C43", "=EBITDA_Bridge_Model!D43", "Expands from 12.01% to 16.57% (Base)", FMT_PCT2, True),
        ("Margin Expansion (bps)", 0, "=EBITDA_Bridge_Model!B44", "=EBITDA_Bridge_Model!C44", "=EBITDA_Bridge_Model!D44", "+267 bps to +636 bps range", FMT_BPS, False),
        ("Working Capital Cash Released ($)", 0, "=EBITDA_Bridge_Model!B46", "=EBITDA_Bridge_Model!C46", "=EBITDA_Bridge_Model!D46", "6-15 day DIO cut via ABC-XYZ", FMT_CURR, True),
    ]
    r = 10
    for idx, (lbl, f0, f1, f2, f3, note, fmt, is_b) in enumerate(exec_rows):
        fill = FILL_GREEN if lbl == "Annualized EBITDA ($)" else (FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE)
        ws_exec.cell(r, 1, lbl).font = FONT_BOLD if is_b else FONT_DATA
        for col_i, form in [(2, f0), (3, f1), (4, f2), (5, f3)]:
            c = ws_exec.cell(r, col_i, form)
            c.number_format = fmt
            c.font = FONT_GREEN if lbl == "Annualized EBITDA ($)" else (FONT_BOLD if is_b else FONT_DATA)
        ws_exec.cell(r, 6, note).font = FONT_NOTE
        for c_i in range(1, 7):
            cell = ws_exec.cell(r, c_i)
            cell.fill = fill
            cell.border = BORDER_TOTAL if lbl in ("Annualized EBITDA ($)", "Working Capital Cash Released ($)") else BORDER_DATA
            cell.alignment = Alignment(horizontal="left" if c_i in (1, 6) else "right", vertical="center")
        ws_exec.row_dimensions[r].height = 21
        r += 1

    # Implementation Roadmap Table (Rows 17-24)
    style_section_header(ws_exec, 17, 1, 6, "3. PRIORITIZED OPERATIONS TRANSFORMATION ROADMAP (QUICK WINS vs. STRUCTURAL FIXES)")
    style_table_header(ws_exec, 18, [
        "Horizon", "Strategic Initiative", "Target Bottleneck (Diagnosed)", "Base EBITDA Lift ($)", "Workstream Owner", "Execution Risk"
    ])
    roadmap = [
        ("Quick Win\n(0-3 Months)", "Automated TRANSFER Payment Gate", "100% of fraud (4,062) & cancels (3,692) in TRANSFER", "=EBITDA_Bridge_Model!C37", "Risk & FinOps", "Low"),
        ("Quick Win\n(0-3 Months)", "ERP >15% Discount Approval Cap", "$2.27M (60.8% of discounts) granted in >15% tiers", "=EBITDA_Bridge_Model!C34", "Commercial / Pricing", "Low"),
        ("Medium-Term\n(3-6 Months)", "Minimum Margin Floor on Loss Tail", "18.7% orders negative profit (-$3.88M cumulative)", "=EBITDA_Bridge_Model!C35", "Commercial & Finance", "Medium"),
        ("Medium-Term\n(3-6 Months)", "First & Second Class SLA Re-Engineering", "95.3% First Class & 76.6% Second Class late rates", "=EBITDA_Bridge_Model!C36", "Logistics & 3PL Ops", "Medium"),
        ("Structural\n(6-12 Months)", "ABC-XYZ Portfolio & MTO Conversion", "26 B-Z/C-Z categories = 52% portfolio, 7.6% rev", "=EBITDA_Bridge_Model!C38", "Supply Chain Planning", "Medium-High"),
    ]
    r = 19
    for idx, (hor, init, bot, f_lift, owner, risk) in enumerate(roadmap):
        fill = FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE
        vals = [hor, init, bot, f_lift, owner, risk]
        for c_i, v in enumerate(vals, 1):
            c = ws_exec.cell(r, c_i, v)
            c.font = FONT_BOLD if c_i in (1, 2, 4) else FONT_DATA
            c.fill = FILL_GREEN if (c_i == 6 and risk == "Low") else (FILL_AMBER if (c_i == 6 and "Medium" in risk) else fill)
            c.border = BORDER_DATA
            if c_i == 4:
                c.number_format = FMT_CURR
            c.alignment = Alignment(
                horizontal="right" if c_i == 4 else ("center" if c_i in (1, 6) else "left"),
                vertical="center", wrap_text=True
            )
        ws_exec.row_dimensions[r].height = 28
        r += 1

    # Total Roadmap Row (Row 24)
    ws_exec.cell(r, 1, "Total")
    ws_exec.cell(r, 2, "5 Integrated Transformation Workstreams")
    ws_exec.cell(r, 3, "Plus $197,420 One-Time Working Capital Cash Freed")
    ws_exec.cell(r, 4, "=SUM(D19:D23)").number_format = FMT_CURR
    ws_exec.cell(r, 5, "PMO / Transformation Office")
    ws_exec.cell(r, 6, "Balanced")
    for c_i in range(1, 7):
        cell = ws_exec.cell(r, c_i)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
        cell.alignment = Alignment(horizontal="right" if c_i == 4 else ("center" if c_i in (1, 6) else "left"), vertical="center")
    ws_exec.row_dimensions[r].height = 22

    # Embedded Native Excel Chart on Executive Summary: Baseline vs Scenarios EBITDA ($)
    ch_exec = BarChart()
    ch_exec.type = "col"
    ch_exec.style = 10
    ch_exec.title = "Annualized EBITDA Expansion Across Transformation Scenarios ($)"
    ch_exec.y_axis.title = "Annualized EBITDA ($)"
    ch_exec.y_axis.number_format = "$#,##0"
    ch_exec.legend = None
    data_e = Reference(ws_exec, min_col=2, max_col=5, min_row=11, max_row=11)
    cats_e = Reference(ws_exec, min_col=2, max_col=5, min_row=9, max_row=9)
    ch_exec.add_data(data_e, from_rows=True, titles_from_data=False)
    ch_exec.set_categories(cats_e)
    ch_exec.series[0].graphicalProperties.solidFill = "005EB8"
    ch_exec.width = 23
    ch_exec.height = 10.5
    ws_exec.add_chart(ch_exec, "A26")

    ws_exec.freeze_panes = "A4"
    set_col_widths(ws_exec, {"A": 26, "B": 28, "C": 34, "D": 20, "E": 22, "F": 32})

    out_path = os.path.join(DELIV_DIR, "EBITDA_Transformation_Model.xlsx")
    wb.save(out_path)
    print(f"Successfully built Excel model: {out_path} ({os.path.getsize(out_path)/1024:.1f} KB)")


if __name__ == "__main__":
    build_workbook()
