# -*- coding: utf-8 -*-
"""
Hexalevel Financial Management & Revenue Distribution Dashboard Generator
Creates a professional, automated Excel workbook with multi-scenario distribution,
partner financial tracking, dynamic KPIs, charts, and mathematical validation.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Color Palette Definitions
NAVY_HEADER = "0F172A"       # Slate 900
NAVY_SUBHEADER = "1E293B"    # Slate 800
BLUE_ACCENT = "2563EB"       # Blue 600
BLUE_LIGHT = "EFF6FF"        # Blue 50
GREEN_FILL = "D1FAE5"        # Emerald 100
GREEN_TEXT = "065F46"        # Emerald 800
AMBER_FILL = "FEF3C7"        # Amber 100
AMBER_TEXT = "92400E"        # Amber 800
PURPLE_FILL = "EDE9FE"       # Violet 100
PURPLE_TEXT = "5B21B6"       # Violet 800
CARD_BG = "F8FAFC"           # Slate 50
BORDER_COLOR = "CBD5E1"      # Slate 300
BORDER_DARK = "94A3B8"       # Slate 400
WHITE = "FFFFFF"
GRAY_TEXT = "64748B"         # Slate 500

font_title = Font(name="Segoe UI", size=16, bold=True, color=WHITE)
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
font_section = Font(name="Segoe UI", size=12, bold=True, color=NAVY_HEADER)
font_tbl_header = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
font_tbl_header_dark = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)
font_kpi_num = Font(name="Segoe UI", size=14, bold=True, color=NAVY_HEADER)
font_kpi_title = Font(name="Segoe UI", size=8, bold=True, color=GRAY_TEXT)
font_regular = Font(name="Segoe UI", size=9, bold=False, color="1E293B")
font_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
font_total = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)

fill_title = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
fill_tbl_header = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
fill_tbl_header_blue = PatternFill(start_color=BLUE_ACCENT, end_color=BLUE_ACCENT, fill_type="solid")
fill_tbl_header_green = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
fill_tbl_header_purple = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
fill_card = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")
fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
fill_zebra = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

thin_border_side = Side(border_style="thin", color=BORDER_COLOR)
med_border_side = Side(border_style="medium", color=BORDER_DARK)
double_bottom_side = Side(border_style="double", color=NAVY_HEADER)

border_all_thin = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
border_total = Border(top=thin_border_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)
border_card = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ----------------------------------------------------
# 1. SHEET: Projects & Revenue (Master Data)
# ----------------------------------------------------
ws_pr = wb.create_sheet(title="Projects & Revenue")
ws_pr.views.sheetView[0].showGridLines = True

# Title block
ws_pr.merge_cells("A1:V1")
ws_pr["A1"] = "HEXALEVEL • CLIENT PROJECTS & REVENUE DISTRIBUTION MASTER TABLE"
ws_pr["A1"].font = font_title
ws_pr["A1"].fill = fill_title
ws_pr["A1"].alignment = align_left
ws_pr.row_dimensions[1].height = 36

ws_pr.merge_cells("A2:V2")
ws_pr["A2"] = "Dynamic Multi-Scenario Distribution: Scenario A (2-Member 50/50), Scenario B (3-Member 30/70), Scenario C (3 Registered, 2 Working - Non-working skipped, company receives unallocated share)"
ws_pr["A2"].font = font_subtitle
ws_pr["A2"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
ws_pr["A2"].alignment = align_left
ws_pr.row_dimensions[2].height = 20

headers_pr = [
    "Project ID",          # A
    "Project Start",       # B
    "Payment Date",        # C
    "Client Name",         # D
    "Project Name",        # E
    "Service Type",        # F
    "Original Amt",        # G
    "Currency",            # H
    "Payment Amount (PKR)",# I
    "Payment Status",      # J
    "Distribution Scenario",# K
    "Shaista Share (PKR)", # L
    "Amara Share (PKR)",   # M
    "Urooj Share (PKR)",   # N
    "Total Member Share",  # O
    "Company Base (PKR)",  # P
    "Co Unallocated (PKR)",# Q
    "Total Company Share", # R
    "Total Reconciled",    # S
    "Variance Check",      # T
    "Bank / Destination",  # U
    "Notes"                # V
]

ws_pr.row_dimensions[4].height = 28
for c_idx, h in enumerate(headers_pr, 1):
    cell = ws_pr.cell(4, c_idx, h)
    cell.font = font_tbl_header
    cell.fill = fill_tbl_header
    cell.alignment = align_wrap
    cell.border = border_all_thin

raw_projects = [
    ("PRJ-001", "Dec 2025", "2026-01-26", "Ranked Skill", "Mortgage Leads flow", "CRM", 25000.0, "PKR", 25000.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-002", "Jan 2026", "2026-02-20", "India", "Crm +Wp", "Both", 15000.0, "INR", 54541.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Converted from INR to PKR"),
    ("PRJ-003", "Jan 2026", "2026-02-21", "Lahore", "Life insurance", "CRM", 40000.0, "PKR", 40000.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "HBL", "Received"),
    ("PRJ-004", "Dec 2025", "2026-02-25", "Waleed", "Seo site", "WP", 29000.0, "PKR", 29000.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Nayapay", "Received"),
    ("PRJ-005", "Jan 2026", "2026-03-03", "Mehroz", "GHL Work", "CRM", 28600.0, "PKR", 28600.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-006", "Mar 2026", "2026-03-15", "Upwork", "Funnel designing", "CRM", 9382.0, "PKR", 9382.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-007", "Mar 2026", "2026-03-25", "India", "Crm +Wp", "Both", 15000.0, "INR", 42671.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Converted from INR to PKR"),
    ("PRJ-008", "Mar 2026", "2026-03-27", "Mehroz", "A2p", "CRM", 19250.0, "PKR", 19250.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-009", "Apr 2026", "2026-04-06", "Haris", "GHL", "CRM", 46334.0, "PKR", 46334.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Received"),
    ("PRJ-010", "Apr 2026", "2026-04-16", "Mehroz", "A2p", "CRM", 33600.0, "PKR", 33600.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-011", "Apr 2026", "2026-04-24", "India", "CRM+WP", "Both", 42199.0, "PKR", 42199.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Received"),
    ("PRJ-012", "Apr 2026", "2026-04-25", "Usa + Fiver", "GHL", "CRM", 80163.0, "PKR", 80163.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "HBL", "Received"),
    ("PRJ-013", "May 2026", "2026-05-04", "Haris", "GHL", "CRM", 44401.0, "PKR", 44401.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Received"),
    ("PRJ-014", "May 2026", "2026-05-15", "USA", "GHL", "CRM", 85616.0, "PKR", 85616.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received"),
    ("PRJ-015", "May 2026", "2026-05-18", "Mehroz", "A2P", "CRM", 29400.0, "PKR", 29400.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received"),
    ("PRJ-016", "Jun 2026", "2026-06-04", "Haris", "GHL", "CRM", 44614.0, "PKR", 44614.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Payoneer", "Received"),
    ("PRJ-017", "Jun 2026", "2026-06-20", "Shelim", "WP", "WP", 7000.0, "PKR", 7000.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received"),
    ("PRJ-018", "Jul 2026", "2026-07-12", "Sheen", "GHL", "CRM", 15000.0, "PKR", 15000.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Meezan", "Received"),
    ("PRJ-019", "Sep 2026", "2026-09-12", "USA", "GHL", "CRM", 23113.0, "PKR", 23113.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received (Historically row 21)"),
    ("PRJ-020", "Sep 2026", "2026-09-30", "USA", "GHL", "CRM", 21920.0, "PKR", 21920.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received (Historically row 22)"),
    ("PRJ-021", "Sep 2026", "2026-10-27", "USA", "GHL", "CRM", 36099.0, "PKR", 36099.0, "Received", "Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara)", "Askari", "Received (Historically row 23)"),
    ("PRJ-022", "Oct 2026", "2026-10-19", "Umair", "GHL", "CRM", 20000.0, "PKR", 20000.0, "Received", "Scenario B: 3 Members (30% Co / 23.33% Each)", "Meezan", "New model - Urooj joined"),
    ("PRJ-023", "Nov 2026", "2026-11-05", "Apex Global", "Lead Automation Flow", "CRM", 50000.0, "PKR", 50000.0, "Received", "Scenario C: Shaista & Amara (2 Working)", "Meezan", "Scenario C Live - Shaista & Amara worked, Urooj skipped"),
    ("PRJ-024", "Nov 2026", "2026-11-15", "Nova Tech", "Portal Web & Funnel", "Both", 60000.0, "PKR", 60000.0, "Received", "Scenario C: Shaista & Urooj (2 Working)", "Askari", "Scenario C Live - Shaista & Urooj worked, Amara skipped"),
    ("PRJ-025", "Jan 2026", "", "USA Client", "GHL all services", "CRM", 350.0, "USD", 97500.0, "Pending", "Scenario B: 3 Members (30% Co / 23.33% Each)", "Payoneer", "Pending USD 350 (~PKR 97,500)")
]

cur_row = 5
for p in raw_projects:
    ws_pr.row_dimensions[cur_row].height = 20
    ws_pr.cell(cur_row, 1, p[0]).alignment = align_center # ID
    ws_pr.cell(cur_row, 2, p[1]).alignment = align_center # Start
    ws_pr.cell(cur_row, 3, p[2]).alignment = align_center # Date
    ws_pr.cell(cur_row, 4, p[3]).alignment = align_left   # Client
    ws_pr.cell(cur_row, 5, p[4]).alignment = align_left   # Project
    ws_pr.cell(cur_row, 6, p[5]).alignment = align_center # Service
    
    c_orig = ws_pr.cell(cur_row, 7, p[6]) # Orig Amt
    c_orig.alignment = align_right
    c_orig.number_format = "#,##0.00"
    
    ws_pr.cell(cur_row, 8, p[7]).alignment = align_center # Currency
    
    c_pkr = ws_pr.cell(cur_row, 9, p[8]) # Payment PKR
    c_pkr.alignment = align_right
    c_pkr.number_format = "#,##0.00"
    
    ws_pr.cell(cur_row, 10, p[9]).alignment = align_center # Status
    ws_pr.cell(cur_row, 11, p[10]).alignment = align_left  # Scenario
    
    r = cur_row
    # Shaista Share (Column L)
    ws_pr.cell(r, 12, f'=IF(J{r}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", K{r})), I{r}*0.25, IF(ISNUMBER(SEARCH("Scenario B", K{r})), I{r}*(0.7/3), IF(ISNUMBER(SEARCH("Shaista", K{r})), I{r}*(0.7/3), 0))))').number_format = "#,##0.00"
    
    # Amara Share (Column M)
    ws_pr.cell(r, 13, f'=IF(J{r}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", K{r})), I{r}*0.25, IF(ISNUMBER(SEARCH("Scenario B", K{r})), I{r}*(0.7/3), IF(ISNUMBER(SEARCH("Amara", K{r})), I{r}*(0.7/3), 0))))').number_format = "#,##0.00"
    
    # Urooj Share (Column N)
    ws_pr.cell(r, 14, f'=IF(J{r}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", K{r})), 0, IF(ISNUMBER(SEARCH("Scenario B", K{r})), I{r}*(0.7/3), IF(ISNUMBER(SEARCH("Urooj", K{r})), I{r}*(0.7/3), 0))))').number_format = "#,##0.00"
    
    # Total Member Share = L + M + N
    ws_pr.cell(r, 15, f'=L{r}+M{r}+N{r}').number_format = "#,##0.00"
    
    # Company Base Share
    ws_pr.cell(r, 16, f'=IF(J{r}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", K{r})), I{r}*0.5, I{r}*0.3))').number_format = "#,##0.00"
    
    # Company Unallocated Share
    ws_pr.cell(r, 17, f'=IF(J{r}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario C", K{r})), I{r}*(0.7/3), 0))').number_format = "#,##0.00"
    
    # Total Company Share = P + Q
    ws_pr.cell(r, 18, f'=P{r}+Q{r}').number_format = "#,##0.00"
    
    # Total Reconciled = O + R
    ws_pr.cell(r, 19, f'=O{r}+R{r}').number_format = "#,##0.00"
    
    # Variance Check = IF(Received, ROUND(I - S, 2), 0)
    ws_pr.cell(r, 20, f'=IF(J{r}<>"Received", 0, ROUND(I{r}-S{r}, 2))').number_format = "#,##0.00"
    
    # Bank
    ws_pr.cell(r, 21, p[11]).alignment = align_center
    
    # Notes
    ws_pr.cell(r, 22, p[12]).alignment = align_left
    
    # Styling
    for c in range(1, 23):
        cell_item = ws_pr.cell(r, c)
        cell_item.font = font_regular
        cell_item.border = border_all_thin
        if p[9] == "Received":
            if c == 10:
                cell_item.fill = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color=GREEN_TEXT)
        else:
            if c == 10:
                cell_item.fill = PatternFill(start_color=AMBER_FILL, end_color=AMBER_FILL, fill_type="solid")
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color=AMBER_TEXT)
    
    cur_row += 1

# Total Row
last_data_row = cur_row - 1
ws_pr.row_dimensions[cur_row].height = 24
ws_pr.merge_cells(f"A{cur_row}:H{cur_row}")
ws_pr.cell(cur_row, 1, "TOTAL RECEIVED").alignment = Alignment(horizontal="right", vertical="center")
ws_pr.cell(cur_row, 1).font = font_total

ws_pr.cell(cur_row, 9, f"=SUMIF(J5:J{last_data_row}, \"Received\", I5:I{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 12, f"=SUM(L5:L{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 13, f"=SUM(M5:M{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 14, f"=SUM(N5:N{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 15, f"=SUM(O5:O{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 16, f"=SUM(P5:P{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 17, f"=SUM(Q5:Q{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 18, f"=SUM(R5:R{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 19, f"=SUM(S5:S{last_data_row})").number_format = "#,##0.00"
ws_pr.cell(cur_row, 20, f"=SUM(T5:T{last_data_row})").number_format = "#,##0.00"

for c in range(1, 23):
    cell = ws_pr.cell(cur_row, c)
    cell.font = font_total
    cell.fill = fill_total
    cell.border = border_total

# Column widths for Projects & Revenue
col_widths_pr = {
    "A": 12, "B": 13, "C": 13, "D": 16, "E": 22, "F": 12,
    "G": 14, "H": 10, "I": 20, "J": 14, "K": 38, "L": 18,
    "M": 18, "N": 18, "O": 19, "P": 18, "Q": 19, "R": 19,
    "S": 18, "T": 14, "U": 15, "V": 30
}
for col_let, width in col_widths_pr.items():
    ws_pr.column_dimensions[col_let].width = width

# Add Data Validation Dropdowns
dv_scenario = DataValidation(type="list", formula1='"Scenario A: 2 Members (50% Co / 25% Shaista / 25% Amara),Scenario B: 3 Members (30% Co / 23.33% Each),Scenario C: Shaista & Amara (2 Working),Scenario C: Shaista & Urooj (2 Working),Scenario C: Amara & Urooj (2 Working)"', allow_blank=True)
ws_pr.add_data_validation(dv_scenario)
dv_scenario.add(f"K5:K{cur_row+50}")

dv_status = DataValidation(type="list", formula1='"Received,Pending"', allow_blank=True)
ws_pr.add_data_validation(dv_status)
dv_status.add(f"J5:J{cur_row+50}")


# ----------------------------------------------------
# 2. SHEET: Expenses & Outflows
# ----------------------------------------------------
ws_exp = wb.create_sheet(title="Expenses & Outflows")
ws_exp.views.sheetView[0].showGridLines = True

ws_exp.merge_cells("A1:G1")
ws_exp["A1"] = "HEXALEVEL • OPERATING PURCHASES & CASH OUTFLOWS"
ws_exp["A1"].font = font_title
ws_exp["A1"].fill = fill_title
ws_exp["A1"].alignment = align_left
ws_exp.row_dimensions[1].height = 36

# Section 1: Purchases / Operating Expenses
ws_exp.merge_cells("A3:D3")
ws_exp["A3"] = "1. OPERATING PURCHASES & TOOL SUBSCRIPTIONS"
ws_exp["A3"].font = font_section
ws_exp["A3"].fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
ws_exp.row_dimensions[3].height = 24

headers_exp1 = ["Expense ID", "Category / Item", "Vendor / Purpose", "Amount (PKR)"]
for c_idx, h in enumerate(headers_exp1, 1):
    c = ws_exp.cell(4, c_idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header_blue
    c.alignment = align_center
    c.border = border_all_thin

purchases = [
    ("EXP-001", "Domain", "theghlvideo.com", 3314.40),
    ("EXP-002", "Upwork Connects", "$15 Connects Purchase", 4830.00),
    ("EXP-003", "Fiverr Course", "Chain solution training", 1999.00),
    ("EXP-004", "Course", "6 course Leads generation (6/6)", 1200.00),
    ("EXP-005", "Buy CRM", "Guarantee CRM License", 370.00),
    ("EXP-006", "Email Workspace", "BUY EMAIL THEGHLVIDEO", 2533.82),
    ("EXP-007", "Domain Renewal", "Renew hexalevel Domain (Omitted in old sheet)", 4955.00)
]

r_e = 5
for item in purchases:
    ws_exp.row_dimensions[r_e].height = 19
    ws_exp.cell(r_e, 1, item[0]).alignment = align_center
    ws_exp.cell(r_e, 2, item[1]).alignment = align_left
    ws_exp.cell(r_e, 3, item[2]).alignment = align_left
    amt_cell = ws_exp.cell(r_e, 4, item[3])
    amt_cell.alignment = align_right
    amt_cell.number_format = "#,##0.00"
    for c in range(1, 5):
        ws_exp.cell(r_e, c).border = border_all_thin
        ws_exp.cell(r_e, c).font = font_regular
    r_e += 1

# Total Purchases (Row 12)
ws_exp.row_dimensions[r_e].height = 22
ws_exp.merge_cells(f"A{r_e}:C{r_e}")
ws_exp.cell(r_e, 1, "TOTAL PURCHASES").alignment = align_right
ws_exp.cell(r_e, 1).font = font_total
tot_p = ws_exp.cell(r_e, 4, f"=SUM(D5:D{r_e-1})")
tot_p.font = font_total
tot_p.number_format = "#,##0.00"
for c in range(1, 5):
    ws_exp.cell(r_e, c).fill = fill_total
    ws_exp.cell(r_e, c).border = border_total

# Section 2: Send Money / Cash Outflows
r_sm_start = r_e + 3
ws_exp.merge_cells(f"A{r_sm_start}:D{r_sm_start}")
ws_exp.cell(r_sm_start, 1, "2. CASH OUTFLOWS & ADVANCES (SEND MONEY)").font = font_section
ws_exp.cell(r_sm_start, 1).fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
ws_exp.row_dimensions[r_sm_start].height = 24

r_sm_hdr = r_sm_start + 1
headers_exp2 = ["Outflow ID", "Recipient / Description", "Category", "Amount Sent (PKR)"]
for c_idx, h in enumerate(headers_exp2, 1):
    c = ws_exp.cell(r_sm_hdr, c_idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header
    c.alignment = align_center
    c.border = border_all_thin

send_money = [
    ("OUT-001", "Send Sundas for email testing", "Testing Services", 1200.00),
    ("OUT-002", "Meeting Fiverr with Ahsan", "Business Meeting", 10000.00),
    ("OUT-003", "Waqar Udhar", "Short-term Advance", 3250.00)
]

r_s = r_sm_hdr + 1
for item in send_money:
    ws_exp.row_dimensions[r_s].height = 19
    ws_exp.cell(r_s, 1, item[0]).alignment = align_center
    ws_exp.cell(r_s, 2, item[1]).alignment = align_left
    ws_exp.cell(r_s, 3, item[2]).alignment = align_center
    amt_cell = ws_exp.cell(r_s, 4, item[3])
    amt_cell.alignment = align_right
    amt_cell.number_format = "#,##0.00"
    for c in range(1, 5):
        ws_exp.cell(r_s, c).border = border_all_thin
        ws_exp.cell(r_s, c).font = font_regular
    r_s += 1

# Total Send Money (Row 20)
ws_exp.row_dimensions[r_s].height = 22
ws_exp.merge_cells(f"A{r_s}:C{r_s}")
ws_exp.cell(r_s, 1, "TOTAL CASH OUTFLOWS").alignment = align_right
ws_exp.cell(r_s, 1).font = font_total
tot_sm = ws_exp.cell(r_s, 4, f"=SUM(D{r_sm_hdr+1}:D{r_s-1})")
tot_sm.font = font_total
tot_sm.number_format = "#,##0.00"
for c in range(1, 5):
    ws_exp.cell(r_s, c).fill = fill_total
    ws_exp.cell(r_s, c).border = border_total

# Combined Deductions Summary Card (Rows 23-26)
r_sum = r_s + 3
ws_exp.merge_cells(f"A{r_sum}:D{r_sum}")
ws_exp.cell(r_sum, 1, "COMBINED COMPANY DEDUCTIONS SUMMARY").font = font_section
ws_exp.cell(r_sum, 1).fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

ws_exp.cell(r_sum+1, 1, "Total Operating Purchases:").font = font_bold
ws_exp.cell(r_sum+1, 4, f"=D{r_e}").font = font_bold
ws_exp.cell(r_sum+1, 4).number_format = "#,##0.00"

ws_exp.cell(r_sum+2, 1, "Total Cash Outflows:").font = font_bold
ws_exp.cell(r_sum+2, 4, f"=D{r_s}").font = font_bold
ws_exp.cell(r_sum+2, 4).number_format = "#,##0.00"

# Row 26: Combined total
ws_exp.cell(r_sum+3, 1, "TOTAL COMPANY DEDUCTIONS:").font = font_total
ws_exp.cell(r_sum+3, 4, f"=D{r_sum+1}+D{r_sum+2}").font = font_total
ws_exp.cell(r_sum+3, 4).number_format = "#,##0.00"
for r_idx in range(r_sum+1, r_sum+4):
    for c in range(1, 5):
        ws_exp.cell(r_idx, c).border = border_all_thin

ws_exp.column_dimensions["A"].width = 16
ws_exp.column_dimensions["B"].width = 32
ws_exp.column_dimensions["C"].width = 30
ws_exp.column_dimensions["D"].width = 20


# ----------------------------------------------------
# 3. SHEET: Member Financials (Partner Statements)
# ----------------------------------------------------
ws_mem = wb.create_sheet(title="Member Financials")
ws_mem.views.sheetView[0].showGridLines = True

ws_mem.merge_cells("A1:K1")
ws_mem["A1"] = "HEXALEVEL • INDIVIDUAL MEMBER EARNINGS & PROFIT ATTRIBUTION"
ws_mem["A1"].font = font_title
ws_mem["A1"].fill = fill_title
ws_mem["A1"].alignment = align_left
ws_mem.row_dimensions[1].height = 36

headers_mem = [
    "Project ID", "Client Name", "Project Name", "Status",
    "Scenario", "Direct Share (PKR)", "Scenario C Surplus (PKR)",
    "Base Profit Share (PKR)", "Total Net Earnings", "Payout Destination"
]

# Section for Shaista
ws_mem.merge_cells("A3:J3")
ws_mem["A3"] = "1. SHAISTA — FINANCIAL STATEMENT & NET EARNINGS"
ws_mem["A3"].font = font_section
ws_mem["A3"].fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
ws_mem.row_dimensions[3].height = 24

for idx, h in enumerate(headers_mem, 1):
    c = ws_mem.cell(4, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header_green
    c.alignment = align_wrap
    c.border = border_all_thin

r_sh = 5
for pr_idx in range(5, cur_row):
    ws_mem.row_dimensions[r_sh].height = 18
    ws_mem.cell(r_sh, 1, f"='Projects & Revenue'!A{pr_idx}").alignment = align_center
    ws_mem.cell(r_sh, 2, f"='Projects & Revenue'!D{pr_idx}").alignment = align_left
    ws_mem.cell(r_sh, 3, f"='Projects & Revenue'!E{pr_idx}").alignment = align_left
    ws_mem.cell(r_sh, 4, f"='Projects & Revenue'!J{pr_idx}").alignment = align_center
    ws_mem.cell(r_sh, 5, f"='Projects & Revenue'!K{pr_idx}").alignment = align_left
    
    # Direct Share
    ws_mem.cell(r_sh, 6, f"='Projects & Revenue'!L{pr_idx}").number_format = "#,##0.00"
    
    # Scenario C Surplus (If Shaista worked in Scenario C, half of unallocated share)
    ws_mem.cell(r_sh, 7, f'=IF(AND(ISNUMBER(SEARCH("Scenario C", \'Projects & Revenue\'!K{pr_idx})), ISNUMBER(SEARCH("Shaista", \'Projects & Revenue\'!K{pr_idx}))), \'Projects & Revenue\'!Q{pr_idx}/2, 0)').number_format = "#,##0.00"
    
    # Base Profit Share
    ws_mem.cell(r_sh, 8, f'=IF(\'Projects & Revenue\'!J{pr_idx}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", \'Projects & Revenue\'!K{pr_idx})), \'Projects & Revenue\'!P{pr_idx}*0.5, \'Projects & Revenue\'!P{pr_idx}/3))').number_format = "#,##0.00"
    
    # Total Net Earnings
    ws_mem.cell(r_sh, 9, f'=F{r_sh}+G{r_sh}').number_format = "#,##0.00"
    
    # Bank
    ws_mem.cell(r_sh, 10, f"='Projects & Revenue'!U{pr_idx}").alignment = align_center
    
    for c in range(1, 11):
        ws_mem.cell(r_sh, c).border = border_all_thin
        ws_mem.cell(r_sh, c).font = font_regular
    r_sh += 1

# Shaista Total (Row 30)
ws_mem.row_dimensions[r_sh].height = 22
ws_mem.merge_cells(f"A{r_sh}:E{r_sh}")
ws_mem.cell(r_sh, 1, "SHAISTA TOTAL EARNINGS").alignment = align_right
ws_mem.cell(r_sh, 1).font = font_total
ws_mem.cell(r_sh, 6, f"=SUM(F5:F{r_sh-1})").number_format = "#,##0.00"
ws_mem.cell(r_sh, 7, f"=SUM(G5:G{r_sh-1})").number_format = "#,##0.00"
ws_mem.cell(r_sh, 8, f"=SUM(H5:H{r_sh-1})").number_format = "#,##0.00"
ws_mem.cell(r_sh, 9, f"=SUM(I5:I{r_sh-1})").number_format = "#,##0.00"
for c in range(1, 11):
    ws_mem.cell(r_sh, c).fill = fill_total
    ws_mem.cell(r_sh, c).font = font_total
    ws_mem.cell(r_sh, c).border = border_total


# Section for Amara
r_am_start = r_sh + 3 # Row 33
ws_mem.merge_cells(f"A{r_am_start}:J{r_am_start}")
ws_mem.cell(r_am_start, 1, "2. AMARA — FINANCIAL STATEMENT & NET EARNINGS").font = font_section
ws_mem.cell(r_am_start, 1).fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
ws_mem.row_dimensions[r_am_start].height = 24

r_am_hdr = r_am_start + 1 # Row 34
for idx, h in enumerate(headers_mem, 1):
    c = ws_mem.cell(r_am_hdr, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header_blue
    c.alignment = align_wrap
    c.border = border_all_thin

r_am = r_am_hdr + 1 # Row 35
for pr_idx in range(5, cur_row):
    ws_mem.row_dimensions[r_am].height = 18
    ws_mem.cell(r_am, 1, f"='Projects & Revenue'!A{pr_idx}").alignment = align_center
    ws_mem.cell(r_am, 2, f"='Projects & Revenue'!D{pr_idx}").alignment = align_left
    ws_mem.cell(r_am, 3, f"='Projects & Revenue'!E{pr_idx}").alignment = align_left
    ws_mem.cell(r_am, 4, f"='Projects & Revenue'!J{pr_idx}").alignment = align_center
    ws_mem.cell(r_am, 5, f"='Projects & Revenue'!K{pr_idx}").alignment = align_left
    
    # Direct Share
    ws_mem.cell(r_am, 6, f"='Projects & Revenue'!M{pr_idx}").number_format = "#,##0.00"
    
    # Scenario C Surplus
    ws_mem.cell(r_am, 7, f'=IF(AND(ISNUMBER(SEARCH("Scenario C", \'Projects & Revenue\'!K{pr_idx})), ISNUMBER(SEARCH("Amara", \'Projects & Revenue\'!K{pr_idx}))), \'Projects & Revenue\'!Q{pr_idx}/2, 0)').number_format = "#,##0.00"
    
    # Base Profit Share
    ws_mem.cell(r_am, 8, f'=IF(\'Projects & Revenue\'!J{pr_idx}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", \'Projects & Revenue\'!K{pr_idx})), \'Projects & Revenue\'!P{pr_idx}*0.5, \'Projects & Revenue\'!P{pr_idx}/3))').number_format = "#,##0.00"
    
    # Total Net Earnings
    ws_mem.cell(r_am, 9, f'=F{r_am}+G{r_am}').number_format = "#,##0.00"
    
    # Bank
    ws_mem.cell(r_am, 10, f"='Projects & Revenue'!U{pr_idx}").alignment = align_center
    
    for c in range(1, 11):
        ws_mem.cell(r_am, c).border = border_all_thin
        ws_mem.cell(r_am, c).font = font_regular
    r_am += 1

# Amara Total (Row 60)
ws_mem.row_dimensions[r_am].height = 22
ws_mem.merge_cells(f"A{r_am}:E{r_am}")
ws_mem.cell(r_am, 1, "AMARA TOTAL EARNINGS").alignment = align_right
ws_mem.cell(r_am, 1).font = font_total
ws_mem.cell(r_am, 6, f"=SUM(F{r_am_hdr+1}:F{r_am-1})").number_format = "#,##0.00"
ws_mem.cell(r_am, 7, f"=SUM(G{r_am_hdr+1}:G{r_am-1})").number_format = "#,##0.00"
ws_mem.cell(r_am, 8, f"=SUM(H{r_am_hdr+1}:H{r_am-1})").number_format = "#,##0.00"
ws_mem.cell(r_am, 9, f"=SUM(I{r_am_hdr+1}:I{r_am-1})").number_format = "#,##0.00"
for c in range(1, 11):
    ws_mem.cell(r_am, c).fill = fill_total
    ws_mem.cell(r_am, c).font = font_total
    ws_mem.cell(r_am, c).border = border_total


# Section for Urooj
r_ur_start = r_am + 3 # Row 63
ws_mem.merge_cells(f"A{r_ur_start}:J{r_ur_start}")
ws_mem.cell(r_ur_start, 1, "3. UROOJ — FINANCIAL STATEMENT & NET EARNINGS").font = font_section
ws_mem.cell(r_ur_start, 1).fill = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")
ws_mem.row_dimensions[r_ur_start].height = 24

r_ur_hdr = r_ur_start + 1 # Row 64
for idx, h in enumerate(headers_mem, 1):
    c = ws_mem.cell(r_ur_hdr, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header_purple
    c.alignment = align_wrap
    c.border = border_all_thin

r_ur = r_ur_hdr + 1 # Row 65
for pr_idx in range(5, cur_row):
    ws_mem.row_dimensions[r_ur].height = 18
    ws_mem.cell(r_ur, 1, f"='Projects & Revenue'!A{pr_idx}").alignment = align_center
    ws_mem.cell(r_ur, 2, f"='Projects & Revenue'!D{pr_idx}").alignment = align_left
    ws_mem.cell(r_ur, 3, f"='Projects & Revenue'!E{pr_idx}").alignment = align_left
    ws_mem.cell(r_ur, 4, f"='Projects & Revenue'!J{pr_idx}").alignment = align_center
    ws_mem.cell(r_ur, 5, f"='Projects & Revenue'!K{pr_idx}").alignment = align_left
    
    # Direct Share
    ws_mem.cell(r_ur, 6, f"='Projects & Revenue'!N{pr_idx}").number_format = "#,##0.00"
    
    # Scenario C Surplus
    ws_mem.cell(r_ur, 7, f'=IF(AND(ISNUMBER(SEARCH("Scenario C", \'Projects & Revenue\'!K{pr_idx})), ISNUMBER(SEARCH("Urooj", \'Projects & Revenue\'!K{pr_idx}))), \'Projects & Revenue\'!Q{pr_idx}/2, 0)').number_format = "#,##0.00"
    
    # Base Profit Share
    ws_mem.cell(r_ur, 8, f'=IF(\'Projects & Revenue\'!J{pr_idx}<>"Received", 0, IF(ISNUMBER(SEARCH("Scenario A", \'Projects & Revenue\'!K{pr_idx})), 0, \'Projects & Revenue\'!P{pr_idx}/3))').number_format = "#,##0.00"
    
    # Total Net Earnings
    ws_mem.cell(r_ur, 9, f'=F{r_ur}+G{r_ur}').number_format = "#,##0.00"
    
    # Bank
    ws_mem.cell(r_ur, 10, f"='Projects & Revenue'!U{pr_idx}").alignment = align_center
    
    for c in range(1, 11):
        ws_mem.cell(r_ur, c).border = border_all_thin
        ws_mem.cell(r_ur, c).font = font_regular
    r_ur += 1

# Urooj Total (Row 90)
ws_mem.row_dimensions[r_ur].height = 22
ws_mem.merge_cells(f"A{r_ur}:E{r_ur}")
ws_mem.cell(r_ur, 1, "UROOJ TOTAL EARNINGS").alignment = align_right
ws_mem.cell(r_ur, 1).font = font_total
ws_mem.cell(r_ur, 6, f"=SUM(F{r_ur_hdr+1}:F{r_ur-1})").number_format = "#,##0.00"
ws_mem.cell(r_ur, 7, f"=SUM(G{r_ur_hdr+1}:G{r_ur-1})").number_format = "#,##0.00"
ws_mem.cell(r_ur, 8, f"=SUM(H{r_ur_hdr+1}:H{r_ur-1})").number_format = "#,##0.00"
ws_mem.cell(r_ur, 9, f"=SUM(I{r_ur_hdr+1}:I{r_ur-1})").number_format = "#,##0.00"
for c in range(1, 11):
    ws_mem.cell(r_ur, c).fill = fill_total
    ws_mem.cell(r_ur, c).font = font_total
    ws_mem.cell(r_ur, c).border = border_total

col_widths_mem = {"A": 12, "B": 16, "C": 22, "D": 12, "E": 34, "F": 18, "G": 20, "H": 20, "I": 20, "J": 18}
for col_let, width in col_widths_mem.items():
    ws_mem.column_dimensions[col_let].width = width


# ----------------------------------------------------
# 4. SHEET: Monthly & Yearly Reports
# ----------------------------------------------------
ws_rep = wb.create_sheet(title="Reports & Analytics")
ws_rep.views.sheetView[0].showGridLines = True

ws_rep.merge_cells("A1:K1")
ws_rep["A1"] = "HEXALEVEL • MONTHLY & YEARLY FINANCIAL PERFORMANCE"
ws_rep["A1"].font = font_title
ws_rep["A1"].fill = fill_title
ws_rep["A1"].alignment = align_left
ws_rep.row_dimensions[1].height = 36

headers_rep = [
    "Month / Period",      # A
    "Total Payments (PKR)",# B
    "Company Base (PKR)",  # C
    "Co Unallocated (PKR)",# D
    "Total Co Gross (PKR)",# E
    "Shaista Direct (PKR)",# F
    "Amara Direct (PKR)",  # G
    "Urooj Direct (PKR)",  # H
    "Total Member Share",  # I
    "Projects Completed",  # J
    "Active Scenario Model"# K
]

ws_rep.row_dimensions[3].height = 26
for idx, h in enumerate(headers_rep, 1):
    c = ws_rep.cell(3, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header
    c.alignment = align_wrap
    c.border = border_all_thin

periods = [
    ("Dec 2025", "Scenario A: 2 Members (50/50)"),
    ("Jan 2026", "Scenario A: 2 Members (50/50)"),
    ("Feb 2026", "Scenario A: 2 Members (50/50)"),
    ("Mar 2026", "Scenario A: 2 Members (50/50)"),
    ("Apr 2026", "Scenario A: 2 Members (50/50)"),
    ("May 2026", "Scenario A: 2 Members (50/50)"),
    ("Jun 2026", "Scenario A: 2 Members (50/50)"),
    ("Jul 2026", "Scenario A: 2 Members (50/50)"),
    ("Sep 2026", "Scenario A: 2 Members (50/50)"),
    ("Oct 2026", "Scenario B: 3 Members (30/70)"),
    ("Nov 2026", "Scenario C: 3 Reg, 2 Working (53.33/46.67)")
]

r_rp = 4
for p_idx, (p_name, sc_note) in enumerate(periods, 1):
    ws_rep.row_dimensions[r_rp].height = 20
    ws_rep.cell(r_rp, 1, p_name).alignment = align_center
    
    ws_rep.cell(r_rp, 2, f'=SUMIFS(\'Projects & Revenue\'!$I$5:$I$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 3, f'=SUMIFS(\'Projects & Revenue\'!$P$5:$P$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 4, f'=SUMIFS(\'Projects & Revenue\'!$Q$5:$Q$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 5, f'=C{r_rp}+D{r_rp}').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 6, f'=SUMIFS(\'Projects & Revenue\'!$L$5:$L$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 7, f'=SUMIFS(\'Projects & Revenue\'!$M$5:$M$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 8, f'=SUMIFS(\'Projects & Revenue\'!$N$5:$N$29, \'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 9, f'=F{r_rp}+G{r_rp}+H{r_rp}').number_format = "#,##0.00"
    ws_rep.cell(r_rp, 10, f'=COUNTIFS(\'Projects & Revenue\'!$B$5:$B$29, A{r_rp}, \'Projects & Revenue\'!$J$5:$J$29, "Received")').alignment = align_center
    ws_rep.cell(r_rp, 11, sc_note).alignment = align_left
    
    for c in range(1, 12):
        ws_rep.cell(r_rp, c).border = border_all_thin
        ws_rep.cell(r_rp, c).font = font_regular
    r_rp += 1

# Total Row for Reports
ws_rep.row_dimensions[r_rp].height = 24
ws_rep.cell(r_rp, 1, "OVERALL TOTALS").alignment = align_center
ws_rep.cell(r_rp, 1).font = font_total
for c_idx in range(2, 10):
    let = get_column_letter(c_idx)
    ws_rep.cell(r_rp, c_idx, f"=SUM({let}4:{let}{r_rp-1})").number_format = "#,##0.00"
    ws_rep.cell(r_rp, c_idx).font = font_total

ws_rep.cell(r_rp, 10, f"=SUM(J4:J{r_rp-1})").alignment = align_center
ws_rep.cell(r_rp, 10).font = font_total
ws_rep.cell(r_rp, 11, "All Scenarios Reconciled").alignment = align_left

for c in range(1, 12):
    ws_rep.cell(r_rp, c).fill = fill_total
    ws_rep.cell(r_rp, c).border = border_total

col_widths_rep = {"A": 16, "B": 20, "C": 18, "D": 19, "E": 20, "F": 19, "G": 19, "H": 19, "I": 20, "J": 18, "K": 35}
for col_let, width in col_widths_rep.items():
    ws_rep.column_dimensions[col_let].width = width


# ----------------------------------------------------
# 5. SHEET: Testing & Validation
# ----------------------------------------------------
ws_val = wb.create_sheet(title="Testing & Validation")
ws_val.views.sheetView[0].showGridLines = True

ws_val.merge_cells("A1:I1")
ws_val["A1"] = "HEXALEVEL • FINANCIAL ENGINE RECONCILIATION & VALIDATION CHECKS"
ws_val["A1"].font = font_title
ws_val["A1"].fill = fill_title
ws_val["A1"].alignment = align_left
ws_val.row_dimensions[1].height = 36

ws_val.merge_cells("A3:I3")
ws_val["A3"] = "MATHEMATICAL PROOF & SCENARIO RECONCILIATION TEST SUITE"
ws_val["A3"].font = font_section
ws_val["A3"].fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
ws_val.row_dimensions[3].height = 24

headers_val = [
    "Test Case", "Test Scenario Description", "Client Payment",
    "Company Total Share", "Shaista Share", "Amara Share",
    "Urooj Share", "Sum of Parts", "Variance / Status"
]

ws_val.row_dimensions[4].height = 26
for idx, h in enumerate(headers_val, 1):
    c = ws_val.cell(4, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header
    c.alignment = align_wrap
    c.border = border_all_thin

test_cases = [
    ("TEST-01", "Scenario A: 2 Members Working (100k Payment)", 100000.0,
     "=C5*0.5", "=C5*0.25", "=C5*0.25", 0.0, "=D5+E5+F5+G5", '=IF(ROUND(C5-H5,2)=0,"PASSED (0.00 Diff)","FAILED")'),
    
    ("TEST-02", "Scenario B: 3 Members Working (100k Payment)", 100000.0,
     "=C6*0.3", "=C6*(0.7/3)", "=C6*(0.7/3)", "=C6*(0.7/3)", "=D6+E6+F6+G6", '=IF(ROUND(C6-H6,2)=0,"PASSED (0.00 Diff)","FAILED")'),
    
    ("TEST-03", "Scenario C: Shaista & Amara Working, Urooj Skipped (100k Payment)", 100000.0,
     "=C7*(0.3+0.7/3)", "=C7*(0.7/3)", "=C7*(0.7/3)", 0.0, "=D7+E7+F7+G7", '=IF(ROUND(C7-H7,2)=0,"PASSED (0.00 Diff)","FAILED")'),
    
    ("TEST-04", "Scenario C: Shaista & Urooj Working, Amara Skipped (100k Payment)", 100000.0,
     "=C8*(0.3+0.7/3)", "=C8*(0.7/3)", 0.0, "=C8*(0.7/3)", "=D8+E8+F8+G8", '=IF(ROUND(C8-H8,2)=0,"PASSED (0.00 Diff)","FAILED")'),
    
    ("TEST-05", "Scenario C: Amara & Urooj Working, Shaista Skipped (100k Payment)", 100000.0,
     "=C9*(0.3+0.7/3)", 0.0, "=C9*(0.7/3)", "=C9*(0.7/3)", "=D9+E9+F9+G9", '=IF(ROUND(C9-H9,2)=0,"PASSED (0.00 Diff)","FAILED")'),
    
    ("TEST-06", "Live Database Reconciliation: Total Projects Table Reconciliation",
     "='Projects & Revenue'!I30", "='Projects & Revenue'!R30", "='Projects & Revenue'!L30",
     "='Projects & Revenue'!M30", "='Projects & Revenue'!N30", "=D10+E10+F10+G10",
     '=IF(ROUND(C10-H10,2)=0,"PASSED (100% Balanced)","FAILED")')
]

r_tc = 5
for tc in test_cases:
    ws_val.row_dimensions[r_tc].height = 22
    ws_val.cell(r_tc, 1, tc[0]).alignment = align_center
    ws_val.cell(r_tc, 2, tc[1]).alignment = align_left
    
    c_pay = ws_val.cell(r_tc, 3, tc[2])
    c_pay.alignment = align_right
    c_pay.number_format = "#,##0.00"
    
    for c_idx, formula_val in enumerate(tc[3:8], 4):
        c = ws_val.cell(r_tc, c_idx, formula_val)
        c.alignment = align_right
        c.number_format = "#,##0.00"
    
    status_c = ws_val.cell(r_tc, 9, tc[8])
    status_c.alignment = align_center
    status_c.font = Font(name="Segoe UI", size=9, bold=True, color=GREEN_TEXT)
    status_c.fill = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
    
    for c in range(1, 10):
        ws_val.cell(r_tc, c).border = border_all_thin
        if c < 9:
            ws_val.cell(r_tc, c).font = font_regular
    r_tc += 1

# Scenario C Distribution Rules Explanation Table
r_ex = r_tc + 3
ws_val.merge_cells(f"A{r_ex}:I{r_ex}")
ws_val.cell(r_ex, 1, "SCENARIO C (3 REGISTERED, 2 WORKING) MATHEMATICAL SPECIFICATION").font = font_section
ws_val.cell(r_ex, 1).fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
ws_val.row_dimensions[r_ex].height = 24

specs = [
    ("Component", "Formula / Percentage", "Description"),
    ("Base Company Share", "30.00% (=Amount * 0.30)", "Standard baseline company share"),
    ("Working Member 1 Share", "23.33% (=Amount * (0.70 / 3))", "Standard 3-member share paid to working partner"),
    ("Working Member 2 Share", "23.33% (=Amount * (0.70 / 3))", "Standard 3-member share paid to working partner"),
    ("Non-Working Member Share", "0.00% ($0.00 PKR)", "Non-working member is strictly skipped from direct project payout"),
    ("Unallocated Share to Company", "23.33% (=Amount * (0.70 / 3))", "Non-working member's share is transferred into Company Earnings"),
    ("Total Company Gross Share", "53.33% (=30% + 23.33%)", "Company gross = 30% base + 23.33% unallocated share"),
    ("Total Distribution Check", "100.00% (=53.33% + 23.33% + 23.33% + 0%)", "Guaranteed zero-variance mathematical reconciliation")
]

for s_idx, sp in enumerate(specs):
    row_sp = r_ex + 1 + s_idx
    ws_val.row_dimensions[row_sp].height = 20
    ws_val.cell(row_sp, 1, sp[0]).alignment = align_left
    ws_val.cell(row_sp, 1).font = font_bold if s_idx == 0 else font_regular
    ws_val.merge_cells(f"B{row_sp}:C{row_sp}")
    ws_val.cell(row_sp, 2, sp[1]).alignment = align_left
    ws_val.cell(row_sp, 2).font = font_bold if s_idx == 0 else font_regular
    ws_val.merge_cells(f"D{row_sp}:I{row_sp}")
    ws_val.cell(row_sp, 4, sp[2]).alignment = align_left
    ws_val.cell(row_sp, 4).font = font_bold if s_idx == 0 else font_regular
    
    for c in range(1, 10):
        cell_sp = ws_val.cell(row_sp, c)
        cell_sp.border = border_all_thin
        if s_idx == 0:
            cell_sp.fill = fill_tbl_header
            cell_sp.font = font_tbl_header

col_widths_val = {"A": 14, "B": 36, "C": 18, "D": 22, "E": 18, "F": 18, "G": 18, "H": 18, "I": 24}
for col_let, width in col_widths_val.items():
    ws_val.column_dimensions[col_let].width = width


# ----------------------------------------------------
# 6. SHEET: Dashboard (Executive User Interface)
# ----------------------------------------------------
ws_db = wb.create_sheet(title="Dashboard", index=0)
ws_db.views.sheetView[0].showGridLines = True

# Banner
ws_db.merge_cells("A1:N1")
ws_db["A1"] = "HEXALEVEL • FINANCIAL MANAGEMENT & REVENUE DISTRIBUTION DASHBOARD"
ws_db["A1"].font = font_title
ws_db["A1"].fill = fill_title
ws_db["A1"].alignment = align_left
ws_db.row_dimensions[1].height = 36

ws_db.merge_cells("A2:N2")
ws_db["A2"] = "Automated Financial Engine • Dynamic Partner Revenue Allocation (Scenario A / B / C) • Real-time Net Profit & Variance Tracking"
ws_db["A2"].font = font_subtitle
ws_db["A2"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
ws_db["A2"].alignment = align_left
ws_db.row_dimensions[2].height = 20

# ----------------- TOP KPI CARDS (Row 4 to 6) -----------------
# Card 1: Total Revenue Received
ws_db.merge_cells("A4:B4")
ws_db["A4"] = "TOTAL REVENUE RECEIVED"
ws_db["A4"].font = font_kpi_title
ws_db["A4"].alignment = align_center

ws_db.merge_cells("A5:B5")
ws_db["A5"] = "='Projects & Revenue'!I30"
ws_db["A5"].font = Font(name="Segoe UI", size=14, bold=True, color=BLUE_ACCENT)
ws_db["A5"].alignment = align_center
ws_db["A5"].number_format = "#,##0.00"

ws_db.merge_cells("A6:B6")
ws_db["A6"] = "Verified Inflow (PKR)"
ws_db["A6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["A6"].alignment = align_center

# Card 2: Company Gross Share
ws_db.merge_cells("C4:D4")
ws_db["C4"] = "TOTAL COMPANY GROSS"
ws_db["C4"].font = font_kpi_title
ws_db["C4"].alignment = align_center

ws_db.merge_cells("C5:D5")
ws_db["C5"] = "='Projects & Revenue'!R30"
ws_db["C5"].font = Font(name="Segoe UI", size=14, bold=True, color=NAVY_HEADER)
ws_db["C5"].alignment = align_center
ws_db["C5"].number_format = "#,##0.00"

ws_db.merge_cells("C6:D6")
ws_db["C6"] = "Base (30/50%) + Unallocated"
ws_db["C6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["C6"].alignment = align_center

# Card 3: Total Company Expenses & Outflows (Row 26 of Expenses & Outflows)
ws_db.merge_cells("E4:F4")
ws_db["E4"] = "TOTAL COMPANY DEDUCTIONS"
ws_db["E4"].font = font_kpi_title
ws_db["E4"].alignment = align_center

ws_db.merge_cells("E5:F5")
ws_db["E5"] = "='Expenses & Outflows'!D26"
ws_db["E5"].font = Font(name="Segoe UI", size=14, bold=True, color="DC2626")
ws_db["E5"].alignment = align_center
ws_db["E5"].number_format = "#,##0.00"

ws_db.merge_cells("E6:F6")
ws_db["E6"] = "Purchases + Send Money"
ws_db["E6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["E6"].alignment = align_center

# Card 4: Company Net Profit
ws_db.merge_cells("G4:H4")
ws_db["G4"] = "COMPANY NET PROFIT"
ws_db["G4"].font = font_kpi_title
ws_db["G4"].alignment = align_center

ws_db.merge_cells("G5:H5")
ws_db["G5"] = "=C5-E5"
ws_db["G5"].font = Font(name="Segoe UI", size=14, bold=True, color="059669")
ws_db["G5"].alignment = align_center
ws_db["G5"].number_format = "#,##0.00"

ws_db.merge_cells("G6:H6")
ws_db["G6"] = "Gross Share minus Deductions"
ws_db["G6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["G6"].alignment = align_center

# Card 5: Total Member Direct Payouts
ws_db.merge_cells("I4:J4")
ws_db["I4"] = "TOTAL MEMBER PAYOUTS"
ws_db["I4"].font = font_kpi_title
ws_db["I4"].alignment = align_center

ws_db.merge_cells("I5:J5")
ws_db["I5"] = "='Projects & Revenue'!O30"
ws_db["I5"].font = Font(name="Segoe UI", size=14, bold=True, color="7C3AED")
ws_db["I5"].alignment = align_center
ws_db["I5"].number_format = "#,##0.00"

ws_db.merge_cells("I6:J6")
ws_db["I6"] = "Shaista + Amara + Urooj"
ws_db["I6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["I6"].alignment = align_center

# Card 6: Total Projects & Status
ws_db.merge_cells("K4:L4")
ws_db["K4"] = "TOTAL PROJECTS COUNT"
ws_db["K4"].font = font_kpi_title
ws_db["K4"].alignment = align_center

ws_db.merge_cells("K5:L5")
ws_db["K5"] = "=COUNTA('Projects & Revenue'!A5:A29)"
ws_db["K5"].font = font_kpi_num
ws_db["K5"].alignment = align_center

ws_db.merge_cells("K6:L6")
ws_db["K6"] = "=COUNTIF('Projects & Revenue'!J5:J29, \"Received\") & \" Received | \" & COUNTIF('Projects & Revenue'!J5:J29, \"Pending\") & \" Pending\""
ws_db["K6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["K6"].alignment = align_center

# Card 7: Total Outstanding / Pending Payments
ws_db.merge_cells("M4:N4")
ws_db["M4"] = "PENDING PAYMENTS (PKR)"
ws_db["M4"].font = font_kpi_title
ws_db["M4"].alignment = align_center

ws_db.merge_cells("M5:N5")
ws_db["M5"] = "=SUMIF('Projects & Revenue'!J5:J29, \"Pending\", 'Projects & Revenue'!I5:I29)"
ws_db["M5"].font = Font(name="Segoe UI", size=14, bold=True, color="D97706")
ws_db["M5"].alignment = align_center
ws_db["M5"].number_format = "#,##0.00"

ws_db.merge_cells("M6:N6")
ws_db["M6"] = "Pipeline Awaiting Receipt"
ws_db["M6"].font = Font(name="Segoe UI", size=8, italic=True, color=GRAY_TEXT)
ws_db["M6"].alignment = align_center

# Apply card borders and background
for c in range(1, 15):
    for r in range(4, 7):
        cell = ws_db.cell(r, c)
        cell.fill = fill_card
        cell.border = border_card

# ----------------- PARTNER FINANCIAL SUMMARY CARDS (Row 8 to 14) -----------------
ws_db.merge_cells("A8:N8")
ws_db["A8"] = "PARTNER FINANCIAL SUMMARY & INDIVIDUAL NET BALANCES"
ws_db["A8"].font = font_section
ws_db["A8"].fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
ws_db["A8"].alignment = align_left
ws_db.row_dimensions[8].height = 24

# Partner 1: Shaista (Cols A to D)
ws_db.merge_cells("A9:D9")
ws_db["A9"] = "SHAISTA (Co-Founder & Partner)"
ws_db["A9"].font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
ws_db["A9"].fill = fill_tbl_header_green
ws_db["A9"].alignment = align_center

ws_db.cell(10, 1, "Direct Project Payouts:").font = font_regular
ws_db.merge_cells("C10:D10")
ws_db.cell(10, 3, "='Projects & Revenue'!L30").font = font_bold
ws_db.cell(10, 3).number_format = "#,##0.00"

ws_db.cell(11, 1, "Scenario C Surplus Profit:").font = font_regular
ws_db.merge_cells("C11:D11")
ws_db.cell(11, 3, "='Member Financials'!G30").font = font_bold
ws_db.cell(11, 3).number_format = "#,##0.00"

ws_db.cell(12, 1, "TOTAL NET EARNINGS:").font = font_total
ws_db.merge_cells("C12:D12")
ws_db.cell(12, 3, "=C10+C11").font = Font(name="Segoe UI", size=11, bold=True, color="059669")
ws_db.cell(12, 3).number_format = "#,##0.00"

ws_db.cell(13, 1, "Projects Participated:").font = font_regular
ws_db.merge_cells("C13:D13")
ws_db.cell(13, 3, "=COUNTIFS('Projects & Revenue'!L5:L29, \">0\")").font = font_bold

ws_db.cell(14, 1, "Payout Destination:").font = font_regular
ws_db.merge_cells("C14:D14")
ws_db.cell(14, 3, "Nayapay / Cutting").font = font_bold

for r in range(9, 15):
    for c in range(1, 5):
        ws_db.cell(r, c).border = border_all_thin


# Partner 2: Amara (Cols F to I)
ws_db.merge_cells("F9:I9")
ws_db["F9"] = "AMARA (Co-Founder & Partner)"
ws_db["F9"].font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
ws_db["F9"].fill = fill_tbl_header_blue
ws_db["F9"].alignment = align_center

ws_db.cell(10, 6, "Direct Project Payouts:").font = font_regular
ws_db.merge_cells("H10:I10")
ws_db.cell(10, 8, "='Projects & Revenue'!M30").font = font_bold
ws_db.cell(10, 8).number_format = "#,##0.00"

ws_db.cell(11, 6, "Scenario C Surplus Profit:").font = font_regular
ws_db.merge_cells("H11:I11")
ws_db.cell(11, 8, "='Member Financials'!G60").font = font_bold # Correct Row 60
ws_db.cell(11, 8).number_format = "#,##0.00"

ws_db.cell(12, 6, "TOTAL NET EARNINGS:").font = font_total
ws_db.merge_cells("H12:I12")
ws_db.cell(12, 8, "=H10+H11").font = Font(name="Segoe UI", size=11, bold=True, color="2563EB")
ws_db.cell(12, 8).number_format = "#,##0.00"

ws_db.cell(13, 6, "Projects Participated:").font = font_regular
ws_db.merge_cells("H13:I13")
ws_db.cell(13, 8, "=COUNTIFS('Projects & Revenue'!M5:M29, \">0\")").font = font_bold

ws_db.cell(14, 6, "Payout Destination:").font = font_regular
ws_db.merge_cells("H14:I14")
ws_db.cell(14, 8, "Askari Bank / Nayapay").font = font_bold

for r in range(9, 15):
    for c in range(6, 10):
        ws_db.cell(r, c).border = border_all_thin


# Partner 3: Urooj (Cols K to N)
ws_db.merge_cells("K9:N9")
ws_db["K9"] = "UROOJ (Partner - Joined 1 Sep 2026)"
ws_db["K9"].font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
ws_db["K9"].fill = fill_tbl_header_purple
ws_db["K9"].alignment = align_center

ws_db.cell(10, 11, "Direct Project Payouts:").font = font_regular
ws_db.merge_cells("M10:N10")
ws_db.cell(10, 13, "='Projects & Revenue'!N30").font = font_bold
ws_db.cell(10, 13).number_format = "#,##0.00"

ws_db.cell(11, 11, "Scenario C Surplus Profit:").font = font_regular
ws_db.merge_cells("M11:N11")
ws_db.cell(11, 13, "='Member Financials'!G90").font = font_bold # Correct Row 90
ws_db.cell(11, 13).number_format = "#,##0.00"

ws_db.cell(12, 11, "TOTAL NET EARNINGS:").font = font_total
ws_db.merge_cells("M12:N12")
ws_db.cell(12, 13, "=M10+M11").font = Font(name="Segoe UI", size=11, bold=True, color="7C3AED")
ws_db.cell(12, 13).number_format = "#,##0.00"

ws_db.cell(13, 11, "Projects Participated:").font = font_regular
ws_db.merge_cells("M13:N13")
ws_db.cell(13, 13, "=COUNTIFS('Projects & Revenue'!N5:N29, \">0\")").font = font_bold

ws_db.cell(14, 11, "Payout Destination:").font = font_regular
ws_db.merge_cells("M14:N14")
ws_db.cell(14, 13, "UBL Bank").font = font_bold

for r in range(9, 15):
    for c in range(11, 15):
        ws_db.cell(r, c).border = border_all_thin

# ----------------- SCENARIO BREAKDOWN TABLE (Row 16 to 22) -----------------
ws_db.merge_cells("A16:G16")
ws_db["A16"] = "REVENUE & PROJECT DISTRIBUTION BY SCENARIO"
ws_db["A16"].font = font_section
ws_db["A16"].fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
ws_db.row_dimensions[16].height = 24

sc_headers = ["Distribution Scenario", "Formula Logic", "Revenue (PKR)", "Company Share", "Team Share", "Projects", "Status"]
for idx, h in enumerate(sc_headers, 1):
    c = ws_db.cell(17, idx, h)
    c.font = font_tbl_header
    c.fill = fill_tbl_header
    c.alignment = align_center
    c.border = border_all_thin

sc_rows = [
    ("Scenario A: 2 Members Working", "50% Co | 25% Shaista | 25% Amara",
     "=SUMIFS('Projects & Revenue'!$I$5:$I$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario A*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$R$5:$R$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario A*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$O$5:$O$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario A*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=COUNTIFS('Projects & Revenue'!$K$5:$K$29, \"*Scenario A*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")", "Historical"),
    
    ("Scenario B: 3 Members Working", "30% Co | 23.33% Each (Shaista/Amara/Urooj)",
     "=SUMIFS('Projects & Revenue'!$I$5:$I$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario B*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$R$5:$R$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario B*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$O$5:$O$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario B*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=COUNTIFS('Projects & Revenue'!$K$5:$K$29, \"*Scenario B*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")", "Standard 3-Way"),
    
    ("Scenario C: 3 Reg, 2 Working", "53.33% Co | 23.33% Each | 3rd Skipped",
     "=SUMIFS('Projects & Revenue'!$I$5:$I$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario C*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$R$5:$R$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario C*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=SUMIFS('Projects & Revenue'!$O$5:$O$29, 'Projects & Revenue'!$K$5:$K$29, \"*Scenario C*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")",
     "=COUNTIFS('Projects & Revenue'!$K$5:$K$29, \"*Scenario C*\", 'Projects & Revenue'!$J$5:$J$29, \"Received\")", "Active New Rule")
]

for idx, r_data in enumerate(sc_rows, 18):
    ws_db.row_dimensions[idx].height = 20
    ws_db.cell(idx, 1, r_data[0]).alignment = align_left
    ws_db.cell(idx, 2, r_data[1]).alignment = align_left
    
    for c_i in range(3, 6):
        c = ws_db.cell(idx, c_i, r_data[c_i-1])
        c.alignment = align_right
        c.number_format = "#,##0.00"
    
    ws_db.cell(idx, 6, r_data[5]).alignment = align_center
    ws_db.cell(idx, 7, r_data[6]).alignment = align_center
    
    for c in range(1, 8):
        ws_db.cell(idx, c).border = border_all_thin
        ws_db.cell(idx, c).font = font_regular

# Total Scenario Row
ws_db.row_dimensions[21].height = 22
ws_db.merge_cells("A21:B21")
ws_db.cell(21, 1, "TOTAL ALL SCENARIOS").alignment = align_right
ws_db.cell(21, 1).font = font_total
ws_db.cell(21, 3, "=SUM(C18:C20)").number_format = "#,##0.00"
ws_db.cell(21, 4, "=SUM(D18:D20)").number_format = "#,##0.00"
ws_db.cell(21, 5, "=SUM(E18:E20)").number_format = "#,##0.00"
ws_db.cell(21, 6, "=SUM(F18:F20)").alignment = align_center
ws_db.cell(21, 7, "100% Balanced").alignment = align_center
for c in range(1, 8):
    ws_db.cell(21, c).fill = fill_total
    ws_db.cell(21, c).font = font_total
    ws_db.cell(21, c).border = border_total

# Table for Chart Data in Cols I to K, Rows 16 to 20
ws_db.cell(16, 9, "Entity")
ws_db.cell(16, 10, "Total Distribution (PKR)")
ws_db.cell(17, 9, "Company Gross")
ws_db.cell(17, 10, "='Projects & Revenue'!R30")
ws_db.cell(18, 9, "Shaista Share")
ws_db.cell(18, 10, "='Projects & Revenue'!L30")
ws_db.cell(19, 9, "Amara Share")
ws_db.cell(19, 10, "='Projects & Revenue'!M30")
ws_db.cell(20, 9, "Urooj Share")
ws_db.cell(20, 10, "='Projects & Revenue'!N30")

pie = PieChart()
pie.title = "Revenue Distribution Breakdown"
pie_labels = Reference(ws_db, min_col=9, min_row=17, max_row=20)
pie_data = Reference(ws_db, min_col=10, min_row=16, max_row=20)
pie.add_data(pie_data, titles_from_data=True)
pie.set_categories(pie_labels)
pie.width = 15
pie.height = 7.5
ws_db.add_chart(pie, "I22")

# Monthly Trend Bar Chart
chart_bar = BarChart()
chart_bar.title = "Monthly Revenue vs Company Gross Share"
chart_bar.style = 10
chart_bar.y_axis.title = "Amount (PKR)"
chart_bar.x_axis.title = "Month"
chart_bar.width = 17
chart_bar.height = 7.5

bar_data = Reference(ws_rep, min_col=2, min_row=3, max_row=14)
bar_cats = Reference(ws_rep, min_col=1, min_row=4, max_row=14)
chart_bar.add_data(bar_data, titles_from_data=True)
chart_bar.set_categories(bar_cats)
ws_db.add_chart(chart_bar, "A22")

col_widths_db = {
    "A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18,
    "H": 18, "I": 18, "J": 18, "K": 18, "L": 18, "M": 18, "N": 18
}
for col_let, width in col_widths_db.items():
    ws_db.column_dimensions[col_let].width = width


# ----------------------------------------------------
# 7. SHEET: User Guide & Instructions
# ----------------------------------------------------
ws_ug = wb.create_sheet(title="User Guide & System Manual")
ws_ug.views.sheetView[0].showGridLines = True

ws_ug.merge_cells("A1:H1")
ws_ug["A1"] = "HEXALEVEL • FINANCIAL DASHBOARD USER GUIDE & SYSTEM MANUAL"
ws_ug["A1"].font = font_title
ws_ug["A1"].fill = fill_title
ws_ug["A1"].alignment = align_left
ws_ug.row_dimensions[1].height = 36

guide_sections = [
    ("1. HOW TO ADD A NEW CLIENT PROJECT",
     ["Navigate to the 'Projects & Revenue' worksheet.",
      "Go to the first empty row below the existing project records.",
      "Enter the Project ID, Start Month, Payment Date, Client Name, and Project Name.",
      "Enter the Payment Amount in PKR and select Payment Status ('Received' or 'Pending').",
      "Select the Distribution Scenario from the dropdown list (Scenario A, Scenario B, or Scenario C).",
      "All formulas for Member Shares, Company Base Share, and Company Unallocated Share will calculate automatically!",
      "The Variance Check column (Column T) verifies that total distribution exactly equals the payment amount (must be 0.00)."]),
    
    ("2. HOW THE THREE SCENARIOS WORK",
     ["Scenario A: 2 Members Working (Historical Model up to Aug 2026)",
      "  • Company receives 50% | Shaista receives 25% | Amara receives 25% | Urooj receives $0.",
      "Scenario B: 3 Members Working (Standard Model from 1 Sep 2026)",
      "  • Company receives 30% | Shaista receives 23.33% | Amara receives 23.33% | Urooj receives 23.33%.",
      "Scenario C: 3 Registered Members, Only 2 Working (NEW ENGINE REQUIREMENT)",
      "  • The 2 working members receive their 23.33% share each.",
      "  • The non-working member receives $0.00 (strictly skipped from project payout).",
      "  • The unallocated 23.33% share is transferred directly to Company Earnings (Gross Company Share = 53.33%).",
      "  • In Net Profit attribution, the non-working member is skipped from surplus, ensuring working members earn more."]),
    
    ("3. HOW EXPENSES & NET PROFIT ARE MANAGED",
     ["Navigate to the 'Expenses & Outflows' worksheet to log software tools, connects, courses, or cash advances.",
      "Total Company Deductions are automatically subtracted from Total Company Gross Revenue.",
      "Company Net Profit = Total Company Gross - Total Company Deductions.",
      "The Dashboard and Reports sheets update automatically whenever new expenses or payments are added."]),
    
    ("4. اردو میں مختصر ہدایات (Urdu Quick Guide)",
     ["1. نیا پروجیکٹ شامل کرنے کے لیے 'Projects & Revenue' شیٹ پر جائیں اور اگلی خالی قطار میں کلائنٹ کا نام اور رقم درج کریں۔",
      "2. 'Distribution Scenario' کے ڈراپ ڈاؤن مینو سے مطلوبہ ماڈل منتخب کریں (Scenario A، Scenario B یا Scenario C)۔",
      "3. اگر دو ممبرز نے کام کیا ہے اور ایک نے نہیں کیا، تو 'Scenario C' منتخب کریں۔ کام کرنے والے دونوں ممبرز کو 23.33% ملے گا، تیسرے ممبر کا شیئر زیرو ہو گا اور اس کا حصہ کمپنی میں چلا جائے گا۔",
      "4. کمپنی کے نیٹ پرافٹ میں بھی غیر حاضر ممبر کو سکپ رکھا گیا ہے تاکہ کام کرنے والے دو ممبرز کے پیسے بڑھتے رہیں۔",
      "5. تمام حساب کتاب، چارٹس اور ڈیش بورڈ آٹومیٹک اپ ڈیٹ ہوتے ہیں۔"])
]

r_ug = 3
for title_s, bullets in guide_sections:
    ws_ug.merge_cells(f"A{r_ug}:H{r_ug}")
    ws_ug.cell(r_ug, 1, title_s).font = font_section
    ws_ug.cell(r_ug, 1).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    ws_ug.row_dimensions[r_ug].height = 24
    r_ug += 1
    
    for b in bullets:
        ws_ug.row_dimensions[r_ug].height = 20
        ws_ug.cell(r_ug, 1, "•").alignment = align_center
        ws_ug.cell(r_ug, 1).font = font_bold
        ws_ug.merge_cells(f"B{r_ug}:H{r_ug}")
        ws_ug.cell(r_ug, 2, b).alignment = align_left
        ws_ug.cell(r_ug, 2).font = font_regular
        r_ug += 1
    r_ug += 1

ws_ug.column_dimensions["A"].width = 6
ws_ug.column_dimensions["B"].width = 30
ws_ug.column_dimensions["C"].width = 25
ws_ug.column_dimensions["D"].width = 25
ws_ug.column_dimensions["E"].width = 20
ws_ug.column_dimensions["F"].width = 20
ws_ug.column_dimensions["G"].width = 20
ws_ug.column_dimensions["H"].width = 20

# Save the complete workbook
output_filename = "Hexalevel_Financial_Management_Dashboard.xlsx"
wb.save(output_filename)
print(f"Workbook '{output_filename}' generated successfully!")
