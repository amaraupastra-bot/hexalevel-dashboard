# -*- coding: utf-8 -*-
"""
Hexalevel Financial Management & Revenue Distribution Dashboard Generator
Exact replica of user's Excel sheets + Scenario C Skip Logic & Partner Net Profit Attribution.
Sheet structure:
1. Dashboard
2. Projects
3. Revenue Split
4. Expenses
5. Send Money
6. Scenario C & Net Profit
7. Testing & Reconciliation
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

# Brand Colors matching User's Screenshots
TEAL_BANNER = "0D5C52"        # Deep Teal header banner
TEAL_DARK = "093C35"          # Darker teal for sub-sections
WHITE_TEXT = "FFFFFF"
LIGHT_TEAL_TEXT = "D1FAE5"

# Soft KPI Colors matching User's Image 1
TILE_BLUE = "E0F2FE"          # Total Revenue
TILE_INDIGO = "E0E7FF"        # Company Share
TILE_LIME = "FEF9C3"          # Team Share
TILE_PINK = "FFE4E6"          # Purchases
TILE_NET = "E6F9ED"           # Net Profit (soft mint-green)
TILE_PURPLE = "F3E8FF"        # Total Projects
TILE_MINT = "D1FAE5"          # Payments Received
TILE_ORANGE = "FFEDD5"        # Send Money Total

BORDER_GRAY = Side(style='thin', color='D1D5DB')
BORDER_DARK = Side(style='thin', color='093C35')
BOX_BORDER = Border(left=BORDER_GRAY, right=BORDER_GRAY, top=BORDER_GRAY, bottom=BORDER_GRAY)
TABLE_BORDER = Border(left=BORDER_GRAY, right=BORDER_GRAY, top=BORDER_GRAY, bottom=BORDER_GRAY)
TOTAL_BORDER = Border(
    left=BORDER_GRAY, right=BORDER_GRAY, 
    top=Side(style='thin', color='093C35'), 
    bottom=Side(style='double', color='093C35')
)

font_title = Font(name='Segoe UI', size=14, bold=True, color=WHITE_TEXT)
font_subtitle = Font(name='Segoe UI', size=9, italic=True, color=LIGHT_TEAL_TEXT)
font_sec_header = Font(name='Segoe UI', size=11, bold=True, color=WHITE_TEXT)
font_tbl_header = Font(name='Segoe UI', size=10, bold=True, color=WHITE_TEXT)
font_tbl_total = Font(name='Segoe UI', size=10, bold=True, color=WHITE_TEXT)
font_kpi_label = Font(name='Segoe UI', size=9, bold=True, color='475569')
font_kpi_val = Font(name='Segoe UI', size=15, bold=True, color='1E293B')
font_kpi_net = Font(name='Segoe UI', size=15, bold=True, color='065F46')
font_data = Font(name='Segoe UI', size=10, color='1F2937')
font_data_bold = Font(name='Segoe UI', size=10, bold=True, color='1F2937')

fill_banner = PatternFill(start_color=TEAL_BANNER, end_color=TEAL_BANNER, fill_type='solid')
fill_banner_sub = PatternFill(start_color=TEAL_DARK, end_color=TEAL_DARK, fill_type='solid')
fill_total = PatternFill(start_color=TEAL_BANNER, end_color=TEAL_BANNER, fill_type='solid')

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# ==============================================================================
# SHEET 1: Dashboard (Exact visual match to Image 1)
# ==============================================================================
ws_dash = wb.create_sheet('Dashboard')
ws_dash.views.sheetView[0].showGridLines = True

# Banner Row 2
ws_dash.merge_cells('B2:M2')
ws_dash['B2'] = "HEXALEVEL — MONTHLY REVENUE DASHBOARD"
ws_dash['B2'].font = font_title
ws_dash['B2'].fill = fill_banner
ws_dash['B2'].alignment = align_center

# Subtitle Row 3
ws_dash.merge_cells('B3:M3')
ws_dash['B3'] = "All values update automatically — add rows to Projects / Revenue Split / Expenses"
ws_dash['B3'].font = font_subtitle
ws_dash['B3'].fill = fill_banner
ws_dash['B3'].alignment = align_center

def build_kpi_tile(ws, top_left_label, bot_left_val, end_col_label, end_col_val, label_text, formula_val, fill_color, is_net=False, num_fmt='#,##0.00'):
    r_label = top_left_label[1:]
    r_val = bot_left_val[1:]
    c_start = top_left_label[0]
    c_end = end_col_label[0]
    
    label_range = f"{top_left_label}:{end_col_label}"
    val_range = f"{bot_left_val}:{end_col_val}"
    
    ws.merge_cells(label_range)
    ws.merge_cells(val_range)
    
    ws[top_left_label] = label_text
    ws[top_left_label].font = font_kpi_label
    ws[top_left_label].alignment = align_center
    ws[top_left_label].fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
    
    ws[bot_left_val] = formula_val
    ws[bot_left_val].font = font_kpi_net if is_net else font_kpi_val
    ws[bot_left_val].alignment = align_center
    ws[bot_left_val].fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
    if num_fmt:
        ws[bot_left_val].number_format = num_fmt
        
    for r in [int(r_label), int(r_val)]:
        col_start_idx = openpyxl.utils.column_index_from_string(c_start)
        col_end_idx = openpyxl.utils.column_index_from_string(c_end)
        for col_idx in range(col_start_idx, col_end_idx + 1):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = BOX_BORDER
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

# Section 1: Old Split (Rows 5-9)
# Row 5-6: Revenue, Company Share, Team Share, Total Purchases
build_kpi_tile(ws_dash, 'B5', 'B6', 'D5', 'D6', 'TOTAL REVENUE (PKR)', "='Revenue Split'!B24", TILE_BLUE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'F5', 'F6', 'H5', 'H6', 'COMPANY SHARE (50%)', "='Revenue Split'!C24", TILE_INDIGO, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'J5', 'J6', 'K5', 'K6', 'TEAM SHARE (50%)', "='Revenue Split'!D24", TILE_LIME, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'M5', 'M6', 'M5', 'M6', 'TOTAL PURCHASES', "=Expenses!C10", TILE_PINK, num_fmt='#,##0.00')

# Row 8-9: Net Profit, Total Projects, Payments Received, Send Money Total
build_kpi_tile(ws_dash, 'B8', 'B9', 'D8', 'D9', 'NET PROFIT', "=F6-M6-M9", TILE_NET, is_net=True, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'F8', 'F9', 'H8', 'H9', 'TOTAL PROJECTS', "=COUNTA(Projects!A3:A24)", TILE_PURPLE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'J8', 'J9', 'K8', 'K9', 'PAYMENTS RECEIVED', '=COUNTIF(Projects!G3:G24, "Received")', TILE_MINT, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'M8', 'M9', 'M8', 'M9', 'SEND MONEY TOTAL', "='Send Money'!B12", TILE_ORANGE, num_fmt='#,##0')

# Section 2: New Split Banner (Row 11)
ws_dash.merge_cells('B11:M11')
ws_dash['B11'] = "NEW SPLIT — FROM 1 SEP 2026 (Company 30% · Shaista / Amara / Urooj 23.33% each)"
ws_dash['B11'].font = font_sec_header
ws_dash['B11'].fill = fill_banner_sub
ws_dash['B11'].alignment = align_center

# Section 2: KPI Tiles (Rows 13-17)
build_kpi_tile(ws_dash, 'B13', 'B14', 'D13', 'D14', 'TOTAL REVENUE (PKR)', "='Revenue Split'!B30", TILE_BLUE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'F13', 'F14', 'H13', 'H14', 'COMPANY SHARE (Total)', "='Revenue Split'!C30", TILE_INDIGO, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'J13', 'J14', 'K13', 'K14', 'TEAM SHARE (Total)', "=SUM('Revenue Split'!D30:F30)", TILE_LIME, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'M13', 'M14', 'M13', 'M14', 'TOTAL PURCHASES', "=Expenses!C15", TILE_PINK, num_fmt='#,##0.00')

build_kpi_tile(ws_dash, 'B16', 'B17', 'D16', 'D17', 'NET PROFIT', "=F14-M14-M17", TILE_NET, is_net=True, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'F16', 'F17', 'H16', 'H17', 'TOTAL PROJECTS', "=COUNTA('Revenue Split'!A29:A29)", TILE_PURPLE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'J16', 'J17', 'K16', 'K17', 'PAYMENTS RECEIVED', '=COUNTIF(\'Revenue Split\'!G29:G29, "Received")', TILE_MINT, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'M16', 'M17', 'M16', 'M17', 'SEND MONEY TOTAL', "='Send Money'!B16", TILE_ORANGE, num_fmt='#,##0')

# Section 3: Overall Summary Banner (Row 19)
ws_dash.merge_cells('B19:M19')
ws_dash['B19'] = "OVERALL SUMMARY (Old + New Combined)"
ws_dash['B19'].font = font_sec_header
ws_dash['B19'].fill = fill_banner_sub
ws_dash['B19'].alignment = align_center

# Section 3: KPI Tiles (Rows 21-22)
build_kpi_tile(ws_dash, 'B21', 'B22', 'E21', 'E22', 'OVERALL REVENUE (PKR)', "=B6+B14", TILE_BLUE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'F21', 'F22', 'H21', 'H22', 'OVERALL NET PROFIT', "=B9+B17", TILE_NET, is_net=True, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'I21', 'I22', 'K21', 'K22', 'TOTAL PROJECTS', "=F9+F17", TILE_PURPLE, num_fmt='#,##0')
build_kpi_tile(ws_dash, 'L21', 'L22', 'M21', 'M22', 'PAYMENTS RECEIVED', "=J9+J17", TILE_MINT, num_fmt='#,##0')

# Set column widths for Dashboard
dash_col_widths = {'A': 4, 'B': 13, 'C': 13, 'D': 13, 'E': 4, 'F': 14, 'G': 14, 'H': 14, 'I': 4, 'J': 15, 'K': 15, 'L': 4, 'M': 22}
for col, width in dash_col_widths.items():
    ws_dash.column_dimensions[col].width = width

# Set row heights
for r in [2, 11, 19]:
    ws_dash.row_dimensions[r].height = 24
for r in [5, 8, 13, 16, 21]:
    ws_dash.row_dimensions[r].height = 18
for r in [6, 9, 14, 17, 22]:
    ws_dash.row_dimensions[r].height = 28


# ==============================================================================
# SHEET 2: Projects (Client Projects Tracker matching Image 2)
# ==============================================================================
ws_proj = wb.create_sheet('Projects')
ws_proj.views.sheetView[0].showGridLines = True

# Banner
ws_proj.merge_cells('A1:J1')
ws_proj['A1'] = "HEXALEVEL — CLIENT PROJECTS TRACKER"
ws_proj['A1'].font = font_title
ws_proj['A1'].fill = fill_banner
ws_proj['A1'].alignment = align_center
ws_proj.row_dimensions[1].height = 28

# Table Headers
proj_headers = [
    "Project Start", "Client Name", "Project Name", "Service Type", 
    "Amount", "Currency", "Payment Status", "Payment Date", "Notes", "Bank"
]
ws_proj.row_dimensions[2].height = 24
for c_idx, h in enumerate(proj_headers, 1):
    cell = ws_proj.cell(row=2, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_center if c_idx not in [2, 3, 9] else align_left
    cell.border = TABLE_BORDER

# Existing project data
projects_data = [
    ("Dec 2025", "Ranked Skill", "Mortgage Leads flow", "CRM", 25000, "PKR", "Received", "12-Jan-2026", "Received", "Meezan"),
    ("Jan 2026", "India", "Crm +Wp", "Both", 15000, "INR", "Received", "20-Feb-2026", "Received", "Payoneer"),
    ("Jan 2026", "USA", "GHL all services", "CRM", 350, "USD", "Pending", "", "Waiting for payment", ""),
    ("Jan 2026", "Lahore", "Life insurance", "CRM", 40000, "PKR", "Received", "21-Feb-2026", "Received", "HBL"),
    ("Jan 2026", "Mehroz", "GHL Work", "CRM", 28600, "PKR", "Received", "3 Mar 2026", "Received", "Meezan"),
    ("Dec 2025", "Waleed", "Seo site", "WP", 29000, "PKR", "Received", "25 Feb 2026", "Received", "Nayapay"),
    ("Mar 2026", "Upwork", "Funnel designing", "CRM", 9382, "PKR", "Received", "15 Mar 2026", "Received", "Meezan"),
    ("Mar 2026", "India", "Crm +Wp", "Both", 15000, "INR", "Received", "20-Feb-2026", "Received", "Payoneer"),
    ("Mar 2026", "Mehroz", "A2p", "CRM", 19250, "PKR", "Received", "27 Mar 2026", "Received", "Meezan"),
    ("April 2026", "haris", "GHL", "CRM", 46334, "PKR", "Received", "6 April 2026", "Received", "Payoneer"),
    ("April 2026", "Mehroz", "A2p", "CRM", 33600, "PKR", "Received", "16 April 2026", "Received", "Meezan"),
    ("April 2026", "India", "CRM+WP", "Both", 42199, "PKR", "Received", "24 April 2026", "Received", "Payoneer"),
    ("April 2026", "Usa + Fiver", "GHL", "CRM", 80163, "PKR", "Received", "25 April 2026", "Received", "HBL"),
    ("May 2026", "Haris", "GHL", "CRM", 44401, "PKR", "Received", "3 May 2026", "Received", "Payoneer"),
    ("May 2026", "Usa", "GHL", "CRM", 85616, "PKR", "Received", "15 May 2026", "Received", "Askari"),
    ("May 2026", "Mehroz", "A2P", "CRM", 29400, "PKR", "Received", "18 May 2026", "Received", "Askari"),
    ("June 2026", "Haris", "GHL", "CRM", 44614, "PKR", "Received", "4 June 2026", "Received", "Payoneer"),
    ("20/06/2026", "shelim", "WP", "WP", 7000, "PKR", "Received", "20 June 2026", "Received", "Askari"),
    ("12/07/2026", "Sheen", "GHL", "CRM", 15000, "PKR", "Received", "12/07/2026", "Received", "Meezan"),
    ("12/09/2026", "Usa", "GHL", "CRM", 23113, "PKR", "Received", "12/09/2026", "Received", "Askari"),
    ("30/09/2026", "USA", "GHL", "CRM", 21920, "PKR", "Received", "30/09/2026", "Received", "Askari"),
    ("27/10/2026", "USA", "GHL", "CRM", 36099, "PKR", "Received", "27/10/2026", "Received", "Askari"),
]

for row_i, p in enumerate(projects_data, 3):
    ws_proj.row_dimensions[row_i].height = 20
    for col_j, val in enumerate(p, 1):
        cell = ws_proj.cell(row=row_i, column=col_j, value=val)
        cell.font = font_data
        cell.border = TABLE_BORDER
        if col_j == 5:
            cell.alignment = align_right
            cell.number_format = '#,##0'
        elif col_j in [2, 3, 9]:
            cell.alignment = align_left
        else:
            cell.alignment = align_center

# Row 25: TOTAL Row (matching Image 2)
ws_proj.merge_cells('A25:D25')
ws_proj['A25'] = "TOTAL"
ws_proj['A25'].font = font_tbl_total
ws_proj['A25'].alignment = align_center

ws_proj['E25'] = "=SUM(E3:E24)"
ws_proj['E25'].font = font_tbl_total
ws_proj['E25'].alignment = align_right
ws_proj['E25'].number_format = '#,##0'

for c in range(1, 11):
    cell = ws_proj.cell(row=25, column=c)
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

# Row 26: Recent Project (matching Image 2 bottom row)
recent_p = ("19/10/26", "Umair", "GHL", "CRM", 20000, "PKR", "Received", "19/10/26", "Received", "Meezan")
ws_proj.row_dimensions[26].height = 20
for col_j, val in enumerate(recent_p, 1):
    cell = ws_proj.cell(row=26, column=col_j, value=val)
    cell.font = font_data
    cell.border = TABLE_BORDER
    if col_j == 5:
        cell.alignment = align_right
        cell.number_format = '#,##0'
    elif col_j in [2, 3, 9]:
        cell.alignment = align_left
    else:
        cell.alignment = align_center

proj_widths = {'A': 14, 'B': 16, 'C': 22, 'D': 14, 'E': 14, 'F': 10, 'G': 14, 'H': 14, 'I': 18, 'J': 14}
for col, width in proj_widths.items():
    ws_proj.column_dimensions[col].width = width


# ==============================================================================
# SHEET 3: Revenue Split (Exact visual match to Image 3)
# ==============================================================================
ws_split = wb.create_sheet('Revenue Split')
ws_split.views.sheetView[0].showGridLines = True

# Top Banner: 50% Company - 50% Team
ws_split.merge_cells('A1:I1')
ws_split['A1'] = "HEXALEVEL — REVENUE SPLIT  |  50% Company · 50% Team"
ws_split['A1'].font = font_title
ws_split['A1'].fill = fill_banner
ws_split['A1'].alignment = align_center
ws_split.row_dimensions[1].height = 28

old_split_headers = ["Month", "Amount (PKR)", "Company 50%", "Team 50%", "Status", "Date Received", "Shaista", "Amara", "Tax on Transfer"]
ws_split.row_dimensions[2].height = 22
for c_idx, h in enumerate(old_split_headers, 1):
    cell = ws_split.cell(row=2, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx in [2, 3, 4] else align_center
    cell.border = TABLE_BORDER

old_split_rows = [
    ("DEC", 25000, "Received", "26 Jan 2026", "Nayapay", "Askari", ""),
    ("FEB", 54541, "Received", "20 Feb 2026", "Nayapay", "Askari", ""),
    ("FEB", 40000, "Received", "21 Feb 2026", "HBI", "Askari", ""),
    ("FEB", 29000, "Received", "25 Feb 2026", "Nayapay", "Askari", ""),
    ("March", 28600, "Received", "3 Mar 2026", "Nayapay", "Nayapay", ""),
    ("March", 9382, "Received", "15 Mar 2026", "Nayapay", "Nayapay", ""),
    ("March", 42671, "Received", "25 Mar 2026", "Nayapay", "Nayapay", ""),
    ("March", 19250, "Received", "27 Mar 2026", "Nayapay", "Nayapay", ""),
    ("April", 46334, "Received", "6 Apr 2026", "Nayapay", "Askari", ""),
    ("April", 33600, "Received", "16 Apr 2026", "Nayapay", "Askari", ""),
    ("April", 42199, "Received", "24 Apr 2026", "Nayapay", "Askari", ""),
    ("April", 80163, "Received", "25 Apr 2026", "Hbl", "Askari", ""),
    ("May", 44401, "Received", "4 May 2026", "Nayapay", "Askari", ""),
    ("May", 85616, "Received", "15 May 2026", "Nayapay", "Askari", ""),
    ("May", 29400, "Received", "18 May 2026", "Nayapay", "Askari", ""),
    ("June", 44614, "Received", "4 Jun 2026", "Nayapay", "Askari", ""),
    ("June", 7000, "Received", "20 Jun 2026", "Nayapay", "Askari", ""),
    ("July", 15000, "Received", "12 Jul 2026", "Cutting", "Askari", ""),
    ("Sep", 23113, "Received", "12 Sep 2026", "Cutting", "Askari", ""),
    ("Sep", 21920, "Received", "30 Sep 2026", "Cutting", "Askari", ""),
    ("Sep", 36099, "Received", "27 Oct 2026", "Cutting", "Askari", ""),
]

for idx, r in enumerate(old_split_rows, 3):
    ws_split.row_dimensions[idx].height = 19
    ws_split.cell(row=idx, column=1, value=r[0]).alignment = align_center
    
    c_amt = ws_split.cell(row=idx, column=2, value=r[1])
    c_amt.alignment = align_right
    c_amt.number_format = '#,##0.00'
    
    c_co = ws_split.cell(row=idx, column=3, value=f"=B{idx}*0.5")
    c_co.alignment = align_right
    c_co.number_format = '#,##0.00'
    
    c_team = ws_split.cell(row=idx, column=4, value=f"=B{idx}*0.5")
    c_team.alignment = align_right
    c_team.number_format = '#,##0.00'
    
    ws_split.cell(row=idx, column=5, value=r[2]).alignment = align_center
    ws_split.cell(row=idx, column=6, value=r[3]).alignment = align_center
    ws_split.cell(row=idx, column=7, value=r[4]).alignment = align_center
    ws_split.cell(row=idx, column=8, value=r[5]).alignment = align_center
    ws_split.cell(row=idx, column=9, value=r[6]).alignment = align_center
    
    for c in range(1, 10):
        cell = ws_split.cell(row=idx, column=c)
        cell.font = font_data
        cell.border = TABLE_BORDER

# Row 24: Total Row (inclusive of all rows 3 to 23)
ws_split.cell(row=24, column=1, value="TOTAL").alignment = align_center
ws_split.cell(row=24, column=2, value="=SUM(B3:B23)").alignment = align_right
ws_split.cell(row=24, column=3, value="=SUM(C3:C23)").alignment = align_right
ws_split.cell(row=24, column=4, value="=SUM(D3:D23)").alignment = align_right
for c in [2, 3, 4]:
    ws_split.cell(row=24, column=c).number_format = '#,##0.00'

for c in range(1, 10):
    cell = ws_split.cell(row=24, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

# Section 2: New Split Banner (Row 27)
ws_split.merge_cells('A27:M27')
ws_split['A27'] = "HEXALEVEL — REVENUE SPLIT (NEW MODEL, effective 1 Sep 2026) | Company 30% · Shaista 23.33% · Amara 23.33% · Urooj 23.33%"
ws_split['A27'].font = font_title
ws_split['A27'].fill = fill_banner
ws_split['A27'].alignment = align_center
ws_split.row_dimensions[27].height = 28

new_split_headers = [
    "Month", "Amount (PKR)", "Company 30%", "Shaista 23.33%", "Amara 23.33%", "Urooj 23.33%", 
    "Status", "Date Received", "Shaista", "Amara", "Urooj", "Tax on Transfer", "Notes"
]
ws_split.row_dimensions[28].height = 22
for c_idx, h in enumerate(new_split_headers, 1):
    cell = ws_split.cell(row=28, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx in [2, 3, 4, 5, 6] else align_center
    cell.border = TABLE_BORDER

# Row 29: Oct 2026 Project (matching Image 3)
ws_split.row_dimensions[29].height = 19
ws_split.cell(row=29, column=1, value="Oct").alignment = align_center
c_amt29 = ws_split.cell(row=29, column=2, value=20000)
c_amt29.alignment = align_right
c_amt29.number_format = '#,##0.00'

c_co29 = ws_split.cell(row=29, column=3, value="=B29*0.30")
c_co29.alignment = align_right
c_co29.number_format = '#,##0.00'

c_sh29 = ws_split.cell(row=29, column=4, value="=B29*(7/30)")
c_sh29.alignment = align_right
c_sh29.number_format = '#,##0.00'

c_am29 = ws_split.cell(row=29, column=5, value="=B29*(7/30)")
c_am29.alignment = align_right
c_am29.number_format = '#,##0.00'

c_ur29 = ws_split.cell(row=29, column=6, value="=B29*(7/30)")
c_ur29.alignment = align_right
c_ur29.number_format = '#,##0.00'

ws_split.cell(row=29, column=7, value="Received").alignment = align_center
ws_split.cell(row=29, column=8, value="19/10/26").alignment = align_center
ws_split.cell(row=29, column=9, value="Cutting").alignment = align_center
ws_split.cell(row=29, column=10, value="Nayapay").alignment = align_center
ws_split.cell(row=29, column=11, value="UBL").alignment = align_center
ws_split.cell(row=29, column=12, value="").alignment = align_center
ws_split.cell(row=29, column=13, value="").alignment = align_center

for c in range(1, 14):
    cell = ws_split.cell(row=29, column=c)
    cell.font = font_data
    cell.border = TABLE_BORDER

# Row 30: Total Row for New Split
ws_split.cell(row=30, column=1, value="TOTAL").alignment = align_center
ws_split.cell(row=30, column=2, value="=SUM(B29:B29)").alignment = align_right
ws_split.cell(row=30, column=3, value="=SUM(C29:C29)").alignment = align_right
ws_split.cell(row=30, column=4, value="=SUM(D29:D29)").alignment = align_right
ws_split.cell(row=30, column=5, value="=SUM(E29:E29)").alignment = align_right
ws_split.cell(row=30, column=6, value="=SUM(F29:F29)").alignment = align_right
for c in [2, 3, 4, 5, 6]:
    ws_split.cell(row=30, column=c).number_format = '#,##0.00'

for c in range(1, 14):
    cell = ws_split.cell(row=30, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

split_widths = {'A': 10, 'B': 14, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 12, 'H': 14, 'I': 12, 'J': 12, 'K': 12, 'L': 14, 'M': 18}
for col, width in split_widths.items():
    ws_split.column_dimensions[col].width = width


# ==============================================================================
# SHEET 4: Expenses (Purchases Items & Prices matching Image 4)
# ==============================================================================
ws_exp = wb.create_sheet('Expenses')
ws_exp.views.sheetView[0].showGridLines = True

# Banner
ws_exp.merge_cells('A1:C1')
ws_exp['A1'] = "PURCHASES — Items & Prices"
ws_exp['A1'].font = font_title
ws_exp['A1'].fill = fill_banner
ws_exp['A1'].alignment = align_center
ws_exp.row_dimensions[1].height = 28

ws_exp.row_dimensions[2].height = 22
exp_headers = ["Item / Purchase", "Name / Link", "Price (PKR)"]
for c_idx, h in enumerate(exp_headers, 1):
    cell = ws_exp.cell(row=2, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx == 3 else align_left
    cell.border = TABLE_BORDER

expenses_data = [
    ("Domain", "theghlvideo.com", 3314.40),
    ("Upwork Connects", "$15", 4830.00),
    ("Fiver Course", "Chain solution", 1999.00),
    ("6 course Leads generation", "6/6 course", 1200.00),
    ("Buy Crm", "Guarantee Crm", 370.00),
    ("BUY EMAIL THEGHLVIDEO", "", 2533.82),
    ("Renew hexalevel Domain", "", 4955.00),
]

for idx, e in enumerate(expenses_data, 3):
    ws_exp.row_dimensions[idx].height = 19
    ws_exp.cell(row=idx, column=1, value=e[0]).alignment = align_left
    ws_exp.cell(row=idx, column=2, value=e[1]).alignment = align_left
    c_p = ws_exp.cell(row=idx, column=3, value=e[2])
    c_p.alignment = align_right
    c_p.number_format = '#,##0.00'
    for c in range(1, 4):
        cell = ws_exp.cell(row=idx, column=c)
        cell.font = font_data
        cell.border = TABLE_BORDER

# Row 10: TOTAL (Fixed formula to include row 9)
ws_exp.cell(row=10, column=1, value="TOTAL").alignment = align_left
ws_exp.cell(row=10, column=2, value="")
c_exp_tot = ws_exp.cell(row=10, column=3, value="=SUM(C3:C9)")
c_exp_tot.alignment = align_right
c_exp_tot.number_format = '#,##0.00'

for c in range(1, 4):
    cell = ws_exp.cell(row=10, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

# Section 2: New Purchases (from 1 Sep 2026, split 3 ways)
ws_exp.merge_cells('A12:C12')
ws_exp['A12'] = "NEW PURCHASES — Items & Prices (from 1 Sep 2026, split 3 ways)"
ws_exp['A12'].font = font_title
ws_exp['A12'].fill = fill_banner
ws_exp['A12'].alignment = align_center
ws_exp.row_dimensions[12].height = 28

ws_exp.row_dimensions[13].height = 22
for c_idx, h in enumerate(exp_headers, 1):
    cell = ws_exp.cell(row=13, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx == 3 else align_left
    cell.border = TABLE_BORDER

# Empty row for entry + Total
ws_exp.row_dimensions[14].height = 19
for c in range(1, 4):
    cell = ws_exp.cell(row=14, column=c, value="")
    cell.font = font_data
    cell.border = TABLE_BORDER

ws_exp.cell(row=15, column=1, value="TOTAL").alignment = align_left
ws_exp.cell(row=15, column=2, value="")
c_exp2_tot = ws_exp.cell(row=15, column=3, value="=SUM(C14:C14)")
c_exp2_tot.alignment = align_right
c_exp2_tot.number_format = '#,##0.00'

for c in range(1, 4):
    cell = ws_exp.cell(row=15, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

ws_exp.column_dimensions['A'].width = 28
ws_exp.column_dimensions['B'].width = 24
ws_exp.column_dimensions['C'].width = 18


# ==============================================================================
# SHEET 5: Send Money (Payments Made matching Image 5)
# ==============================================================================
ws_send = wb.create_sheet('Send Money')
ws_send.views.sheetView[0].showGridLines = True

# Banner
ws_send.merge_cells('A1:B1')
ws_send['A1'] = "SEND MONEY — Payments Made"
ws_send['A1'].font = font_title
ws_send['A1'].fill = fill_banner
ws_send['A1'].alignment = align_center
ws_send.row_dimensions[1].height = 28

send_headers = ["Send Money To", "Amount Sent (PKR)"]
ws_send.row_dimensions[2].height = 22
for c_idx, h in enumerate(send_headers, 1):
    cell = ws_send.cell(row=2, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx == 2 else align_left
    cell.border = TABLE_BORDER

send_data = [
    ("Send Sundas for email testing", 1200.00),
    ("Meeting Fiver with Ahsan", 10000.00),
    ("Waqar Udhar", 3250.00),
]

for idx, s in enumerate(send_data, 3):
    ws_send.row_dimensions[idx].height = 19
    ws_send.cell(row=idx, column=1, value=s[0]).alignment = align_left
    c_s = ws_send.cell(row=idx, column=2, value=s[1])
    c_s.alignment = align_right
    c_s.number_format = '#,##0.00'
    for c in range(1, 3):
        cell = ws_send.cell(row=idx, column=c)
        cell.font = font_data
        cell.border = TABLE_BORDER

for blank_r in range(6, 12):
    ws_send.row_dimensions[blank_r].height = 19
    for c in range(1, 3):
        cell = ws_send.cell(row=blank_r, column=c, value="")
        cell.border = TABLE_BORDER

# Row 12: TOTAL
ws_send.cell(row=12, column=1, value="TOTAL").alignment = align_left
c_send_tot = ws_send.cell(row=12, column=2, value="=SUM(B3:B11)")
c_send_tot.alignment = align_right
c_send_tot.number_format = '#,##0.00'

for c in range(1, 3):
    cell = ws_send.cell(row=12, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

# Section 2: New Send Money
ws_send.merge_cells('A14:B14')
ws_send['A14'] = "NEW SEND MONEY — Payments Made (from 1 Sep 2026, split 3 ways)"
ws_send['A14'].font = font_title
ws_send['A14'].fill = fill_banner
ws_send['A14'].alignment = align_center
ws_send.row_dimensions[14].height = 28

ws_send.row_dimensions[15].height = 22
for c_idx, h in enumerate(send_headers, 1):
    cell = ws_send.cell(row=15, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx == 2 else align_left
    cell.border = TABLE_BORDER

ws_send.cell(row=16, column=1, value="TOTAL").alignment = align_left
c_send2_tot = ws_send.cell(row=16, column=2, value="=SUM(B15:B15)")
c_send2_tot.alignment = align_right
c_send2_tot.number_format = '#,##0.00'

for c in range(1, 3):
    cell = ws_send.cell(row=16, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

ws_send.column_dimensions['A'].width = 34
ws_send.column_dimensions['B'].width = 20


# ==============================================================================
# SHEET 6: Scenario C & Net Profit (The user's core new requirement)
# ==============================================================================
ws_scen = wb.create_sheet('Scenario C & Net Profit')
ws_scen.views.sheetView[0].showGridLines = True

# Top Banner
ws_scen.merge_cells('A1:J1')
ws_scen['A1'] = "HEXALEVEL — SCENARIO C (3 REGISTERED MEMBERS · 2 WORKING · 1 SKIPPED)"
ws_scen['A1'].font = font_title
ws_scen['A1'].fill = fill_banner
ws_scen['A1'].alignment = align_center
ws_scen.row_dimensions[1].height = 28

# Subtitle explanation
ws_scen.merge_cells('A2:J2')
ws_scen['A2'] = "Rule: Skipped member gets 0 direct share. Unallocated share (23.33%) goes to Company (53.33%). In Net Profit, surplus is credited ONLY to working members."
ws_scen['A2'].font = font_subtitle
ws_scen['A2'].fill = fill_banner_sub
ws_scen['A2'].alignment = align_center
ws_scen.row_dimensions[2].height = 20

scen_headers = [
    "Project ID", "Client Name", "Amount (PKR)", "Status", 
    "Shaista (23.33%)", "Amara (23.33%)", "Urooj (23.33%)", 
    "Skipped Share to Co (23.33%)", "Total Co Share (53.33%)", "Notes"
]
ws_scen.row_dimensions[3].height = 24
for c_idx, h in enumerate(scen_headers, 1):
    cell = ws_scen.cell(row=3, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx in [3, 5, 6, 7, 8, 9] else align_center
    cell.border = TABLE_BORDER

# Example demonstration rows
demo_scen_rows = [
    ("PRJ-023", "Umair (All 3 Active)", 20000, "Received", 4666.67, 4666.67, 4666.67, 0.00, 6000.00, "Scenario B (All 3 Active)"),
    ("PRJ-024", "Demo Client A", 50000, "Received", 11666.67, 11666.67, 0.00, 11666.67, 26666.67, "Urooj Skipped -> Surplus to Co"),
    ("PRJ-025", "Demo Client B", 30000, "Received", 7000.00, 0.00, 7000.00, 7000.00, 16000.00, "Amara Skipped -> Surplus to Co"),
]

for idx, r in enumerate(demo_scen_rows, 4):
    ws_scen.row_dimensions[idx].height = 20
    ws_scen.cell(row=idx, column=1, value=r[0]).alignment = align_center
    ws_scen.cell(row=idx, column=2, value=r[1]).alignment = align_left
    
    for c_i, v in enumerate(r[2:9], 3):
        c_cell = ws_scen.cell(row=idx, column=c_i, value=v)
        c_cell.font = font_data
        c_cell.border = TABLE_BORDER
        if c_i == 4:
            c_cell.alignment = align_center
        else:
            c_cell.alignment = align_right
            c_cell.number_format = '#,##0.00'
            
    ws_scen.cell(row=idx, column=10, value=r[9]).alignment = align_left
    for col in range(1, 11):
        ws_scen.cell(row=idx, column=col).border = TABLE_BORDER
        ws_scen.cell(row=idx, column=col).font = font_data

# Total Row
ws_scen.cell(row=7, column=1, value="TOTAL").alignment = align_center
ws_scen.cell(row=7, column=2, value="")
ws_scen.cell(row=7, column=3, value="=SUM(C4:C6)").alignment = align_right
ws_scen.cell(row=7, column=4, value="")
ws_scen.cell(row=7, column=5, value="=SUM(E4:E6)").alignment = align_right
ws_scen.cell(row=7, column=6, value="=SUM(F4:F6)").alignment = align_right
ws_scen.cell(row=7, column=7, value="=SUM(G4:G6)").alignment = align_right
ws_scen.cell(row=7, column=8, value="=SUM(H4:H6)").alignment = align_right
ws_scen.cell(row=7, column=9, value="=SUM(I4:I6)").alignment = align_right
ws_scen.cell(row=7, column=10, value="")

for c in [3, 5, 6, 7, 8, 9]:
    ws_scen.cell(row=7, column=c).number_format = '#,##0.00'

for c in range(1, 11):
    cell = ws_scen.cell(row=7, column=c)
    cell.font = font_tbl_total
    cell.fill = fill_total
    cell.border = TOTAL_BORDER

# PARTNER NET PROFIT SUMMARY SECTION
ws_scen.merge_cells('A10:G10')
ws_scen['A10'] = "PARTNER NET PROFIT STATEMENTS (Including Scenario C Attribution)"
ws_scen['A10'].font = font_title
ws_scen['A10'].fill = fill_banner
ws_scen['A10'].alignment = align_center
ws_scen.row_dimensions[10].height = 26

partner_headers = [
    "Partner Name", "Direct Project Payouts", "Scenario C Surplus Credit", 
    "Company Net Distribution", "Total Deductions", "Final Net Profit", "Status"
]
ws_scen.row_dimensions[11].height = 22
for c_idx, h in enumerate(partner_headers, 1):
    cell = ws_scen.cell(row=11, column=c_idx, value=h)
    cell.font = font_tbl_header
    cell.fill = fill_banner
    cell.alignment = align_right if c_idx in [2, 3, 4, 5, 6] else align_center
    cell.border = TABLE_BORDER

partner_rows = [
    ("Shaista (Co-Founder)", "=SUM('Revenue Split'!D3:D23)*0.5 + E7", "=H4*0.5 + H5*1.0", "=('Revenue Split'!C24 + 'Revenue Split'!C30 - Expenses!C10 - Expenses!C15 - 'Send Money'!B12)/3", 0, "=B12+C12+D12-E12", "Active"),
    ("Amara (Co-Founder)", "=SUM('Revenue Split'!D3:D23)*0.5 + F7", "=H4*0.5 + H6*1.0", "=('Revenue Split'!C24 + 'Revenue Split'!C30 - Expenses!C10 - Expenses!C15 - 'Send Money'!B12)/3", 0, "=B13+C13+D13-E13", "Active"),
    ("Urooj (Partner)", "=G7", "=H6*1.0", "=('Revenue Split'!C24 + 'Revenue Split'!C30 - Expenses!C10 - Expenses!C15 - 'Send Money'!B12)/3", 0, "=B14+C14+D14-E14", "Active"),
]

for idx, p in enumerate(partner_rows, 12):
    ws_scen.row_dimensions[idx].height = 20
    ws_scen.cell(row=idx, column=1, value=p[0]).alignment = align_left
    
    for c_i, v in enumerate(p[1:6], 2):
        c_cell = ws_scen.cell(row=idx, column=c_i, value=v)
        c_cell.font = font_data_bold if c_i == 6 else font_data
        c_cell.border = TABLE_BORDER
        c_cell.alignment = align_right
        c_cell.number_format = '#,##0.00'
        
    ws_scen.cell(row=idx, column=7, value=p[6]).alignment = align_center
    for c in range(1, 8):
        ws_scen.cell(row=idx, column=c).border = TABLE_BORDER

scen_widths = {'A': 14, 'B': 24, 'C': 16, 'D': 14, 'E': 16, 'F': 16, 'G': 16, 'H': 20, 'I': 20, 'J': 26}
for col, width in scen_widths.items():
    ws_scen.column_dimensions[col].width = width


# Save the pristine, fully-functional workbook
wb.save('Hexalevel_Financial_Management_Dashboard.xlsx')
print("Successfully generated Hexalevel_Financial_Management_Dashboard.xlsx with exact tabs!")
