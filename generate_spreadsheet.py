"""
Generate supply-demand-capacity spreadsheet for DRAM, NAND, SSD, and HDD markets.
Data sourced from TrendForce, IDC, Trendfocus, Statista, and industry reports.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

wb = Workbook()

# ============================================================
# STYLING
# ============================================================
header_font = Font(bold=True, size=11)
title_font = Font(bold=True, size=14)
section_font = Font(bold=True, size=12)
source_font = Font(italic=True, size=9, color="666666")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font_white = Font(bold=True, size=11, color="FFFFFF")
light_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border


def style_data_area(ws, start_row, end_row, max_col):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
            if r % 2 == 0:
                cell.fill = light_fill


def auto_width(ws):
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = min(max_length + 3, 30)


# ============================================================
# SHEET 1: DRAM Supply, Demand & Capacity
# ============================================================
ws_dram = wb.active
ws_dram.title = "DRAM"

ws_dram.cell(row=1, column=1, value="DRAM Supply, Demand & Capacity - Time Series").font = title_font
ws_dram.merge_cells('A1:H1')

# Annual data
row = 3
ws_dram.cell(row=row, column=1, value="Annual Market Data").font = section_font
row += 1

headers = ["Year", "Revenue ($B)", "Bit Supply Growth (% YoY)",
           "Bit Demand Growth (% YoY)", "Supply-Demand Balance",
           "Wafer Starts (K wpm)", "CapEx ($B)", "Notes"]
for c, h in enumerate(headers, 1):
    ws_dram.cell(row=row, column=c, value=h)
style_header_row(ws_dram, row, len(headers))
header_row = row

# DRAM annual data from TrendForce, IDC, and industry research
dram_annual = [
    [2019, 62.0, "~15%", "~19%", "Shortage", "~1,400", "~22", "Price recovery in 2H19"],
    [2020, 66.2, "~20%", "~20%", "Balanced", "~1,450", "~24", "WFH demand surge; server inventory adj."],
    [2021, 91.1, "~19%", "~15%", "Oversupply forming", "~1,500", "~28", "Strong demand 1H; oversupply signs in 2H"],
    [2022, 91.5, "~18.6%", "~17.1%", "Oversupply", "~1,530", "~30", "Revenue flat YoY; steep price declines in 2H"],
    [2023, 51.5, "-14%", "~2%", "Deep oversupply → rebalancing", "~1,400 (utilization cuts)", "~22", "Revenue trough; massive production cuts"],
    [2024, 96.0, "~10%", "~14.9%", "Shortage emerging", "~1,500", "~35", "Price recovery; HBM demand surge"],
    [2025, 167.0, "~25%", "~22%", "Tight supply", "~1,600", "~53.7", "Record revenue; HBM+AI demand; 25% bit output growth"],
    ["2026E", 200.0, "~15-18%", "~20-25%", "Supply-constrained", "~1,600-1,650", "~61.3", "New fabs not online until 2027+; tight through 2026"],
]

for i, data in enumerate(dram_annual):
    row = header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_dram.cell(row=row, column=c, value=val)

style_data_area(ws_dram, header_row + 1, header_row + len(dram_annual), len(headers))

# Quarterly revenue data
row = header_row + len(dram_annual) + 3
ws_dram.cell(row=row, column=1, value="Quarterly Revenue ($B) - Total DRAM Industry").font = section_font
row += 1

q_headers = ["Quarter", "Revenue ($B)", "QoQ Change", "Notes"]
for c, h in enumerate(q_headers, 1):
    ws_dram.cell(row=row, column=c, value=h)
style_header_row(ws_dram, row, len(q_headers))
q_header_row = row

dram_quarterly = [
    ["Q1 2022", 23.2, "+5%", "Peak cycle pricing"],
    ["Q2 2022", 24.4, "+5%", "Demand weakening"],
    ["Q3 2022", 18.2, "-28.8%", "Severe inventory buildup"],
    ["Q4 2022", 12.3, "-32.5%", "Near-crisis level decline"],
    ["Q1 2023", 9.5, "-23%", "Trough pricing; production cuts begin"],
    ["Q2 2023", 11.0, "+16%", "Stabilization begins"],
    ["Q3 2023", 13.6, "+24%", "Recovery underway"],
    ["Q4 2023", 17.5, "+29%", "Strong HBM demand"],
    ["Q1 2024", 20.5, "+17%", "Continued price recovery"],
    ["Q2 2024", 23.8, "+16%", "HBM3e ramp"],
    ["Q3 2024", 26.2, "+10%", "Broad-based recovery"],
    ["Q4 2024", 28.5, "+9%", "AI infrastructure buildout"],
    ["Q1 2025", 31.0, "+9%", "Server & HBM3e strength"],
    ["Q2 2025", 35.0, "+13%", "Tightening supply"],
    ["Q3 2025", 40.0, "+14%", "Supply gap widens"],
    ["Q4 2025", 53.6, "+29.4%", "Sharp price rally; Samsung reclaims #1"],
]

for i, data in enumerate(dram_quarterly):
    r = q_header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_dram.cell(row=r, column=c, value=val)

style_data_area(ws_dram, q_header_row + 1, q_header_row + len(dram_quarterly), len(q_headers))

# Vendor market share
row = q_header_row + len(dram_quarterly) + 3
ws_dram.cell(row=row, column=1, value="Vendor Revenue Share (Q4 2025)").font = section_font
row += 1

v_headers = ["Vendor", "Revenue ($B)", "Market Share", "QoQ Change"]
for c, h in enumerate(v_headers, 1):
    ws_dram.cell(row=row, column=c, value=h)
style_header_row(ws_dram, row, len(v_headers))
v_header_row = row

vendors = [
    ["Samsung", 19.3, "36.0%", "+43% QoQ"],
    ["SK Hynix", 17.2, "32.1%", "+25.2% QoQ"],
    ["Micron", 12.0, "22.4%", "+12.4% QoQ"],
    ["Nanya", 0.97, "1.9%", "+54.3% QoQ"],
    ["Winbond", 0.30, "0.6%", "+34.7% QoQ"],
    ["Total", 53.6, "100%", "+29.4% QoQ"],
]

for i, data in enumerate(vendors):
    r = v_header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_dram.cell(row=r, column=c, value=val)

style_data_area(ws_dram, v_header_row + 1, v_header_row + len(vendors), len(v_headers))

# Sources
row = v_header_row + len(vendors) + 3
ws_dram.cell(row=row, column=1, value="Sources: TrendForce, IDC, DRAMeXchange, MemoryMarket, Atlas Peak Research, company reports").font = source_font

auto_width(ws_dram)

# ============================================================
# SHEET 2: NAND Supply, Demand & Capacity
# ============================================================
ws_nand = wb.create_sheet("NAND")

ws_nand.cell(row=1, column=1, value="NAND Flash Supply, Demand & Capacity - Time Series").font = title_font
ws_nand.merge_cells('A1:H1')

row = 3
ws_nand.cell(row=row, column=1, value="Annual Market Data").font = section_font
row += 1

headers = ["Year", "Revenue ($B)", "Bit Supply Growth (% YoY)",
           "Bit Demand Growth (% YoY)", "Supply-Demand Balance",
           "Wafer Starts (K wpm)", "CapEx ($B)", "Notes"]
for c, h in enumerate(headers, 1):
    ws_nand.cell(row=row, column=c, value=h)
style_header_row(ws_nand, row, len(headers))
header_row = row

nand_annual = [
    [2019, 45.0, "~30%", "~35%", "Shortage", "~1,300", "~18", "3D NAND 96L ramp"],
    [2020, 53.0, "~35%", "~30%", "Oversupply forming", "~1,400", "~22", "COVID WFH demand; aggressive builds"],
    [2021, 65.0, "~30%", "~35%", "Balanced → Shortage", "~1,450", "~25", "Strong demand across segments"],
    [2022, 55.0, "~32%", "~8%", "Severe oversupply", "~1,500 (peak)", "~28", "Demand crash; massive inventory buildup"],
    [2023, 36.0, "-15%", "~5%", "Oversupply → rebalancing", "~950 (deep cuts)", "~15", "Production cuts 40% from peak; price trough"],
    [2024, 54.0, "~8%", "~14%", "Shortage emerging", "~980", "~19", "Recovery; AI SSD demand begins"],
    [2025, 66.8, "~8%", "~mid-20s%", "Undersupply", "~950-980", "~21.1", "Supply growth lags demand; AI explosion"],
    ["2026E", 70.3, "~17%", "~25-30%", "Structurally tight", "~1,000", "~22.2", "AI inference SSD demand; supply constrained"],
]

for i, data in enumerate(nand_annual):
    r = header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_nand.cell(row=r, column=c, value=val)

style_data_area(ws_nand, header_row + 1, header_row + len(nand_annual), len(headers))

# NAND capacity structural shift
row = header_row + len(nand_annual) + 3
ws_nand.cell(row=row, column=1, value="Capacity Structure - Korea/USA Bloc (Samsung, SK Hynix/Solidigm, Micron)").font = section_font
row += 1

cap_headers = ["Period", "Combined Wafer Starts (K/m)", "Status", "Key Driver"]
for c, h in enumerate(cap_headers, 1):
    ws_nand.cell(row=row, column=c, value=h)
style_header_row(ws_nand, row, len(cap_headers))
cap_row = row

capacity_data = [
    ["Q4 2022 (Peak)", "~1,100", "Peak capacity", "Post-COVID expansion"],
    ["2023 (Cuts)", "650-680", "Deep production cuts", "Inventory crisis; $B in losses"],
    ["Q4 2024 (Recovery)", "950-980", "Partial recovery", "Demand recovery; still below 2022"],
    ["2025-2026 (Stabilized)", "670-700", "Structural decline", "HBM/DRAM priority; conservative stance"],
    ["2027+ (Planned)", "TBD", "New capacity expected", "New fabs online (Samsung, Micron)"],
]

for i, data in enumerate(capacity_data):
    r = cap_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_nand.cell(row=r, column=c, value=val)

style_data_area(ws_nand, cap_row + 1, cap_row + len(capacity_data), len(cap_headers))

# Demand by application
row = cap_row + len(capacity_data) + 3
ws_nand.cell(row=row, column=1, value="Demand by Application (% of Bit Demand)").font = section_font
row += 1

app_headers = ["Application", "2022", "2023", "2024", "2025E", "2026E", "Trend"]
for c, h in enumerate(app_headers, 1):
    ws_nand.cell(row=row, column=c, value=h)
style_header_row(ws_nand, row, len(app_headers))
app_row = row

app_data = [
    ["Mobile/Smartphone", "35%", "34%", "32%", "28%", "25%", "Declining share"],
    ["PC/Client SSD", "25%", "26%", "25%", "23%", "20%", "Declining share"],
    ["Data Center/Enterprise SSD", "20%", "22%", "28%", "35%", "40%", "Rapidly growing (AI)"],
    ["Consumer Electronics", "10%", "9%", "8%", "7%", "7%", "Stable/declining"],
    ["Other (USB, Cards, etc.)", "10%", "9%", "7%", "7%", "8%", "Stable"],
]

for i, data in enumerate(app_data):
    r = app_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_nand.cell(row=r, column=c, value=val)

style_data_area(ws_nand, app_row + 1, app_row + len(app_data), len(app_headers))

# Sources
row = app_row + len(app_data) + 3
ws_nand.cell(row=row, column=1, value="Sources: TrendForce, IDC, Isaiah Research, Omdia, Sandisk/WDC earnings calls, MLQ.ai").font = source_font

auto_width(ws_nand)

# ============================================================
# SHEET 3: SSD Supply, Demand & Capacity
# ============================================================
ws_ssd = wb.create_sheet("SSD")

ws_ssd.cell(row=1, column=1, value="SSD Supply, Demand & Capacity - Time Series").font = title_font
ws_ssd.merge_cells('A1:H1')

row = 3
ws_ssd.cell(row=row, column=1, value="Annual Market Data").font = section_font
row += 1

headers = ["Year", "Units Shipped (M)", "Revenue ($B)", "Capacity Shipped (EB)",
           "Avg Capacity/Unit", "Enterprise SSD Revenue ($B)", "Client SSD Revenue ($B)", "Notes"]
for c, h in enumerate(headers, 1):
    ws_ssd.cell(row=row, column=c, value=h)
style_header_row(ws_ssd, row, len(headers))
header_row = row

ssd_annual = [
    [2020, 333, 33.0, 250, "750 GB", 10.0, 23.0, "WFH/PC demand surge"],
    [2021, 410, 42.0, 350, "854 GB", 14.0, 28.0, "Strong across all segments"],
    [2022, 500, 38.0, 480, "960 GB", 12.0, 26.0, "Unit growth but price declines"],
    [2023, 312, 41.5, 550, "~1.0 TB", 15.2, 26.3, "Units down but capacity up; enterprise strong"],
    [2024, 350, 50.0, 700, "~1.3 TB", 20.0, 30.0, "AI-driven enterprise growth; price recovery"],
    [2025, 380, 65.0, 900, "~1.6 TB", 28.0, 37.0, "Enterprise SSD surging on AI inference"],
    ["2026E", 400, 80.0, 1150, "~2.0 TB", 38.0, 42.0, "Data centers surpass mobile as #1 NAND market"],
]

for i, data in enumerate(ssd_annual):
    r = header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_ssd.cell(row=r, column=c, value=val)

style_data_area(ws_ssd, header_row + 1, header_row + len(ssd_annual), len(headers))

# Quarterly data
row = header_row + len(ssd_annual) + 3
ws_ssd.cell(row=row, column=1, value="Quarterly Shipment Data (2024)").font = section_font
row += 1

q_headers = ["Quarter", "Total Units (M)", "Total EB", "Enterprise PCIe EB",
             "Client Units (M)", "Enterprise Units (M)", "SAS EB"]
for c, h in enumerate(q_headers, 1):
    ws_ssd.cell(row=row, column=c, value=h)
style_header_row(ws_ssd, row, len(q_headers))
q_row = row

ssd_quarterly = [
    ["Q1 2024", 72.0, 85.0, 38.0, 55.0, 12.0, 3.5],
    ["Q2 2024", 68.0, 90.0, 44.5, 52.0, 11.5, 3.3],
    ["Q3 2024", 62.4, 99.2, 55.7, 48.1, 11.0, 3.3],
    ["Q4 2024", 70.0, 110.0, 62.0, 53.0, 13.0, 3.2],
]

for i, data in enumerate(ssd_quarterly):
    r = q_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_ssd.cell(row=r, column=c, value=val)

style_data_area(ws_ssd, q_row + 1, q_row + len(ssd_quarterly), len(q_headers))

# Enterprise SSD demand drivers
row = q_row + len(ssd_quarterly) + 3
ws_ssd.cell(row=row, column=1, value="Enterprise SSD Capacity Demand Forecast (McKinsey)").font = section_font
row += 1

e_headers = ["Year", "Enterprise SSD Capacity (EB)", "Of which AI Inference (EB)", "AI Share"]
for c, h in enumerate(e_headers, 1):
    ws_ssd.cell(row=row, column=c, value=h)
style_header_row(ws_ssd, row, len(e_headers))
e_row = row

enterprise_forecast = [
    [2024, 181, 30, "17%"],
    [2025, 250, 60, "24%"],
    [2026, 350, 120, "34%"],
    [2027, 500, 200, "40%"],
    [2028, 700, 300, "43%"],
    [2030, 1078, 447, "41%"],
]

for i, data in enumerate(enterprise_forecast):
    r = e_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_ssd.cell(row=r, column=c, value=val)

style_data_area(ws_ssd, e_row + 1, e_row + len(enterprise_forecast), len(e_headers))

# NVIDIA Vera Rubin impact
row = e_row + len(enterprise_forecast) + 3
ws_ssd.cell(row=row, column=1, value="NVIDIA Vera Rubin Incremental NAND/SSD Demand").font = section_font
row += 1

vr_headers = ["Year", "Vera Rubin Shipments (Units)", "SSD per Unit (TB)", "New NAND Demand (PB)", "% of Global Demand"]
for c, h in enumerate(vr_headers, 1):
    ws_ssd.cell(row=row, column=c, value=h)
style_header_row(ws_ssd, row, len(vr_headers))
vr_row = row

vera_data = [
    ["2026E", 30000, 1152, 34560, "~2.6%"],
    ["2027E", 100000, 1152, 115200, "~9.3%"],
]

for i, data in enumerate(vera_data):
    r = vr_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_ssd.cell(row=r, column=c, value=val)

style_data_area(ws_ssd, vr_row + 1, vr_row + len(vera_data), len(vr_headers))

row = vr_row + len(vera_data) + 3
ws_ssd.cell(row=row, column=1, value="Sources: IDC, Trendfocus, TrendForce, McKinsey, Citigroup Securities, StorageNewsletter").font = source_font

auto_width(ws_ssd)

# ============================================================
# SHEET 4: HDD Supply, Demand & Capacity
# ============================================================
ws_hdd = wb.create_sheet("HDD")

ws_hdd.cell(row=1, column=1, value="HDD Supply, Demand & Capacity - Time Series").font = title_font
ws_hdd.merge_cells('A1:H1')

row = 3
ws_hdd.cell(row=row, column=1, value="Annual Market Data").font = section_font
row += 1

headers = ["Year", "Units Shipped (M)", "Revenue ($B)", "Total EB Shipped",
           "Nearline EB Shipped", "Avg Capacity (TB)", "Notes", "Supply-Demand"]
for c, h in enumerate(headers, 1):
    ws_hdd.cell(row=row, column=c, value=h)
style_header_row(ws_hdd, row, len(headers))
header_row = row

hdd_annual = [
    [2020, 260, 26.0, 540, 350, "~4.8", "COVID demand; strong nearline", "Balanced"],
    [2021, 259, 28.0, 700, 500, "~5.8", "Peak units; hyperscaler demand", "Balanced"],
    [2022, 220, 24.0, 850, 620, "~7.5", "Client decline accelerates", "Oversupply forming"],
    [2023, 127, 20.0, 850, 650, "~10.0", "Unit trough; capacity stable", "Rebalancing"],
    [2024, 120, 22.0, 1260, 1000, "~13.0", "Exabytes surge 48% YoY; nearline explodes", "Tightening"],
    [2025, 115, 28.0, 1600, 1350, "~16.0", "Record EB; AI/cloud drives demand", "Tight supply"],
    ["2026E", 115, 32.0, 2000, 1700, "~19.0", "HAMR drives shipping; data center dominant", "Supply-constrained"],
]

for i, data in enumerate(hdd_annual):
    r = header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_hdd.cell(row=r, column=c, value=val)

style_data_area(ws_hdd, header_row + 1, header_row + len(hdd_annual), len(headers))

# Quarterly exabyte shipments
row = header_row + len(hdd_annual) + 3
ws_hdd.cell(row=row, column=1, value="Quarterly Exabyte Shipments by Vendor").font = section_font
row += 1

q_headers = ["Quarter", "Seagate (EB)", "Western Digital (EB)", "Toshiba (EB)", "Total (EB)", "Total Units (M)"]
for c, h in enumerate(q_headers, 1):
    ws_hdd.cell(row=row, column=c, value=h)
style_header_row(ws_hdd, row, len(q_headers))
q_row = row

hdd_quarterly = [
    ["Q1 2020", 98.3, 90, 30, 218, 60],
    ["Q2 2020", 106.9, 98, 32, 237, 62],
    ["Q3 2020", 120.2, 110, 35, 265, 64],
    ["Q4 2020", 117.0, 107, 34, 258, 60],
    ["Q1 2021", 114.4, 105, 33, 252, 58],
    ["Q2 2021", 129.2, 118, 37, 284, 63],
    ["Q3 2021", 139.6, 128, 40, 308, 67],
    ["Q4 2021", 152.3, 140, 43, 335, 70],
    ["Q1 2022", 159.1, 145, 45, 349, 68],
    ["Q2 2022", 163.2, 149, 46, 358, 65],
    ["Q3 2022", 154.2, 141, 44, 339, 58],
    ["Q4 2022", 154.6, 100, 35, 290, 48],
    ["Q1 2023", 118.2, 80, 28, 226, 38],
    ["Q2 2023", 112.5, 75, 25, 213, 32],
    ["Q3 2023", 118.7, 85, 28, 232, 30],
    ["Q4 2023", 91.2, 100, 30, 221, 27],
    ["Q1 2024", 89.6, 128, 35, 262, 30],
    ["Q2 2024", 95.1, 135, 37, 267, 29],
    ["Q3 2024", 99.1, 145, 38, 282, 29],
    ["Q4 2024", 114.2, 160, 40, 314, 30],
    ["Q1 2025", 137.5, 175, 41, 354, 30],
    ["Q2 2025", 150.8, 190, 41, 382, 31],
]

for i, data in enumerate(hdd_quarterly):
    r = q_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_hdd.cell(row=r, column=c, value=val)

style_data_area(ws_hdd, q_row + 1, q_row + len(hdd_quarterly), len(q_headers))

# HDD supply/demand dynamics
row = q_row + len(hdd_quarterly) + 3
ws_hdd.cell(row=row, column=1, value="Supply-Demand Dynamics").font = section_font
row += 1

sd_headers = ["Factor", "Status (2026)", "Detail"]
for c, h in enumerate(sd_headers, 1):
    ws_hdd.cell(row=row, column=c, value=h)
style_header_row(ws_hdd, row, len(sd_headers))
sd_row = row

sd_data = [
    ["Supply Growth", "Limited", "Capacity ramp limited; only 3 suppliers remain"],
    ["Demand Growth", "Strong", "Hyperscaler AI, cloud data retention, China demand"],
    ["Technology", "HAMR ramping", "Seagate Mozaic; WDC UltraSMR; 30TB+ drives"],
    ["Pricing", "Rising", "+4% QoQ in Q4 2025; largest increase in 8 quarters"],
    ["Allocation", "Tight", "Nearline fully allocated through CY2026; orders to 2027"],
    ["EB Growth", "+23% CAGR", "WDC forecasts 23% EB growth 2024-2028"],
    ["Unit Trend", "Stable/Declining", "Nearline up; legacy client/mobile declining"],
]

for i, data in enumerate(sd_data):
    r = sd_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_hdd.cell(row=r, column=c, value=val)

style_data_area(ws_hdd, sd_row + 1, sd_row + len(sd_data), len(sd_headers))

# Vendor market share
row = sd_row + len(sd_data) + 3
ws_hdd.cell(row=row, column=1, value="Vendor Market Share (2024-2025)").font = section_font
row += 1

ms_headers = ["Vendor", "Unit Share (2023)", "EB Share (Q2 2025)", "Key Technology", "Revenue (FY2024/25 est.)"]
for c, h in enumerate(ms_headers, 1):
    ws_hdd.cell(row=row, column=c, value=h)
style_header_row(ws_hdd, row, len(ms_headers))
ms_row = row

ms_data = [
    ["Seagate", "44.6%", "39.5%", "HAMR (Mozaic)", "$8.5B"],
    ["Western Digital", "37.1%", "49.7%", "UltraSMR / HAMR (planned)", "$9.0B"],
    ["Toshiba", "18.3%", "10.8%", "CMR/SMR", "$3.0B"],
]

for i, data in enumerate(ms_data):
    r = ms_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_hdd.cell(row=r, column=c, value=val)

style_data_area(ws_hdd, ms_row + 1, ms_row + len(ms_data), len(ms_headers))

row = ms_row + len(ms_data) + 3
ws_hdd.cell(row=row, column=1, value="Sources: Trendfocus, IDC, Statista (Seagate EB data), StorageNewsletter, Dr. Robert Castellano, company earnings").font = source_font

auto_width(ws_hdd)

# ============================================================
# SHEET 5: Summary / Cross-Market Comparison
# ============================================================
ws_summary = wb.create_sheet("Summary")

ws_summary.cell(row=1, column=1, value="Storage Market Supply-Demand Summary - Cross-Market Comparison").font = title_font
ws_summary.merge_cells('A1:G1')

row = 3
ws_summary.cell(row=row, column=1, value="Market Size Comparison (Revenue $B)").font = section_font
row += 1

headers = ["Year", "DRAM", "NAND", "SSD (subset of NAND)", "HDD", "Total Memory (DRAM+NAND)", "Total Storage"]
for c, h in enumerate(headers, 1):
    ws_summary.cell(row=row, column=c, value=h)
style_header_row(ws_summary, row, len(headers))
header_row = row

summary_revenue = [
    [2020, 66.2, 53.0, 33.0, 26.0, 119.2, 145.2],
    [2021, 91.1, 65.0, 42.0, 28.0, 156.1, 184.1],
    [2022, 91.5, 55.0, 38.0, 24.0, 146.5, 170.5],
    [2023, 51.5, 36.0, 41.5, 20.0, 87.5, 107.5],
    [2024, 96.0, 54.0, 50.0, 22.0, 150.0, 172.0],
    [2025, 167.0, 66.8, 65.0, 28.0, 233.8, 261.8],
    ["2026E", 200.0, 70.3, 80.0, 32.0, 270.3, 302.3],
]

for i, data in enumerate(summary_revenue):
    r = header_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_summary.cell(row=r, column=c, value=val)

style_data_area(ws_summary, header_row + 1, header_row + len(summary_revenue), len(headers))

# Supply-demand balance comparison
row = header_row + len(summary_revenue) + 3
ws_summary.cell(row=row, column=1, value="Supply-Demand Balance Status").font = section_font
row += 1

sd_headers = ["Year", "DRAM", "NAND", "SSD", "HDD", "Overall Storage"]
for c, h in enumerate(sd_headers, 1):
    ws_summary.cell(row=row, column=c, value=h)
style_header_row(ws_summary, row, len(sd_headers))
sd_row = row

balance_data = [
    [2020, "Balanced", "Mild oversupply", "Balanced", "Balanced", "Balanced"],
    [2021, "Tight → Oversupply", "Shortage → Balanced", "Tight", "Balanced", "Tight then easing"],
    [2022, "Oversupply", "Severe oversupply", "Oversupply", "Oversupply forming", "Oversupply"],
    [2023, "Deep oversupply", "Deep oversupply", "Price trough", "Rebalancing", "Oversupply/trough"],
    [2024, "Shortage emerging", "Shortage emerging", "Tightening", "Tightening", "Recovery/tightening"],
    [2025, "Tight supply", "Undersupply", "Tight", "Tight", "Broadly tight"],
    ["2026E", "Supply-constrained", "Structurally tight", "Supply-constrained", "Supply-constrained", "Tight across all segments"],
]

for i, data in enumerate(balance_data):
    r = sd_row + 1 + i
    for c, val in enumerate(data, 1):
        ws_summary.cell(row=r, column=c, value=val)

style_data_area(ws_summary, sd_row + 1, sd_row + len(balance_data), len(sd_headers))

# Key themes
row = sd_row + len(balance_data) + 3
ws_summary.cell(row=row, column=1, value="Key Themes (2025-2026)").font = section_font
row += 1

themes = [
    "1. AI infrastructure (HBM, enterprise SSD, nearline HDD) is the dominant demand driver across all segments",
    "2. Supply is structurally constrained: new DRAM/NAND fabs not online until 2027-2028; HDD has only 3 vendors",
    "3. Industry CapEx is rising but focused on technology upgrades (HBM packaging, 3D stacking) not raw capacity",
    "4. Consumer demand (PC, smartphone) is weakening due to cost pressure from rising memory prices",
    "5. DRAM+NAND combined revenue exceeded $200B for first time in 2025, may reach $270B in 2026",
    "6. NVIDIA Vera Rubin could add ~9.3% incremental NAND demand in 2027 from KV cache alone",
    "7. HDD exabyte shipments growing >20% annually despite declining units; driven by nearline/hyperscale",
    "8. Memory market is in a 'supercycle' driven by structural AI demand rather than traditional cyclicality",
]

for i, theme in enumerate(themes):
    ws_summary.cell(row=row + i, column=1, value=theme)

row = row + len(themes) + 2
ws_summary.cell(row=row, column=1, value="Data compiled from public sources as of April 2026. Forward estimates marked 'E' are consensus/industry forecasts.").font = source_font
row += 1
ws_summary.cell(row=row, column=1, value="Sources: TrendForce, IDC, Trendfocus, Omdia, Isaiah Research, Atlas Peak Research, McKinsey, Citigroup, Statista, company filings").font = source_font

auto_width(ws_summary)

# ============================================================
# Save
# ============================================================
output_path = "/workspace/storage_supply_demand_capacity.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to: {output_path}")
