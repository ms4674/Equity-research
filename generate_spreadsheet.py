"""
Generate an Excel spreadsheet with Apple Services Revenue breakdown and time series data.
Sources: Apple 10-K filings, quarterly earnings reports, ycharts.com, TechLila, BusinessOfApps.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Color palette & styles ──────────────────────────────────────────────────
DARK_GRAY = "333333"
APPLE_BLUE = "007AFF"
LIGHT_BLUE = "D6EAFF"
HEADER_FILL = PatternFill(start_color=DARK_GRAY, end_color=DARK_GRAY, fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color=DARK_GRAY)
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=11, color=DARK_GRAY)
DATA_FONT = Font(name="Calibri", size=11)
NUM_FMT_B = '#,##0.00"B"'
NUM_FMT_PCT = '0.0"%"'
THIN_BORDER = Border(
    bottom=Side(style="thin", color="CCCCCC"),
)
ALT_FILL = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_data_area(ws, start_row, end_row, max_col, num_fmt=None):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            if c == 1:
                cell.alignment = Alignment(horizontal="left")
            else:
                cell.alignment = Alignment(horizontal="center")
                if num_fmt and cell.value is not None:
                    cell.number_format = num_fmt
            if (r - start_row) % 2 == 1:
                cell.fill = ALT_FILL


def auto_width(ws, max_col, min_width=12):
    for col in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = max(min_width, 16)


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — Annual Services Revenue (FY 2012 – FY 2025)
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Annual Services Revenue"
ws1.sheet_properties.tabColor = "007AFF"

ws1.merge_cells("A1:F1")
ws1["A1"] = "Apple Services Revenue — Annual Time Series (Fiscal Year)"
ws1["A1"].font = TITLE_FONT

ws1.merge_cells("A2:F2")
ws1["A2"] = "Sources: Apple 10-K filings, ycharts, BusinessOfApps, TechLila. All figures in USD billions."
ws1["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")

headers = ["Fiscal Year", "Services Revenue ($B)", "Total Revenue ($B)",
           "Services % of Total", "YoY Growth (%)", "Services Gross Margin (%)"]
for i, h in enumerate(headers, 1):
    ws1.cell(row=4, column=i, value=h)
style_header_row(ws1, 4, len(headers))

annual_data = [
    # (FY, services_rev, total_rev, pct_of_total, yoy_growth, gross_margin)
    ("FY 2012", 12.89, 156.51, None, None, None),
    ("FY 2013", 16.05, 170.91, None, 24.5, None),
    ("FY 2014", 18.06, 182.80, None, 12.5, None),
    ("FY 2015", 19.91, 233.72, 8.5, 10.2, None),
    ("FY 2016", 24.35, 215.64, 11.3, 22.3, None),
    ("FY 2017", 32.70, 229.23, 14.3, 34.3, None),
    ("FY 2018", 39.75, 265.60, 15.0, 21.6, None),
    ("FY 2019", 46.29, 260.17, 17.8, 16.5, 63.7),
    ("FY 2020", 53.77, 274.52, 19.6, 16.2, 66.0),
    ("FY 2021", 68.42, 365.82, 18.7, 27.3, 69.7),
    ("FY 2022", 78.13, 394.33, 19.8, 14.2, 71.7),
    ("FY 2023", 85.20, 383.29, 22.2, 9.1, 70.8),
    ("FY 2024", 96.17, 391.04, 24.6, 12.9, 74.0),
    ("FY 2025", 109.16, 416.16, 26.2, 13.5, 75.3),
]

for idx, row in enumerate(annual_data):
    r = 5 + idx
    ws1.cell(row=r, column=1, value=row[0])
    ws1.cell(row=r, column=2, value=row[1])
    ws1.cell(row=r, column=3, value=row[2])
    if row[3] is not None:
        ws1.cell(row=r, column=4, value=row[3])
    if row[4] is not None:
        ws1.cell(row=r, column=5, value=row[4])
    if row[5] is not None:
        ws1.cell(row=r, column=6, value=row[5])

style_data_area(ws1, 5, 5 + len(annual_data) - 1, 6, num_fmt="#,##0.00")
auto_width(ws1, 6)

# Line chart — annual services revenue
chart1 = LineChart()
chart1.title = "Apple Annual Services Revenue ($B)"
chart1.style = 10
chart1.y_axis.title = "Revenue ($B)"
chart1.x_axis.title = "Fiscal Year"
chart1.width = 28
chart1.height = 14

cats = Reference(ws1, min_col=1, min_row=5, max_row=5 + len(annual_data) - 1)
vals = Reference(ws1, min_col=2, min_row=4, max_row=5 + len(annual_data) - 1)
chart1.add_data(vals, titles_from_data=True)
chart1.set_categories(cats)
s = chart1.series[0]
s.graphicalProperties.line.solidFill = APPLE_BLUE
s.graphicalProperties.line.width = 28000
chart1.legend = None

ws1.add_chart(chart1, "A21")

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — Quarterly Services Revenue
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Quarterly Services Revenue")
ws2.sheet_properties.tabColor = "34C759"

ws2.merge_cells("A1:E1")
ws2["A1"] = "Apple Services Revenue — Quarterly Time Series"
ws2["A1"].font = TITLE_FONT

ws2.merge_cells("A2:E2")
ws2["A2"] = "Apple fiscal quarters: Q1=Oct-Dec, Q2=Jan-Mar, Q3=Apr-Jun, Q4=Jul-Sep. Figures in USD billions."
ws2["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")

q_headers = ["Fiscal Quarter", "Calendar Period", "Services Revenue ($B)",
             "QoQ Change (%)", "YoY Change (%)"]
for i, h in enumerate(q_headers, 1):
    ws2.cell(row=4, column=i, value=h)
style_header_row(ws2, 4, len(q_headers))

quarterly_data = [
    ("FY20 Q1", "Oct–Dec 2019", 12.72, None, None),
    ("FY20 Q2", "Jan–Mar 2020", 13.35, 5.0, None),
    ("FY20 Q3", "Apr–Jun 2020", 13.16, -1.4, None),
    ("FY20 Q4", "Jul–Sep 2020", 14.55, 10.6, None),
    ("FY21 Q1", "Oct–Dec 2020", 15.76, 8.3, 23.9),
    ("FY21 Q2", "Jan–Mar 2021", 16.90, 7.2, 26.6),
    ("FY21 Q3", "Apr–Jun 2021", 17.49, 3.5, 32.9),
    ("FY21 Q4", "Jul–Sep 2021", 18.28, 4.5, 25.6),
    ("FY22 Q1", "Oct–Dec 2021", 19.52, 6.8, 23.9),
    ("FY22 Q2", "Jan–Mar 2022", 19.82, 1.5, 17.3),
    ("FY22 Q3", "Apr–Jun 2022", 19.60, -1.1, 12.1),
    ("FY22 Q4", "Jul–Sep 2022", 19.19, -2.1, 5.0),
    ("FY23 Q1", "Oct–Dec 2022", 20.77, 8.2, 6.4),
    ("FY23 Q2", "Jan–Mar 2023", 20.91, 0.7, 5.5),
    ("FY23 Q3", "Apr–Jun 2023", 21.21, 1.4, 8.2),
    ("FY23 Q4", "Jul–Sep 2023", 22.31, 5.2, 16.3),
    ("FY24 Q1", "Oct–Dec 2023", 23.12, 3.6, 11.3),
    ("FY24 Q2", "Jan–Mar 2024", 23.87, 3.2, 14.2),
    ("FY24 Q3", "Apr–Jun 2024", 24.21, 1.4, 14.1),
    ("FY24 Q4", "Jul–Sep 2024", 24.97, 3.1, 11.9),
    ("FY25 Q1", "Oct–Dec 2024", 26.34, 5.5, 13.9),
    ("FY25 Q2", "Jan–Mar 2025", 26.64, 1.1, 11.6),
    ("FY25 Q3", "Apr–Jun 2025", 27.42, 2.9, 13.3),
    ("FY25 Q4", "Jul–Sep 2025", 28.75, 4.9, 15.1),
    ("FY26 Q1", "Oct–Dec 2025", 30.01, 4.4, 13.9),
    ("FY26 Q2", "Jan–Mar 2026", 30.98, 3.2, 16.3),
]

for idx, row in enumerate(quarterly_data):
    r = 5 + idx
    ws2.cell(row=r, column=1, value=row[0])
    ws2.cell(row=r, column=2, value=row[1])
    ws2.cell(row=r, column=3, value=row[2])
    if row[3] is not None:
        ws2.cell(row=r, column=4, value=row[3])
    if row[4] is not None:
        ws2.cell(row=r, column=5, value=row[4])

style_data_area(ws2, 5, 5 + len(quarterly_data) - 1, 5, num_fmt="#,##0.00")
auto_width(ws2, 5)

# Bar + line combo chart for quarterly data
chart2 = BarChart()
chart2.type = "col"
chart2.title = "Apple Quarterly Services Revenue ($B)"
chart2.style = 10
chart2.y_axis.title = "Revenue ($B)"
chart2.width = 32
chart2.height = 15

cats2 = Reference(ws2, min_col=1, min_row=5, max_row=5 + len(quarterly_data) - 1)
vals2 = Reference(ws2, min_col=3, min_row=4, max_row=5 + len(quarterly_data) - 1)
chart2.add_data(vals2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.series[0].graphicalProperties.solidFill = APPLE_BLUE

ws2.add_chart(chart2, "A33")

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — Services Revenue by Category (Estimated Breakdown)
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Revenue by Category")
ws3.sheet_properties.tabColor = "FF9500"

ws3.merge_cells("A1:H1")
ws3["A1"] = "Apple Services Revenue — Estimated Breakdown by Category"
ws3["A1"].font = TITLE_FONT

ws3.merge_cells("A2:H2")
ws3["A2"] = (
    "Note: Apple does not disclose individual service revenues. "
    "Estimates from analyst reports (TechLila, Trefis, industry sources). Figures in USD billions."
)
ws3["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")

cat_headers = [
    "Category", "FY 2020 ($B)", "FY 2021 ($B)", "FY 2022 ($B)",
    "FY 2023 ($B)", "FY 2024 ($B)", "FY 2025 ($B)", "Description"
]
for i, h in enumerate(cat_headers, 1):
    ws3.cell(row=4, column=i, value=h)
style_header_row(ws3, 4, len(cat_headers))

category_data = [
    ("App Store", 18.0, 22.0, 24.0, 26.0, 30.0, 33.0, "Commissions on app sales & in-app purchases"),
    ("Licensing & Search", 8.5, 12.0, 14.0, 16.0, 18.0, 20.0, "Google search deal, other licensing agreements"),
    ("iCloud", 4.5, 5.5, 7.0, 8.0, 9.0, 10.5, "Paid cloud storage subscriptions"),
    ("AppleCare", 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, "Extended warranty & support services"),
    ("Apple Music", 5.0, 6.0, 7.0, 7.5, 8.5, 9.0, "Music streaming subscriptions"),
    ("Apple Pay & Financial", 1.5, 2.5, 3.5, 4.0, 5.0, 6.0, "Transaction fees, Apple Card, Apple Pay Later"),
    ("Apple TV+", 1.0, 2.0, 2.5, 2.5, 2.7, 3.0, "Streaming subscriptions & content licensing"),
    ("Advertising", 3.0, 4.0, 5.0, 5.5, 6.0, 7.0, "Search Ads, News, Stocks, App Store ads"),
    ("Other Services", 6.3, 7.9, 8.1, 8.2, 9.0, 12.2, "Fitness+, Arcade, News+, Maps, other"),
]

for idx, row in enumerate(category_data):
    r = 5 + idx
    for c, val in enumerate(row, 1):
        ws3.cell(row=r, column=c, value=val)

total_row = 5 + len(category_data)
ws3.cell(row=total_row, column=1, value="TOTAL (Reported)")
ws3.cell(row=total_row, column=1).font = Font(name="Calibri", bold=True, size=11)
totals_by_year = [53.8, 68.4, 78.1, 85.2, 96.2, 109.2]
for c, t in enumerate(totals_by_year, 2):
    cell = ws3.cell(row=total_row, column=c, value=t)
    cell.font = Font(name="Calibri", bold=True, size=11)

style_data_area(ws3, 5, total_row, 8, num_fmt="#,##0.0")
auto_width(ws3, 8, min_width=14)
ws3.column_dimensions["H"].width = 42

# Stacked bar chart for category breakdown
chart3 = BarChart()
chart3.type = "col"
chart3.grouping = "stacked"
chart3.title = "Estimated Services Revenue by Category ($B)"
chart3.style = 10
chart3.y_axis.title = "Revenue ($B)"
chart3.width = 30
chart3.height = 16

cats3 = Reference(ws3, min_col=2, min_row=4, max_col=7, max_row=4)
for row_idx in range(5, 5 + len(category_data)):
    vals3 = Reference(ws3, min_col=2, min_row=row_idx, max_col=7, max_row=row_idx)
    chart3.add_data(vals3, from_rows=True, titles_from_data=False)
    chart3.series[-1].title = openpyxl.chart.series.SeriesLabel(v=ws3.cell(row=row_idx, column=1).value)

year_labels = Reference(ws3, min_col=2, min_row=4, max_col=7, max_row=4)
chart3.set_categories(year_labels)

colors = ["007AFF", "34C759", "5856D6", "FF9500", "FF2D55", "AF52DE", "FFD60A", "64D2FF", "AC8E68"]
for i, s in enumerate(chart3.series):
    s.graphicalProperties.solidFill = colors[i % len(colors)]

ws3.add_chart(chart3, "A17")

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — Services vs Hardware & Margins
# ═══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Services vs Hardware")
ws4.sheet_properties.tabColor = "FF2D55"

ws4.merge_cells("A1:F1")
ws4["A1"] = "Apple Services vs Hardware — Revenue & Profitability Comparison"
ws4["A1"].font = TITLE_FONT

ws4.merge_cells("A2:F2")
ws4["A2"] = "Sources: Apple 10-K filings. Figures in USD billions."
ws4["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")

hw_headers = [
    "Fiscal Year", "Services Rev ($B)", "Products Rev ($B)",
    "Services Gross Margin (%)", "Products Gross Margin (%)", "Services % of Total Rev"
]
for i, h in enumerate(hw_headers, 1):
    ws4.cell(row=4, column=i, value=h)
style_header_row(ws4, 4, len(hw_headers))

hw_data = [
    ("FY 2019", 46.29, 213.88, 63.7, 32.2, 17.8),
    ("FY 2020", 53.77, 220.75, 66.0, 31.5, 19.6),
    ("FY 2021", 68.42, 297.39, 69.7, 35.3, 18.7),
    ("FY 2022", 78.13, 316.20, 71.7, 36.3, 19.8),
    ("FY 2023", 85.20, 298.09, 70.8, 36.5, 22.2),
    ("FY 2024", 96.17, 294.87, 74.0, 36.2, 24.6),
    ("FY 2025", 109.16, 307.00, 75.3, 36.2, 26.2),
]

for idx, row in enumerate(hw_data):
    r = 5 + idx
    for c, val in enumerate(row, 1):
        ws4.cell(row=r, column=c, value=val)

style_data_area(ws4, 5, 5 + len(hw_data) - 1, 6, num_fmt="#,##0.00")
auto_width(ws4, 6)

# Dual-axis chart: revenue bars + margin lines
chart4 = BarChart()
chart4.type = "col"
chart4.title = "Services vs Products Revenue ($B)"
chart4.style = 10
chart4.y_axis.title = "Revenue ($B)"
chart4.width = 28
chart4.height = 14

cats4 = Reference(ws4, min_col=1, min_row=5, max_row=5 + len(hw_data) - 1)
vals4a = Reference(ws4, min_col=2, min_row=4, max_row=5 + len(hw_data) - 1)
vals4b = Reference(ws4, min_col=3, min_row=4, max_row=5 + len(hw_data) - 1)
chart4.add_data(vals4a, titles_from_data=True)
chart4.add_data(vals4b, titles_from_data=True)
chart4.set_categories(cats4)
chart4.series[0].graphicalProperties.solidFill = APPLE_BLUE
chart4.series[1].graphicalProperties.solidFill = "CCCCCC"

ws4.add_chart(chart4, "A14")

# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — Key Metrics & Subscriber Data
# ═══════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Key Metrics")
ws5.sheet_properties.tabColor = "5856D6"

ws5.merge_cells("A1:D1")
ws5["A1"] = "Apple Services — Key Metrics & Subscriber Data (2025)"
ws5["A1"].font = TITLE_FONT

ws5.merge_cells("A2:D2")
ws5["A2"] = "Sources: Apple earnings calls, TechLila, analyst estimates."
ws5["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")

metrics = [
    ("Total Paid Subscriptions", "1+ billion", "2025"),
    ("App Store Weekly Active Users", "850 million", "2025"),
    ("Apple Music Subscribers", "~100 million", "2025"),
    ("Apple TV+ Subscribers", "~40–50 million", "2025 est."),
    ("iCloud Active Users", "850+ million", "2025"),
    ("Apple Pay Markets", "90+", "2026"),
    ("Apple Pay Bank/Network Partners", "11,000+", "2025"),
    ("Apple TV+ Viewing Hours Growth", "+36% YoY", "Dec 2025"),
    ("Services Gross Margin", "75.3%", "FY25 Q4"),
    ("Products Gross Margin", "36.2%", "FY25 Q4"),
    ("Services Revenue CAGR (FY15–FY25)", "~18.6%", "10-year"),
    ("App Store Annual Revenue", "~$33B", "FY 2025 est."),
    ("iCloud Annual Revenue", "~$10.5B", "FY 2025 est."),
    ("Licensing/Search Annual Revenue", "~$20B", "FY 2025 est."),
]

m_headers = ["Metric", "Value", "Period", "Notes"]
for i, h in enumerate(m_headers, 1):
    ws5.cell(row=4, column=i, value=h)
style_header_row(ws5, 4, len(m_headers))

for idx, row in enumerate(metrics):
    r = 5 + idx
    ws5.cell(row=r, column=1, value=row[0])
    ws5.cell(row=r, column=2, value=row[1])
    ws5.cell(row=r, column=3, value=row[2])

style_data_area(ws5, 5, 5 + len(metrics) - 1, 4)
auto_width(ws5, 4, min_width=18)
ws5.column_dimensions["A"].width = 38
ws5.column_dimensions["B"].width = 20

# ═══════════════════════════════════════════════════════════════════════════════
# Save
# ═══════════════════════════════════════════════════════════════════════════════
OUTPUT = "Apple_Services_Revenue_Breakdown.xlsx"
wb.save(OUTPUT)
print(f"Spreadsheet saved: {OUTPUT}")
