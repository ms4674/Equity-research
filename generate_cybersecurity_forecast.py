import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from copy import copy

HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
ACCENT_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
GROWTH_POSITIVE_FONT = Font(name="Calibri", color="006100", size=11)
GROWTH_NEGATIVE_FONT = Font(name="Calibri", color="9C0006", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="2F5496")
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=12, color="2F5496")
NORMAL_FONT = Font(name="Calibri", size=11)
BOLD_FONT = Font(name="Calibri", bold=True, size=11)
PCT_FORMAT = "0.0%"
USD_FORMAT = '#,##0.0'
USD_FORMAT_INT = '#,##0'
THIN_BORDER = Border(
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    bottom=Side(style="thin", color="B4C6E7"),
)

YEARS = list(range(2019, 2031))
HISTORICAL_END = 2024
FORECAST_LABEL = "2025E–2030E"

SEGMENTS = {
    "SIEM": {
        "full_name": "Security Information & Event Management (SIEM)",
        "market_size": {
            2019: 3.6, 2020: 3.9, 2021: 4.4, 2022: 5.0, 2023: 5.7,
            2024: 6.4, 2025: 7.2, 2026: 8.1, 2027: 9.2, 2028: 10.4,
            2029: 11.8, 2030: 13.4,
        },
        "companies": {
            "Splunk (Cisco)":    {2019: 0.90, 2020: 1.01, 2021: 1.14, 2022: 1.25, 2023: 1.37, 2024: 1.47, 2025: 1.58, 2026: 1.70, 2027: 1.84, 2028: 1.99, 2029: 2.12, 2030: 2.28},
            "Microsoft":        {2019: 0.18, 2020: 0.27, 2021: 0.44, 2022: 0.65, 2023: 0.86, 2024: 1.09, 2025: 1.37, 2026: 1.62, 2027: 1.93, 2028: 2.29, 2029: 2.71, 2030: 3.22},
            "IBM":              {2019: 0.54, 2020: 0.55, 2021: 0.55, 2022: 0.55, 2023: 0.54, 2024: 0.54, 2025: 0.54, 2026: 0.53, 2027: 0.53, 2028: 0.52, 2029: 0.52, 2030: 0.51},
            "Palo Alto (XSIAM)":{2019: 0.00, 2020: 0.00, 2021: 0.00, 2022: 0.05, 2023: 0.17, 2024: 0.38, 2025: 0.58, 2026: 0.81, 2027: 1.01, 2028: 1.25, 2029: 1.53, 2030: 1.87},
            "Securonix":        {2019: 0.11, 2020: 0.14, 2021: 0.18, 2022: 0.22, 2023: 0.26, 2024: 0.29, 2025: 0.32, 2026: 0.36, 2027: 0.41, 2028: 0.47, 2029: 0.53, 2030: 0.60},
            "Exabeam":          {2019: 0.07, 2020: 0.10, 2021: 0.13, 2022: 0.16, 2023: 0.18, 2024: 0.19, 2025: 0.22, 2026: 0.24, 2027: 0.28, 2028: 0.31, 2029: 0.35, 2030: 0.40},
            "LogRhythm":        {2019: 0.14, 2020: 0.14, 2021: 0.14, 2022: 0.13, 2023: 0.12, 2024: 0.11, 2025: 0.11, 2026: 0.10, 2027: 0.10, 2028: 0.09, 2029: 0.09, 2030: 0.08},
        },
    },
    "Vulnerability Assessment": {
        "full_name": "Vulnerability Assessment & Management",
        "market_size": {
            2019: 9.0, 2020: 9.8, 2021: 10.8, 2022: 11.9, 2023: 13.2,
            2024: 14.9, 2025: 16.1, 2026: 17.5, 2027: 19.0, 2028: 20.6,
            2029: 22.3, 2030: 24.1,
        },
        "companies": {
            "Tenable":          {2019: 0.35, 2020: 0.44, 2021: 0.54, 2022: 0.68, 2023: 0.79, 2024: 0.90, 2025: 1.01, 2026: 1.13, 2027: 1.27, 2028: 1.42, 2029: 1.58, 2030: 1.76},
            "Qualys":           {2019: 0.32, 2020: 0.36, 2021: 0.41, 2022: 0.49, 2023: 0.55, 2024: 0.62, 2025: 0.69, 2026: 0.77, 2027: 0.86, 2028: 0.96, 2029: 1.07, 2030: 1.19},
            "Rapid7":           {2019: 0.33, 2020: 0.41, 2021: 0.54, 2022: 0.68, 2023: 0.81, 2024: 0.88, 2025: 0.96, 2026: 1.05, 2027: 1.14, 2028: 1.24, 2029: 1.35, 2030: 1.47},
            "Microsoft":        {2019: 0.10, 2020: 0.15, 2021: 0.22, 2022: 0.33, 2023: 0.46, 2024: 0.60, 2025: 0.77, 2026: 0.96, 2027: 1.14, 2028: 1.34, 2029: 1.56, 2030: 1.81},
            "CrowdStrike":      {2019: 0.05, 2020: 0.08, 2021: 0.14, 2022: 0.24, 2023: 0.36, 2024: 0.52, 2025: 0.66, 2026: 0.79, 2027: 0.95, 2028: 1.14, 2029: 1.37, 2030: 1.64},
            "Ivanti":           {2019: 0.18, 2020: 0.20, 2021: 0.23, 2022: 0.25, 2023: 0.28, 2024: 0.30, 2025: 0.32, 2026: 0.35, 2027: 0.38, 2028: 0.41, 2029: 0.45, 2030: 0.48},
        },
    },
    "Identity (IAM)": {
        "full_name": "Identity & Access Management (IAM)",
        "market_size": {
            2019: 10.5, 2020: 12.0, 2021: 14.0, 2022: 16.5, 2023: 19.5,
            2024: 22.9, 2025: 26.5, 2026: 29.5, 2027: 32.0, 2028: 34.3,
            2029: 37.0, 2030: 40.0,
        },
        "companies": {
            "Microsoft":        {2019: 2.10, 2020: 2.64, 2021: 3.36, 2022: 4.29, 2023: 5.46, 2024: 6.87, 2025: 8.22, 2026: 9.44, 2027: 10.56, 2028: 11.66, 2029: 12.95, 2030: 14.40},
            "Okta":             {2019: 0.47, 2020: 0.59, 2021: 0.84, 2022: 1.32, 2023: 1.86, 2024: 2.52, 2025: 2.92, 2026: 3.25, 2027: 3.52, 2028: 3.77, 2029: 4.07, 2030: 4.40},
            "CyberArk":         {2019: 0.39, 2020: 0.42, 2021: 0.47, 2022: 0.56, 2023: 0.72, 2024: 0.94, 2025: 1.19, 2026: 1.48, 2027: 1.73, 2028: 1.89, 2029: 2.04, 2030: 2.20},
            "SailPoint":        {2019: 0.28, 2020: 0.32, 2021: 0.38, 2022: 0.45, 2023: 0.55, 2024: 0.64, 2025: 0.74, 2026: 0.83, 2027: 0.93, 2028: 1.03, 2029: 1.11, 2030: 1.20},
            "Ping Identity":    {2019: 0.24, 2020: 0.26, 2021: 0.30, 2022: 0.35, 2023: 0.42, 2024: 0.50, 2025: 0.58, 2026: 0.65, 2027: 0.74, 2028: 0.82, 2029: 0.89, 2030: 0.96},
            "IBM":              {2019: 0.63, 2020: 0.66, 2021: 0.70, 2022: 0.74, 2023: 0.78, 2024: 0.82, 2025: 0.85, 2026: 0.88, 2027: 0.90, 2028: 0.93, 2029: 0.96, 2030: 1.00},
            "Oracle":           {2019: 0.53, 2020: 0.54, 2021: 0.56, 2022: 0.59, 2023: 0.63, 2024: 0.69, 2025: 0.74, 2026: 0.80, 2027: 0.86, 2028: 0.93, 2029: 1.00, 2030: 1.08},
        },
    },
    "Endpoint Security": {
        "full_name": "Endpoint Security (EPP / EDR / XDR)",
        "market_size": {
            2019: 12.8, 2020: 13.5, 2021: 14.6, 2022: 16.1, 2023: 17.8,
            2024: 18.4, 2025: 19.8, 2026: 21.3, 2027: 22.9, 2028: 24.7,
            2029: 26.6, 2030: 28.6,
        },
        "companies": {
            "CrowdStrike":      {2019: 0.46, 2020: 0.68, 2021: 1.07, 2022: 1.73, 2023: 2.45, 2024: 3.06, 2025: 3.52, 2026: 3.94, 2027: 4.35, 2028: 4.69, 2029: 5.05, 2030: 5.43},
            "Microsoft":        {2019: 1.15, 2020: 1.49, 2021: 1.90, 2022: 2.42, 2023: 3.02, 2024: 3.50, 2025: 3.96, 2026: 4.47, 2027: 5.04, 2028: 5.68, 2029: 6.38, 2030: 7.15},
            "Palo Alto Networks":{2019: 0.38, 2020: 0.54, 2021: 0.73, 2022: 0.97, 2023: 1.25, 2024: 1.47, 2025: 1.68, 2026: 1.92, 2027: 2.18, 2028: 2.47, 2029: 2.79, 2030: 3.15},
            "SentinelOne":      {2019: 0.05, 2020: 0.10, 2021: 0.20, 2022: 0.42, 2023: 0.62, 2024: 0.74, 2025: 0.85, 2026: 0.96, 2027: 1.07, 2028: 1.19, 2029: 1.33, 2030: 1.49},
            "Trellix (McAfee)": {2019: 1.41, 2020: 1.35, 2021: 1.31, 2022: 1.21, 2023: 1.12, 2024: 1.01, 2025: 0.95, 2026: 0.89, 2027: 0.82, 2028: 0.77, 2029: 0.72, 2030: 0.66},
            "Trend Micro":      {2019: 0.90, 2020: 0.90, 2021: 0.92, 2022: 0.93, 2023: 0.93, 2024: 0.92, 2025: 0.93, 2026: 0.94, 2027: 0.96, 2028: 0.99, 2029: 1.01, 2030: 1.03},
            "Broadcom (Symantec)":{2019: 1.28, 2020: 1.15, 2021: 1.05, 2022: 0.97, 2023: 0.89, 2024: 0.83, 2025: 0.77, 2026: 0.72, 2027: 0.69, 2028: 0.64, 2029: 0.61, 2030: 0.57},
        },
    },
    "Network Security": {
        "full_name": "Network Security (Firewall / IDS/IPS / UTM)",
        "market_size": {
            2019: 15.5, 2020: 16.6, 2021: 18.3, 2022: 20.5, 2023: 23.8,
            2024: 26.6, 2025: 28.7, 2026: 31.0, 2027: 33.5, 2028: 36.2,
            2029: 39.1, 2030: 42.3,
        },
        "companies": {
            "Palo Alto Networks":{2019: 2.79, 2020: 3.07, 2021: 3.66, 2022: 4.31, 2023: 5.24, 2024: 5.99, 2025: 6.60, 2026: 7.28, 2027: 8.04, 2028: 8.69, 2029: 9.39, 2030: 10.15},
            "Fortinet":         {2019: 2.02, 2020: 2.32, 2021: 2.93, 2022: 3.59, 2023: 4.21, 2024: 4.71, 2025: 5.17, 2026: 5.58, 2027: 6.03, 2028: 6.51, 2029: 7.04, 2030: 7.61},
            "Cisco":            {2019: 2.17, 2020: 2.22, 2021: 2.38, 2022: 2.56, 2023: 2.85, 2024: 2.93, 2025: 3.07, 2026: 3.22, 2027: 3.35, 2028: 3.44, 2029: 3.52, 2030: 3.59},
            "Check Point":      {2019: 1.86, 2020: 1.94, 2021: 2.06, 2022: 2.15, 2023: 2.14, 2024: 2.13, 2025: 2.15, 2026: 2.17, 2027: 2.18, 2028: 2.18, 2029: 2.18, 2030: 2.15},
            "Juniper Networks":  {2019: 0.62, 2020: 0.63, 2021: 0.64, 2022: 0.72, 2023: 0.81, 2024: 0.85, 2025: 0.89, 2026: 0.93, 2027: 0.97, 2028: 1.01, 2029: 1.05, 2030: 1.10},
            "Zscaler":          {2019: 0.19, 2020: 0.31, 2021: 0.51, 2022: 0.82, 2023: 1.14, 2024: 1.46, 2025: 1.72, 2026: 2.01, 2027: 2.35, 2028: 2.75, 2029: 3.21, 2030: 3.76},
        },
    },
    "SASE": {
        "full_name": "Secure Access Service Edge (SASE)",
        "market_size": {
            2019: 1.2, 2020: 2.0, 2021: 3.4, 2022: 5.6, 2023: 8.4,
            2024: 9.6, 2025: 10.9, 2026: 12.4, 2027: 14.1, 2028: 16.0,
            2029: 18.2, 2030: 20.7,
        },
        "companies": {
            "Zscaler":          {2019: 0.22, 2020: 0.39, 2021: 0.67, 2022: 1.18, 2023: 1.85, 2024: 2.02, 2025: 2.29, 2026: 2.60, 2027: 2.96, 2028: 3.36, 2029: 3.82, 2030: 4.35},
            "Cisco":            {2019: 0.14, 2020: 0.26, 2021: 0.48, 2022: 0.84, 2023: 1.34, 2024: 1.54, 2025: 1.74, 2026: 1.98, 2027: 2.26, 2028: 2.56, 2029: 2.91, 2030: 3.31},
            "Palo Alto Networks":{2019: 0.08, 2020: 0.18, 2021: 0.37, 2022: 0.73, 2023: 1.18, 2024: 1.34, 2025: 1.53, 2026: 1.74, 2027: 1.97, 2028: 2.24, 2029: 2.55, 2030: 2.90},
            "Broadcom":         {2019: 0.06, 2020: 0.12, 2021: 0.24, 2022: 0.45, 2023: 0.76, 2024: 0.86, 2025: 0.98, 2026: 1.12, 2027: 1.27, 2028: 1.44, 2029: 1.64, 2030: 1.86},
            "Fortinet":         {2019: 0.05, 2020: 0.10, 2021: 0.20, 2022: 0.39, 2023: 0.59, 2024: 0.67, 2025: 0.76, 2026: 0.87, 2027: 0.99, 2028: 1.12, 2029: 1.27, 2030: 1.45},
            "Netskope":         {2019: 0.04, 2020: 0.08, 2021: 0.17, 2022: 0.34, 2023: 0.50, 2024: 0.58, 2025: 0.65, 2026: 0.74, 2027: 0.85, 2028: 0.96, 2029: 1.09, 2030: 1.24},
            "Cloudflare":       {2019: 0.02, 2020: 0.04, 2021: 0.10, 2022: 0.22, 2023: 0.34, 2024: 0.43, 2025: 0.54, 2026: 0.68, 2027: 0.85, 2028: 1.06, 2029: 1.33, 2030: 1.66},
        },
    },
}


def apply_cell_style(cell, font=None, fill=None, alignment=None, border=None, number_format=None):
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format


def style_range(ws, row, col_start, col_end, **kwargs):
    for c in range(col_start, col_end + 1):
        apply_cell_style(ws.cell(row=row, column=c), **kwargs)


def write_segment_sheet(wb, segment_key, data):
    ws = wb.create_sheet(title=segment_key)

    ws.cell(row=1, column=1, value=data["full_name"])
    apply_cell_style(ws.cell(row=1, column=1), font=TITLE_FONT)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(YEARS) + 2)

    ws.cell(row=2, column=1, value="All figures in USD Billions. Sources: IDC, Gartner, Dell'Oro, company filings. 2025-2030 are estimates.")
    apply_cell_style(ws.cell(row=2, column=1), font=Font(name="Calibri", italic=True, size=10, color="666666"))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(YEARS) + 2)

    # --- SECTION 1: Total Market Size ---
    row = 4
    ws.cell(row=row, column=1, value="Total Market Size ($B)")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="center"))
    style_range(ws, row, 1, len(YEARS) + 1, border=THIN_BORDER)

    row = 5
    ws.cell(row=row, column=1, value="Market Size")
    apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, border=THIN_BORDER)
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=data["market_size"][yr])
        fmt = USD_FORMAT
        apply_cell_style(cell, font=NORMAL_FONT, border=THIN_BORDER, number_format=fmt, alignment=Alignment(horizontal="center"))
        if yr > HISTORICAL_END:
            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    row = 6
    ws.cell(row=row, column=1, value="YoY Growth (%)")
    apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, border=THIN_BORDER)
    for i, yr in enumerate(YEARS):
        if i == 0:
            ws.cell(row=row, column=i + 2, value="—")
            apply_cell_style(ws.cell(row=row, column=i + 2), font=NORMAL_FONT, border=THIN_BORDER, alignment=Alignment(horizontal="center"))
            continue
        prev = data["market_size"][YEARS[i - 1]]
        cur = data["market_size"][yr]
        growth = (cur - prev) / prev
        cell = ws.cell(row=row, column=i + 2, value=growth)
        gfont = GROWTH_POSITIVE_FONT if growth >= 0 else GROWTH_NEGATIVE_FONT
        apply_cell_style(cell, font=gfont, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))
        if yr > HISTORICAL_END:
            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # --- SECTION 2: Company Revenue ---
    row = 8
    ws.cell(row=row, column=1, value="Company Revenue ($B)")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="center"))
    style_range(ws, row, 1, len(YEARS) + 1, border=THIN_BORDER)

    company_start_row = row + 1
    for ci, (company, rev_data) in enumerate(data["companies"].items()):
        row = company_start_row + ci
        ws.cell(row=row, column=1, value=company)
        f = ACCENT_FILL if ci % 2 == 0 else None
        apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, fill=f, border=THIN_BORDER)
        for i, yr in enumerate(YEARS):
            cell = ws.cell(row=row, column=i + 2, value=rev_data[yr])
            apply_cell_style(cell, font=NORMAL_FONT, fill=f, border=THIN_BORDER, number_format=USD_FORMAT, alignment=Alignment(horizontal="center"))
            if yr > HISTORICAL_END:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Others row
    row += 1
    ws.cell(row=row, column=1, value="Others")
    apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, border=THIN_BORDER)
    for i, yr in enumerate(YEARS):
        company_total = sum(rev[yr] for rev in data["companies"].values())
        others = data["market_size"][yr] - company_total
        cell = ws.cell(row=row, column=i + 2, value=round(others, 2))
        apply_cell_style(cell, font=NORMAL_FONT, border=THIN_BORDER, number_format=USD_FORMAT, alignment=Alignment(horizontal="center"))
        if yr > HISTORICAL_END:
            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # --- SECTION 3: Company Market Share ---
    row += 2
    share_header_row = row
    ws.cell(row=row, column=1, value="Market Share (%)")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="center"))
    style_range(ws, row, 1, len(YEARS) + 1, border=THIN_BORDER)

    for ci, (company, rev_data) in enumerate(data["companies"].items()):
        row = share_header_row + 1 + ci
        ws.cell(row=row, column=1, value=company)
        f = ACCENT_FILL if ci % 2 == 0 else None
        apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, fill=f, border=THIN_BORDER)
        for i, yr in enumerate(YEARS):
            share = rev_data[yr] / data["market_size"][yr]
            cell = ws.cell(row=row, column=i + 2, value=share)
            apply_cell_style(cell, font=NORMAL_FONT, fill=f, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))
            if yr > HISTORICAL_END:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Others share
    row += 1
    ws.cell(row=row, column=1, value="Others")
    apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, border=THIN_BORDER)
    for i, yr in enumerate(YEARS):
        company_total = sum(rev[yr] for rev in data["companies"].values())
        others_share = (data["market_size"][yr] - company_total) / data["market_size"][yr]
        cell = ws.cell(row=row, column=i + 2, value=others_share)
        apply_cell_style(cell, font=NORMAL_FONT, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))
        if yr > HISTORICAL_END:
            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # --- SECTION 4: Company Revenue Growth ---
    row += 2
    growth_header_row = row
    ws.cell(row=row, column=1, value="Company Revenue YoY Growth (%)")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="center"))
    style_range(ws, row, 1, len(YEARS) + 1, border=THIN_BORDER)

    for ci, (company, rev_data) in enumerate(data["companies"].items()):
        row = growth_header_row + 1 + ci
        ws.cell(row=row, column=1, value=company)
        f = ACCENT_FILL if ci % 2 == 0 else None
        apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, fill=f, border=THIN_BORDER)
        for i, yr in enumerate(YEARS):
            if i == 0:
                ws.cell(row=row, column=i + 2, value="—")
                apply_cell_style(ws.cell(row=row, column=i + 2), font=NORMAL_FONT, fill=f, border=THIN_BORDER, alignment=Alignment(horizontal="center"))
                continue
            prev_val = rev_data[YEARS[i - 1]]
            cur_val = rev_data[yr]
            if prev_val == 0:
                ws.cell(row=row, column=i + 2, value="N/A")
                apply_cell_style(ws.cell(row=row, column=i + 2), font=NORMAL_FONT, fill=f, border=THIN_BORDER, alignment=Alignment(horizontal="center"))
                continue
            growth = (cur_val - prev_val) / prev_val
            cell = ws.cell(row=row, column=i + 2, value=growth)
            gfont = Font(name="Calibri", color="006100", size=11) if growth >= 0 else Font(name="Calibri", color="9C0006", size=11)
            apply_cell_style(cell, font=gfont, fill=f, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))
            if yr > HISTORICAL_END:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # --- Add chart: Market Size bar + line ---
    chart_row = row + 3
    chart = BarChart()
    chart.type = "col"
    chart.title = f"{segment_key} Market Size ($B)"
    chart.y_axis.title = "USD Billions"
    chart.x_axis.title = "Year"
    chart.style = 10
    chart.width = 28
    chart.height = 14

    cats = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=4)
    vals = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=5)
    chart.add_data(vals, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].title = openpyxl.chart.series.SeriesLabel(v="Market Size")
    chart.series[0].graphicalProperties.solidFill = "4472C4"

    ws.add_chart(chart, f"A{chart_row}")

    # Column widths
    ws.column_dimensions["A"].width = 28
    for i in range(len(YEARS)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12

    ws.sheet_properties.tabColor = "2F5496"


def write_summary_sheet(wb):
    ws = wb.create_sheet(title="Summary", index=0)

    ws.cell(row=1, column=1, value="Cybersecurity Market Segments — Historical & Forecast Overview")
    apply_cell_style(ws.cell(row=1, column=1), font=TITLE_FONT)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(YEARS) + 2)

    ws.cell(row=2, column=1, value="All figures in USD Billions. Yellow-highlighted columns are forecasts (2025E–2030E).")
    apply_cell_style(ws.cell(row=2, column=1), font=Font(name="Calibri", italic=True, size=10, color="666666"))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(YEARS) + 2)

    # --- Header row ---
    row = 4
    ws.cell(row=row, column=1, value="Segment")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="center"))
    cagr_col = len(YEARS) + 2
    ws.cell(row=row, column=cagr_col, value="CAGR\n'19-'30")
    apply_cell_style(ws.cell(row=row, column=cagr_col), font=HEADER_FONT, fill=HEADER_FILL, alignment=Alignment(horizontal="center", wrap_text=True))
    style_range(ws, row, 1, cagr_col, border=THIN_BORDER)

    # --- Market size rows ---
    seg_order = ["SIEM", "Vulnerability Assessment", "Identity (IAM)", "Endpoint Security", "Network Security", "SASE"]
    total_by_year = {yr: 0 for yr in YEARS}

    for si, seg_key in enumerate(seg_order):
        row = 5 + si
        d = SEGMENTS[seg_key]
        f = ACCENT_FILL if si % 2 == 0 else None
        ws.cell(row=row, column=1, value=seg_key)
        apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, fill=f, border=THIN_BORDER)
        for i, yr in enumerate(YEARS):
            val = d["market_size"][yr]
            total_by_year[yr] += val
            cell = ws.cell(row=row, column=i + 2, value=val)
            apply_cell_style(cell, font=NORMAL_FONT, fill=f, border=THIN_BORDER, number_format=USD_FORMAT, alignment=Alignment(horizontal="center"))
            if yr > HISTORICAL_END:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        cagr = (d["market_size"][2030] / d["market_size"][2019]) ** (1 / 11) - 1
        cell = ws.cell(row=row, column=cagr_col, value=cagr)
        apply_cell_style(cell, font=BOLD_FONT, fill=f, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))

    # Total row
    row = 5 + len(seg_order)
    ws.cell(row=row, column=1, value="TOTAL")
    apply_cell_style(ws.cell(row=row, column=1), font=Font(name="Calibri", bold=True, size=11, color="FFFFFF"), fill=PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid"), border=THIN_BORDER)
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=round(total_by_year[yr], 1))
        apply_cell_style(cell, font=Font(name="Calibri", bold=True, size=11, color="FFFFFF"),
                         fill=PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid"),
                         border=THIN_BORDER, number_format=USD_FORMAT, alignment=Alignment(horizontal="center"))
    total_cagr = (total_by_year[2030] / total_by_year[2019]) ** (1 / 11) - 1
    cell = ws.cell(row=row, column=cagr_col, value=total_cagr)
    apply_cell_style(cell, font=Font(name="Calibri", bold=True, size=11, color="FFFFFF"),
                     fill=PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid"),
                     border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))

    # --- Growth rates section ---
    row += 2
    growth_section_start = row
    ws.cell(row=row, column=1, value="YoY Growth Rate (%)")
    apply_cell_style(ws.cell(row=row, column=1), font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="left"))
    for i, yr in enumerate(YEARS):
        cell = ws.cell(row=row, column=i + 2, value=yr)
        apply_cell_style(cell, font=HEADER_FONT, fill=SUBHEADER_FILL, alignment=Alignment(horizontal="center"))
    style_range(ws, row, 1, len(YEARS) + 1, border=THIN_BORDER)

    for si, seg_key in enumerate(seg_order):
        row = growth_section_start + 1 + si
        d = SEGMENTS[seg_key]
        f = ACCENT_FILL if si % 2 == 0 else None
        ws.cell(row=row, column=1, value=seg_key)
        apply_cell_style(ws.cell(row=row, column=1), font=BOLD_FONT, fill=f, border=THIN_BORDER)
        for i, yr in enumerate(YEARS):
            if i == 0:
                ws.cell(row=row, column=i + 2, value="—")
                apply_cell_style(ws.cell(row=row, column=i + 2), font=NORMAL_FONT, fill=f, border=THIN_BORDER, alignment=Alignment(horizontal="center"))
                continue
            prev = d["market_size"][YEARS[i - 1]]
            cur = d["market_size"][yr]
            growth = (cur - prev) / prev
            cell = ws.cell(row=row, column=i + 2, value=growth)
            gfont = GROWTH_POSITIVE_FONT if growth >= 0 else GROWTH_NEGATIVE_FONT
            apply_cell_style(cell, font=gfont, fill=f, border=THIN_BORDER, number_format=PCT_FORMAT, alignment=Alignment(horizontal="center"))
            if yr > HISTORICAL_END:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # --- Chart ---
    chart_row = row + 3
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.title = "Cybersecurity Segments — Market Size ($B)"
    chart.y_axis.title = "USD Billions"
    chart.x_axis.title = "Year"
    chart.style = 10
    chart.width = 30
    chart.height = 16

    cats = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=4)
    colors = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    for si, seg_key in enumerate(seg_order):
        vals = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=5 + si)
        chart.add_data(vals, from_rows=True, titles_from_data=False)
        chart.set_categories(cats)
        chart.series[si].title = openpyxl.chart.series.SeriesLabel(v=seg_key)
        chart.series[si].graphicalProperties.solidFill = colors[si % len(colors)]

    ws.add_chart(chart, f"A{chart_row}")

    # Column widths
    ws.column_dimensions["A"].width = 28
    for i in range(len(YEARS)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12
    ws.column_dimensions[get_column_letter(cagr_col)].width = 12

    ws.sheet_properties.tabColor = "1F3864"


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_summary_sheet(wb)

    for seg_key in ["SIEM", "Vulnerability Assessment", "Identity (IAM)", "Endpoint Security", "Network Security", "SASE"]:
        write_segment_sheet(wb, seg_key, SEGMENTS[seg_key])

    output = "cybersecurity_segments_forecast.xlsx"
    wb.save(output)
    print(f"Workbook saved to {output}")


if __name__ == "__main__":
    main()
