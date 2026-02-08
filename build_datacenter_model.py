#!/usr/bin/env python3
"""
Datacenter Model Builder
========================
Creates a comprehensive Excel-based datacenter model covering:
- Capex Spend ($B)
- Server Counts (thousands)
- Datacenter Counts
- GW Capacity
- Power Generation & Availability
- Revenue/MW Analysis

Covers hyperscalers (AWS, Azure, Google Cloud, Meta, Oracle, Apple) and
neocloud vendors (CoreWeave, Lambda, Crusoe, Voltage Park, Together AI, Applied Digital).

Historical data: 2018-2024 (estimated from public filings, earnings, industry reports)
Forecast: 2025E-2030E
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.chart import BarChart, LineChart, Reference, BarChart3D
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from copy import copy

# ============================================================
# COLOR PALETTE & STYLE CONSTANTS
# ============================================================
DARK_BLUE = "1B2A4A"
MED_BLUE = "2E5090"
LIGHT_BLUE = "D6E4F0"
ACCENT_BLUE = "4472C4"
ACCENT_ORANGE = "ED7D31"
ACCENT_GREEN = "70AD47"
ACCENT_GOLD = "FFC000"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
BORDER_GRAY = "B4B4B4"
FORECAST_BG = "FFF2CC"  # Light yellow for forecast columns
HEADER_FONT_COLOR = WHITE

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=WHITE)
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color=DARK_BLUE)
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color=WHITE)
DATA_FONT = Font(name="Calibri", size=10)
DATA_FONT_BOLD = Font(name="Calibri", size=10, bold=True)
COMPANY_FONT = Font(name="Calibri", size=10, bold=True, color=DARK_BLUE)
TOTAL_FONT = Font(name="Calibri", size=10, bold=True, color=WHITE)
NOTE_FONT = Font(name="Calibri", size=8, italic=True, color="666666")
PCT_FONT = Font(name="Calibri", size=9, italic=True, color="666666")

TITLE_FILL = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
HEADER_FILL = PatternFill(start_color=MED_BLUE, end_color=MED_BLUE, fill_type="solid")
LIGHT_FILL = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
ALT_ROW_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
TOTAL_FILL = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
FORECAST_FILL = PatternFill(start_color=FORECAST_BG, end_color=FORECAST_BG, fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color=BORDER_GRAY),
    right=Side(style="thin", color=BORDER_GRAY),
    top=Side(style="thin", color=BORDER_GRAY),
    bottom=Side(style="thin", color=BORDER_GRAY),
)

# Years
YEARS = list(range(2018, 2031))
HIST_YEARS = list(range(2018, 2025))
FCST_YEARS = list(range(2025, 2031))
YEAR_LABELS = [str(y) if y < 2025 else f"{y}E" for y in YEARS]

# Companies
HYPERSCALERS = ["Amazon (AWS)", "Microsoft (Azure)", "Google (GCP)", "Meta", "Oracle Cloud", "Apple"]
NEOCLOUDS = ["CoreWeave", "Lambda", "Crusoe Energy", "Voltage Park", "Together AI", "Applied Digital"]
ALL_COMPANIES = HYPERSCALERS + NEOCLOUDS

# ============================================================
# DATA TABLES  (All estimates; sources: public filings,
# earnings calls, industry analyst reports, press releases)
# Units noted per section
# ============================================================

# --- CAPEX SPEND ($B) ---
CAPEX_DATA = {
    # Historical 2018-2024, Forecast 2025E-2030E
    "Amazon (AWS)":       [13.4, 16.1, 21.0, 27.0, 36.0, 48.4, 75.0, 100.0, 120.0, 138.0, 150.0, 155.0, 158.0],
    "Microsoft (Azure)":  [11.6, 13.9, 15.4, 22.0, 28.0, 32.0, 55.7, 80.0, 95.0, 110.0, 120.0, 125.0, 128.0],
    "Google (GCP)":       [10.1, 12.0, 15.0, 21.0, 26.0, 32.3, 52.5, 75.0, 85.0, 95.0, 100.0, 105.0, 108.0],
    "Meta":               [ 6.7,  8.4, 11.0, 15.0, 18.0, 28.0, 38.0, 60.0, 68.0, 75.0,  78.0,  80.0,  82.0],
    "Oracle Cloud":       [ 1.8,  1.9,  2.4,  3.0,  4.0,  6.9, 13.0, 18.0, 22.0, 25.0,  27.0,  28.0,  29.0],
    "Apple":              [ 4.5,  5.0,  5.5,  6.0,  6.5,  7.0,  9.3, 11.0, 13.0, 15.0,  16.0,  17.0,  18.0],
    "CoreWeave":          [ 0.0,  0.0,  0.0,  0.0,  0.1,  0.8,  3.5,  8.0, 12.0, 15.0,  17.0,  18.0,  19.0],
    "Lambda":             [ 0.0,  0.0,  0.0,  0.0,  0.05, 0.2,  0.8,  2.0,  3.5,  5.0,   6.0,   7.0,   7.5],
    "Crusoe Energy":      [ 0.0,  0.0,  0.0,  0.0,  0.05, 0.3,  1.0,  2.5,  4.0,  5.5,   7.0,   8.0,   8.5],
    "Voltage Park":       [ 0.0,  0.0,  0.0,  0.0,  0.0,  0.1,  0.5,  1.5,  2.5,  3.5,   4.5,   5.0,   5.5],
    "Together AI":        [ 0.0,  0.0,  0.0,  0.0,  0.0,  0.1,  0.3,  0.8,  1.5,  2.5,   3.5,   4.0,   4.5],
    "Applied Digital":    [ 0.0,  0.0,  0.0,  0.0,  0.02, 0.15, 0.6,  1.5,  2.5,  3.5,   4.5,   5.0,   5.5],
}

# --- REVENUE ($B) --- used for Revenue/MW calculation
REVENUE_DATA = {
    "Amazon (AWS)":       [25.7, 35.0, 45.4, 62.2, 80.1, 90.8, 107.6, 130.0, 155.0, 182.0, 210.0, 240.0, 270.0],
    "Microsoft (Azure)":  [23.2, 33.7, 43.1, 60.0, 75.0, 96.8, 125.0, 160.0, 195.0, 230.0, 265.0, 300.0, 335.0],
    "Google (GCP)":       [ 5.8,  8.9, 13.1, 19.2, 26.3, 33.7,  43.2,  55.0,  68.0,  82.0,  97.0, 112.0, 128.0],
    "Meta":               [55.8, 70.7, 86.0, 117.9, 116.6, 134.9, 164.5, 195.0, 225.0, 255.0, 285.0, 315.0, 345.0],
    "Oracle Cloud":       [ 6.8,  7.0,  7.1,  7.5,  8.4, 12.5,  18.0,  24.0,  30.0,  36.0,  42.0,  48.0,  54.0],
    "Apple":              [265.6, 260.2, 274.5, 365.8, 394.3, 383.3, 391.0, 410.0, 430.0, 450.0, 470.0, 490.0, 510.0],
    "CoreWeave":          [ 0.0,  0.0,  0.0,  0.0,  0.02, 0.2,   0.8,   2.5,   5.0,   8.0,  11.0,  14.0,  17.0],
    "Lambda":             [ 0.0,  0.0,  0.0,  0.0,  0.01, 0.05,  0.3,   0.8,   1.5,   2.5,   3.5,   4.5,   5.5],
    "Crusoe Energy":      [ 0.0,  0.0,  0.0,  0.0,  0.01, 0.05,  0.2,   0.6,   1.2,   2.0,   3.0,   4.0,   5.0],
    "Voltage Park":       [ 0.0,  0.0,  0.0,  0.0,  0.0,  0.02,  0.1,   0.4,   0.8,   1.5,   2.5,   3.5,   4.5],
    "Together AI":        [ 0.0,  0.0,  0.0,  0.0,  0.0,  0.03,  0.1,   0.3,   0.7,   1.2,   2.0,   3.0,   4.0],
    "Applied Digital":    [ 0.0,  0.0,  0.0,  0.0,  0.005,0.04,  0.15,  0.5,   1.0,   1.8,   2.8,   3.8,   4.8],
}

# --- SERVER COUNT (thousands of servers) ---
SERVER_DATA = {
    "Amazon (AWS)":       [2800, 3200, 3800, 4500, 5200, 5900, 7000, 8500, 10000, 11500, 13000, 14500, 16000],
    "Microsoft (Azure)":  [2200, 2600, 3000, 3800, 4500, 5200, 6500, 8000,  9500, 11000, 12500, 14000, 15000],
    "Google (GCP)":       [2000, 2300, 2700, 3200, 3800, 4500, 5500, 7000,  8500, 10000, 11500, 13000, 14000],
    "Meta":               [ 600,  750,  900, 1200, 1500, 2200, 3000, 4000,  5000,  6000,  7000,  8000,  9000],
    "Oracle Cloud":       [ 200,  250,  350,  500,  700, 1000, 1500, 2200,  3000,  3800,  4500,  5000,  5500],
    "Apple":              [ 150,  180,  220,  280,  340,  400,  500,  650,   800,   950,  1100,  1250,  1400],
    "CoreWeave":          [   0,    0,    0,    0,    2,   15,   60,  150,   280,   400,   520,   640,   750],
    "Lambda":             [   0,    0,    0,    0,    1,    5,   20,   50,   100,   160,   220,   280,   340],
    "Crusoe Energy":      [   0,    0,    0,    0,    1,    8,   25,   60,   120,   180,   250,   320,   380],
    "Voltage Park":       [   0,    0,    0,    0,    0,    3,   12,   35,    70,   110,   155,   200,   240],
    "Together AI":        [   0,    0,    0,    0,    0,    2,    8,   20,    45,    75,   110,   150,   190],
    "Applied Digital":    [   0,    0,    0,    0,    1,    4,   15,   40,    80,   130,   180,   230,   280],
}

# --- DATACENTER COUNT ---
DC_COUNT_DATA = {
    "Amazon (AWS)":       [ 60,  69,  76,  81,  87,  96, 110, 130, 150, 170, 185, 200, 215],
    "Microsoft (Azure)":  [ 50,  54,  58,  60,  65,  72,  82,  95, 110, 125, 140, 155, 170],
    "Google (GCP)":       [ 21,  23,  26,  29,  33,  38,  45,  55,  65,  75,  85,  95, 105],
    "Meta":               [ 10,  12,  14,  16,  18,  22,  28,  35,  42,  50,  58,  65,  72],
    "Oracle Cloud":       [ 16,  20,  28,  33,  40,  47,  55,  66,  78,  90, 100, 108, 115],
    "Apple":              [  5,   5,   6,   6,   7,   8,  10,  12,  14,  16,  18,  20,  22],
    "CoreWeave":          [  0,   0,   0,   0,   1,   4,  10,  18,  28,  38,  48,  56,  62],
    "Lambda":             [  0,   0,   0,   0,   1,   2,   4,   8,  14,  20,  26,  32,  38],
    "Crusoe Energy":      [  0,   0,   0,   0,   1,   3,   6,  12,  20,  28,  36,  42,  48],
    "Voltage Park":       [  0,   0,   0,   0,   0,   1,   3,   6,  10,  15,  20,  25,  30],
    "Together AI":        [  0,   0,   0,   0,   0,   1,   2,   4,   7,  11,  16,  21,  26],
    "Applied Digital":    [  0,   0,   0,   0,   1,   2,   4,   8,  13,  18,  24,  30,  35],
}

# --- GW CAPACITY (IT Load, GW) ---
GW_CAPACITY_DATA = {
    "Amazon (AWS)":       [2.5, 3.0, 3.6, 4.5, 5.5, 6.8, 8.5, 11.0, 14.0, 17.0, 20.0, 23.0, 26.0],
    "Microsoft (Azure)":  [2.0, 2.4, 3.0, 3.8, 4.8, 6.0, 7.5, 10.0, 13.0, 16.0, 19.0, 22.0, 25.0],
    "Google (GCP)":       [1.8, 2.1, 2.6, 3.2, 4.0, 5.0, 6.5,  8.5, 11.0, 13.5, 16.0, 18.5, 21.0],
    "Meta":               [0.8, 1.0, 1.3, 1.8, 2.3, 3.5, 5.0,  7.0,  9.0, 11.0, 13.0, 15.0, 17.0],
    "Oracle Cloud":       [0.3, 0.4, 0.5, 0.7, 1.0, 1.5, 2.5,  3.8,  5.0,  6.5,  8.0,  9.5, 11.0],
    "Apple":              [0.2, 0.3, 0.3, 0.4, 0.5, 0.6, 0.8,  1.0,  1.3,  1.6,  1.9,  2.2,  2.5],
    "CoreWeave":          [0.0, 0.0, 0.0, 0.0, 0.01,0.08,0.35, 0.9,  1.6,  2.5,  3.5,  4.5,  5.5],
    "Lambda":             [0.0, 0.0, 0.0, 0.0, 0.005,0.03,0.12,0.30, 0.60, 1.00, 1.40, 1.80, 2.20],
    "Crusoe Energy":      [0.0, 0.0, 0.0, 0.0, 0.005,0.05,0.15,0.40, 0.80, 1.20, 1.70, 2.20, 2.70],
    "Voltage Park":       [0.0, 0.0, 0.0, 0.0, 0.0, 0.02,0.08,0.20, 0.45, 0.75, 1.10, 1.45, 1.80],
    "Together AI":        [0.0, 0.0, 0.0, 0.0, 0.0, 0.01,0.05,0.12, 0.28, 0.50, 0.80, 1.10, 1.40],
    "Applied Digital":    [0.0, 0.0, 0.0, 0.0, 0.003,0.03,0.10,0.25, 0.50, 0.80, 1.20, 1.60, 2.00],
}

# --- POWER CONTRACTED / AVAILABLE (GW) - total power secured (may exceed IT load) ---
POWER_AVAILABLE_DATA = {
    "Amazon (AWS)":       [3.0, 3.6, 4.5, 5.5, 7.0, 9.0, 12.0, 15.0, 19.0, 23.0, 27.0, 31.0, 35.0],
    "Microsoft (Azure)":  [2.5, 3.0, 3.8, 5.0, 6.2, 8.0, 10.5, 14.0, 18.0, 22.0, 26.0, 30.0, 34.0],
    "Google (GCP)":       [2.2, 2.6, 3.2, 4.0, 5.2, 6.5,  9.0, 12.0, 15.0, 18.5, 22.0, 25.5, 29.0],
    "Meta":               [1.0, 1.3, 1.7, 2.3, 3.0, 5.0,  7.0, 10.0, 13.0, 16.0, 19.0, 22.0, 25.0],
    "Oracle Cloud":       [0.4, 0.5, 0.7, 1.0, 1.4, 2.2,  3.5,  5.5,  7.5, 10.0, 12.5, 14.5, 16.5],
    "Apple":              [0.3, 0.4, 0.4, 0.5, 0.6, 0.8,  1.1,  1.5,  2.0,  2.5,  3.0,  3.5,  4.0],
    "CoreWeave":          [0.0, 0.0, 0.0, 0.0, 0.015,0.12,0.5,  1.3,  2.5,  4.0,  5.5,  7.0,  8.5],
    "Lambda":             [0.0, 0.0, 0.0, 0.0, 0.008,0.05,0.18, 0.45, 0.90, 1.50, 2.10, 2.70, 3.30],
    "Crusoe Energy":      [0.0, 0.0, 0.0, 0.0, 0.008,0.08,0.25, 0.60, 1.20, 1.80, 2.50, 3.20, 4.00],
    "Voltage Park":       [0.0, 0.0, 0.0, 0.0, 0.0, 0.03,0.12, 0.30, 0.65, 1.10, 1.60, 2.10, 2.60],
    "Together AI":        [0.0, 0.0, 0.0, 0.0, 0.0, 0.02,0.08, 0.18, 0.40, 0.75, 1.15, 1.60, 2.05],
    "Applied Digital":    [0.0, 0.0, 0.0, 0.0, 0.005,0.05,0.15, 0.40, 0.75, 1.20, 1.80, 2.40, 3.00],
}

# --- POWER GENERATION MIX (% Renewable of total power) ---
RENEWABLE_PCT = {
    "Amazon (AWS)":       [0.30, 0.36, 0.42, 0.50, 0.65, 0.80, 0.90, 0.95, 1.00, 1.00, 1.00, 1.00, 1.00],
    "Microsoft (Azure)":  [0.35, 0.40, 0.45, 0.50, 0.58, 0.68, 0.78, 0.85, 0.92, 0.97, 1.00, 1.00, 1.00],
    "Google (GCP)":       [0.50, 0.60, 0.70, 0.80, 0.90, 0.95, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00],
    "Meta":               [0.25, 0.30, 0.40, 0.50, 0.60, 0.72, 0.82, 0.90, 0.95, 1.00, 1.00, 1.00, 1.00],
    "Oracle Cloud":       [0.10, 0.12, 0.15, 0.20, 0.28, 0.38, 0.50, 0.60, 0.70, 0.80, 0.88, 0.93, 0.97],
    "Apple":              [0.80, 0.85, 0.90, 0.95, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00],
    "CoreWeave":          [0.00, 0.00, 0.00, 0.00, 0.10, 0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75, 0.85],
    "Lambda":             [0.00, 0.00, 0.00, 0.00, 0.10, 0.15, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80],
    "Crusoe Energy":      [0.00, 0.00, 0.00, 0.00, 0.05, 0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80],
    "Voltage Park":       [0.00, 0.00, 0.00, 0.00, 0.00, 0.10, 0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75],
    "Together AI":        [0.00, 0.00, 0.00, 0.00, 0.00, 0.10, 0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75],
    "Applied Digital":    [0.00, 0.00, 0.00, 0.00, 0.05, 0.10, 0.18, 0.28, 0.38, 0.48, 0.58, 0.68, 0.78],
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

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


def write_title_row(ws, row, title, num_cols):
    """Merge cells and write a title bar."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    cell = ws.cell(row=row, column=1, value=title)
    apply_cell_style(cell, font=TITLE_FONT, fill=TITLE_FILL,
                     alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[row].height = 32


def write_section_header(ws, row, title, num_cols):
    """Write a section sub-header."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    cell = ws.cell(row=row, column=1, value=title)
    apply_cell_style(cell, font=SECTION_FONT, fill=LIGHT_FILL,
                     alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[row].height = 22


def write_column_headers(ws, row, labels, start_col=1):
    """Write column header row."""
    for i, label in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + i, value=label)
        apply_cell_style(cell, font=HEADER_FONT, fill=HEADER_FILL,
                         alignment=Alignment(horizontal="center", vertical="center"),
                         border=THIN_BORDER)
    ws.row_dimensions[row].height = 20


def write_data_row(ws, row, company, values, start_col=1, fmt="0.0", is_total=False, alt=False):
    """Write a data row with company name and values."""
    # Company name
    cell = ws.cell(row=row, column=start_col, value=company)
    if is_total:
        apply_cell_style(cell, font=TOTAL_FONT, fill=TOTAL_FILL,
                         alignment=Alignment(horizontal="left", vertical="center"),
                         border=THIN_BORDER)
    else:
        apply_cell_style(cell, font=COMPANY_FONT,
                         fill=ALT_ROW_FILL if alt else None,
                         alignment=Alignment(horizontal="left", vertical="center"),
                         border=THIN_BORDER)
    # Values
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=start_col + 1 + i, value=v)
        year_idx = i
        is_forecast = (year_idx >= len(HIST_YEARS))
        if is_total:
            apply_cell_style(cell, font=TOTAL_FONT, fill=TOTAL_FILL,
                             alignment=Alignment(horizontal="center", vertical="center"),
                             border=THIN_BORDER, number_format=fmt)
        else:
            fill = FORECAST_FILL if is_forecast else (ALT_ROW_FILL if alt else None)
            apply_cell_style(cell, font=DATA_FONT,
                             fill=fill,
                             alignment=Alignment(horizontal="center", vertical="center"),
                             border=THIN_BORDER, number_format=fmt)


def write_yoy_row(ws, row, values_row_data, start_col=1, alt=False):
    """Write YoY growth % row (skip first year)."""
    cell = ws.cell(row=row, column=start_col, value="  YoY Growth %")
    apply_cell_style(cell, font=PCT_FONT, fill=ALT_ROW_FILL if alt else None,
                     alignment=Alignment(horizontal="left", vertical="center", indent=2),
                     border=THIN_BORDER)
    for i in range(len(values_row_data)):
        if i == 0 or values_row_data[i-1] == 0:
            cell = ws.cell(row=row, column=start_col + 1 + i, value="—")
            apply_cell_style(cell, font=PCT_FONT, fill=ALT_ROW_FILL if alt else None,
                             alignment=Alignment(horizontal="center"), border=THIN_BORDER)
        else:
            growth = (values_row_data[i] - values_row_data[i-1]) / values_row_data[i-1]
            cell = ws.cell(row=row, column=start_col + 1 + i, value=growth)
            is_forecast = (i >= len(HIST_YEARS))
            fill = FORECAST_FILL if is_forecast else (ALT_ROW_FILL if alt else None)
            apply_cell_style(cell, font=PCT_FONT, fill=fill,
                             alignment=Alignment(horizontal="center"), border=THIN_BORDER,
                             number_format="0.0%")


def compute_totals(data_dict, companies):
    """Sum across a subset of companies for each year."""
    totals = [0.0] * len(YEARS)
    for c in companies:
        for i, v in enumerate(data_dict[c]):
            totals[i] += v
    return [round(t, 2) for t in totals]


def add_line_chart(ws, title, cat_row, data_rows, labels, min_col, max_col, chart_row, chart_col, width=22, height=12):
    """Add a line chart to the worksheet."""
    chart = LineChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.y_axis.title = ""
    chart.x_axis.title = ""
    
    cats = Reference(ws, min_col=min_col, min_row=cat_row, max_col=max_col)
    chart.set_categories(cats)
    
    colors = ["4472C4", "ED7D31", "70AD47", "FFC000", "5B9BD5", "FF6384",
              "9966FF", "FF9F40", "36A2EB", "C9CBCF", "FF6384", "4BC0C0"]
    
    for idx, (drow, label) in enumerate(zip(data_rows, labels)):
        data = Reference(ws, min_col=min_col, min_row=drow, max_col=max_col)
        chart.add_data(data, from_rows=True, titles_from_data=False)
        chart.series[idx].name = label  # Set the series name directly
        chart.series[idx].graphicalProperties.line.width = 22000
        if idx < len(colors):
            from openpyxl.chart.series import SeriesLabel
            chart.series[idx].graphicalProperties.line.solidFill = colors[idx]
    
    chart.legend.position = 'b'
    ws.add_chart(chart, f"{get_column_letter(chart_col)}{chart_row}")
    return chart


def add_bar_chart(ws, title, cat_row, data_rows, labels, min_col, max_col, chart_row, chart_col, width=22, height=12, stacked=False):
    """Add a bar chart to the worksheet."""
    chart = BarChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    if stacked:
        chart.grouping = "stacked"
        chart.overlap = 100
    
    cats = Reference(ws, min_col=min_col, min_row=cat_row, max_col=max_col)
    chart.set_categories(cats)
    
    colors = ["4472C4", "ED7D31", "70AD47", "FFC000", "5B9BD5", "FF6384",
              "9966FF", "FF9F40", "36A2EB", "C9CBCF", "FF6384", "4BC0C0"]
    
    for idx, (drow, label) in enumerate(zip(data_rows, labels)):
        data = Reference(ws, min_col=min_col, min_row=drow, max_col=max_col)
        chart.add_data(data, from_rows=True, titles_from_data=False)
        chart.series[idx].name = label
        if idx < len(colors):
            chart.series[idx].graphicalProperties.solidFill = colors[idx]
    
    chart.legend.position = 'b'
    ws.add_chart(chart, f"{get_column_letter(chart_col)}{chart_row}")
    return chart


def build_data_sheet(wb, sheet_name, title, data_dict, unit_label, fmt="0.0"):
    """Build a standard data sheet with hyperscaler and neocloud sections."""
    ws = wb.create_sheet(title=sheet_name)
    num_cols = 1 + len(YEARS)  # company + years
    
    # Column widths
    ws.column_dimensions['A'].width = 22
    for i in range(len(YEARS)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12
    
    row = 1
    # Title
    write_title_row(ws, row, f"{title} ({unit_label})", num_cols)
    row += 1
    
    # Subtitle/note
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    note_cell = ws.cell(row=row, column=1, value="Historical: 2018-2024  |  Forecast: 2025E-2030E (yellow shading)  |  Sources: Public filings, industry estimates")
    apply_cell_style(note_cell, font=NOTE_FONT, alignment=Alignment(horizontal="left"))
    row += 1
    
    # === HYPERSCALERS ===
    write_section_header(ws, row, "Hyperscalers", num_cols)
    row += 1
    
    header_row = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    hyper_data_rows = []
    for idx, company in enumerate(HYPERSCALERS):
        write_data_row(ws, row, company, data_dict[company], fmt=fmt, alt=(idx % 2 == 1))
        hyper_data_rows.append(row)
        row += 1
    
    # Hyperscaler total
    hyper_totals = compute_totals(data_dict, HYPERSCALERS)
    write_data_row(ws, row, "Hyperscaler Total", hyper_totals, fmt=fmt, is_total=True)
    hyper_total_row = row
    row += 1
    
    row += 1  # spacer
    
    # === NEOCLOUDS ===
    write_section_header(ws, row, "Neocloud Vendors", num_cols)
    row += 1
    
    neo_header_row = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    neo_data_rows = []
    for idx, company in enumerate(NEOCLOUDS):
        write_data_row(ws, row, company, data_dict[company], fmt=fmt, alt=(idx % 2 == 1))
        neo_data_rows.append(row)
        row += 1
    
    # Neocloud total
    neo_totals = compute_totals(data_dict, NEOCLOUDS)
    write_data_row(ws, row, "Neocloud Total", neo_totals, fmt=fmt, is_total=True)
    neo_total_row = row
    row += 1
    
    row += 1  # spacer
    
    # === COMBINED TOTAL ===
    all_totals = compute_totals(data_dict, ALL_COMPANIES)
    write_data_row(ws, row, "GRAND TOTAL", all_totals, fmt=fmt, is_total=True)
    grand_total_row = row
    row += 2
    
    # === CHARTS ===
    chart_start_row = row
    
    # Hyperscaler line chart
    add_line_chart(
        ws, f"{title} - Hyperscalers",
        header_row, hyper_data_rows, HYPERSCALERS,
        min_col=2, max_col=1 + len(YEARS),
        chart_row=chart_start_row, chart_col=1,
        width=24, height=14
    )
    
    # Neocloud line chart
    add_line_chart(
        ws, f"{title} - Neoclouds",
        neo_header_row, neo_data_rows, NEOCLOUDS,
        min_col=2, max_col=1 + len(YEARS),
        chart_row=chart_start_row, chart_col=9,
        width=24, height=14
    )
    
    # Totals comparison (Hyperscaler vs Neocloud bar chart)
    totals_bar_row = chart_start_row + 16
    # Write helper rows for totals chart
    ws.cell(row=totals_bar_row, column=1, value="Segment")
    for i, yl in enumerate(YEAR_LABELS):
        ws.cell(row=totals_bar_row, column=2+i, value=yl)
    ws.cell(row=totals_bar_row+1, column=1, value="Hyperscalers")
    for i, v in enumerate(hyper_totals):
        ws.cell(row=totals_bar_row+1, column=2+i, value=v)
    ws.cell(row=totals_bar_row+2, column=1, value="Neoclouds")
    for i, v in enumerate(neo_totals):
        ws.cell(row=totals_bar_row+2, column=2+i, value=v)
    
    add_bar_chart(
        ws, f"Total {title}: Hyperscalers vs Neoclouds",
        totals_bar_row, [totals_bar_row+1, totals_bar_row+2],
        ["Hyperscalers", "Neoclouds"],
        min_col=2, max_col=1+len(YEARS),
        chart_row=totals_bar_row+4, chart_col=1,
        width=28, height=14, stacked=True
    )
    
    # Freeze panes
    ws.freeze_panes = "B1"
    
    return ws, hyper_data_rows, neo_data_rows, hyper_total_row, neo_total_row, grand_total_row


def build_power_sheet(wb):
    """Build the Power Generation & Availability sheet."""
    ws = wb.create_sheet(title="Power & Generation")
    num_cols = 1 + len(YEARS)
    
    ws.column_dimensions['A'].width = 22
    for i in range(len(YEARS)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12
    
    row = 1
    write_title_row(ws, row, "Power Capacity, Availability & Generation Mix", num_cols)
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    note = ws.cell(row=row, column=1, value="IT Load (GW) | Power Contracted (GW) | Utilization (%) | Renewable Mix (%)  |  Yellow = Forecast")
    apply_cell_style(note, font=NOTE_FONT)
    row += 2
    
    # --- Section 1: IT Load (GW) ---
    write_section_header(ws, row, "IT Load Capacity (GW)", num_cols)
    row += 1
    header_row_gw = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    gw_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        write_data_row(ws, row, company, GW_CAPACITY_DATA[company], fmt="0.00", alt=(idx % 2 == 1))
        gw_rows.append(row)
        row += 1
    all_gw_totals = compute_totals(GW_CAPACITY_DATA, ALL_COMPANIES)
    write_data_row(ws, row, "TOTAL IT LOAD", all_gw_totals, fmt="0.0", is_total=True)
    gw_total_row = row
    row += 2
    
    # --- Section 2: Power Contracted (GW) ---
    write_section_header(ws, row, "Total Power Contracted / Available (GW)", num_cols)
    row += 1
    header_row_pwr = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    pwr_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        write_data_row(ws, row, company, POWER_AVAILABLE_DATA[company], fmt="0.00", alt=(idx % 2 == 1))
        pwr_rows.append(row)
        row += 1
    all_pwr_totals = compute_totals(POWER_AVAILABLE_DATA, ALL_COMPANIES)
    write_data_row(ws, row, "TOTAL CONTRACTED", all_pwr_totals, fmt="0.0", is_total=True)
    pwr_total_row = row
    row += 2
    
    # --- Section 3: Power Utilization (IT Load / Contracted) ---
    write_section_header(ws, row, "Power Utilization Rate (IT Load / Contracted)", num_cols)
    row += 1
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    util_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        utils = []
        for y in range(len(YEARS)):
            if POWER_AVAILABLE_DATA[company][y] > 0:
                utils.append(GW_CAPACITY_DATA[company][y] / POWER_AVAILABLE_DATA[company][y])
            else:
                utils.append(0)
        write_data_row(ws, row, company, utils, fmt="0.0%", alt=(idx % 2 == 1))
        util_rows.append(row)
        row += 1
    row += 1
    
    # --- Section 4: Renewable Energy Mix ---
    write_section_header(ws, row, "Renewable Energy Mix (%)", num_cols)
    row += 1
    renew_header = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    renew_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        write_data_row(ws, row, company, RENEWABLE_PCT[company], fmt="0%", alt=(idx % 2 == 1))
        renew_rows.append(row)
        row += 1
    row += 2
    
    # Charts
    chart_row = row
    
    # IT Load chart - Hyperscalers
    add_line_chart(
        ws, "IT Load (GW) - Hyperscalers",
        header_row_gw, gw_rows[:6], HYPERSCALERS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=chart_row, chart_col=1,
        width=24, height=14
    )
    
    # IT Load chart - Neoclouds
    add_line_chart(
        ws, "IT Load (GW) - Neoclouds",
        header_row_gw, gw_rows[6:], NEOCLOUDS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=chart_row, chart_col=9,
        width=24, height=14
    )
    
    # Renewable mix chart
    renew_chart_row = chart_row + 16
    add_line_chart(
        ws, "Renewable Energy Mix (%) - All Companies",
        renew_header, renew_rows[:6], HYPERSCALERS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=renew_chart_row, chart_col=1,
        width=24, height=14
    )
    
    add_line_chart(
        ws, "Renewable Energy Mix (%) - Neoclouds",
        renew_header, renew_rows[6:], NEOCLOUDS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=renew_chart_row, chart_col=9,
        width=24, height=14
    )
    
    ws.freeze_panes = "B1"
    return ws


def build_revenue_per_mw_sheet(wb):
    """Build Revenue/MW analysis sheet."""
    ws = wb.create_sheet(title="Revenue per MW")
    num_cols = 1 + len(YEARS)
    
    ws.column_dimensions['A'].width = 22
    for i in range(len(YEARS)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12
    
    row = 1
    write_title_row(ws, row, "Revenue per MW Analysis ($M / MW)", num_cols)
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    note = ws.cell(row=row, column=1, value="Revenue ($B) / IT Load (GW) = Revenue per MW ($M/MW)  |  Higher = more revenue-efficient  |  Yellow = Forecast")
    apply_cell_style(note, font=NOTE_FONT)
    row += 2
    
    # --- Revenue ($B) ---
    write_section_header(ws, row, "Revenue ($B)", num_cols)
    row += 1
    rev_header = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    rev_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        write_data_row(ws, row, company, REVENUE_DATA[company], fmt="0.0", alt=(idx % 2 == 1))
        rev_rows.append(row)
        row += 1
    rev_totals = compute_totals(REVENUE_DATA, ALL_COMPANIES)
    write_data_row(ws, row, "TOTAL", rev_totals, fmt="0.0", is_total=True)
    row += 2
    
    # --- IT Load (GW) ---
    write_section_header(ws, row, "IT Load Capacity (GW)", num_cols)
    row += 1
    gw_header = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    gw_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        write_data_row(ws, row, company, GW_CAPACITY_DATA[company], fmt="0.00", alt=(idx % 2 == 1))
        gw_rows.append(row)
        row += 1
    row += 1
    
    # --- Revenue / MW ($M per MW) ---
    # Revenue ($B) / GW = $B/GW. Since 1 GW = 1000 MW, $B/GW = $M/MW
    write_section_header(ws, row, "Revenue per MW ($M / MW)", num_cols)
    row += 1
    rpm_header = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    rpm_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        rpm_values = []
        for y in range(len(YEARS)):
            gw = GW_CAPACITY_DATA[company][y]
            rev = REVENUE_DATA[company][y]
            if gw > 0:
                # Revenue ($B) / GW = $B/GW = $M/MW  (since 1 GW = 1000 MW, and $B = 1000 $M)
                rpm_values.append(round(rev / gw, 2))
            else:
                rpm_values.append(0)
        write_data_row(ws, row, company, rpm_values, fmt="#,##0.0", alt=(idx % 2 == 1))
        rpm_rows.append(row)
        row += 1
    row += 1
    
    # --- Capex / MW ($M per MW) ---
    write_section_header(ws, row, "Capex per MW ($M / MW)", num_cols)
    row += 1
    cpm_header = row
    write_column_headers(ws, row, ["Company"] + YEAR_LABELS)
    row += 1
    
    cpm_rows = []
    for idx, company in enumerate(ALL_COMPANIES):
        cpm_values = []
        for y in range(len(YEARS)):
            gw = GW_CAPACITY_DATA[company][y]
            capex = CAPEX_DATA[company][y]
            if gw > 0:
                cpm_values.append(round(capex / gw, 2))
            else:
                cpm_values.append(0)
        write_data_row(ws, row, company, cpm_values, fmt="#,##0.0", alt=(idx % 2 == 1))
        cpm_rows.append(row)
        row += 1
    row += 2
    
    # Charts
    chart_row = row
    
    # Revenue/MW - Hyperscalers
    add_line_chart(
        ws, "Revenue / MW ($M/MW) - Hyperscalers",
        rpm_header, rpm_rows[:6], HYPERSCALERS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=chart_row, chart_col=1,
        width=24, height=14
    )
    
    # Revenue/MW - Neoclouds
    add_line_chart(
        ws, "Revenue / MW ($M/MW) - Neoclouds",
        rpm_header, rpm_rows[6:], NEOCLOUDS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=chart_row, chart_col=9,
        width=24, height=14
    )
    
    # Capex/MW charts
    cap_chart_row = chart_row + 16
    add_line_chart(
        ws, "Capex / MW ($M/MW) - Hyperscalers",
        cpm_header, cpm_rows[:6], HYPERSCALERS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=cap_chart_row, chart_col=1,
        width=24, height=14
    )
    
    add_line_chart(
        ws, "Capex / MW ($M/MW) - Neoclouds",
        cpm_header, cpm_rows[6:], NEOCLOUDS,
        min_col=2, max_col=1+len(YEARS),
        chart_row=cap_chart_row, chart_col=9,
        width=24, height=14
    )
    
    ws.freeze_panes = "B1"
    return ws


def build_dashboard(wb):
    """Build the summary dashboard sheet as the first sheet."""
    ws = wb.create_sheet(title="Dashboard", index=0)
    num_cols = 16
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 24
    for c in range(3, num_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13
    
    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    cell = ws.cell(row=row, column=1, value="DATACENTER INFRASTRUCTURE MODEL — EXECUTIVE DASHBOARD")
    apply_cell_style(cell, font=Font(name="Calibri", size=18, bold=True, color=WHITE),
                     fill=TITLE_FILL, alignment=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[row].height = 40
    row += 1
    
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    cell = ws.cell(row=row, column=1, value="Hyperscalers & Neocloud Vendors  |  2018-2030E  |  Capex, Servers, Datacenters, Power, Revenue/MW")
    apply_cell_style(cell, font=Font(name="Calibri", size=10, color="999999"),
                     alignment=Alignment(horizontal="center"))
    row += 2
    
    # === KEY METRICS SUMMARY TABLE ===
    write_section_header(ws, row, "Key Metrics Summary (All Companies)", num_cols)
    row += 1
    
    metrics = [
        ("Total Capex ($B)", CAPEX_DATA, "0.0"),
        ("Total Servers (000s)", SERVER_DATA, "#,##0"),
        ("Total Datacenters", DC_COUNT_DATA, "#,##0"),
        ("Total IT Load (GW)", GW_CAPACITY_DATA, "0.0"),
        ("Total Power Contracted (GW)", POWER_AVAILABLE_DATA, "0.0"),
        ("Total Revenue ($B)", REVENUE_DATA, "0.0"),
    ]
    
    summary_header = row
    write_column_headers(ws, row, ["", "Metric"] + YEAR_LABELS, start_col=1)
    row += 1
    
    summary_data_rows = []
    for idx, (label, data, fmt) in enumerate(metrics):
        totals = compute_totals(data, ALL_COMPANIES)
        cell_a = ws.cell(row=row, column=1, value="")
        apply_cell_style(cell_a, border=THIN_BORDER)
        write_data_row(ws, row, label, totals, start_col=2, fmt=fmt, alt=(idx % 2 == 1))
        summary_data_rows.append(row)
        row += 1
    
    # Avg Revenue/MW
    rev_totals = compute_totals(REVENUE_DATA, ALL_COMPANIES)
    gw_totals = compute_totals(GW_CAPACITY_DATA, ALL_COMPANIES)
    avg_rpm = [round(r/g, 1) if g > 0 else 0 for r, g in zip(rev_totals, gw_totals)]
    cell_a = ws.cell(row=row, column=1, value="")
    apply_cell_style(cell_a, border=THIN_BORDER)
    write_data_row(ws, row, "Avg Revenue/MW ($M/MW)", avg_rpm, start_col=2, fmt="0.0", alt=True)
    summary_data_rows.append(row)
    row += 1
    
    # Avg Utilization
    pwr_totals = compute_totals(POWER_AVAILABLE_DATA, ALL_COMPANIES)
    avg_util = [round(g/p, 3) if p > 0 else 0 for g, p in zip(gw_totals, pwr_totals)]
    cell_a = ws.cell(row=row, column=1, value="")
    apply_cell_style(cell_a, border=THIN_BORDER)
    write_data_row(ws, row, "Avg Power Utilization", avg_util, start_col=2, fmt="0.0%", alt=False)
    summary_data_rows.append(row)
    row += 2
    
    # === HYPERSCALER vs NEOCLOUD SPLIT ===
    write_section_header(ws, row, "Hyperscaler vs Neocloud — Capex Split ($B)", num_cols)
    row += 1
    split_header = row
    write_column_headers(ws, row, ["", "Segment"] + YEAR_LABELS, start_col=1)
    row += 1
    
    hyper_capex = compute_totals(CAPEX_DATA, HYPERSCALERS)
    neo_capex = compute_totals(CAPEX_DATA, NEOCLOUDS)
    
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Hyperscalers", hyper_capex, start_col=2, fmt="0.0")
    hyper_capex_row = row
    row += 1
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Neoclouds", neo_capex, start_col=2, fmt="0.0", alt=True)
    neo_capex_row = row
    row += 1
    
    total_capex = compute_totals(CAPEX_DATA, ALL_COMPANIES)
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Total", total_capex, start_col=2, fmt="0.0", is_total=True)
    row += 1
    
    # Neocloud % share
    cell_a = ws.cell(row=row, column=1, value="")
    neo_share = [round(n/t, 3) if t > 0 else 0 for n, t in zip(neo_capex, total_capex)]
    write_data_row(ws, row, "Neocloud Share %", neo_share, start_col=2, fmt="0.0%", alt=True)
    row += 2
    
    # === GW CAPACITY SPLIT ===
    write_section_header(ws, row, "Hyperscaler vs Neocloud — IT Load (GW)", num_cols)
    row += 1
    gw_split_header = row
    write_column_headers(ws, row, ["", "Segment"] + YEAR_LABELS, start_col=1)
    row += 1
    
    hyper_gw = compute_totals(GW_CAPACITY_DATA, HYPERSCALERS)
    neo_gw = compute_totals(GW_CAPACITY_DATA, NEOCLOUDS)
    
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Hyperscalers", hyper_gw, start_col=2, fmt="0.0")
    hyper_gw_row = row
    row += 1
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Neoclouds", neo_gw, start_col=2, fmt="0.0", alt=True)
    neo_gw_row = row
    row += 1
    
    total_gw = [round(h + n, 2) for h, n in zip(hyper_gw, neo_gw)]
    cell_a = ws.cell(row=row, column=1, value="")
    write_data_row(ws, row, "Total", total_gw, start_col=2, fmt="0.0", is_total=True)
    row += 2
    
    # === CHARTS ON DASHBOARD ===
    chart_row = row
    
    # Chart 1: Total Capex stacked bar
    # Write helper data for chart
    helper_row = 120  # far below visible area
    ws.cell(row=helper_row, column=2, value="Year")
    for i, yl in enumerate(YEAR_LABELS):
        ws.cell(row=helper_row, column=3+i, value=yl)
    ws.cell(row=helper_row+1, column=2, value="Hyperscalers")
    for i, v in enumerate(hyper_capex):
        ws.cell(row=helper_row+1, column=3+i, value=v)
    ws.cell(row=helper_row+2, column=2, value="Neoclouds")
    for i, v in enumerate(neo_capex):
        ws.cell(row=helper_row+2, column=3+i, value=v)
    
    add_bar_chart(
        ws, "Total Capex ($B) — Hyperscalers vs Neoclouds",
        helper_row, [helper_row+1, helper_row+2],
        ["Hyperscalers", "Neoclouds"],
        min_col=3, max_col=2+len(YEARS),
        chart_row=chart_row, chart_col=2,
        width=26, height=14, stacked=True
    )
    
    # Chart 2: IT Load
    ws.cell(row=helper_row+4, column=2, value="Year")
    for i, yl in enumerate(YEAR_LABELS):
        ws.cell(row=helper_row+4, column=3+i, value=yl)
    ws.cell(row=helper_row+5, column=2, value="Hyperscalers")
    for i, v in enumerate(hyper_gw):
        ws.cell(row=helper_row+5, column=3+i, value=v)
    ws.cell(row=helper_row+6, column=2, value="Neoclouds")
    for i, v in enumerate(neo_gw):
        ws.cell(row=helper_row+6, column=3+i, value=v)
    
    add_bar_chart(
        ws, "Total IT Load (GW) — Hyperscalers vs Neoclouds",
        helper_row+4, [helper_row+5, helper_row+6],
        ["Hyperscalers", "Neoclouds"],
        min_col=3, max_col=2+len(YEARS),
        chart_row=chart_row, chart_col=10,
        width=26, height=14, stacked=True
    )
    
    # Chart 3: Revenue/MW trend
    chart_row2 = chart_row + 16
    ws.cell(row=helper_row+8, column=2, value="Year")
    for i, yl in enumerate(YEAR_LABELS):
        ws.cell(row=helper_row+8, column=3+i, value=yl)
    ws.cell(row=helper_row+9, column=2, value="Avg Rev/MW")
    for i, v in enumerate(avg_rpm):
        ws.cell(row=helper_row+9, column=3+i, value=v)
    
    add_line_chart(
        ws, "Average Revenue per MW ($M/MW)",
        helper_row+8, [helper_row+9],
        ["Avg Rev/MW"],
        min_col=3, max_col=2+len(YEARS),
        chart_row=chart_row2, chart_col=2,
        width=26, height=14
    )
    
    # Chart 4: Total Servers
    hyper_servers = compute_totals(SERVER_DATA, HYPERSCALERS)
    neo_servers = compute_totals(SERVER_DATA, NEOCLOUDS)
    ws.cell(row=helper_row+11, column=2, value="Year")
    for i, yl in enumerate(YEAR_LABELS):
        ws.cell(row=helper_row+11, column=3+i, value=yl)
    ws.cell(row=helper_row+12, column=2, value="Hyperscalers")
    for i, v in enumerate(hyper_servers):
        ws.cell(row=helper_row+12, column=3+i, value=v)
    ws.cell(row=helper_row+13, column=2, value="Neoclouds")
    for i, v in enumerate(neo_servers):
        ws.cell(row=helper_row+13, column=3+i, value=v)
    
    add_bar_chart(
        ws, "Total Servers (000s) — Hyperscalers vs Neoclouds",
        helper_row+11, [helper_row+12, helper_row+13],
        ["Hyperscalers", "Neoclouds"],
        min_col=3, max_col=2+len(YEARS),
        chart_row=chart_row2, chart_col=10,
        width=26, height=14, stacked=True
    )
    
    ws.freeze_panes = "C5"
    ws.sheet_properties.tabColor = ACCENT_BLUE
    return ws


def build_assumptions_sheet(wb):
    """Build an Assumptions & Sources sheet."""
    ws = wb.create_sheet(title="Assumptions & Sources")
    
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 80
    
    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1, value="MODEL ASSUMPTIONS & METHODOLOGY")
    apply_cell_style(cell, font=TITLE_FONT, fill=TITLE_FILL,
                     alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[row].height = 32
    row += 2
    
    assumptions = [
        ("Data Sources", [
            "Historical capex and revenue from public SEC filings (10-K, 10-Q) and earnings calls",
            "Server counts estimated from industry reports (Synergy Research, Dell'Oro, Omdia)",
            "Datacenter counts from company announcements, press releases, and analyst estimates",
            "Power capacity from sustainability reports, PPA announcements, and utility filings",
            "Neocloud data from funding announcements, press coverage, and industry estimates",
        ]),
        ("Key Assumptions — Forecasts (2025E-2030E)", [
            "AI infrastructure buildout continues to accelerate through 2027, moderating thereafter",
            "Hyperscaler capex growth driven by AI/ML workloads, sovereign cloud, and edge expansion",
            "Neocloud vendors benefit from GPU-as-a-service demand but face capital constraints",
            "Power availability is the key constraint on datacenter buildout after 2026",
            "Renewable energy targets: most hyperscalers reach 100% by 2027-2030",
            "Server density improves 8-12% annually (more compute per rack unit)",
            "Average PUE improves from ~1.3 (2024) to ~1.15 (2030) industry-wide",
        ]),
        ("Definitions", [
            "IT Load (GW): Total electrical power consumed by IT equipment (servers, storage, networking)",
            "Power Contracted (GW): Total power capacity secured via PPAs, utility contracts, and on-site generation",
            "Power Utilization: IT Load / Power Contracted (higher = more efficient use of secured power)",
            "Revenue/MW: Annual revenue ($B) / IT Load (GW) — since 1 GW = 1000 MW and $1B = 1000 $M, result is in $M/MW",
            "Capex/MW: Annual capex ($B) / IT Load (GW) — investment intensity per MW of IT capacity",
            "Servers (000s): Total server count in thousands across all datacenter locations",
        ]),
        ("Caveats", [
            "All neocloud figures are estimates based on limited public disclosure",
            "Company-specific DC counts may include leased/colocation facilities",
            "Apple revenue is total company revenue (not cloud-specific); Apple DC figures are for internal use",
            "Meta revenue is total company revenue; DC infrastructure supports ads + AI workloads",
            "Forecasts represent base-case scenario; upside/downside cases not included",
            "Currency: All figures in USD",
        ]),
    ]
    
    for section_title, items in assumptions:
        cell = ws.cell(row=row, column=1, value="")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        cell = ws.cell(row=row, column=1, value=section_title)
        apply_cell_style(cell, font=SECTION_FONT, fill=LIGHT_FILL,
                         alignment=Alignment(horizontal="left"))
        ws.row_dimensions[row].height = 22
        row += 1
        
        for item in items:
            ws.cell(row=row, column=1, value="•")
            cell = ws.cell(row=row, column=2, value=item)
            apply_cell_style(cell, font=DATA_FONT)
            row += 1
        row += 1
    
    ws.sheet_properties.tabColor = "999999"
    return ws


# ============================================================
# MAIN BUILD
# ============================================================

def main():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    print("Building Dashboard...")
    build_dashboard(wb)
    
    print("Building Capex sheet...")
    build_data_sheet(wb, "Capex Spend", "Capital Expenditure", CAPEX_DATA, "$B", fmt="#,##0.0")
    
    print("Building Revenue sheet...")
    build_data_sheet(wb, "Revenue", "Revenue", REVENUE_DATA, "$B", fmt="#,##0.0")
    
    print("Building Servers sheet...")
    build_data_sheet(wb, "Servers", "Server Count", SERVER_DATA, "000s", fmt="#,##0")
    
    print("Building Datacenters sheet...")
    build_data_sheet(wb, "Datacenters", "Datacenter Count", DC_COUNT_DATA, "Facilities", fmt="#,##0")
    
    print("Building GW Capacity sheet...")
    build_data_sheet(wb, "GW Capacity", "IT Load Capacity", GW_CAPACITY_DATA, "GW", fmt="0.00")
    
    print("Building Power & Generation sheet...")
    build_power_sheet(wb)
    
    print("Building Revenue per MW sheet...")
    build_revenue_per_mw_sheet(wb)
    
    print("Building Assumptions sheet...")
    build_assumptions_sheet(wb)
    
    output_path = "/workspace/Datacenter_Infrastructure_Model.xlsx"
    wb.save(output_path)
    print(f"\nModel saved to: {output_path}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    main()
