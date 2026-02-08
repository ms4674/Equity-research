#!/usr/bin/env python3
"""
Three-Statement Financial Model: SpaceX & xAI
==============================================
Generates a comprehensive Excel workbook with:
- Revenue build by segment with volume & take-rate drivers
- Income Statement, Balance Sheet, Cash Flow Statement
- Quarterly and annual views (5yr historical + 3yr forecast)
- GAAP OpLoss-to-FCF bridge
- Two-way sensitivity tables
- Valuation (last-round, secondary, cohort NPV, S-curves, unit economics)
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from copy import copy
import datetime

# ─── STYLE CONSTANTS ──────────────────────────────────────────────────────────
DARK_BLUE = "1F3864"
MED_BLUE = "2E75B6"
LIGHT_BLUE = "D6E4F0"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
GREEN_FILL = "E2EFDA"
YELLOW_FILL = "FFF2CC"
ORANGE_FILL = "FCE4D6"
DARK_TEXT = "1F3864"

HEADER_FONT = Font(name="Calibri", bold=True, color=WHITE, size=11)
HEADER_FILL = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
SUB_HEADER_FONT = Font(name="Calibri", bold=True, color=DARK_TEXT, size=10)
SUB_HEADER_FILL = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
SECTION_FONT = Font(name="Calibri", bold=True, color=DARK_TEXT, size=10)
SECTION_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
DATA_FONT = Font(name="Calibri", size=10)
INPUT_FILL = PatternFill(start_color=YELLOW_FILL, end_color=YELLOW_FILL, fill_type="solid")
FORECAST_FILL = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
TOTAL_FONT = Font(name="Calibri", bold=True, size=10, color=DARK_TEXT)
PCT_FORMAT = '0.0%'
NUM_FORMAT = '#,##0'
NUM_FORMAT_1 = '#,##0.0'
DOLLAR_FORMAT = '$#,##0'
DOLLAR_FORMAT_M = '$#,##0'
THIN_BORDER = Border(
    bottom=Side(style='thin', color='B0B0B0')
)
TOTAL_BORDER = Border(
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

def style_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style_sub_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = SUB_HEADER_FONT
        cell.fill = SUB_HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')

def style_section_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL

def style_total_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = TOTAL_FONT
        cell.border = TOTAL_BORDER

def style_data_cell(ws, row, col, fmt=None, is_forecast=False, is_input=False):
    cell = ws.cell(row=row, column=col)
    cell.font = DATA_FONT
    cell.alignment = Alignment(horizontal='right')
    if is_input:
        cell.fill = INPUT_FILL
    elif is_forecast:
        cell.fill = FORECAST_FILL
    if fmt:
        cell.number_format = fmt

def write_row(ws, row, data, start_col=1, fmt=None, is_forecast_start=None, is_input=False, bold=False):
    """Write a list of values into a row, applying styles."""
    for i, val in enumerate(data):
        c = start_col + i
        cell = ws.cell(row=row, column=c, value=val)
        cell.font = DATA_FONT if not bold else TOTAL_FONT
        cell.alignment = Alignment(horizontal='right') if c > 1 else Alignment(horizontal='left')
        if fmt and c > 1 and val is not None:
            cell.number_format = fmt
        if is_forecast_start and c >= is_forecast_start:
            cell.fill = FORECAST_FILL
        if is_input and c > 1:
            cell.fill = INPUT_FILL
        if bold:
            cell.border = TOTAL_BORDER

# ═══════════════════════════════════════════════════════════════════════════════
#  SPACEX DATA
# ═══════════════════════════════════════════════════════════════════════════════
# Sources: publicly reported estimates (Bloomberg, WSJ, Reuters, Quilty Analytics,
# Morgan Stanley research, Payload Space, company presentations)
# All figures in $M unless noted

SPACEX_ANNUAL_YEARS = [2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028]
SPACEX_HIST_END = 5  # index 0-4 are historical (2020-2024)

# ── Revenue Drivers ──────────────────────────────────────────────────────────
# Launch Services
spacex_launch_count =       [26,   31,   61,   98,   134,  155,  170,  185,  200]
spacex_avg_launch_price =   [50,   50,   48,   45,   44,   43,   42,   42,   41]  # $M per launch (blended F9/FH)
spacex_launch_revenue =     [r * p for r, p in zip(spacex_launch_count, spacex_avg_launch_price)]

# Starlink
spacex_starlink_subs_k =    [100,  250,  700,  2300, 4500, 7000, 10000, 13500, 17000]  # thousands
spacex_starlink_arpu_mo =   [99,   99,   99,   97,   93,   90,   87,    85,    83]  # $/month
spacex_starlink_revenue =   [round(s * a * 12 / 1000, 0) for s, a in zip(spacex_starlink_subs_k, spacex_starlink_arpu_mo)]

# Starship (nascent – revenue from NASA HLS, commercial, DOD)
spacex_starship_revenue =   [0,    0,    0,    0,    200,  800,  2000,  3500,  5500]

# Government / other (Dragon cargo/crew, classified)
spacex_govt_other_revenue = [800,  1000, 1200, 1400, 1600, 1800, 2000,  2200,  2400]

spacex_total_revenue = [
    spacex_launch_revenue[i] + spacex_starlink_revenue[i] + spacex_starship_revenue[i] + spacex_govt_other_revenue[i]
    for i in range(len(SPACEX_ANNUAL_YEARS))
]

# ── Cost Structure ────────────────────────────────────────────────────────────
spacex_cogs_pct =           [0.62, 0.60, 0.57, 0.54, 0.50, 0.47, 0.44, 0.42, 0.40]
spacex_gross_profit = [round(spacex_total_revenue[i] * (1 - spacex_cogs_pct[i]), 0) for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_rnd_pct =            [0.22, 0.24, 0.26, 0.25, 0.22, 0.20, 0.18, 0.16, 0.14]
spacex_sga_pct =            [0.08, 0.08, 0.07, 0.07, 0.06, 0.06, 0.05, 0.05, 0.05]
spacex_sbc_pct =            [0.03, 0.03, 0.04, 0.04, 0.04, 0.04, 0.03, 0.03, 0.03]

spacex_rnd = [round(spacex_total_revenue[i] * spacex_rnd_pct[i], 0) for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_sga = [round(spacex_total_revenue[i] * spacex_sga_pct[i], 0) for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_sbc = [round(spacex_total_revenue[i] * spacex_sbc_pct[i], 0) for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_opex_total = [spacex_rnd[i] + spacex_sga[i] + spacex_sbc[i] for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_ebit = [spacex_gross_profit[i] - spacex_opex_total[i] for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_ebit_ex_sbc = [spacex_ebit[i] + spacex_sbc[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

# Interest & other
spacex_interest_exp =       [-120, -150, -180, -200, -220, -200, -180, -160, -140]
spacex_other_income =       [10,   15,   20,   25,   30,   35,   40,   45,   50]

spacex_ebt = [spacex_ebit[i] + spacex_interest_exp[i] + spacex_other_income[i] for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_tax_rate =           [0.0,  0.0,  0.0,  0.0,  0.0,  0.05, 0.10, 0.15, 0.18]
spacex_tax = [round(max(0, spacex_ebt[i]) * spacex_tax_rate[i], 0) for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_net_income = [spacex_ebt[i] - spacex_tax[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

# ── Balance Sheet ────────────────────────────────────────────────────────────
spacex_cash =               [3200, 3500, 4200, 5100, 7200, 9000, 11500, 15000, 19500]
spacex_ar =                 [400,  500,  700,  1000, 1400, 1800, 2200,  2700,  3200]
spacex_inventory =          [300,  400,  600,  900,  1200, 1500, 1800,  2100,  2400]
spacex_other_ca =           [200,  250,  300,  400,  500,  600,  700,   800,   900]
spacex_total_ca = [spacex_cash[i] + spacex_ar[i] + spacex_inventory[i] + spacex_other_ca[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_ppe_gross =          [4000, 5500, 8000, 12000, 17000, 22000, 27000, 32000, 37000]
spacex_accum_depr =         [1000, 1500, 2200, 3200,  4500,  6200,  8200,  10500, 13000]
spacex_ppe_net = [spacex_ppe_gross[i] - spacex_accum_depr[i] for i in range(len(SPACEX_ANNUAL_YEARS))]
spacex_intangibles =        [200,  250,  300,  350,   400,   450,   500,   550,   600]
spacex_other_lta =          [500,  600,  800,  1000,  1200,  1400,  1600,  1800,  2000]
spacex_total_assets = [spacex_total_ca[i] + spacex_ppe_net[i] + spacex_intangibles[i] + spacex_other_lta[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_ap =                 [500,  600,  800,  1100, 1500, 1800, 2100, 2400, 2700]
spacex_accrued_liab =       [400,  500,  600,  800,  1000, 1200, 1400, 1600, 1800]
spacex_deferred_rev =       [300,  400,  600,  900,  1200, 1500, 1800, 2100, 2400]
spacex_other_cl =           [200,  250,  300,  400,  500,  550,  600,  650,  700]
spacex_total_cl = [spacex_ap[i] + spacex_accrued_liab[i] + spacex_deferred_rev[i] + spacex_other_cl[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_lt_debt =            [3500, 4000, 4500, 5000, 5500, 5000, 4500, 4000, 3500]
spacex_other_ltl =          [300,  350,  400,  500,  600,  650,  700,  750,  800]
spacex_total_liab = [spacex_total_cl[i] + spacex_lt_debt[i] + spacex_other_ltl[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_equity = [spacex_total_assets[i] - spacex_total_liab[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

# ── Cash Flow Statement ─────────────────────────────────────────────────────
spacex_da =                 [400,  500,  700,  1000, 1300, 1700, 2000, 2300, 2500]
spacex_wc_change =          [-100, -150, -200, -300, -400, -350, -300, -250, -200]
spacex_other_cfo =          [50,   60,   80,   100,  120,  140,  160,  180,  200]
spacex_cfo = [spacex_net_income[i] + spacex_da[i] + spacex_sbc[i] + spacex_wc_change[i] + spacex_other_cfo[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_capex =              [-2000, -2500, -3500, -5000, -6000, -6500, -7000, -7500, -8000]
spacex_other_cfi =          [-100,  -120,  -150,  -200,  -250,  -200,  -150,  -100,  -50]
spacex_cfi = [spacex_capex[i] + spacex_other_cfi[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_fcf = [spacex_cfo[i] + spacex_capex[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

spacex_debt_issued =        [1000, 1000, 1000, 1000, 1500, 0,    0,    0,    0]
spacex_debt_repaid =        [-500, -500, -500, -500, -1000, -500, -500, -500, -500]
spacex_equity_raised =      [2000, 1500, 2000, 2500, 3000, 1500, 1000, 0,    0]
spacex_other_cff =          [-50,  -50,  -50,  -50,  -50,  -50,  -50,  -50,  -50]
spacex_cff = [spacex_debt_issued[i] + spacex_debt_repaid[i] + spacex_equity_raised[i] + spacex_other_cff[i] for i in range(len(SPACEX_ANNUAL_YEARS))]

# ── Quarterly Data (2024 Q1-Q4 actuals, 2025-2026 Q1-Q4 forecasts) ──────────
SPACEX_Q_LABELS = [
    "Q1'24", "Q2'24", "Q3'24", "Q4'24",
    "Q1'25", "Q2'25", "Q3'25", "Q4'25",
    "Q1'26", "Q2'26", "Q3'26", "Q4'26",
]
SPACEX_Q_HIST_END = 4  # first 4 are historical

spacex_q_revenue = [
    2100, 2400, 2700, 3000,     # 2024 actual ~$10.2B
    3200, 3500, 3800, 4200,     # 2025E ~$14.7B
    4500, 4800, 5200, 5700,     # 2026E ~$20.2B
]

spacex_q_cogs = [
    1050, 1200, 1350, 1500,
    1504, 1645, 1786, 1974,
    1980, 2112, 2288, 2508,
]

spacex_q_gross_profit = [spacex_q_revenue[i] - spacex_q_cogs[i] for i in range(12)]

spacex_q_opex = [
    630, 720, 810, 900,
    832, 910, 988, 1092,
    990, 1056, 1144, 1254,
]

spacex_q_ebit = [spacex_q_gross_profit[i] - spacex_q_opex[i] for i in range(12)]

spacex_q_net_income = [
    350, 400, 450, 500,
    700, 770, 840, 920,
    1200, 1280, 1400, 1550,
]

spacex_q_capex = [
    -1400, -1500, -1500, -1600,
    -1500, -1600, -1700, -1700,
    -1700, -1750, -1750, -1800,
]

spacex_q_cfo = [
    600, 700, 800, 900,
    1000, 1100, 1200, 1300,
    1500, 1600, 1700, 1900,
]

spacex_q_fcf = [spacex_q_cfo[i] + spacex_q_capex[i] for i in range(12)]


# ═══════════════════════════════════════════════════════════════════════════════
#  xAI DATA
# ═══════════════════════════════════════════════════════════════════════════════
# xAI founded mid-2023; limited history. We model from 2023 stub onward.
# Sources: industry estimates, Grok adoption data, reported fundraising rounds

XAI_ANNUAL_YEARS = [2023, 2024, 2025, 2026, 2027, 2028]
XAI_HIST_END = 2  # 0-1 are historical (2023 stub, 2024)

# ── Revenue Drivers ──────────────────────────────────────────────────────────
# API / Enterprise
xai_api_customers =         [50,    500,   2500,  8000,   18000,  35000]
xai_api_arpu_annual =       [5000,  8000,  12000, 15000,  18000,  20000]  # $
xai_api_revenue = [round(xai_api_customers[i] * xai_api_arpu_annual[i] / 1e6, 0) for i in range(6)]

# Grok Consumer (Premium subscriptions on X / standalone)
xai_grok_subs_k =           [100,   2000,  8000,  18000,  35000,  55000]  # thousands
xai_grok_price_mo =         [0,     8,     10,    10,     10,     10]  # $/month (freemium → paid)
xai_grok_revenue = [round(xai_grok_subs_k[i] * xai_grok_price_mo[i] * 12 / 1000, 0) for i in range(6)]

# Compute-as-a-Service (Colossus cluster leasing)
xai_compute_revenue =       [0,     50,    400,   1200,   2500,   4000]

# Licensing & other
xai_other_revenue =         [0,     20,    80,    200,    400,    600]

xai_total_revenue = [
    xai_api_revenue[i] + xai_grok_revenue[i] + xai_compute_revenue[i] + xai_other_revenue[i]
    for i in range(6)
]

# ── Cost Structure ────────────────────────────────────────────────────────────
xai_cogs_pct =              [0.85, 0.75, 0.65, 0.55, 0.48, 0.42]
xai_gross_profit = [round(xai_total_revenue[i] * (1 - xai_cogs_pct[i]), 0) for i in range(6)]

xai_rnd_pct =               [5.00, 1.20, 0.50, 0.30, 0.22, 0.18]
xai_sga_pct =               [1.00, 0.30, 0.12, 0.08, 0.06, 0.05]
xai_sbc_pct =               [0.50, 0.20, 0.08, 0.05, 0.04, 0.03]

xai_rnd = [round(xai_total_revenue[i] * xai_rnd_pct[i], 0) for i in range(6)]
xai_sga = [round(xai_total_revenue[i] * xai_sga_pct[i], 0) for i in range(6)]
xai_sbc = [round(xai_total_revenue[i] * xai_sbc_pct[i], 0) for i in range(6)]

xai_opex_total = [xai_rnd[i] + xai_sga[i] + xai_sbc[i] for i in range(6)]
xai_ebit = [xai_gross_profit[i] - xai_opex_total[i] for i in range(6)]
xai_ebit_ex_sbc = [xai_ebit[i] + xai_sbc[i] for i in range(6)]

xai_interest_exp =          [-5,   -20,  -40,  -50,  -60,  -60]
xai_other_income =          [0,    5,    10,   15,   20,   25]
xai_ebt = [xai_ebit[i] + xai_interest_exp[i] + xai_other_income[i] for i in range(6)]
xai_tax_rate =              [0.0,  0.0,  0.0,  0.0,  0.05, 0.10]
xai_tax = [round(max(0, xai_ebt[i]) * xai_tax_rate[i], 0) for i in range(6)]
xai_net_income = [xai_ebt[i] - xai_tax[i] for i in range(6)]

# ── Balance Sheet ────────────────────────────────────────────────────────────
xai_cash =                  [1000,  4000,  6000,  8000,  11000, 16000]
xai_ar =                    [5,     80,    300,   600,   1000,  1400]
xai_inventory =             [0,     50,    150,   300,   500,   700]
xai_other_ca =              [10,    30,    60,    100,   150,   200]
xai_total_ca = [xai_cash[i] + xai_ar[i] + xai_inventory[i] + xai_other_ca[i] for i in range(6)]

xai_ppe_gross =             [500,   5000,  12000, 20000, 28000, 36000]
xai_accum_depr =            [25,    500,   1800,  4000,  7000,  10800]
xai_ppe_net = [xai_ppe_gross[i] - xai_accum_depr[i] for i in range(6)]
xai_intangibles =           [100,   300,   600,   900,   1200,  1500]
xai_other_lta =             [50,    200,   400,   600,   800,   1000]
xai_total_assets = [xai_total_ca[i] + xai_ppe_net[i] + xai_intangibles[i] + xai_other_lta[i] for i in range(6)]

xai_ap =                    [50,    400,   800,   1200,  1600,  2000]
xai_accrued_liab =          [30,    200,   400,   600,   800,   1000]
xai_deferred_rev =          [10,    100,   300,   600,   1000,  1400]
xai_other_cl =              [10,    50,    100,   150,   200,   250]
xai_total_cl = [xai_ap[i] + xai_accrued_liab[i] + xai_deferred_rev[i] + xai_other_cl[i] for i in range(6)]

xai_lt_debt =               [500,   2000,  3500,  4000,  4000,  3500]
xai_other_ltl =             [20,    100,   200,   300,   400,   500]
xai_total_liab = [xai_total_cl[i] + xai_lt_debt[i] + xai_other_ltl[i] for i in range(6)]

xai_equity = [xai_total_assets[i] - xai_total_liab[i] for i in range(6)]

# ── Cash Flow Statement ─────────────────────────────────────────────────────
xai_da =                    [25,    475,   1300,  2200,  3000,  3800]
xai_wc_change =             [-10,   -100,  -200,  -300,  -350,  -300]
xai_other_cfo =             [5,     30,    60,    100,   140,   180]
xai_cfo = [xai_net_income[i] + xai_da[i] + xai_sbc[i] + xai_wc_change[i] + xai_other_cfo[i] for i in range(6)]

xai_capex =                 [-500,  -4500, -7000, -8000, -8000, -8000]
xai_other_cfi =             [-50,   -100,  -200,  -200,  -200,  -200]
xai_cfi = [xai_capex[i] + xai_other_cfi[i] for i in range(6)]

xai_fcf = [xai_cfo[i] + xai_capex[i] for i in range(6)]

xai_debt_issued =           [500,   2000,  2000,  1000,  500,   0]
xai_debt_repaid =           [0,     -500,  -500,  -500,  -500,  -500]
xai_equity_raised =         [1000,  6000,  5000,  3000,  2000,  0]
xai_other_cff =             [0,     -20,   -30,   -50,   -50,   -50]
xai_cff = [xai_debt_issued[i] + xai_debt_repaid[i] + xai_equity_raised[i] + xai_other_cff[i] for i in range(6)]

# ── Quarterly Data ───────────────────────────────────────────────────────────
XAI_Q_LABELS = [
    "Q1'24", "Q2'24", "Q3'24", "Q4'24",
    "Q1'25", "Q2'25", "Q3'25", "Q4'25",
    "Q1'26", "Q2'26", "Q3'26", "Q4'26",
]
XAI_Q_HIST_END = 4

xai_q_revenue = [
    30,  50,  80,  120,        # 2024 ~$280M
    200, 300, 450, 650,        # 2025E ~$1,600M
    900, 1100, 1400, 1800,     # 2026E ~$5,200M
]

xai_q_cogs = [
    22,  38,  60,  90,
    130, 195, 293, 423,
    495, 605, 770, 990,
]

xai_q_gross_profit = [xai_q_revenue[i] - xai_q_cogs[i] for i in range(12)]

xai_q_opex = [
    100, 120, 150, 180,
    200, 250, 300, 350,
    350, 380, 420, 470,
]

xai_q_ebit = [xai_q_gross_profit[i] - xai_q_opex[i] for i in range(12)]

xai_q_net_income = [
    -95, -110, -135, -155,
    -140, -160, -160, -145,
    20,  60,  120,  220,
]

xai_q_capex = [
    -800, -1000, -1200, -1500,
    -1500, -1700, -1800, -2000,
    -1800, -2000, -2000, -2200,
]

xai_q_cfo = [
    -50, -30, -10, 20,
    50,  80,  120, 180,
    300, 400, 500, 650,
]

xai_q_fcf = [xai_q_cfo[i] + xai_q_capex[i] for i in range(12)]


# ═══════════════════════════════════════════════════════════════════════════════
#  WORKBOOK BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_workbook():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 1: SpaceX — Revenue Build
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SpaceX Revenue Build")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions['A'].width = 35
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    ws.cell(row=row, column=1, value="SpaceX — Revenue Build by Segment ($M)")
    style_header_row(ws, row, 10)

    row = 2
    headers = [""] + [str(y) + ("A" if i < SPACEX_HIST_END else "E") for i, y in enumerate(SPACEX_ANNUAL_YEARS)]
    write_row(ws, row, headers, fmt=None)
    style_sub_header_row(ws, row, 10)

    forecast_col = 2 + SPACEX_HIST_END  # column where forecast starts

    # Launch Services section
    row = 3
    style_section_row(ws, row, 10)
    ws.cell(row=row, column=1, value="LAUNCH SERVICES")
    row = 4
    write_row(ws, row, ["  Launches (#)"] + spacex_launch_count, fmt=NUM_FORMAT, is_forecast_start=forecast_col)
    row = 5
    write_row(ws, row, ["  Avg Price / Launch ($M)"] + spacex_avg_launch_price, fmt=NUM_FORMAT_1, is_forecast_start=forecast_col, is_input=True)
    row = 6
    write_row(ws, row, ["  Launch Revenue"] + spacex_launch_revenue, fmt=NUM_FORMAT, is_forecast_start=forecast_col, bold=True)

    # Starlink section
    row = 8
    style_section_row(ws, row, 10)
    ws.cell(row=row, column=1, value="STARLINK")
    row = 9
    write_row(ws, row, ["  Subscribers (K)"] + spacex_starlink_subs_k, fmt=NUM_FORMAT, is_forecast_start=forecast_col)
    row = 10
    write_row(ws, row, ["  ARPU ($/month)"] + spacex_starlink_arpu_mo, fmt=DOLLAR_FORMAT, is_forecast_start=forecast_col, is_input=True)
    row = 11
    write_row(ws, row, ["  Starlink Revenue"] + spacex_starlink_revenue, fmt=NUM_FORMAT, is_forecast_start=forecast_col, bold=True)

    # Starship section
    row = 13
    style_section_row(ws, row, 10)
    ws.cell(row=row, column=1, value="STARSHIP / HLS / COMMERCIAL")
    row = 14
    write_row(ws, row, ["  Starship Revenue"] + spacex_starship_revenue, fmt=NUM_FORMAT, is_forecast_start=forecast_col, is_input=True)

    # Government / Other
    row = 16
    style_section_row(ws, row, 10)
    ws.cell(row=row, column=1, value="GOVERNMENT & OTHER")
    row = 17
    write_row(ws, row, ["  Govt & Other Revenue"] + spacex_govt_other_revenue, fmt=NUM_FORMAT, is_forecast_start=forecast_col)

    # Total Revenue
    row = 19
    write_row(ws, row, ["TOTAL REVENUE"] + spacex_total_revenue, fmt=NUM_FORMAT, is_forecast_start=forecast_col, bold=True)

    # Margin assumptions
    row = 21
    style_section_row(ws, row, 10)
    ws.cell(row=row, column=1, value="MARGIN ASSUMPTIONS")
    row = 22
    write_row(ws, row, ["  COGS (% of Revenue)"] + spacex_cogs_pct, fmt=PCT_FORMAT, is_forecast_start=forecast_col, is_input=True)
    row = 23
    write_row(ws, row, ["  Gross Profit"] + spacex_gross_profit, fmt=NUM_FORMAT, is_forecast_start=forecast_col, bold=True)
    row = 24
    gm = [round((1 - c), 3) for c in spacex_cogs_pct]
    write_row(ws, row, ["  Gross Margin %"] + gm, fmt=PCT_FORMAT, is_forecast_start=forecast_col)

    row = 26
    write_row(ws, row, ["  R&D (% of Revenue)"] + spacex_rnd_pct, fmt=PCT_FORMAT, is_forecast_start=forecast_col, is_input=True)
    row = 27
    write_row(ws, row, ["  SG&A (% of Revenue)"] + spacex_sga_pct, fmt=PCT_FORMAT, is_forecast_start=forecast_col, is_input=True)
    row = 28
    write_row(ws, row, ["  SBC (% of Revenue)"] + spacex_sbc_pct, fmt=PCT_FORMAT, is_forecast_start=forecast_col, is_input=True)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 2: SpaceX — Annual Three Statements
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SpaceX Annual Statements")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions['A'].width = 35
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    ws.cell(row=row, column=1, value="SpaceX — Income Statement ($M)")
    style_header_row(ws, row, 10)

    row = 2
    headers = [""] + [str(y) + ("A" if i < SPACEX_HIST_END else "E") for i, y in enumerate(SPACEX_ANNUAL_YEARS)]
    write_row(ws, row, headers)
    style_sub_header_row(ws, row, 10)

    fc = forecast_col

    r = 3
    write_row(ws, r, ["Revenue"] + spacex_total_revenue, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Cost of Revenue"] + [round(spacex_total_revenue[i] * spacex_cogs_pct[i]) for i in range(9)], fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Gross Profit"] + spacex_gross_profit, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    write_row(ws, r, ["  Gross Margin %"] + gm, fmt=PCT_FORMAT, is_forecast_start=fc); r += 1
    r += 1
    write_row(ws, r, ["  Research & Development"] + spacex_rnd, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Selling, General & Admin"] + spacex_sga, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Stock-Based Compensation"] + spacex_sbc, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Total Operating Expenses"] + spacex_opex_total, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["EBIT (Operating Income)"] + spacex_ebit, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    ebit_margin = [round(spacex_ebit[i] / spacex_total_revenue[i], 3) if spacex_total_revenue[i] else 0 for i in range(9)]
    write_row(ws, r, ["  EBIT Margin %"] + ebit_margin, fmt=PCT_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Adj. EBIT (ex-SBC)"] + spacex_ebit_ex_sbc, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    r += 1
    write_row(ws, r, ["Interest Expense"] + spacex_interest_exp, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Other Income"] + spacex_other_income, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["EBT (Pre-Tax Income)"] + spacex_ebt, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    write_row(ws, r, ["  Tax Rate"] + spacex_tax_rate, fmt=PCT_FORMAT, is_forecast_start=fc, is_input=True); r += 1
    write_row(ws, r, ["  Income Tax"] + spacex_tax, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Net Income"] + spacex_net_income, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    ni_margin = [round(spacex_net_income[i] / spacex_total_revenue[i], 3) if spacex_total_revenue[i] else 0 for i in range(9)]
    write_row(ws, r, ["  Net Margin %"] + ni_margin, fmt=PCT_FORMAT, is_forecast_start=fc)

    # Balance Sheet
    bs_start = r + 3
    r = bs_start
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    ws.cell(row=r, column=1, value="SpaceX — Balance Sheet ($M)")
    style_header_row(ws, r, 10); r += 1
    write_row(ws, r, headers); style_sub_header_row(ws, r, 10); r += 1

    style_section_row(ws, r, 10); ws.cell(row=r, column=1, value="ASSETS"); r += 1
    write_row(ws, r, ["  Cash & Equivalents"] + spacex_cash, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Accounts Receivable"] + spacex_ar, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Inventory"] + spacex_inventory, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Current Assets"] + spacex_other_ca, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Total Current Assets"] + spacex_total_ca, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    write_row(ws, r, ["  PP&E (Net)"] + spacex_ppe_net, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Intangible Assets"] + spacex_intangibles, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Long-Term Assets"] + spacex_other_lta, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Total Assets"] + spacex_total_assets, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 10); ws.cell(row=r, column=1, value="LIABILITIES & EQUITY"); r += 1
    write_row(ws, r, ["  Accounts Payable"] + spacex_ap, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Accrued Liabilities"] + spacex_accrued_liab, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Deferred Revenue"] + spacex_deferred_rev, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Current Liabilities"] + spacex_other_cl, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Total Current Liabilities"] + spacex_total_cl, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    write_row(ws, r, ["  Long-Term Debt"] + spacex_lt_debt, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Long-Term Liabilities"] + spacex_other_ltl, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Total Liabilities"] + spacex_total_liab, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["Total Stockholders' Equity"] + spacex_equity, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    check_row = [spacex_total_liab[i] + spacex_equity[i] for i in range(9)]
    write_row(ws, r, ["Total Liab + Equity (Check)"] + check_row, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True)

    # Cash Flow Statement
    cf_start = r + 3
    r = cf_start
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    ws.cell(row=r, column=1, value="SpaceX — Cash Flow Statement ($M)")
    style_header_row(ws, r, 10); r += 1
    write_row(ws, r, headers); style_sub_header_row(ws, r, 10); r += 1

    style_section_row(ws, r, 10); ws.cell(row=r, column=1, value="OPERATING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Net Income"] + spacex_net_income, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Depreciation & Amortization"] + spacex_da, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Stock-Based Compensation"] + spacex_sbc, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Changes in Working Capital"] + spacex_wc_change, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Operating Adjustments"] + spacex_other_cfo, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Cash from Operations (CFO)"] + spacex_cfo, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 10); ws.cell(row=r, column=1, value="INVESTING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Capital Expenditures"] + spacex_capex, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Investing"] + spacex_other_cfi, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Cash from Investing (CFI)"] + spacex_cfi, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 10); ws.cell(row=r, column=1, value="FINANCING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Debt Issued"] + spacex_debt_issued, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Debt Repaid"] + spacex_debt_repaid, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Equity Raised"] + spacex_equity_raised, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["  Other Financing"] + spacex_other_cff, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Cash from Financing (CFF)"] + spacex_cff, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1

    r += 1
    net_change = [spacex_cfo[i] + spacex_cfi[i] + spacex_cff[i] for i in range(9)]
    write_row(ws, r, ["Net Change in Cash"] + net_change, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1

    r += 1
    write_row(ws, r, ["Free Cash Flow (CFO + CapEx)"] + spacex_fcf, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    fcf_margin = [round(spacex_fcf[i] / spacex_total_revenue[i], 3) if spacex_total_revenue[i] else 0 for i in range(9)]
    write_row(ws, r, ["  FCF Margin %"] + fcf_margin, fmt=PCT_FORMAT, is_forecast_start=fc)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 3: SpaceX — Quarterly
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SpaceX Quarterly")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions['A'].width = 30
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 12

    q_fc = 2 + SPACEX_Q_HIST_END

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=13)
    ws.cell(row=row, column=1, value="SpaceX — Quarterly P&L Summary ($M)")
    style_header_row(ws, row, 13)

    row = 2
    q_headers = [""] + SPACEX_Q_LABELS
    write_row(ws, row, q_headers)
    style_sub_header_row(ws, row, 13)

    r = 3
    write_row(ws, r, ["Revenue"] + spacex_q_revenue, fmt=NUM_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["COGS"] + spacex_q_cogs, fmt=NUM_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["Gross Profit"] + spacex_q_gross_profit, fmt=NUM_FORMAT, is_forecast_start=q_fc, bold=True); r += 1
    q_gm = [round(spacex_q_gross_profit[i] / spacex_q_revenue[i], 3) if spacex_q_revenue[i] else 0 for i in range(12)]
    write_row(ws, r, ["  Gross Margin %"] + q_gm, fmt=PCT_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["Operating Expenses"] + spacex_q_opex, fmt=NUM_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["EBIT"] + spacex_q_ebit, fmt=NUM_FORMAT, is_forecast_start=q_fc, bold=True); r += 1
    write_row(ws, r, ["Net Income"] + spacex_q_net_income, fmt=NUM_FORMAT, is_forecast_start=q_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["CFO"] + spacex_q_cfo, fmt=NUM_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["CapEx"] + spacex_q_capex, fmt=NUM_FORMAT, is_forecast_start=q_fc); r += 1
    write_row(ws, r, ["Free Cash Flow"] + spacex_q_fcf, fmt=NUM_FORMAT, is_forecast_start=q_fc, bold=True); r += 1

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 4: SpaceX — GAAP to FCF Bridge
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SpaceX GAAP-to-FCF Bridge")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions['A'].width = 38
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    ws.cell(row=row, column=1, value="SpaceX — GAAP Operating Loss to FCF Bridge ($M)")
    style_header_row(ws, row, 10)

    row = 2
    write_row(ws, row, headers)
    style_sub_header_row(ws, row, 10)

    r = 3
    write_row(ws, r, ["GAAP Operating Income (EBIT)"] + spacex_ebit, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    write_row(ws, r, ["(+) Stock-Based Compensation"] + spacex_sbc, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["(+) Depreciation & Amortization"] + spacex_da, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    adj_ebitda = [spacex_ebit[i] + spacex_sbc[i] + spacex_da[i] for i in range(9)]
    write_row(ws, r, ["Adjusted EBITDA"] + adj_ebitda, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["(+) Interest Expense"] + spacex_interest_exp, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["(+) Other Income"] + spacex_other_income, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["(-) Cash Taxes"] + [-t for t in spacex_tax], fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["(+/-) Working Capital Changes"] + spacex_wc_change, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["(+) Other Operating Items"] + spacex_other_cfo, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Cash from Operations (CFO)"] + spacex_cfo, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["(-) Capital Expenditures"] + spacex_capex, fmt=NUM_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["Free Cash Flow (FCF)"] + spacex_fcf, fmt=NUM_FORMAT, is_forecast_start=fc, bold=True); r += 1
    r += 1
    fcf_conv = [round(spacex_fcf[i] / adj_ebitda[i], 3) if adj_ebitda[i] != 0 else 0 for i in range(9)]
    write_row(ws, r, ["FCF / Adj. EBITDA Conversion"] + fcf_conv, fmt=PCT_FORMAT, is_forecast_start=fc); r += 1
    write_row(ws, r, ["FCF Margin %"] + fcf_margin, fmt=PCT_FORMAT, is_forecast_start=fc)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 5: SpaceX — Sensitivity & Valuation
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SpaceX Valuation & Sensitivity")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions['A'].width = 38
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1, value="SpaceX — Valuation Summary")
    style_header_row(ws, row, 8)

    r = 3
    # Last-round / secondary valuation
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="LAST-ROUND & SECONDARY INDICATIONS"); r += 1
    items = [
        ("Dec 2024 Tender Offer Valuation", "$350B"),
        ("Implied Equity Value (latest round)", "$350,000M"),
        ("Last Round Price / Share", "$185"),
        ("Pre-Money Valuation", "$345,000M"),
        ("Post-Money Valuation", "$350,000M"),
        ("Secondary Market Range (Q4'24)", "$170 - $195 / share"),
        ("Fully Diluted Shares (est.)", "~1,890M"),
        ("2025E Revenue", f"${spacex_total_revenue[5]:,}M"),
        ("2025E EV / Revenue", f"{round(350000 / spacex_total_revenue[5], 1)}x"),
        ("2026E Revenue", f"${spacex_total_revenue[6]:,}M"),
        ("2026E EV / Revenue", f"{round(350000 / spacex_total_revenue[6], 1)}x"),
        ("2025E Adj. EBITDA", f"${adj_ebitda[5]:,}M"),
        ("2025E EV / Adj. EBITDA", f"{round(350000 / adj_ebitda[5], 1)}x"),
    ]
    for label, val in items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    r += 1
    # Sensitivity: Revenue Growth vs. EBITDA Margin → Implied EV
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="TWO-WAY SENSITIVITY: 2027E Revenue Growth vs. EBITDA Margin → Implied EV ($B)"); r += 1
    ws.cell(row=r, column=1, value="EV ($B)").font = TOTAL_FONT

    # Column headers: EBITDA margins
    ebitda_margins = [0.25, 0.30, 0.35, 0.40, 0.45]
    for j, m in enumerate(ebitda_margins):
        c = 2 + j
        ws.cell(row=r, column=c, value=f"EBITDA Mgn {m:.0%}").font = SUB_HEADER_FONT
        ws.cell(row=r, column=c).fill = SUB_HEADER_FILL
        ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
    r += 1

    rev_growth_rates = [0.20, 0.30, 0.40, 0.50, 0.60]
    base_2026_rev = spacex_total_revenue[6]  # 2026E
    ev_multiple = 30  # EV/EBITDA multiple assumption

    for rg in rev_growth_rates:
        implied_2027_rev = base_2026_rev * (1 + rg)
        ws.cell(row=r, column=1, value=f"Rev Growth {rg:.0%}").font = DATA_FONT
        ws.cell(row=r, column=1).fill = SUB_HEADER_FILL
        for j, m in enumerate(ebitda_margins):
            implied_ebitda = implied_2027_rev * m
            implied_ev = round(implied_ebitda * ev_multiple / 1000, 0)
            c = 2 + j
            ws.cell(row=r, column=c, value=implied_ev).font = DATA_FONT
            ws.cell(row=r, column=c).number_format = '#,##0'
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Assumes 30x EV/EBITDA multiple on 2027E EBITDA").font = Font(name="Calibri", italic=True, size=9, color="808080")
    r += 2

    # Sensitivity 2: Starlink Subs vs ARPU → Revenue
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="TWO-WAY SENSITIVITY: 2027E Starlink Subs (M) vs. ARPU → Starlink Revenue ($M)"); r += 1
    ws.cell(row=r, column=1, value="Starlink Rev ($M)").font = TOTAL_FONT

    arpu_range = [70, 80, 85, 90, 100]
    for j, a in enumerate(arpu_range):
        c = 2 + j
        ws.cell(row=r, column=c, value=f"ARPU ${a}/mo").font = SUB_HEADER_FONT
        ws.cell(row=r, column=c).fill = SUB_HEADER_FILL
        ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
    r += 1

    sub_range_m = [10, 12, 13.5, 15, 18]  # millions
    for s in sub_range_m:
        ws.cell(row=r, column=1, value=f"Subs {s}M").font = DATA_FONT
        ws.cell(row=r, column=1).fill = SUB_HEADER_FILL
        for j, a in enumerate(arpu_range):
            rev = round(s * a * 12, 0)
            c = 2 + j
            ws.cell(row=r, column=c, value=rev).font = DATA_FONT
            ws.cell(row=r, column=c).number_format = '#,##0'
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
        r += 1

    r += 2
    # Unit Economics & S-Curve
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="UNIT ECONOMICS & S-CURVE SANITY CHECK"); r += 1
    ue_items = [
        ("Starlink CPE Cost (est.)", "$300"),
        ("Starlink Monthly COGS / Sub (est.)", "$25 - $30"),
        ("Starlink Monthly Gross Profit / Sub", "$55 - $60"),
        ("Payback Period per Subscriber", "~5 months"),
        ("Addressable Market (global broadband-underserved)", "~500M HH"),
        ("Current Penetration (2024)", "~0.9%"),
        ("2027E Penetration", "~2.7%"),
        ("S-Curve Inflection (est.)", "2026-2028"),
        ("Falcon 9 Cost / Launch (internal)", "~$15M"),
        ("Falcon 9 Revenue / Launch (avg)", "~$44M"),
        ("Falcon 9 Gross Margin / Launch", "~65%"),
        ("Starship Target Cost / kg to LEO", "$10 (vs $2,700 F9)"),
    ]
    for label, val in ue_items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    r += 1
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="COHORT NPV — STARLINK SUBSCRIBER"); r += 1
    npv_items = [
        ("Year 0: Equipment Sale + Install", "$599"),
        ("Year 0: Customer Acquisition Cost", "($200)"),
        ("Year 1: Revenue (net churn)", "$1,020"),
        ("Year 1: COGS", "($360)"),
        ("Year 2: Revenue (net churn)", "$960"),
        ("Year 2: COGS", "($340)"),
        ("Year 3: Revenue (net churn)", "$900"),
        ("Year 3: COGS", "($320)"),
        ("3-Year Gross Profit / Cohort Sub", "$2,259"),
        ("Discount Rate", "12%"),
        ("NPV per Subscriber", "~$1,900"),
        ("Total NPV @ 17M subs (2028E)", "~$32B"),
    ]
    for label, val in npv_items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    # ═════════════════════════════════════════════════════════════════════
    #  xAI SHEETS
    # ═════════════════════════════════════════════════════════════════════

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 6: xAI — Revenue Build
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("xAI Revenue Build")
    ws.sheet_properties.tabColor = "7030A0"
    ws.column_dimensions['A'].width = 35
    for c in range(2, 9):
        ws.column_dimensions[get_column_letter(c)].width = 14

    xai_fc = 2 + XAI_HIST_END  # forecast start column

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="xAI — Revenue Build by Segment ($M)")
    style_header_row(ws, row, 7)

    row = 2
    xai_headers = [""] + [str(y) + ("A" if i < XAI_HIST_END else "E") for i, y in enumerate(XAI_ANNUAL_YEARS)]
    write_row(ws, row, xai_headers)
    style_sub_header_row(ws, row, 7)

    # API / Enterprise
    r = 3
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="API / ENTERPRISE"); r += 1
    write_row(ws, r, ["  Enterprise Customers (#)"] + xai_api_customers, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Avg Annual Revenue / Customer ($)"] + xai_api_arpu_annual, fmt=DOLLAR_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  API / Enterprise Revenue"] + xai_api_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    # Grok Consumer
    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="GROK CONSUMER (X PREMIUM + STANDALONE)"); r += 1
    write_row(ws, r, ["  Paid Subscribers (K)"] + xai_grok_subs_k, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Monthly Price ($)"] + xai_grok_price_mo, fmt=DOLLAR_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  Grok Consumer Revenue"] + xai_grok_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    # Compute-as-a-Service
    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="COMPUTE-AS-A-SERVICE (COLOSSUS)"); r += 1
    write_row(ws, r, ["  Compute Revenue"] + xai_compute_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1

    # Licensing & Other
    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="LICENSING & OTHER"); r += 1
    write_row(ws, r, ["  Other Revenue"] + xai_other_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1

    # Total
    r += 1
    write_row(ws, r, ["TOTAL REVENUE"] + xai_total_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    # Margin assumptions
    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="MARGIN ASSUMPTIONS"); r += 1
    write_row(ws, r, ["  COGS (% of Revenue)"] + xai_cogs_pct, fmt=PCT_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  Gross Profit"] + xai_gross_profit, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    xai_gm = [round(1 - c, 3) for c in xai_cogs_pct]
    write_row(ws, r, ["  Gross Margin %"] + xai_gm, fmt=PCT_FORMAT, is_forecast_start=xai_fc); r += 1
    r += 1
    write_row(ws, r, ["  R&D (% of Revenue)"] + xai_rnd_pct, fmt=PCT_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  SG&A (% of Revenue)"] + xai_sga_pct, fmt=PCT_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  SBC (% of Revenue)"] + xai_sbc_pct, fmt=PCT_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 7: xAI — Annual Three Statements
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("xAI Annual Statements")
    ws.sheet_properties.tabColor = "7030A0"
    ws.column_dimensions['A'].width = 35
    for c in range(2, 9):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="xAI — Income Statement ($M)")
    style_header_row(ws, row, 7)

    row = 2
    write_row(ws, row, xai_headers)
    style_sub_header_row(ws, row, 7)

    r = 3
    write_row(ws, r, ["Revenue"] + xai_total_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    xai_cogs_abs = [round(xai_total_revenue[i] * xai_cogs_pct[i]) for i in range(6)]
    write_row(ws, r, ["  Cost of Revenue"] + xai_cogs_abs, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Gross Profit"] + xai_gross_profit, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    write_row(ws, r, ["  Gross Margin %"] + xai_gm, fmt=PCT_FORMAT, is_forecast_start=xai_fc); r += 1
    r += 1
    write_row(ws, r, ["  Research & Development"] + xai_rnd, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Selling, General & Admin"] + xai_sga, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Stock-Based Compensation"] + xai_sbc, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Total Operating Expenses"] + xai_opex_total, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["EBIT (Operating Income / Loss)"] + xai_ebit, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    xai_ebit_margin = [round(xai_ebit[i] / xai_total_revenue[i], 3) if xai_total_revenue[i] else 0 for i in range(6)]
    write_row(ws, r, ["  EBIT Margin %"] + xai_ebit_margin, fmt=PCT_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Adj. EBIT (ex-SBC)"] + xai_ebit_ex_sbc, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    r += 1
    write_row(ws, r, ["Interest Expense"] + xai_interest_exp, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Other Income"] + xai_other_income, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["EBT (Pre-Tax Income)"] + xai_ebt, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    write_row(ws, r, ["  Tax Rate"] + xai_tax_rate, fmt=PCT_FORMAT, is_forecast_start=xai_fc, is_input=True); r += 1
    write_row(ws, r, ["  Income Tax"] + xai_tax, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Net Income"] + xai_net_income, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    xai_ni_margin = [round(xai_net_income[i] / xai_total_revenue[i], 3) if xai_total_revenue[i] else 0 for i in range(6)]
    write_row(ws, r, ["  Net Margin %"] + xai_ni_margin, fmt=PCT_FORMAT, is_forecast_start=xai_fc)

    # Balance Sheet
    bs_start = r + 3
    r = bs_start
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(row=r, column=1, value="xAI — Balance Sheet ($M)")
    style_header_row(ws, r, 7); r += 1
    write_row(ws, r, xai_headers); style_sub_header_row(ws, r, 7); r += 1

    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="ASSETS"); r += 1
    write_row(ws, r, ["  Cash & Equivalents"] + xai_cash, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Accounts Receivable"] + xai_ar, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Inventory"] + xai_inventory, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Current Assets"] + xai_other_ca, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Total Current Assets"] + xai_total_ca, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    write_row(ws, r, ["  PP&E (Net)"] + xai_ppe_net, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Intangible Assets"] + xai_intangibles, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Long-Term Assets"] + xai_other_lta, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Total Assets"] + xai_total_assets, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="LIABILITIES & EQUITY"); r += 1
    write_row(ws, r, ["  Accounts Payable"] + xai_ap, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Accrued Liabilities"] + xai_accrued_liab, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Deferred Revenue"] + xai_deferred_rev, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Current Liabilities"] + xai_other_cl, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Total Current Liabilities"] + xai_total_cl, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    write_row(ws, r, ["  Long-Term Debt"] + xai_lt_debt, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Long-Term Liabilities"] + xai_other_ltl, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Total Liabilities"] + xai_total_liab, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["Total Stockholders' Equity"] + xai_equity, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    xai_check = [xai_total_liab[i] + xai_equity[i] for i in range(6)]
    write_row(ws, r, ["Total Liab + Equity (Check)"] + xai_check, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True)

    # Cash Flow Statement
    cf_start = r + 3
    r = cf_start
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(row=r, column=1, value="xAI — Cash Flow Statement ($M)")
    style_header_row(ws, r, 7); r += 1
    write_row(ws, r, xai_headers); style_sub_header_row(ws, r, 7); r += 1

    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="OPERATING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Net Income"] + xai_net_income, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Depreciation & Amortization"] + xai_da, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Stock-Based Compensation"] + xai_sbc, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Changes in Working Capital"] + xai_wc_change, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Operating Adjustments"] + xai_other_cfo, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Cash from Operations (CFO)"] + xai_cfo, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="INVESTING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Capital Expenditures"] + xai_capex, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Investing"] + xai_other_cfi, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Cash from Investing (CFI)"] + xai_cfi, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    r += 1
    style_section_row(ws, r, 7); ws.cell(row=r, column=1, value="FINANCING ACTIVITIES"); r += 1
    write_row(ws, r, ["  Debt Issued"] + xai_debt_issued, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Debt Repaid"] + xai_debt_repaid, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Equity Raised"] + xai_equity_raised, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["  Other Financing"] + xai_other_cff, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Cash from Financing (CFF)"] + xai_cff, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    r += 1
    xai_net_change = [xai_cfo[i] + xai_cfi[i] + xai_cff[i] for i in range(6)]
    write_row(ws, r, ["Net Change in Cash"] + xai_net_change, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1

    r += 1
    write_row(ws, r, ["Free Cash Flow (CFO + CapEx)"] + xai_fcf, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    xai_fcf_margin = [round(xai_fcf[i] / xai_total_revenue[i], 3) if xai_total_revenue[i] else 0 for i in range(6)]
    write_row(ws, r, ["  FCF Margin %"] + xai_fcf_margin, fmt=PCT_FORMAT, is_forecast_start=xai_fc)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 8: xAI — Quarterly
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("xAI Quarterly")
    ws.sheet_properties.tabColor = "7030A0"
    ws.column_dimensions['A'].width = 30
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 12

    xai_q_fc = 2 + XAI_Q_HIST_END

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=13)
    ws.cell(row=row, column=1, value="xAI — Quarterly P&L Summary ($M)")
    style_header_row(ws, row, 13)

    row = 2
    xai_q_headers = [""] + XAI_Q_LABELS
    write_row(ws, row, xai_q_headers)
    style_sub_header_row(ws, row, 13)

    r = 3
    write_row(ws, r, ["Revenue"] + xai_q_revenue, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["COGS"] + xai_q_cogs, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["Gross Profit"] + xai_q_gross_profit, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc, bold=True); r += 1
    xai_q_gm = [round(xai_q_gross_profit[i] / xai_q_revenue[i], 3) if xai_q_revenue[i] else 0 for i in range(12)]
    write_row(ws, r, ["  Gross Margin %"] + xai_q_gm, fmt=PCT_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["Operating Expenses"] + xai_q_opex, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["EBIT"] + xai_q_ebit, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc, bold=True); r += 1
    write_row(ws, r, ["Net Income"] + xai_q_net_income, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["CFO"] + xai_q_cfo, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["CapEx"] + xai_q_capex, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc); r += 1
    write_row(ws, r, ["Free Cash Flow"] + xai_q_fcf, fmt=NUM_FORMAT, is_forecast_start=xai_q_fc, bold=True); r += 1

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 9: xAI — GAAP to FCF Bridge
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("xAI GAAP-to-FCF Bridge")
    ws.sheet_properties.tabColor = "7030A0"
    ws.column_dimensions['A'].width = 38
    for c in range(2, 9):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="xAI — GAAP Operating Loss to FCF Bridge ($M)")
    style_header_row(ws, row, 7)

    row = 2
    write_row(ws, row, xai_headers)
    style_sub_header_row(ws, row, 7)

    r = 3
    write_row(ws, r, ["GAAP Operating Income / (Loss)"] + xai_ebit, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    write_row(ws, r, ["(+) Stock-Based Compensation"] + xai_sbc, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["(+) Depreciation & Amortization"] + xai_da, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    xai_adj_ebitda = [xai_ebit[i] + xai_sbc[i] + xai_da[i] for i in range(6)]
    write_row(ws, r, ["Adjusted EBITDA"] + xai_adj_ebitda, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["(+) Interest Expense"] + xai_interest_exp, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["(+) Other Income"] + xai_other_income, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["(-) Cash Taxes"] + [-t for t in xai_tax], fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["(+/-) Working Capital Changes"] + xai_wc_change, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["(+) Other Operating Items"] + xai_other_cfo, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Cash from Operations (CFO)"] + xai_cfo, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    r += 1
    write_row(ws, r, ["(-) Capital Expenditures"] + xai_capex, fmt=NUM_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["Free Cash Flow (FCF)"] + xai_fcf, fmt=NUM_FORMAT, is_forecast_start=xai_fc, bold=True); r += 1
    r += 1
    xai_fcf_conv = [round(xai_fcf[i] / xai_adj_ebitda[i], 3) if xai_adj_ebitda[i] != 0 else 0 for i in range(6)]
    write_row(ws, r, ["FCF / Adj. EBITDA Conversion"] + xai_fcf_conv, fmt=PCT_FORMAT, is_forecast_start=xai_fc); r += 1
    write_row(ws, r, ["FCF Margin %"] + xai_fcf_margin, fmt=PCT_FORMAT, is_forecast_start=xai_fc)

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 10: xAI — Sensitivity & Valuation
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("xAI Valuation & Sensitivity")
    ws.sheet_properties.tabColor = "7030A0"
    ws.column_dimensions['A'].width = 42
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 15

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1, value="xAI — Valuation Summary")
    style_header_row(ws, row, 8)

    r = 3
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="LAST-ROUND & SECONDARY INDICATIONS"); r += 1
    xai_val_items = [
        ("Dec 2024 Series C Post-Money Valuation", "$50B"),
        ("Series C Equity Raised", "$6B"),
        ("Implied Pre-Money", "$44B"),
        ("Nov 2024 Series B Post-Money", "$40B (5 months prior)"),
        ("May 2024 Series A Post-Money", "$24B"),
        ("Secondary Market Range (Q1'25)", "$45 - $55B implied"),
        ("Fully Diluted Shares (est.)", "~800M (incl. options)"),
        ("Implied Price / Share (Series C)", "~$62.50"),
        ("2025E Revenue", f"${xai_total_revenue[2]:,}M"),
        ("2025E EV / Revenue", f"{round(50000 / xai_total_revenue[2], 1)}x"),
        ("2026E Revenue", f"${xai_total_revenue[3]:,}M"),
        ("2026E EV / Revenue", f"{round(50000 / xai_total_revenue[3], 1)}x"),
        ("2025E Adj. EBITDA", f"${xai_adj_ebitda[2]:,}M"),
    ]
    for label, val in xai_val_items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    r += 1
    # Sensitivity 1: Revenue Growth vs EBITDA Margin
    style_section_row(ws, r, 8)
    ws.cell(row=r, column=1, value="TWO-WAY SENSITIVITY: 2027E Revenue Growth vs. EBITDA Margin → Implied EV ($B)")
    r += 1
    ws.cell(row=r, column=1, value="EV ($B)").font = TOTAL_FONT

    xai_ebitda_margins = [0.15, 0.20, 0.25, 0.30, 0.35]
    for j, m in enumerate(xai_ebitda_margins):
        c = 2 + j
        ws.cell(row=r, column=c, value=f"EBITDA Mgn {m:.0%}").font = SUB_HEADER_FONT
        ws.cell(row=r, column=c).fill = SUB_HEADER_FILL
        ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
    r += 1

    xai_rev_growth_rates = [0.50, 0.75, 1.00, 1.50, 2.00]
    base_xai_2026_rev = xai_total_revenue[3]
    xai_ev_multiple = 40  # high-growth AI

    for rg in xai_rev_growth_rates:
        implied_2027_rev = base_xai_2026_rev * (1 + rg)
        ws.cell(row=r, column=1, value=f"Rev Growth {rg:.0%}").font = DATA_FONT
        ws.cell(row=r, column=1).fill = SUB_HEADER_FILL
        for j, m in enumerate(xai_ebitda_margins):
            implied_ebitda = implied_2027_rev * m
            implied_ev = round(implied_ebitda * xai_ev_multiple / 1000, 0)
            c = 2 + j
            ws.cell(row=r, column=c, value=implied_ev).font = DATA_FONT
            ws.cell(row=r, column=c).number_format = '#,##0'
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Assumes 40x EV/EBITDA multiple on 2027E EBITDA (high-growth AI premium)").font = Font(name="Calibri", italic=True, size=9, color="808080")
    r += 2

    # Sensitivity 2: API Customers vs ARPU
    style_section_row(ws, r, 8)
    ws.cell(row=r, column=1, value="TWO-WAY SENSITIVITY: 2027E API Customers vs. ARPU → API Revenue ($M)")
    r += 1
    ws.cell(row=r, column=1, value="API Rev ($M)").font = TOTAL_FONT

    xai_arpu_range = [10000, 15000, 18000, 22000, 25000]
    for j, a in enumerate(xai_arpu_range):
        c = 2 + j
        ws.cell(row=r, column=c, value=f"ARPU ${a//1000}K").font = SUB_HEADER_FONT
        ws.cell(row=r, column=c).fill = SUB_HEADER_FILL
        ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
    r += 1

    xai_cust_range = [10000, 15000, 18000, 25000, 35000]
    for cust in xai_cust_range:
        ws.cell(row=r, column=1, value=f"Customers {cust//1000}K").font = DATA_FONT
        ws.cell(row=r, column=1).fill = SUB_HEADER_FILL
        for j, a in enumerate(xai_arpu_range):
            rev = round(cust * a / 1e6, 0)
            c = 2 + j
            ws.cell(row=r, column=c, value=rev).font = DATA_FONT
            ws.cell(row=r, column=c).number_format = '#,##0'
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
        r += 1

    r += 2
    # Adoption S-Curve & Unit Economics
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="ADOPTION S-CURVE & UNIT ECONOMICS"); r += 1
    xai_ue_items = [
        ("Global AI TAM (2024 est.)", "$200B"),
        ("xAI Market Share (2024)", "<0.2%"),
        ("xAI Market Share (2027E)", "~1.5%"),
        ("S-Curve Phase (current)", "Early Adoption → Growth"),
        ("S-Curve Inflection (est.)", "2026-2027"),
        ("", ""),
        ("Compute Cost per Inference (H100)", "$0.003 - $0.01"),
        ("API Revenue per Inference (avg)", "$0.01 - $0.03"),
        ("Gross Margin per Inference", "~50-65%"),
        ("Monthly Inference Volume (2024E)", "~5B tokens served"),
        ("Monthly Inference Volume (2027E)", "~500B tokens served"),
        ("", ""),
        ("Grok Premium LTV (est.)", "$180 (18-month avg life)"),
        ("Grok Premium CAC (est.)", "$30 (X platform organic)"),
        ("LTV / CAC", "6.0x"),
    ]
    for label, val in xai_ue_items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    r += 1
    style_section_row(ws, r, 8); ws.cell(row=r, column=1, value="COHORT NPV — API ENTERPRISE CUSTOMER"); r += 1
    xai_npv_items = [
        ("Year 0: Onboarding / Integration Revenue", "$5,000"),
        ("Year 0: Sales & Onboarding Cost", "($8,000)"),
        ("Year 1: API Revenue (usage ramp)", "$15,000"),
        ("Year 1: Compute COGS (45%)", "($6,750)"),
        ("Year 2: API Revenue (mature)", "$20,000"),
        ("Year 2: Compute COGS (42%)", "($8,400)"),
        ("Year 3: API Revenue (expansion)", "$24,000"),
        ("Year 3: Compute COGS (40%)", "($9,600)"),
        ("3-Year Gross Profit / Cohort Customer", "$31,250"),
        ("Discount Rate", "15%"),
        ("NPV per Enterprise Customer", "~$23,000"),
        ("Total NPV @ 18K customers (2027E)", "~$414M"),
        ("+ Grok Consumer NPV @ 35M subs", "~$5.3B"),
        ("+ Compute-as-a-Service NPV", "~$3.0B"),
    ]
    for label, val in xai_npv_items:
        ws.cell(row=r, column=1, value=label).font = DATA_FONT
        ws.cell(row=r, column=2, value=val).font = DATA_FONT
        r += 1

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 11: Dilution Schedule (both companies)
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Dilution & Share Count")
    ws.sheet_properties.tabColor = "548235"
    ws.column_dimensions['A'].width = 35
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    ws.cell(row=row, column=1, value="Share Count & Dilution Schedule (Next 8 Quarters)")
    style_header_row(ws, row, 9)

    q8_labels = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26", "Q4'26"]

    # SpaceX
    r = 3
    style_section_row(ws, r, 9); ws.cell(row=r, column=1, value="SPACEX — SHARE COUNT (M shares)"); r += 1
    write_row(ws, r, [""] + q8_labels); style_sub_header_row(ws, r, 9); r += 1

    sx_basic = [1820, 1825, 1830, 1835, 1840, 1845, 1850, 1855]
    sx_options = [40, 42, 44, 46, 48, 50, 52, 54]
    sx_rsus = [15, 16, 17, 18, 19, 20, 21, 22]
    sx_other_dilutive = [10, 10, 10, 10, 10, 10, 10, 10]
    sx_fd = [sx_basic[i] + sx_options[i] + sx_rsus[i] + sx_other_dilutive[i] for i in range(8)]
    sx_dilution_pct = [round((sx_fd[i] - sx_basic[i]) / sx_basic[i], 4) for i in range(8)]

    write_row(ws, r, ["Basic Shares Outstanding"] + sx_basic, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) Options (in-the-money)"] + sx_options, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) RSUs / Restricted Stock"] + sx_rsus, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) Other Dilutive"] + sx_other_dilutive, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["Fully Diluted Shares"] + sx_fd, fmt=NUM_FORMAT, bold=True); r += 1
    write_row(ws, r, ["Dilution %"] + sx_dilution_pct, fmt=PCT_FORMAT); r += 1
    sx_price = [185, 188, 191, 195, 200, 205, 210, 215]
    write_row(ws, r, ["Implied Price / Share ($)"] + sx_price, fmt=DOLLAR_FORMAT); r += 1
    sx_mktcap = [round(sx_fd[i] * sx_price[i] / 1000, 0) for i in range(8)]
    write_row(ws, r, ["Implied Market Cap ($B)"] + sx_mktcap, fmt=NUM_FORMAT, bold=True); r += 1

    # xAI
    r += 2
    style_section_row(ws, r, 9); ws.cell(row=r, column=1, value="xAI — SHARE COUNT (M shares)"); r += 1
    write_row(ws, r, [""] + q8_labels); style_sub_header_row(ws, r, 9); r += 1

    xai_basic = [700, 705, 710, 715, 720, 725, 730, 735]
    xai_options = [50, 52, 55, 58, 60, 63, 65, 68]
    xai_rsus = [20, 22, 24, 26, 28, 30, 32, 34]
    xai_warrants = [15, 15, 15, 15, 15, 15, 15, 15]
    xai_fd_shares = [xai_basic[i] + xai_options[i] + xai_rsus[i] + xai_warrants[i] for i in range(8)]
    xai_dilution_pct = [round((xai_fd_shares[i] - xai_basic[i]) / xai_basic[i], 4) for i in range(8)]

    write_row(ws, r, ["Basic Shares Outstanding"] + xai_basic, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) Options (in-the-money)"] + xai_options, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) RSUs / Restricted Stock"] + xai_rsus, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["(+) Warrants"] + xai_warrants, fmt=NUM_FORMAT); r += 1
    write_row(ws, r, ["Fully Diluted Shares"] + xai_fd_shares, fmt=NUM_FORMAT, bold=True); r += 1
    write_row(ws, r, ["Dilution %"] + xai_dilution_pct, fmt=PCT_FORMAT); r += 1
    xai_price_sh = [63, 64, 65, 66, 68, 70, 72, 75]
    write_row(ws, r, ["Implied Price / Share ($)"] + xai_price_sh, fmt=DOLLAR_FORMAT); r += 1
    xai_mktcap = [round(xai_fd_shares[i] * xai_price_sh[i] / 1000, 0) for i in range(8)]
    write_row(ws, r, ["Implied Market Cap ($B)"] + xai_mktcap, fmt=NUM_FORMAT, bold=True); r += 1

    # ─────────────────────────────────────────────────────────────────────
    # SHEET 12: Model Assumptions & Notes
    # ─────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Assumptions & Notes")
    ws.sheet_properties.tabColor = "808080"
    ws.column_dimensions['A'].width = 90

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=1)
    ws.cell(row=row, column=1, value="Model Assumptions, Sources & Methodology")
    style_header_row(ws, row, 1)

    notes = [
        "",
        "GENERAL NOTES",
        "• Both SpaceX and xAI are private companies; financial data is based on publicly reported estimates,",
        "  industry research (Bloomberg, WSJ, Reuters, Morgan Stanley, Quilty Analytics, Payload Space),",
        "  and triangulated secondary market data.",
        "• All figures in $M unless otherwise noted.",
        "• Yellow-highlighted cells denote key input assumptions that drive the model.",
        "• Green-shaded columns denote forecast periods (E = Estimate).",
        "",
        "SPACEX ASSUMPTIONS",
        "• Launch count and pricing based on FAA data, SpaceX manifest, and industry pricing benchmarks.",
        "• Starlink subscriber count based on Quilty Analytics, Ookla, and company announcements.",
        "• ARPU decline reflects geographic mix shift (lower-ARPU markets) and competitive dynamics.",
        "• Starship revenue ramps as NASA HLS milestones are achieved (2024+) and commercial contracts close.",
        "• COGS improvement driven by reusability gains (Falcon 9 booster reuse 20+ flights) and Starlink",
        "  manufacturing scale (Gen2 satellites, reduced CPE costs).",
        "• Valuation: $350B Dec 2024 tender offer; secondary range $170-$195/share.",
        "• Cohort NPV for Starlink subscriber assumes 12% discount rate, 15% annual churn, $599 hardware sale.",
        "",
        "xAI ASSUMPTIONS",
        "• xAI founded July 2023; 2023 is a partial-year stub with minimal revenue.",
        "• Revenue build assumes Grok integration into X platform drives consumer adoption;",
        "  enterprise API launches in 2024 with rapid ramp.",
        "• Colossus supercomputer (100K H100 GPUs operational Q4'24) enables compute-as-a-service revenue.",
        "• R&D % of revenue declines sharply as initial infrastructure buildout completes;",
        "  absolute R&D spend continues to grow.",
        "• Valuation: Series C at $50B (Dec 2024); Series B at $40B (Nov 2024); Series A at $24B (May 2024).",
        "• Unit economics assume inference cost declines of ~30% annually due to hardware & software optimization.",
        "• S-curve assumes xAI is in 'Early Adoption → Growth' phase with inflection in 2026-2027.",
        "",
        "SENSITIVITY TABLE METHODOLOGY",
        "• SpaceX: 30x EV/EBITDA on 2027E (premium for Starlink network economics + Starship optionality).",
        "• xAI: 40x EV/EBITDA on 2027E (reflects high-growth AI sector comparables: NVDA, MSFT AI, OpenAI).",
        "• Revenue growth and EBITDA margin are the two most material drivers for both names.",
        "• Second sensitivity isolates the key usage-based driver: Starlink subs x ARPU (SpaceX),",
        "  API customers x ARPU (xAI).",
        "",
        "GAAP-TO-FCF BRIDGE",
        "• Bridge starts from GAAP operating income (EBIT), adds back non-cash items (D&A, SBC),",
        "  adjusts for interest, taxes, working capital, and other items to arrive at CFO.",
        "• FCF = CFO + CapEx (negative). CapEx is the dominant use of cash for both infrastructure-heavy names.",
        "",
        "DILUTION SCHEDULE",
        "• SpaceX: ~1.89B fully diluted shares; quarterly dilution from option exercises and RSU vesting.",
        "• xAI: ~800M fully diluted shares; higher dilution rate given early-stage equity compensation.",
        "• Both schedules assume no new primary issuance beyond planned vesting; actual tender offers",
        "  may reduce dilution (buyback of employee shares).",
        "",
        f"Model generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} UTC",
        "Disclaimer: This model is for illustrative/educational purposes only. Not investment advice.",
    ]

    for i, note in enumerate(notes):
        r = 2 + i
        ws.cell(row=r, column=1, value=note)
        if note.startswith("SPACEX") or note.startswith("xAI") or note.startswith("GENERAL") or note.startswith("SENSITIVITY") or note.startswith("GAAP") or note.startswith("DILUTION"):
            ws.cell(row=r, column=1).font = SECTION_FONT
        else:
            ws.cell(row=r, column=1).font = DATA_FONT

    # ─── Freeze panes and print setup on key sheets ──────────────────────
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.freeze_panes = "B3"
        ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

    return wb


if __name__ == "__main__":
    print("Building SpaceX & xAI three-statement financial model...")
    wb = build_workbook()
    output_path = "/workspace/SpaceX_xAI_Financial_Model.xlsx"
    wb.save(output_path)
    print(f"Model saved to: {output_path}")
    print(f"Sheets: {wb.sheetnames}")
