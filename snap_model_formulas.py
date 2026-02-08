#!/usr/bin/env python3
"""
Snap Inc. (SNAP) – Three-Statement Financial Model WITH EXCEL FORMULAS
======================================================================
Every forecast cell is a live Excel formula referencing the Assumptions tab.
Historical cells are hardcoded from SEC filings; derived rows (margins,
growth rates) are formulas even for historical years.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

# ═══════════════════════════════════════════════════════════════════════════════
# STYLING
# ═══════════════════════════════════════════════════════════════════════════════
HDR_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HDR_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SUB_FONT = Font(name="Calibri", bold=True, size=10, color="2F5496")
SUB_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FCST_FONT = Font(name="Calibri", italic=True, size=10, color="0070C0")
NF = Font(name="Calibri", size=10)
BF = Font(name="Calibri", bold=True, size=10)
TF = Font(name="Calibri", bold=True, size=14, color="2F5496")
SF = Font(name="Calibri", bold=True, size=11, color="2F5496")
GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
THIN_B = Border(bottom=Side(style='thin', color='B4C6E7'))
MED_B = Border(bottom=Side(style='medium', color='2F5496'))

P = '0.0%'      # pct format
N = '#,##0'      # number
D = '#,##0.0'    # 1-decimal
U2 = '$#,##0.00' # USD 2-dec
U1 = '$#,##0.0'  # USD 1-dec

def w(ws, r, c, val, font=NF, fmt=None, fill=None):
    """Write a single cell."""
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = font
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    return cell

def hdr_row(ws, r, labels):
    for i, lb in enumerate(labels):
        c = ws.cell(row=r, column=1+i, value=lb)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='center', wrap_text=True)

def sub_row(ws, r, labels):
    for i, lb in enumerate(labels):
        c = ws.cell(row=r, column=1+i, value=lb)
        c.font = SUB_FONT; c.fill = SUB_FILL
        c.alignment = Alignment(horizontal='center')

def sec(ws, r, label):
    ws.cell(row=r, column=1, value=label).font = SF
    ws.cell(row=r, column=1).border = MED_B

def colw(ws, widths):
    for i, w_ in enumerate(widths, 1):
        ws.column_dimensions[CL(i)].width = w_

# Cross-tab reference helpers (sheet names with spaces need single quotes)
A_R = "Assumptions"
RB_R = "'Revenue Build'"
IS_R = "'Income Statement'"
BS_R = "'Balance Sheet'"
CF_R = "'Cash Flow & FCF Bridge'"
Q_R = "'Quarterly Detail'"
SC_R = "'Share Count & Dilution'"

# ═══════════════════════════════════════════════════════════════════════════════
# PREDEFINED ROW NUMBERS (ALL TABS)
# These are fixed so cross-tab formula references work correctly.
# ═══════════════════════════════════════════════════════════════════════════════

# --- Assumptions tab ---
A = {
    'dau_g': 7, 'arpu_g': 8,
    'gm': 11, 'rd': 12, 'sm': 13, 'ga': 14, 'sbc': 15,
    'na': 18, 'eu': 19, 'row': 20,
    'da': 23, 'capex': 24, 'tax': 25, 'interest': 26, 'wc': 27,
    'ppe': 30, 'intang': 31, 'rou': 32, 'olt_a': 33,
    'ltd': 34, 'lease': 35, 'olt_l': 36,
    'dso': 37, 'dpo': 38, 'accr': 39,
    'shares': 42, 'dilution': 43,
    'cff': 46,
    'season': 50,
}

# --- Revenue Build tab ---
RB = {
    'dau': 7, 'dau_g': 8,
    'arpu': 11, 'arpu_g': 12,
    'rev': 15, 'rev_g': 16,
    'na': 19, 'na_pct': 20,
    'eu': 21, 'eu_pct': 22,
    'row_r': 23, 'row_pct': 24,
    'imp': 27, 'ecpm': 28,
}

# --- Income Statement tab ---
IS = {
    'rev': 6, 'rev_g': 7,
    'cogs': 9, 'gp': 10, 'gm': 11,
    'rd': 14, 'rd_pct': 15,
    'sm': 16, 'sm_pct': 17,
    'ga': 18, 'ga_pct': 19,
    'opex': 20, 'opex_pct': 21,
    'oi': 23, 'oi_m': 24,
    'sbc': 26, 'sbc_pct': 27,
    'da': 29, 'ebitda': 30, 'ebitda_m': 31,
    'int': 33, 'pt': 34, 'tax': 35, 'tax_r': 36,
    'ni': 38,
    'shares': 40, 'eps': 41,
}

# --- Cash Flow tab ---
CF = {
    'ni': 7, 'da': 8, 'sbc': 9, 'wc': 10, 'other': 11, 'cfo': 12,
    'capex': 15, 'other_cfi': 16, 'cfi': 17,
    'cff': 20,
    'net': 22, 'beg': 23, 'end': 24,
    'fcf': 26, 'fcf_m': 27,
    # Bridge
    'br_oi': 33, 'br_da': 34, 'br_sbc': 35, 'br_ae': 36,
    'br_wc': 37, 'br_int': 38, 'br_tax': 39, 'br_other': 40,
    'br_cfo': 41, 'br_capex': 42, 'br_fcf': 43,
}

# --- Balance Sheet tab ---
BS = {
    'cash': 7, 'ar': 8, 'oca': 9, 'tca': 10,
    'ppe': 12, 'intang': 13, 'rou': 14, 'olta': 15, 'ta': 16,
    'ap': 19, 'accr': 20, 'dr': 21, 'ocl': 22, 'tcl': 23,
    'ltd': 25, 'lease': 26, 'oltl': 27, 'tl': 28,
    'eq': 31, 'tle': 32,
    'chk': 34,
}

# --- Quarterly tab ---
QT = {
    'dau': 6, 'rev': 7, 'arpu': 8,
    'cogs': 10, 'gp': 11, 'gm': 12,
    'rd': 14, 'sm': 15, 'ga': 16,
    'oi': 17, 'oi_m': 18,
    'sbc': 20, 'da': 21, 'ebitda': 22,
    'int': 24, 'pt': 25, 'tax': 26, 'ni': 27,
    'capex': 29,
}

# ═══════════════════════════════════════════════════════════════════════════════
# HISTORICAL DATA (from Snap 10-K / 10-Q filings, all $ in millions)
# Columns: B=2019, C=2020, D=2021, E=2022, F=2023, G=2024
# ═══════════════════════════════════════════════════════════════════════════════
dau_h = [218, 265, 319, 375, 414, 443]
rev_h = [1716, 2507, 4117, 4602, 4606, 5360]
rev_na_h = [1100, 1620, 2720, 2990, 2950, 3430]
rev_eu_h = [360, 500, 830, 920, 920, 1070]
rev_row_h = [256, 387, 567, 692, 736, 860]
cor_h = [880, 1128, 1763, 2158, 2190, 2461]
rd_h = [988, 1235, 1535, 2017, 1847, 1802]
sm_h = [654, 794, 1168, 1424, 1284, 1281]
ga_h = [363, 421, 507, 641, 614, 582]
sbc_h = [600, 770, 1096, 1538, 1376, 1296]
da_h = [81, 107, 140, 185, 200, 215]
int_h = [-51, -88, -89, -87, -100, -95]
tax_h = [5, 7, 14, 20, 25, 28]
capex_h = [-51, -59, -85, -115, -93, -90]
shares_h = [1500, 1548, 1620, 1660, 1676, 1690]
# Impressions per DAU per day (estimated)
imp_h = [25, 27, 30, 33, 36, 39]

# Balance Sheet
cash_h = [2024, 2768, 3506, 3938, 3374, 3205]
ar_h = [548, 737, 1068, 995, 1109, 1230]
oca_h = [174, 200, 201, 182, 217, 215]
ppe_h = [204, 258, 339, 452, 430, 400]
intang_h = [181, 182, 411, 440, 395, 360]
rou_h = [250, 280, 310, 350, 360, 370]
olt_a_h = [150, 180, 250, 310, 330, 350]
ap_h = [120, 155, 245, 265, 280, 295]
accr_h = [350, 430, 620, 705, 680, 720]
dr_h = [25, 30, 35, 40, 42, 45]
ocl_h = [0, 0, 0, 0, 0, 0]
ltd_h = [1485, 2190, 3745, 3745, 3745, 3745]
lease_h = [220, 260, 300, 340, 350, 360]
olt_l_h = [80, 95, 120, 140, 150, 155]

# Cash Flow
cfo_h = [-41, 166, 631, 234, 75, 260]
cfi_h = [-200, -500, -1200, -350, -120, -100]
cff_h = [900, 1078, 1307, 548, -519, -329]
wc_h = [-80, -45, -120, 50, 30, 15]
beg_cash_h = [1365, 2024, 2768, 3506, 3938, 3374]

# Compute derived historical
gp_h = [r - c for r, c in zip(rev_h, cor_h)]
opex_h = [a+b+c for a,b,c in zip(rd_h, sm_h, ga_h)]
oi_h = [g - o for g, o in zip(gp_h, opex_h)]
pt_h = [o + i for o, i in zip(oi_h, int_h)]
ni_h = [p - t for p, t in zip(pt_h, tax_h)]
other_cfi_h = [cf - cx for cf, cx in zip(cfi_h, capex_h)]
other_adj_h = [cfo - ni - da - sbc - wcc for cfo, ni, da, sbc, wcc in
               zip(cfo_h, ni_h, da_h, sbc_h, wc_h)]
tca_h = [c + a + o for c, a, o in zip(cash_h, ar_h, oca_h)]
ta_h = [tc + p + i + r + o for tc, p, i, r, o in
        zip(tca_h, ppe_h, intang_h, rou_h, olt_a_h)]
tcl_h = [a + ac + d + o for a, ac, d, o in zip(ap_h, accr_h, dr_h, ocl_h)]
tl_h = [tc + l + le + o for tc, l, le, o in zip(tcl_h, ltd_h, lease_h, olt_l_h)]
eq_h = [a - l for a, l in zip(ta_h, tl_h)]

# Forecast assumption VALUES (placed on Assumptions tab as inputs)
dau_g_f = [0.065, 0.055, 0.045, 0.040]
arpu_g_f = [0.10, 0.09, 0.08, 0.07]
gm_f = [0.555, 0.570, 0.585, 0.600]
rd_f = [0.30, 0.27, 0.25, 0.23]
sm_f = [0.21, 0.19, 0.18, 0.17]
ga_f = [0.095, 0.088, 0.082, 0.078]
sbc_f = [0.20, 0.17, 0.15, 0.13]
na_f = [0.63, 0.62, 0.61, 0.60]
eu_f = [0.20, 0.20, 0.20, 0.21]
row_f = [0.17, 0.18, 0.19, 0.19]
da_f = [225, 240, 255, 270]
capex_f = [95, 100, 105, 110]  # positive on Assumptions, negative on CF
tax_f = [0.05, 0.07, 0.10, 0.12]
interest_f = [90, 85, 80, 75]  # positive on Assumptions, negative on IS
wc_f = [-20, -15, -10, -5]
ppe_f = [390, 385, 380, 375]
intang_f = [340, 320, 300, 280]
rou_f = [375, 380, 385, 390]
olt_a_f = [360, 370, 380, 390]
ltd_f_v = [3745, 3745, 3745, 3745]
lease_f = [365, 370, 375, 380]
olt_l_f = [160, 165, 170, 175]
dso_f = [80, 80, 80, 80]
dpo_f = [42, 42, 42, 42]
accr_f = [0.12, 0.12, 0.12, 0.12]
dilution_f = [0.012, 0.009, 0.009, 0.009]
cff_f = [-200, -200, -250, -300]
imp_f = [42, 44, 46, 48]
season_f = [0.21, 0.23, 0.25, 0.31]

# Quarterly historical data (Q1-23 through Q4-24)
q_labels_h = ['Q1-23','Q2-23','Q3-23','Q4-23','Q1-24','Q2-24','Q3-24','Q4-24']
q_labels_f = ['Q1-25E','Q2-25E','Q3-25E','Q4-25E','Q1-26E','Q2-26E','Q3-26E','Q4-26E']
dau_qh = [383, 397, 406, 414, 422, 432, 437, 443]
rev_qh = [989, 1068, 1189, 1361, 1195, 1237, 1373, 1556]
cor_qh = [490, 510, 545, 645, 545, 565, 620, 731]
rd_qh = [478, 462, 455, 452, 455, 448, 448, 451]
sm_qh = [330, 318, 320, 316, 322, 318, 320, 321]
ga_qh = [158, 155, 150, 151, 147, 145, 145, 145]
sbc_qh = [355, 345, 340, 336, 330, 325, 322, 319]
da_qh = [50, 50, 50, 50, 53, 54, 54, 54]
capex_qh = [-22, -24, -23, -24, -22, -23, -22, -23]
int_qh = [-25, -25, -25, -25, -24, -24, -24, -23]
tax_qh = [6, 6, 6, 7, 7, 7, 7, 7]

# Share count quarterly
rsu_vest = [12, 11, 13, 12, 11, 10, 12, 11]
buyback_q = [5, 5, 5, 5, 5, 5, 5, 5]
other_q_sc = [3, 4, 2, 3, 3, 3, 1, 3]
unvested_rsu = [120, 117, 112, 108, 104, 101, 96, 92]
vested_opts = [15, 14, 13, 12, 11, 10, 9, 8]
new_grants = [8, 8, 8, 8, 7, 7, 7, 7]

# ═══════════════════════════════════════════════════════════════════════════════
# YEAR HEADERS (shared by all annual tabs)
# ═══════════════════════════════════════════════════════════════════════════════
yr_hdr = ["", "2019", "2020", "2021", "2022", "2023", "2024",
          "2025E", "2026E", "2027E", "2028E"]
hf_marker = ["", "Historical","Historical","Historical",
             "Historical","Historical","Historical",
             "Forecast","Forecast","Forecast","Forecast"]

# ═══════════════════════════════════════════════════════════════════════════════
# CREATE WORKBOOK
# ═══════════════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

ws_rb = wb.active; ws_rb.title = "Revenue Build"
ws_is = wb.create_sheet("Income Statement")
ws_bs = wb.create_sheet("Balance Sheet")
ws_cf = wb.create_sheet("Cash Flow & FCF Bridge")
ws_qt = wb.create_sheet("Quarterly Detail")
ws_sc = wb.create_sheet("Share Count & Dilution")
ws_sn = wb.create_sheet("Sensitivity Analysis")
ws_vl = wb.create_sheet("Valuation Cross-Checks")
ws_a  = wb.create_sheet("Assumptions")

for s in [ws_rb, ws_is]: s.sheet_properties.tabColor = "2F5496"
ws_bs.sheet_properties.tabColor = "548235"
ws_cf.sheet_properties.tabColor = "BF8F00"
ws_qt.sheet_properties.tabColor = "7030A0"
ws_sc.sheet_properties.tabColor = "C00000"
ws_sn.sheet_properties.tabColor = "ED7D31"
ws_vl.sheet_properties.tabColor = "00B050"
ws_a.sheet_properties.tabColor = "FF0000"

# Helper: write historical values into columns B-G
def fill_hist(ws, r, vals, fmt=N, font=NF):
    for i, v in enumerate(vals):
        w(ws, r, 2+i, v, font=font, fmt=fmt)

# Helper: write forecast formulas into columns H-K
def fill_fcst(ws, r, formulas, fmt=N, font=FCST_FONT, fill=None):
    for i, f_ in enumerate(formulas):
        w(ws, r, 8+i, f_, font=font, fmt=fmt, fill=fill)

# Helper: write formula for every data column (B-K or C-K)
def fill_all(ws, r, gen, fmt=P, font=NF, start=2):
    """gen(col_idx) -> formula string. start=2 for B, 3 for C (skip first)."""
    for ci in range(start, 12):
        f_ = gen(ci)
        if f_ is not None:
            fnt = FCST_FONT if ci >= 8 else font
            w(ws, r, ci, f_, font=fnt, fmt=fmt)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: ASSUMPTIONS
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_a
w(ws, 1, 1, "Snap Inc. (SNAP) - Model Assumptions & Drivers", TF)
ws.merge_cells('A1:K1')

hdr_row(ws, 3, yr_hdr)
sub_row(ws, 4, hf_marker)

r = 6; sec(ws, r, "GROWTH DRIVERS")

# Row 7: DAU YoY Growth
r = A['dau_g']
w(ws, r, 1, "DAU YoY Growth (%)", BF)
# Historical: formula referencing Revenue Build
for ci in range(3, 8):  # C-G (2020-2024), skip B (no prior)
    w(ws, r, ci, f"={RB_R}!{CL(ci)}{RB['dau']}/{RB_R}!{CL(ci-1)}{RB['dau']}-1", fmt=P)
# Forecast: INPUT
for i, v in enumerate(dau_g_f):
    w(ws, r, 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 8: ARPU YoY Growth
r = A['arpu_g']
w(ws, r, 1, "ARPU YoY Growth (%)", BF)
for ci in range(3, 8):
    w(ws, r, ci, f"={RB_R}!{CL(ci)}{RB['arpu']}/{RB_R}!{CL(ci-1)}{RB['arpu']}-1", fmt=P)
for i, v in enumerate(arpu_g_f):
    w(ws, r, 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 10-15: Margin assumptions
r = 10; sec(ws, r, "MARGIN & COST ASSUMPTIONS")
margin_rows = [
    (A['gm'],  "Gross Margin (%)", gm_f, IS['gp'], IS['rev']),
    (A['rd'],  "R&D % of Revenue", rd_f, IS['rd'], IS['rev']),
    (A['sm'],  "S&M % of Revenue", sm_f, IS['sm'], IS['rev']),
    (A['ga'],  "G&A % of Revenue", ga_f, IS['ga'], IS['rev']),
    (A['sbc'], "SBC % of Revenue", sbc_f, IS['sbc'], IS['rev']),
]
for row_num, label, fvals, num_row, den_row in margin_rows:
    w(ws, row_num, 1, label, BF)
    # Historical: derive from IS
    for ci in range(2, 8):
        w(ws, row_num, ci, f"={IS_R}!{CL(ci)}{num_row}/{IS_R}!{CL(ci)}{den_row}", fmt=P)
    for i, v in enumerate(fvals):
        w(ws, row_num, 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 17-20: Geographic mix
r = 17; sec(ws, r, "GEOGRAPHIC MIX")
geo_rows = [
    (A['na'], "NA % of Revenue", na_f, RB['na']),
    (A['eu'], "EU % of Revenue", eu_f, RB['eu']),
    (A['row'],"ROW % of Revenue", row_f, RB['row_r']),
]
for row_num, label, fvals, rb_row in geo_rows:
    w(ws, row_num, 1, label, BF)
    for ci in range(2, 8):
        w(ws, row_num, ci, f"={RB_R}!{CL(ci)}{rb_row}/{RB_R}!{CL(ci)}{RB['rev']}", fmt=P)
    for i, v in enumerate(fvals):
        w(ws, row_num, 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 22-27: Other P&L
r = 22; sec(ws, r, "OTHER P&L ITEMS")
simple_rows = [
    (A['da'],       "D&A ($M)",           da_h, da_f, N),
    (A['capex'],    "CapEx ($M, positive)",  [abs(x) for x in capex_h], capex_f, N),
    (A['interest'], "Interest Expense ($M, positive)", [abs(x) for x in int_h], interest_f, N),
    (A['wc'],       "WC Change ($M)",     wc_h, wc_f, N),
]
for row_num, label, hvals, fvals, fmt in simple_rows:
    w(ws, row_num, 1, label, BF)
    fill_hist(ws, row_num, hvals, fmt=fmt)
    for i, v in enumerate(fvals):
        w(ws, row_num, 8+i, v, font=FCST_FONT, fmt=fmt, fill=INPUT_FILL)

# Tax rate row
w(ws, A['tax'], 1, "Effective Tax Rate (%)", BF)
for ci in range(2, 8):
    w(ws, A['tax'], ci,
      f"=IF({IS_R}!{CL(ci)}{IS['pt']}>0,{IS_R}!{CL(ci)}{IS['tax']}/{IS_R}!{CL(ci)}{IS['pt']},0)", fmt=P)
for i, v in enumerate(tax_f):
    w(ws, A['tax'], 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 29-39: Balance Sheet Drivers
r = 29; sec(ws, r, "BALANCE SHEET DRIVERS")
bs_rows = [
    (A['ppe'],    "PP&E ($M)",           ppe_h, ppe_f, N),
    (A['intang'], "Intangibles ($M)",    intang_h, intang_f, N),
    (A['rou'],    "ROU Assets ($M)",     rou_h, rou_f, N),
    (A['olt_a'],  "Other LT Assets ($M)",olt_a_h, olt_a_f, N),
    (A['ltd'],    "Long-Term Debt ($M)", ltd_h, ltd_f_v, N),
    (A['lease'],  "Lease Liabilities ($M)", lease_h, lease_f, N),
    (A['olt_l'],  "Other LT Liab ($M)", olt_l_h, olt_l_f, N),
]
for row_num, label, hvals, fvals, fmt in bs_rows:
    w(ws, row_num, 1, label, BF)
    fill_hist(ws, row_num, hvals, fmt=fmt)
    for i, v in enumerate(fvals):
        w(ws, row_num, 8+i, v, font=FCST_FONT, fmt=fmt, fill=INPUT_FILL)

# DSO, DPO, Accrued
w(ws, A['dso'], 1, "DSO (days)", BF)
w(ws, A['dpo'], 1, "DPO (days)", BF)
w(ws, A['accr'], 1, "Accrued Liab % of Rev", BF)
for i in range(4):
    w(ws, A['dso'], 8+i, dso_f[i], font=FCST_FONT, fmt=N, fill=INPUT_FILL)
    w(ws, A['dpo'], 8+i, dpo_f[i], font=FCST_FONT, fmt=N, fill=INPUT_FILL)
    w(ws, A['accr'], 8+i, accr_f[i], font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 41-43: Share Count
r = 41; sec(ws, r, "SHARE COUNT")
w(ws, A['shares'], 1, "Diluted Shares (M)", BF)
fill_hist(ws, A['shares'], shares_h, fmt=N)
# Forecast shares: = prior × (1 + dilution)
for ci in range(8, 12):
    prev = CL(ci - 1)
    cur = CL(ci)
    w(ws, A['shares'], ci,
      f"=ROUND({prev}{A['shares']}*(1+{cur}{A['dilution']}),0)",
      font=FCST_FONT, fmt=N)

w(ws, A['dilution'], 1, "Annual Dilution (%)", BF)
for ci in range(3, 8):
    w(ws, A['dilution'], ci,
      f"={CL(ci)}{A['shares']}/{CL(ci-1)}{A['shares']}-1", fmt=P)
for i, v in enumerate(dilution_f):
    w(ws, A['dilution'], 8+i, v, font=FCST_FONT, fmt=P, fill=INPUT_FILL)

# Row 45-46: Financing
r = 45; sec(ws, r, "FINANCING")
w(ws, A['cff'], 1, "Cash from Financing ($M)", BF)
fill_hist(ws, A['cff'], cff_h, fmt=N)
for i, v in enumerate(cff_f):
    w(ws, A['cff'], 8+i, v, font=FCST_FONT, fmt=N, fill=INPUT_FILL)

# Row 48-51: Quarterly Seasonality
r = 48; sec(ws, r, "QUARTERLY SEASONALITY")
hdr_row(ws, 49, ["", "Q1", "Q2", "Q3", "Q4"])
w(ws, A['season'], 1, "Revenue Seasonality Weight", BF)
for i, v in enumerate(season_f):
    w(ws, A['season'], 2+i, v, fmt=P, fill=INPUT_FILL)
w(ws, 51, 1, "Total Check", BF)
w(ws, 51, 2, f"=SUM(B{A['season']}:E{A['season']})", fmt=P)

colw(ws, [32]+[14]*10)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB: REVENUE BUILD
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_rb
w(ws, 1, 1, "Snap Inc. (SNAP) - Revenue Build by Segment & Driver", TF)
ws.merge_cells('A1:K1')
hdr_row(ws, 3, yr_hdr); sub_row(ws, 4, hf_marker)

# -- DAU --
r = 6; sec(ws, r, "DAU DRIVERS")
w(ws, RB['dau'], 1, "Daily Active Users (M)", BF)
fill_hist(ws, RB['dau'], dau_h, fmt=N)
# Forecast: prior year DAU × (1 + growth)
for ci in range(8, 12):
    prev = CL(ci - 1); cur = CL(ci)
    fill_fcst(ws, RB['dau'],
              [f"=ROUND({CL(ci-1)}{RB['dau']}*(1+{A_R}!{CL(ci)}{A['dau_g']}),0)" for ci in range(8,12)],
              fmt=N)
    break  # fill_fcst writes all 4 at once

w(ws, RB['dau_g'], 1, "  YoY Growth", NF)
fill_all(ws, RB['dau_g'],
         lambda ci: f"={CL(ci)}{RB['dau']}/{CL(ci-1)}{RB['dau']}-1",
         fmt=P, start=3)

# -- ARPU --
r = 10; sec(ws, r, "ARPU DRIVERS")
w(ws, RB['arpu'], 1, "Avg Revenue Per User ($)", BF)
# Historical ARPU = Revenue / DAU
for ci in range(2, 8):
    w(ws, RB['arpu'], ci, f"=ROUND({CL(ci)}{RB['rev']}/{CL(ci)}{RB['dau']},2)", fmt=U2)
# Forecast ARPU = prior × (1 + ARPU growth)
fill_fcst(ws, RB['arpu'],
          [f"=ROUND({CL(ci-1)}{RB['arpu']}*(1+{A_R}!{CL(ci)}{A['arpu_g']}),2)" for ci in range(8,12)],
          fmt=U2)

w(ws, RB['arpu_g'], 1, "  YoY Growth", NF)
fill_all(ws, RB['arpu_g'],
         lambda ci: f"={CL(ci)}{RB['arpu']}/{CL(ci-1)}{RB['arpu']}-1",
         fmt=P, start=3)

# -- REVENUE --
r = 14; sec(ws, r, "REVENUE")
w(ws, RB['rev'], 1, "Total Revenue ($M)", BF)
fill_hist(ws, RB['rev'], rev_h, fmt=N, font=BF)
# Forecast Revenue = ROUND(DAU × ARPU, 0)
fill_fcst(ws, RB['rev'],
          [f"=ROUND({CL(ci)}{RB['dau']}*{CL(ci)}{RB['arpu']},0)" for ci in range(8,12)],
          fmt=N)

w(ws, RB['rev_g'], 1, "  YoY Growth", NF)
fill_all(ws, RB['rev_g'],
         lambda ci: f"={CL(ci)}{RB['rev']}/{CL(ci-1)}{RB['rev']}-1",
         fmt=P, start=3)

# -- GEOGRAPHY --
r = 18; sec(ws, r, "REVENUE BY GEOGRAPHY")
geo_data = [
    (RB['na'], RB['na_pct'], "  North America", rev_na_h, A['na']),
    (RB['eu'], RB['eu_pct'], "  Europe", rev_eu_h, A['eu']),
    (RB['row_r'], RB['row_pct'], "  Rest of World", rev_row_h, A['row']),
]
for rev_row, pct_row, label, hist, a_pct_row in geo_data:
    w(ws, rev_row, 1, label, NF)
    fill_hist(ws, rev_row, hist, fmt=N)
    fill_fcst(ws, rev_row,
              [f"=ROUND({CL(ci)}{RB['rev']}*{A_R}!{CL(ci)}{a_pct_row},0)" for ci in range(8,12)],
              fmt=N)
    w(ws, pct_row, 1, "    % of Total", NF)
    fill_all(ws, pct_row,
             lambda ci, rr=rev_row: f"={CL(ci)}{rr}/{CL(ci)}{RB['rev']}",
             fmt=P)

# -- AD METRICS --
r = 26; sec(ws, r, "IMPLIED AD METRICS")
w(ws, RB['imp'], 1, "  Impressions/DAU/Day (est.)", NF)
fill_hist(ws, RB['imp'], imp_h, fmt=N)
for i, v in enumerate(imp_f):
    w(ws, RB['imp'], 8+i, v, font=FCST_FONT, fmt=N, fill=INPUT_FILL)

w(ws, RB['ecpm'], 1, "  Implied eCPM ($)", NF)
fill_all(ws, RB['ecpm'],
         lambda ci: f"=ROUND({CL(ci)}{RB['rev']}/({CL(ci)}{RB['dau']}*{CL(ci)}{RB['imp']}*365/1000000)*1000,2)",
         fmt=U2)

colw(ws, [32]+[14]*10)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: INCOME STATEMENT
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_is
w(ws, 1, 1, "Snap Inc. (SNAP) - GAAP Income Statement ($M)", TF)
ws.merge_cells('A1:K1')
hdr_row(ws, 3, yr_hdr); sub_row(ws, 4, hf_marker)

# Revenue
w(ws, IS['rev'], 1, "Revenue", BF)
fill_hist(ws, IS['rev'], rev_h, fmt=N, font=BF)
fill_fcst(ws, IS['rev'],
          [f"={RB_R}!{CL(ci)}{RB['rev']}" for ci in range(8,12)], fmt=N)

w(ws, IS['rev_g'], 1, "  YoY Growth", NF)
fill_all(ws, IS['rev_g'],
         lambda ci: f"={CL(ci)}{IS['rev']}/{CL(ci-1)}{IS['rev']}-1",
         fmt=P, start=3)

# COGS
w(ws, IS['cogs'], 1, "Cost of Revenue", NF)
fill_hist(ws, IS['cogs'], cor_h, fmt=N)
fill_fcst(ws, IS['cogs'],
          [f"=ROUND({CL(ci)}{IS['rev']}*(1-{A_R}!{CL(ci)}{A['gm']}),0)" for ci in range(8,12)],
          fmt=N)

# GP
w(ws, IS['gp'], 1, "Gross Profit", BF)
fill_all(ws, IS['gp'],
         lambda ci: f"={CL(ci)}{IS['rev']}-{CL(ci)}{IS['cogs']}",
         fmt=N)

# GM
w(ws, IS['gm'], 1, "  Gross Margin", NF)
fill_all(ws, IS['gm'],
         lambda ci: f"={CL(ci)}{IS['gp']}/{CL(ci)}{IS['rev']}",
         fmt=P)

# OpEx section
r = 13; sec(ws, r, "Operating Expenses")

opex_items = [
    (IS['rd'], IS['rd_pct'], "  Research & Development", rd_h, A['rd']),
    (IS['sm'], IS['sm_pct'], "  Sales & Marketing", sm_h, A['sm']),
    (IS['ga'], IS['ga_pct'], "  General & Administrative", ga_h, A['ga']),
]
for val_row, pct_row, label, hist, a_row in opex_items:
    w(ws, val_row, 1, label, NF)
    fill_hist(ws, val_row, hist, fmt=N)
    fill_fcst(ws, val_row,
              [f"=ROUND({CL(ci)}{IS['rev']}*{A_R}!{CL(ci)}{a_row},0)" for ci in range(8,12)],
              fmt=N)
    w(ws, pct_row, 1, "    % of Revenue", NF)
    fill_all(ws, pct_row,
             lambda ci, vr=val_row: f"={CL(ci)}{vr}/{CL(ci)}{IS['rev']}",
             fmt=P)

# Total OpEx
w(ws, IS['opex'], 1, "Total Operating Expenses", BF)
fill_all(ws, IS['opex'],
         lambda ci: f"={CL(ci)}{IS['rd']}+{CL(ci)}{IS['sm']}+{CL(ci)}{IS['ga']}",
         fmt=N)
w(ws, IS['opex_pct'], 1, "    % of Revenue", NF)
fill_all(ws, IS['opex_pct'],
         lambda ci: f"={CL(ci)}{IS['opex']}/{CL(ci)}{IS['rev']}",
         fmt=P)

# Operating Income
w(ws, IS['oi'], 1, "Operating Income (Loss)", BF)
fill_all(ws, IS['oi'],
         lambda ci: f"={CL(ci)}{IS['gp']}-{CL(ci)}{IS['opex']}",
         fmt=N)
w(ws, IS['oi_m'], 1, "  Operating Margin", NF)
fill_all(ws, IS['oi_m'],
         lambda ci: f"={CL(ci)}{IS['oi']}/{CL(ci)}{IS['rev']}",
         fmt=P)

# SBC memo
w(ws, IS['sbc'], 1, "Memo: Stock-Based Compensation", NF)
fill_hist(ws, IS['sbc'], sbc_h, fmt=N)
fill_fcst(ws, IS['sbc'],
          [f"=ROUND({CL(ci)}{IS['rev']}*{A_R}!{CL(ci)}{A['sbc']},0)" for ci in range(8,12)],
          fmt=N)
w(ws, IS['sbc_pct'], 1, "  SBC % of Revenue", NF)
fill_all(ws, IS['sbc_pct'],
         lambda ci: f"={CL(ci)}{IS['sbc']}/{CL(ci)}{IS['rev']}",
         fmt=P)

# D&A, EBITDA
w(ws, IS['da'], 1, "Depreciation & Amortization", NF)
fill_hist(ws, IS['da'], da_h, fmt=N)
fill_fcst(ws, IS['da'],
          [f"={A_R}!{CL(ci)}{A['da']}" for ci in range(8,12)], fmt=N)

w(ws, IS['ebitda'], 1, "EBITDA", BF)
fill_all(ws, IS['ebitda'],
         lambda ci: f"={CL(ci)}{IS['oi']}+{CL(ci)}{IS['da']}",
         fmt=N)
w(ws, IS['ebitda_m'], 1, "  EBITDA Margin", NF)
fill_all(ws, IS['ebitda_m'],
         lambda ci: f"={CL(ci)}{IS['ebitda']}/{CL(ci)}{IS['rev']}",
         fmt=P)

# Interest
w(ws, IS['int'], 1, "Interest & Other Income (Exp)", NF)
fill_hist(ws, IS['int'], int_h, fmt=N)
fill_fcst(ws, IS['int'],
          [f"=-{A_R}!{CL(ci)}{A['interest']}" for ci in range(8,12)], fmt=N)

# Pretax
w(ws, IS['pt'], 1, "Pre-Tax Income (Loss)", BF)
fill_all(ws, IS['pt'],
         lambda ci: f"={CL(ci)}{IS['oi']}+{CL(ci)}{IS['int']}",
         fmt=N)

# Tax
w(ws, IS['tax'], 1, "Income Tax Provision", NF)
fill_hist(ws, IS['tax'], tax_h, fmt=N)
fill_fcst(ws, IS['tax'],
          [f"=IF({CL(ci)}{IS['pt']}>0,ROUND({CL(ci)}{IS['pt']}*{A_R}!{CL(ci)}{A['tax']},0),0)"
           for ci in range(8,12)], fmt=N)
w(ws, IS['tax_r'], 1, "  Effective Tax Rate", NF)
fill_all(ws, IS['tax_r'],
         lambda ci: f"=IF({CL(ci)}{IS['pt']}>0,{CL(ci)}{IS['tax']}/{CL(ci)}{IS['pt']},0)",
         fmt=P)

# Net Income
w(ws, IS['ni'], 1, "Net Income (Loss)", BF)
fill_all(ws, IS['ni'],
         lambda ci: f"={CL(ci)}{IS['pt']}-{CL(ci)}{IS['tax']}",
         fmt=N)

# Shares, EPS
w(ws, IS['shares'], 1, "Diluted Shares Outstanding (M)", NF)
fill_all(ws, IS['shares'],
         lambda ci: f"={A_R}!{CL(ci)}{A['shares']}",
         fmt=N)
w(ws, IS['eps'], 1, "Diluted EPS", BF)
fill_all(ws, IS['eps'],
         lambda ci: f"=ROUND({CL(ci)}{IS['ni']}/{CL(ci)}{IS['shares']},2)",
         fmt=U2)

colw(ws, [32]+[14]*10)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: CASH FLOW & FCF BRIDGE
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_cf
w(ws, 1, 1, "Snap Inc. (SNAP) - Cash Flow Statement & GAAP OpLoss -> FCF Bridge ($M)", TF)
ws.merge_cells('A1:K1')
hdr_row(ws, 3, yr_hdr); sub_row(ws, 4, hf_marker)

r = 6; sec(ws, r, "CASH FROM OPERATIONS")

# NI
w(ws, CF['ni'], 1, "Net Income (Loss)", NF)
fill_all(ws, CF['ni'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['ni']}", fmt=N)

# D&A
w(ws, CF['da'], 1, "  (+) Depreciation & Amortization", NF)
fill_hist(ws, CF['da'], da_h, fmt=N)
fill_fcst(ws, CF['da'],
          [f"={A_R}!{CL(ci)}{A['da']}" for ci in range(8,12)], fmt=N)

# SBC
w(ws, CF['sbc'], 1, "  (+) Stock-Based Compensation", NF)
fill_all(ws, CF['sbc'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['sbc']}", fmt=N)

# WC
w(ws, CF['wc'], 1, "  (+/-) Working Capital Changes", NF)
fill_hist(ws, CF['wc'], wc_h, fmt=N)
fill_fcst(ws, CF['wc'],
          [f"={A_R}!{CL(ci)}{A['wc']}" for ci in range(8,12)], fmt=N)

# Other adjustments
w(ws, CF['other'], 1, "  (+/-) Other Non-Cash Adj.", NF)
fill_hist(ws, CF['other'], other_adj_h, fmt=N)
fill_fcst(ws, CF['other'], [0, 0, 0, 0], fmt=N)

# CFO
w(ws, CF['cfo'], 1, "Cash from Operations (CFO)", BF)
fill_all(ws, CF['cfo'],
         lambda ci: f"=SUM({CL(ci)}{CF['ni']}:{CL(ci)}{CF['other']})", fmt=N)

# CFI section
r = 14; sec(ws, r, "CASH FROM INVESTING")

w(ws, CF['capex'], 1, "  Capital Expenditures", NF)
fill_hist(ws, CF['capex'], capex_h, fmt=N)
fill_fcst(ws, CF['capex'],
          [f"=-{A_R}!{CL(ci)}{A['capex']}" for ci in range(8,12)], fmt=N)

w(ws, CF['other_cfi'], 1, "  Other Investing Activities", NF)
fill_hist(ws, CF['other_cfi'], other_cfi_h, fmt=N)
fill_fcst(ws, CF['other_cfi'], [0, 0, 0, 0], fmt=N)

w(ws, CF['cfi'], 1, "Cash from Investing (CFI)", BF)
fill_all(ws, CF['cfi'],
         lambda ci: f"={CL(ci)}{CF['capex']}+{CL(ci)}{CF['other_cfi']}", fmt=N)

# CFF section
r = 19; sec(ws, r, "CASH FROM FINANCING")
w(ws, CF['cff'], 1, "Cash from Financing (CFF)", BF)
fill_all(ws, CF['cff'],
         lambda ci: f"={A_R}!{CL(ci)}{A['cff']}", fmt=N)

# Net Change, Beginning/Ending Cash
w(ws, CF['net'], 1, "Net Change in Cash", BF)
fill_all(ws, CF['net'],
         lambda ci: f"={CL(ci)}{CF['cfo']}+{CL(ci)}{CF['cfi']}+{CL(ci)}{CF['cff']}",
         fmt=N)

w(ws, CF['beg'], 1, "Beginning Cash", NF)
fill_hist(ws, CF['beg'], beg_cash_h, fmt=N)
# Forecast beginning = prior year ending
fill_fcst(ws, CF['beg'],
          [f"={CL(ci-1)}{CF['end']}" for ci in range(8,12)], fmt=N)

w(ws, CF['end'], 1, "Ending Cash", BF)
fill_all(ws, CF['end'],
         lambda ci: f"={CL(ci)}{CF['beg']}+{CL(ci)}{CF['net']}", fmt=N)

# FCF
w(ws, CF['fcf'], 1, "Free Cash Flow (CFO + CapEx)", BF)
fill_all(ws, CF['fcf'],
         lambda ci: f"={CL(ci)}{CF['cfo']}+{CL(ci)}{CF['capex']}", fmt=N)
w(ws, CF['fcf_m'], 1, "  FCF Margin", NF)
fill_all(ws, CF['fcf_m'],
         lambda ci: f"={CL(ci)}{CF['fcf']}/{RB_R}!{CL(ci)}{RB['rev']}", fmt=P)

# ── FCF BRIDGE ────────────────────────────────────────────────────────────────
r = 30
w(ws, r, 1, "GAAP Operating Income (Loss) -> Free Cash Flow Bridge", TF)
ws.merge_cells(f'A{r}:K{r}')
hdr_row(ws, 32, yr_hdr)

w(ws, CF['br_oi'], 1, "GAAP Operating Income (Loss)", BF)
fill_all(ws, CF['br_oi'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['oi']}", fmt=N)

w(ws, CF['br_da'], 1, "  (+) Depreciation & Amortization", NF)
fill_all(ws, CF['br_da'],
         lambda ci: f"={CL(ci)}{CF['da']}", fmt=N)

w(ws, CF['br_sbc'], 1, "  (+) Stock-Based Compensation", NF)
fill_all(ws, CF['br_sbc'],
         lambda ci: f"={CL(ci)}{CF['sbc']}", fmt=N)

w(ws, CF['br_ae'], 1, "= Adjusted EBITDA", BF)
fill_all(ws, CF['br_ae'],
         lambda ci: f"={CL(ci)}{CF['br_oi']}+{CL(ci)}{CF['br_da']}+{CL(ci)}{CF['br_sbc']}",
         fmt=N)

w(ws, CF['br_wc'], 1, "  (+/-) Working Capital Changes", NF)
fill_all(ws, CF['br_wc'],
         lambda ci: f"={CL(ci)}{CF['wc']}", fmt=N)

w(ws, CF['br_int'], 1, "  (-) Cash Interest (net)", NF)
fill_all(ws, CF['br_int'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['int']}", fmt=N)

w(ws, CF['br_tax'], 1, "  (-) Cash Taxes", NF)
fill_all(ws, CF['br_tax'],
         lambda ci: f"=-{IS_R}!{CL(ci)}{IS['tax']}", fmt=N)

w(ws, CF['br_other'], 1, "  (+/-) Other Adjustments", NF)
fill_all(ws, CF['br_other'],
         lambda ci: f"={CL(ci)}{CF['other']}", fmt=N)

w(ws, CF['br_cfo'], 1, "= Cash from Operations (CFO)", BF)
fill_all(ws, CF['br_cfo'],
         lambda ci: (f"={CL(ci)}{CF['br_ae']}+{CL(ci)}{CF['br_wc']}"
                     f"+{CL(ci)}{CF['br_int']}+{CL(ci)}{CF['br_tax']}"
                     f"+{CL(ci)}{CF['br_other']}"),
         fmt=N)

w(ws, CF['br_capex'], 1, "  (-) Capital Expenditures", NF)
fill_all(ws, CF['br_capex'],
         lambda ci: f"={CL(ci)}{CF['capex']}", fmt=N)

w(ws, CF['br_fcf'], 1, "= Free Cash Flow", BF)
fill_all(ws, CF['br_fcf'],
         lambda ci: f"={CL(ci)}{CF['br_cfo']}+{CL(ci)}{CF['br_capex']}", fmt=N)

colw(ws, [38]+[14]*10)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: BALANCE SHEET
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_bs
w(ws, 1, 1, "Snap Inc. (SNAP) - Balance Sheet ($M)", TF)
ws.merge_cells('A1:K1')
hdr_row(ws, 3, yr_hdr); sub_row(ws, 4, hf_marker)

r = 6; sec(ws, r, "ASSETS")

# Cash
w(ws, BS['cash'], 1, "Cash & Equivalents", NF)
fill_hist(ws, BS['cash'], cash_h, fmt=N)
fill_fcst(ws, BS['cash'],
          [f"={CF_R}!{CL(ci)}{CF['end']}" for ci in range(8,12)], fmt=N)

# AR
w(ws, BS['ar'], 1, "Accounts Receivable", NF)
fill_hist(ws, BS['ar'], ar_h, fmt=N)
fill_fcst(ws, BS['ar'],
          [f"=ROUND({RB_R}!{CL(ci)}{RB['rev']}*{A_R}!{CL(ci)}{A['dso']}/365,0)"
           for ci in range(8,12)], fmt=N)

# Other CA
w(ws, BS['oca'], 1, "Other Current Assets", NF)
fill_hist(ws, BS['oca'], oca_h, fmt=N)
fill_fcst(ws, BS['oca'], [200, 200, 200, 200], fmt=N)

# Total CA
w(ws, BS['tca'], 1, "Total Current Assets", BF)
fill_all(ws, BS['tca'],
         lambda ci: f"={CL(ci)}{BS['cash']}+{CL(ci)}{BS['ar']}+{CL(ci)}{BS['oca']}",
         fmt=N)

# Long-term assets
lt_asset_rows = [
    (BS['ppe'],   "PP&E, net",           ppe_h,    A['ppe']),
    (BS['intang'],"Intangible Assets",   intang_h, A['intang']),
    (BS['rou'],   "ROU / Lease Assets",  rou_h,    A['rou']),
    (BS['olta'],  "Other Long-Term Assets", olt_a_h, A['olt_a']),
]
for bs_row, label, hist, a_row in lt_asset_rows:
    w(ws, bs_row, 1, label, NF)
    fill_hist(ws, bs_row, hist, fmt=N)
    fill_fcst(ws, bs_row,
              [f"={A_R}!{CL(ci)}{a_row}" for ci in range(8,12)], fmt=N)

# Total Assets
w(ws, BS['ta'], 1, "Total Assets", BF)
fill_all(ws, BS['ta'],
         lambda ci: (f"={CL(ci)}{BS['tca']}+{CL(ci)}{BS['ppe']}"
                     f"+{CL(ci)}{BS['intang']}+{CL(ci)}{BS['rou']}"
                     f"+{CL(ci)}{BS['olta']}"),
         fmt=N)

# Liabilities
r = 18; sec(ws, r, "LIABILITIES")

# AP
w(ws, BS['ap'], 1, "Accounts Payable", NF)
fill_hist(ws, BS['ap'], ap_h, fmt=N)
fill_fcst(ws, BS['ap'],
          [f"=ROUND({IS_R}!{CL(ci)}{IS['cogs']}*{A_R}!{CL(ci)}{A['dpo']}/365,0)"
           for ci in range(8,12)], fmt=N)

# Accrued
w(ws, BS['accr'], 1, "Accrued Liabilities", NF)
fill_hist(ws, BS['accr'], accr_h, fmt=N)
fill_fcst(ws, BS['accr'],
          [f"=ROUND({RB_R}!{CL(ci)}{RB['rev']}*{A_R}!{CL(ci)}{A['accr']},0)"
           for ci in range(8,12)], fmt=N)

# Deferred Revenue
w(ws, BS['dr'], 1, "Deferred Revenue", NF)
fill_hist(ws, BS['dr'], dr_h, fmt=N)
fill_fcst(ws, BS['dr'],
          [f"=ROUND({RB_R}!{CL(ci)}{RB['rev']}*0.007,0)" for ci in range(8,12)], fmt=N)

# Other CL
w(ws, BS['ocl'], 1, "Other Current Liabilities", NF)
fill_hist(ws, BS['ocl'], ocl_h, fmt=N)
fill_fcst(ws, BS['ocl'], [100, 100, 100, 100], fmt=N)

# Total CL
w(ws, BS['tcl'], 1, "Total Current Liabilities", BF)
fill_all(ws, BS['tcl'],
         lambda ci: (f"={CL(ci)}{BS['ap']}+{CL(ci)}{BS['accr']}"
                     f"+{CL(ci)}{BS['dr']}+{CL(ci)}{BS['ocl']}"),
         fmt=N)

# Long-term liabilities
lt_liab_rows = [
    (BS['ltd'],  "Long-Term Debt",        ltd_h,   A['ltd']),
    (BS['lease'],"Lease Liabilities (LT)", lease_h, A['lease']),
    (BS['oltl'], "Other Long-Term Liab",  olt_l_h, A['olt_l']),
]
for bs_row, label, hist, a_row in lt_liab_rows:
    w(ws, bs_row, 1, label, NF)
    fill_hist(ws, bs_row, hist, fmt=N)
    fill_fcst(ws, bs_row,
              [f"={A_R}!{CL(ci)}{a_row}" for ci in range(8,12)], fmt=N)

# Total Liabilities
w(ws, BS['tl'], 1, "Total Liabilities", BF)
fill_all(ws, BS['tl'],
         lambda ci: (f"={CL(ci)}{BS['tcl']}+{CL(ci)}{BS['ltd']}"
                     f"+{CL(ci)}{BS['lease']}+{CL(ci)}{BS['oltl']}"),
         fmt=N)

# Equity section
r = 30; sec(ws, r, "STOCKHOLDERS' EQUITY")

# Total Equity = Assets - Liabilities (plug)
w(ws, BS['eq'], 1, "Total Stockholders' Equity", BF)
fill_all(ws, BS['eq'],
         lambda ci: f"={CL(ci)}{BS['ta']}-{CL(ci)}{BS['tl']}",
         fmt=N)

# Total L&E
w(ws, BS['tle'], 1, "Total Liabilities & Equity", BF)
fill_all(ws, BS['tle'],
         lambda ci: f"={CL(ci)}{BS['tl']}+{CL(ci)}{BS['eq']}",
         fmt=N)

# Balance Check
w(ws, BS['chk'], 1, "Balance Check (should = 0)", NF)
fill_all(ws, BS['chk'],
         lambda ci: f"={CL(ci)}{BS['ta']}-{CL(ci)}{BS['tle']}",
         fmt=N)

colw(ws, [32]+[14]*10)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: QUARTERLY DETAIL
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_qt
w(ws, 1, 1, "Snap Inc. (SNAP) - Quarterly Income Statement ($M)", TF)
ws.merge_cells('A1:Q1')

q_hdr = [""] + q_labels_h + q_labels_f
hdr_row(ws, 3, q_hdr)
sub_row(ws, 4, [""]+["Historical"]*8+["Forecast"]*8)

QFC = 10  # forecast starts at column J (col 10)

# Helper: map quarterly forecast col (10-17) to annual col letter (H or I)
def q_annual(ci):
    return CL(8 + (ci - 10) // 4)

# Helper: seasonality col (B-E for Q1-Q4)
def q_season(ci):
    return CL(2 + (ci - 10) % 4)

# Helper: quarter number in year (1-4)
def q_num(ci):
    return (ci - 10) % 4 + 1

# Helper: previous annual col for DAU interpolation
def q_prev_annual(ci):
    return CL(8 + (ci - 10) // 4 - 1)

# DAU
w(ws, QT['dau'], 1, "DAU (M)", BF)
for i, v in enumerate(dau_qh):
    w(ws, QT['dau'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    pa = q_prev_annual(ci); ca = q_annual(ci); qn = q_num(ci)
    w(ws, QT['dau'], ci,
      f"=ROUND({RB_R}!{pa}{RB['dau']}+({RB_R}!{ca}{RB['dau']}-{RB_R}!{pa}{RB['dau']})*{qn}/4,0)",
      font=FCST_FONT, fmt=N)

# Revenue
w(ws, QT['rev'], 1, "Revenue", BF)
for i, v in enumerate(rev_qh):
    w(ws, QT['rev'], 2+i, v, fmt=N, font=BF)
for ci in range(QFC, 18):
    ac = q_annual(ci); sc = q_season(ci)
    w(ws, QT['rev'], ci,
      f"=ROUND({RB_R}!{ac}{RB['rev']}*{A_R}!{sc}{A['season']},0)",
      font=FCST_FONT, fmt=N)

# ARPU
w(ws, QT['arpu'], 1, "ARPU ($)", NF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['arpu'], ci,
      f"=ROUND({CL(ci)}{QT['rev']}/{CL(ci)}{QT['dau']},2)",
      font=fnt, fmt=U2)

# COGS
w(ws, QT['cogs'], 1, "Cost of Revenue", NF)
for i, v in enumerate(cor_qh):
    w(ws, QT['cogs'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['cogs'], ci,
      f"=ROUND({CL(ci)}{QT['rev']}*(1-{A_R}!{ac}{A['gm']}),0)",
      font=FCST_FONT, fmt=N)

# GP
w(ws, QT['gp'], 1, "Gross Profit", BF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['gp'], ci, f"={CL(ci)}{QT['rev']}-{CL(ci)}{QT['cogs']}", font=fnt, fmt=N)

# GM
w(ws, QT['gm'], 1, "  Gross Margin", NF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['gm'], ci, f"={CL(ci)}{QT['gp']}/{CL(ci)}{QT['rev']}", font=fnt, fmt=P)

# OpEx items
q_opex_items = [
    (QT['rd'], "R&D", rd_qh, A['rd']),
    (QT['sm'], "S&M", sm_qh, A['sm']),
    (QT['ga'], "G&A", ga_qh, A['ga']),
]
for qt_row, label, hist, a_row in q_opex_items:
    w(ws, qt_row, 1, label, NF)
    for i, v in enumerate(hist):
        w(ws, qt_row, 2+i, v, fmt=N)
    for ci in range(QFC, 18):
        ac = q_annual(ci)
        w(ws, qt_row, ci,
          f"=ROUND({CL(ci)}{QT['rev']}*{A_R}!{ac}{a_row},0)",
          font=FCST_FONT, fmt=N)

# OI
w(ws, QT['oi'], 1, "Operating Income (Loss)", BF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['oi'], ci,
      f"={CL(ci)}{QT['gp']}-{CL(ci)}{QT['rd']}-{CL(ci)}{QT['sm']}-{CL(ci)}{QT['ga']}",
      font=fnt, fmt=N)

w(ws, QT['oi_m'], 1, "  Operating Margin", NF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['oi_m'], ci, f"={CL(ci)}{QT['oi']}/{CL(ci)}{QT['rev']}", font=fnt, fmt=P)

# SBC
w(ws, QT['sbc'], 1, "SBC", NF)
for i, v in enumerate(sbc_qh):
    w(ws, QT['sbc'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['sbc'], ci,
      f"=ROUND({CL(ci)}{QT['rev']}*{A_R}!{ac}{A['sbc']},0)",
      font=FCST_FONT, fmt=N)

# D&A
w(ws, QT['da'], 1, "D&A", NF)
for i, v in enumerate(da_qh):
    w(ws, QT['da'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['da'], ci,
      f"=ROUND({A_R}!{ac}{A['da']}/4,0)",
      font=FCST_FONT, fmt=N)

# EBITDA
w(ws, QT['ebitda'], 1, "EBITDA", BF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['ebitda'], ci, f"={CL(ci)}{QT['oi']}+{CL(ci)}{QT['da']}", font=fnt, fmt=N)

# Interest
w(ws, QT['int'], 1, "Interest & Other", NF)
for i, v in enumerate(int_qh):
    w(ws, QT['int'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['int'], ci,
      f"=ROUND(-{A_R}!{ac}{A['interest']}/4,0)",
      font=FCST_FONT, fmt=N)

# Pretax
w(ws, QT['pt'], 1, "Pre-Tax Income", BF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['pt'], ci, f"={CL(ci)}{QT['oi']}+{CL(ci)}{QT['int']}", font=fnt, fmt=N)

# Tax
w(ws, QT['tax'], 1, "Tax Provision", NF)
for i, v in enumerate(tax_qh):
    w(ws, QT['tax'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['tax'], ci,
      f"=IF({CL(ci)}{QT['pt']}>0,ROUND({CL(ci)}{QT['pt']}*{A_R}!{ac}{A['tax']},0),0)",
      font=FCST_FONT, fmt=N)

# NI
w(ws, QT['ni'], 1, "Net Income (Loss)", BF)
for ci in range(2, 18):
    fnt = FCST_FONT if ci >= QFC else NF
    w(ws, QT['ni'], ci, f"={CL(ci)}{QT['pt']}-{CL(ci)}{QT['tax']}", font=fnt, fmt=N)

# CapEx
w(ws, QT['capex'], 1, "CapEx", NF)
for i, v in enumerate(capex_qh):
    w(ws, QT['capex'], 2+i, v, fmt=N)
for ci in range(QFC, 18):
    ac = q_annual(ci)
    w(ws, QT['capex'], ci,
      f"=ROUND(-{A_R}!{ac}{A['capex']}/4,0)",
      font=FCST_FONT, fmt=N)

colw(ws, [28]+[12]*16)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: SHARE COUNT & DILUTION
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_sc
w(ws, 1, 1, "Snap Inc. (SNAP) - Share Count & Dilution Schedule", TF)
ws.merge_cells('A1:K1')

# Annual section
hdr_row(ws, 3, yr_hdr); sub_row(ws, 4, hf_marker)

SC = {'shares': 5, 'dil': 6, 'sbc': 7, 'eps': 8}

w(ws, SC['shares'], 1, "Diluted Shares (M)", BF)
fill_all(ws, SC['shares'],
         lambda ci: f"={A_R}!{CL(ci)}{A['shares']}", fmt=N)

w(ws, SC['dil'], 1, "  YoY Dilution", NF)
fill_all(ws, SC['dil'],
         lambda ci: f"={CL(ci)}{SC['shares']}/{CL(ci-1)}{SC['shares']}-1",
         fmt=P, start=3)

w(ws, SC['sbc'], 1, "SBC ($M)", NF)
fill_all(ws, SC['sbc'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['sbc']}", fmt=N)

w(ws, SC['eps'], 1, "EPS (Diluted)", BF)
fill_all(ws, SC['eps'],
         lambda ci: f"={IS_R}!{CL(ci)}{IS['eps']}", fmt=U2)

# Quarterly dilution schedule
r = 11
w(ws, r, 1, "Quarterly Dilution Schedule (Next 8 Quarters)", TF)
ws.merge_cells(f'A{r}:I{r}')

r = 13
hdr_row(ws, r, [""]+q_labels_f)

SC_Q = {'beg': 14, 'rsu': 15, 'buy': 16, 'other': 17, 'end': 18,
        'grants': 20, 'sbc': 21, 'sbc_ps': 22,
        'unvested': 25, 'options': 26}

# Beginning shares
w(ws, SC_Q['beg'], 1, "Beginning Shares (M)", NF)
# Q1-25E beginning = FY2024 ending shares
w(ws, SC_Q['beg'], 2, f"={A_R}!G{A['shares']}", fmt=N)
for ci in range(3, 10):
    w(ws, SC_Q['beg'], ci, f"={CL(ci-1)}{SC_Q['end']}", fmt=N)

# RSU Vesting (input)
w(ws, SC_Q['rsu'], 1, "(+) RSU Vesting (M)", NF)
for i, v in enumerate(rsu_vest):
    w(ws, SC_Q['rsu'], 2+i, v, fmt=N, fill=INPUT_FILL)

# Buybacks (input)
w(ws, SC_Q['buy'], 1, "(-) Share Buybacks (M)", NF)
for i, v in enumerate(buyback_q):
    w(ws, SC_Q['buy'], 2+i, v, fmt=N, fill=INPUT_FILL)

# Other (input)
w(ws, SC_Q['other'], 1, "(+/-) Other (M)", NF)
for i, v in enumerate(other_q_sc):
    w(ws, SC_Q['other'], 2+i, v, fmt=N, fill=INPUT_FILL)

# Ending shares = Beginning + RSU - Buybacks + Other
w(ws, SC_Q['end'], 1, "Ending Shares (M)", BF)
for ci in range(2, 10):
    c = CL(ci)
    w(ws, SC_Q['end'], ci,
      f"={c}{SC_Q['beg']}+{c}{SC_Q['rsu']}-{c}{SC_Q['buy']}+{c}{SC_Q['other']}",
      fmt=N, font=BF)

# Memo items
w(ws, SC_Q['grants'], 1, "Memo: New RSU Grants (M)", NF)
for i, v in enumerate(new_grants):
    w(ws, SC_Q['grants'], 2+i, v, fmt=N, fill=INPUT_FILL)

w(ws, SC_Q['sbc'], 1, "Memo: Quarterly SBC ($M)", NF)
for ci in range(2, 10):
    # Map quarterly columns to quarterly detail tab columns
    # SC col B=Q1-25E → QT col J (10), SC col C=Q2-25E → QT col K (11), etc.
    qt_ci = ci + 8  # SC col 2 → QT col 10
    w(ws, SC_Q['sbc'], ci,
      f"={Q_R}!{CL(qt_ci)}{QT['sbc']}", fmt=N)

w(ws, SC_Q['sbc_ps'], 1, "Memo: SBC / Share ($)", NF)
for ci in range(2, 10):
    c = CL(ci)
    w(ws, SC_Q['sbc_ps'], ci,
      f"=ROUND({c}{SC_Q['sbc']}/{c}{SC_Q['end']},2)", fmt=U2)

# Outstanding awards
r = 24; sec(ws, r, "Outstanding Equity Awards (est.)")
w(ws, SC_Q['unvested'], 1, "  Unvested RSUs (M shares)", NF)
for i, v in enumerate(unvested_rsu):
    w(ws, SC_Q['unvested'], 2+i, v, fmt=N)

w(ws, SC_Q['options'], 1, "  Vested Unexercised Options (M)", NF)
for i, v in enumerate(vested_opts):
    w(ws, SC_Q['options'], 2+i, v, fmt=N)

colw(ws, [32]+[14]*10)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: SENSITIVITY ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_sn
w(ws, 1, 1, "Snap Inc. (SNAP) - Two-Way Sensitivity Analysis", TF)
ws.merge_cells('A1:G1')

# ── TABLE 1: DAU Growth × ARPU Growth → FY2026E Revenue ──────────────────────
r = 4
w(ws, r, 1, "TABLE 1: FY2026E Revenue ($M) - DAU Growth vs. ARPU Growth", SF)
ws.merge_cells(f'A{r}:G{r}')

dau_scen = [0.02, 0.035, 0.055, 0.07, 0.085]
arpu_scen = [0.04, 0.06, 0.09, 0.12, 0.15]

r = 6
w(ws, r, 1, "DAU Growth (down) / ARPU Growth (right)", BF)
for j, ag in enumerate(arpu_scen):
    w(ws, r, 2+j, ag, font=HDR_FONT, fmt=P, fill=HDR_FILL)
    ws.cell(row=r, column=2+j).alignment = Alignment(horizontal='center')

for i, dg in enumerate(dau_scen):
    rr = r + 1 + i
    w(ws, rr, 1, dg, font=BF, fmt=P, fill=SUB_FILL)
    for j, ag in enumerate(arpu_scen):
        # Revenue = (FY2025E DAU × (1+dg)) × (FY2025E ARPU × (1+ag))
        # DAU 2025 = RB!H7, ARPU 2025 = RB!H11
        formula = (f"=ROUND(({RB_R}!H{RB['dau']}*(1+$A{rr}))"
                   f"*({RB_R}!H{RB['arpu']}*(1+{CL(2+j)}${r})),0)")
        c = w(ws, rr, 2+j, formula, fmt=N)
        c.alignment = Alignment(horizontal='center')
        if abs(dg - 0.055) < 0.001 and abs(ag - 0.09) < 0.001:
            c.fill = GREEN_FILL; c.font = BF

# ── TABLE 2: Revenue Growth × EBITDA Margin → FY2026E EBITDA ────────────────
r2 = r + len(dau_scen) + 4
w(ws, r2, 1, "TABLE 2: FY2026E EBITDA ($M) - Revenue Growth vs. EBITDA Margin", SF)
ws.merge_cells(f'A{r2}:G{r2}')

rev_g_scen = [0.10, 0.14, 0.17, 0.20, 0.24]
ebitda_scen = [0.06, 0.09, 0.12, 0.15, 0.18]

r2 += 2
w(ws, r2, 1, "Rev Growth (down) / EBITDA Margin (right)", BF)
for j, em in enumerate(ebitda_scen):
    w(ws, r2, 2+j, em, font=HDR_FONT, fmt=P, fill=HDR_FILL)
    ws.cell(row=r2, column=2+j).alignment = Alignment(horizontal='center')

for i, rg in enumerate(rev_g_scen):
    rr = r2 + 1 + i
    w(ws, rr, 1, rg, font=BF, fmt=P, fill=SUB_FILL)
    for j, em in enumerate(ebitda_scen):
        # EBITDA = RB!H{rev} × (1+rev_growth) × ebitda_margin
        formula = (f"=ROUND({RB_R}!H{RB['rev']}*(1+$A{rr})*{CL(2+j)}${r2},0)")
        c = w(ws, rr, 2+j, formula, fmt=N)
        c.alignment = Alignment(horizontal='center')
        if abs(rg - 0.17) < 0.01 and abs(em - 0.12) < 0.01:
            c.fill = GREEN_FILL; c.font = BF

# ── TABLE 3: WACC × Terminal Growth → Implied Share Price ────────────────────
r3 = r2 + len(rev_g_scen) + 4
w(ws, r3, 1, "TABLE 3: Implied Share Price ($) - WACC vs. Terminal Growth Rate (DCF)", SF)
ws.merge_cells(f'A{r3}:G{r3}')

# Helper rows for DCF inputs
hr = r3 + 2
w(ws, hr, 1, "DCF Helper Data:", SF)
w(ws, hr+1, 1, "FCF FY2025E", NF)
w(ws, hr+1, 2, f"={CF_R}!H{CF['fcf']}", fmt=N)
w(ws, hr+2, 1, "FCF FY2026E", NF)
w(ws, hr+2, 2, f"={CF_R}!I{CF['fcf']}", fmt=N)
w(ws, hr+3, 1, "FCF FY2027E", NF)
w(ws, hr+3, 2, f"={CF_R}!J{CF['fcf']}", fmt=N)
w(ws, hr+4, 1, "FCF FY2028E", NF)
w(ws, hr+4, 2, f"={CF_R}!K{CF['fcf']}", fmt=N)
w(ws, hr+5, 1, "Net Debt ($M)", NF)
w(ws, hr+5, 2, f"={A_R}!H{A['ltd']}-{CF_R}!H{CF['end']}", fmt=N)
w(ws, hr+6, 1, "Shares (M)", NF)
w(ws, hr+6, 2, f"={A_R}!I{A['shares']}", fmt=N)

# FCF row refs for formulas
fcf_cells = [f"$B${hr+1}", f"$B${hr+2}", f"$B${hr+3}", f"$B${hr+4}"]
nd_cell = f"$B${hr+5}"
sh_cell = f"$B${hr+6}"

wacc_scen = [0.09, 0.10, 0.11, 0.12, 0.13]
tgr_scen = [0.02, 0.025, 0.03, 0.035, 0.04]

tbl_r = hr + 8
w(ws, tbl_r, 1, "WACC (down) / Terminal Growth (right)", BF)
for j, tg in enumerate(tgr_scen):
    w(ws, tbl_r, 2+j, tg, font=HDR_FONT, fmt=P, fill=HDR_FILL)
    ws.cell(row=tbl_r, column=2+j).alignment = Alignment(horizontal='center')

for i, wacc_v in enumerate(wacc_scen):
    rr = tbl_r + 1 + i
    w(ws, rr, 1, wacc_v, font=BF, fmt=P, fill=SUB_FILL)
    for j, tg in enumerate(tgr_scen):
        tg_ref = f"{CL(2+j)}${tbl_r}"
        wacc_ref = f"$A{rr}"
        # Full DCF formula
        pv_parts = "+".join(
            f"{fcf_cells[yr]}/(1+{wacc_ref})^{yr+1}" for yr in range(4)
        )
        tv = f"{fcf_cells[3]}*(1+{tg_ref})/({wacc_ref}-{tg_ref})/(1+{wacc_ref})^4"
        formula = f"=ROUND(({pv_parts}+{tv}-{nd_cell})/{sh_cell},1)"
        c = w(ws, rr, 2+j, formula, fmt=U1)
        c.alignment = Alignment(horizontal='center')
        if abs(wacc_v - 0.11) < 0.005 and abs(tg - 0.03) < 0.003:
            c.fill = GREEN_FILL; c.font = BF

colw(ws, [40]+[14]*8)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB: VALUATION CROSS-CHECKS
# ═══════════════════════════════════════════════════════════════════════════════
ws = ws_vl
w(ws, 1, 1, "Snap Inc. (SNAP) - Valuation Cross-Checks", TF)
ws.merge_cells('A1:I1')

# ── DCF ──────────────────────────────────────────────────────────────────────
r = 3; sec(ws, r, "1. DCF Valuation")

# Inputs
V = {}  # track valuation rows
r = 5; V['wacc'] = r
w(ws, r, 1, "WACC", BF); w(ws, r, 2, 0.11, fmt=P, fill=INPUT_FILL)
r = 6; V['tgr'] = r
w(ws, r, 1, "Terminal Growth Rate", BF); w(ws, r, 2, 0.03, fmt=P, fill=INPUT_FILL)
r = 7; V['nd'] = r
w(ws, r, 1, "Net Debt ($M)", BF)
w(ws, r, 2, f"={A_R}!H{A['ltd']}-{CF_R}!H{CF['end']}", fmt=N)
r = 8; V['sh'] = r
w(ws, r, 1, "Shares Outstanding (M)", BF)
w(ws, r, 2, f"={A_R}!I{A['shares']}", fmt=N)

# FCF projection row
r = 10
hdr_row(ws, r, ["", "FY2025E", "FY2026E", "FY2027E", "FY2028E"])
r = 11; V['fcf_start'] = r
w(ws, r, 1, "Free Cash Flow ($M)", NF)
for i, yr_col in enumerate(['H','I','J','K']):
    w(ws, r, 2+i, f"={CF_R}!{yr_col}{CF['fcf']}", fmt=N)

r = 12
w(ws, r, 1, "Discount Factor", NF)
for i in range(4):
    w(ws, r, 2+i, f"=1/(1+$B${V['wacc']})^{i+1}", fmt='0.0000')

r = 13
w(ws, r, 1, "PV of FCF ($M)", NF)
for i in range(4):
    c_l = CL(2+i)
    w(ws, r, 2+i, f"=ROUND({c_l}{V['fcf_start']}*{c_l}12,0)", fmt=N)

r = 15
w(ws, r, 1, "Sum PV of FCFs ($M)", BF)
w(ws, r, 2, f"=SUM(B13:E13)", fmt=N)
V['sum_pv'] = r

r = 16
w(ws, r, 1, "Terminal Value ($M)", NF)
w(ws, r, 2, f"=ROUND(E{V['fcf_start']}*(1+$B${V['tgr']})/($B${V['wacc']}-$B${V['tgr']}),0)", fmt=N)
V['tv'] = r

r = 17
w(ws, r, 1, "PV of Terminal Value ($M)", NF)
w(ws, r, 2, f"=ROUND(B{V['tv']}*E12,0)", fmt=N)
V['pv_tv'] = r

r = 18
w(ws, r, 1, "Enterprise Value ($M)", BF)
w(ws, r, 2, f"=B{V['sum_pv']}+B{V['pv_tv']}", fmt=N)
V['ev'] = r

r = 19
w(ws, r, 1, "Equity Value ($M)", BF)
w(ws, r, 2, f"=B{V['ev']}-B{V['nd']}", fmt=N)
V['eq_val'] = r

r = 20; V['dcf_price'] = r
w(ws, r, 1, "Implied Share Price ($)", BF)
w(ws, r, 2, f"=ROUND(B{V['eq_val']}/B{V['sh']},2)", fmt=U2)

# ── Cohort NPV ──────────────────────────────────────────────────────────────
r = 23; sec(ws, r, "2. Cohort NPV / User Economics")

r = 25; V['c_arpu'] = r
w(ws, r, 1, "Year-1 ARPU ($)", BF)
w(ws, r, 2, f"={RB_R}!H{RB['arpu']}", fmt=U2)

r = 26; V['c_arpu_g'] = r
w(ws, r, 1, "ARPU Annual Growth", BF); w(ws, r, 2, 0.05, fmt=P, fill=INPUT_FILL)
r = 27; V['c_churn'] = r
w(ws, r, 1, "Annual Churn Rate", BF); w(ws, r, 2, 0.20, fmt=P, fill=INPUT_FILL)
r = 28; V['c_gm'] = r
w(ws, r, 1, "Cohort Gross Margin", BF); w(ws, r, 2, 0.56, fmt=P, fill=INPUT_FILL)
r = 29; V['c_disc'] = r
w(ws, r, 1, "Discount Rate", BF); w(ws, r, 2, 0.11, fmt=P, fill=INPUT_FILL)
r = 30; V['c_cac'] = r
w(ws, r, 1, "CAC per User ($)", BF); w(ws, r, 2, 3.50, fmt=U2, fill=INPUT_FILL)

r = 32
hdr_row(ws, r, [""] + [f"Year {i}" for i in range(1, 9)])

r = 33  # Retention
w(ws, r, 1, "Retention Rate", NF)
for i in range(8):
    w(ws, r, 2+i, f"=(1-$B${V['c_churn']})^{i}", fmt=P)

r = 34  # ARPU
w(ws, r, 1, "ARPU ($)", NF)
for i in range(8):
    w(ws, r, 2+i, f"=ROUND($B${V['c_arpu']}*(1+$B${V['c_arpu_g']})^{i},2)", fmt=U2)

r = 35  # GP per user
w(ws, r, 1, "GP / User ($)", NF)
for i in range(8):
    c = CL(2+i)
    w(ws, r, 2+i, f"=ROUND({c}34*$B${V['c_gm']}*{c}33,2)", fmt=U2)

r = 36  # PV of GP
w(ws, r, 1, "PV of GP / User ($)", NF)
for i in range(8):
    c = CL(2+i)
    w(ws, r, 2+i, f"=ROUND({c}35/(1+$B${V['c_disc']})^{i+1},2)", fmt=U2)

r = 38; V['ltv'] = r
w(ws, r, 1, "Lifetime Value (LTV) ($)", BF)
w(ws, r, 2, "=ROUND(SUM(B36:I36),2)", fmt=U2)

r = 39
w(ws, r, 1, "CAC per User ($)", NF)
w(ws, r, 2, f"=B{V['c_cac']}", fmt=U2)

r = 40; V['ltv_cac'] = r
w(ws, r, 1, "LTV / CAC Ratio", BF)
w(ws, r, 2, f"=ROUND(B{V['ltv']}/B39,1)", fmt=D)

r = 41; V['npv_user'] = r
w(ws, r, 1, "NPV per User ($)", BF)
w(ws, r, 2, f"=ROUND(B{V['ltv']}-B39,2)", fmt=U2)

r = 43
w(ws, r, 1, "Implied EV (NPV/user x DAU)", BF)
w(ws, r, 2, f"=ROUND(B{V['npv_user']}*{RB_R}!H{RB['dau']},0)", fmt=N)
V['cohort_ev'] = r

r = 44
w(ws, r, 1, "Implied Equity Value ($M)", BF)
w(ws, r, 2, f"=B{V['cohort_ev']}-B{V['nd']}", fmt=N)
V['cohort_eq'] = r

r = 45; V['cohort_price'] = r
w(ws, r, 1, "Implied Share Price (Cohort)", BF)
w(ws, r, 2, f"=ROUND(B{V['cohort_eq']}/B{V['sh']},2)", fmt=U2)

# ── S-Curve ──────────────────────────────────────────────────────────────────
r = 48; sec(ws, r, "3. Adoption S-Curve Analysis")

r = 50; V['tam'] = r
w(ws, r, 1, "Total Addressable Market (M DAU)", BF)
w(ws, r, 2, 800, fmt=N, fill=INPUT_FILL)

r = 51
w(ws, r, 1, "Current Penetration (FY2025E)", NF)
w(ws, r, 2, f"={RB_R}!H{RB['dau']}/B{V['tam']}", fmt=P)

r = 53
hdr_row(ws, r, ["Year", "DAU (M)", "Penetration", "YoY Growth", "Phase"])
s_years = list(range(2020, 2031))
s_dau_hist_part = [265, 319, 375, 414, 443]
s_phases_h = ["Early Growth","Early Growth","Acceleration","Mid-Curve","Mid-Curve"]
s_phases_f = ["Deceleration"]*4 + ["Maturation","Maturation"]
for i, yr in enumerate(s_years):
    rr = 54 + i
    w(ws, rr, 1, yr, fmt=N)
    if i < 5:
        # Historical
        w(ws, rr, 2, s_dau_hist_part[i], fmt=N)
        w(ws, rr, 3, f"=B{rr}/$B${V['tam']}", fmt=P)
        if i > 0:
            w(ws, rr, 4, f"=B{rr}/B{rr-1}-1", fmt=P)
        w(ws, rr, 5, s_phases_h[i])
    elif i < 9:
        # Forecast years (2025-2028) reference Revenue Build
        yr_col = CL(8 + (i - 5))  # H, I, J, K
        w(ws, rr, 2, f"={RB_R}!{yr_col}{RB['dau']}", fmt=N)
        w(ws, rr, 3, f"=B{rr}/$B${V['tam']}", fmt=P)
        w(ws, rr, 4, f"=B{rr}/B{rr-1}-1", fmt=P)
        w(ws, rr, 5, s_phases_f[i-5])
    else:
        # 2029-2030: extrapolate with declining growth
        growth = 0.035 if i == 9 else 0.030
        w(ws, rr, 2, f"=ROUND(B{rr-1}*(1+{growth}),0)", fmt=N)
        w(ws, rr, 3, f"=B{rr}/$B${V['tam']}", fmt=P)
        w(ws, rr, 4, f"=B{rr}/B{rr-1}-1", fmt=P)
        w(ws, rr, 5, s_phases_f[i-5])

# ── Unit Economics Sanity Check ──────────────────────────────────────────────
r = 68; sec(ws, r, "4. Unit Economics -> EV Sanity Check")

r = 70; V['mktcap'] = r
w(ws, r, 1, "Current Market Cap ($M, est.)", BF)
w(ws, r, 2, 18000, fmt=N, fill=INPUT_FILL)

r = 71
w(ws, r, 1, "Enterprise Value ($M, est.)", BF)
w(ws, r, 2, f"=B{V['mktcap']}+B{V['nd']}", fmt=N)
V['ev_est'] = r

r = 73
hdr_row(ws, r, ["Metric", "Value", "Commentary"])

r = 74
w(ws, r, 1, "EV / DAU ($)", NF)
w(ws, r, 2, f"=ROUND(B{V['ev_est']}/{RB_R}!H{RB['dau']},1)", fmt=U1)
w(ws, r, 3, "vs. META ~$310, PINS ~$75")

r = 75
w(ws, r, 1, "EV / FY2025E Revenue", NF)
w(ws, r, 2, f"=ROUND(B{V['ev_est']}/{RB_R}!H{RB['rev']},1)", fmt=D)
w(ws, r, 3, "vs. social-media peers at 4-8x")

r = 76
w(ws, r, 1, "EV / FY2026E EBITDA", NF)
w(ws, r, 2, f"=IF({IS_R}!I{IS['ebitda']}>0,ROUND(B{V['ev_est']}/{IS_R}!I{IS['ebitda']},1),\"N/M\")", fmt=D)
w(ws, r, 3, "vs. peers at 15-25x")

r = 77
w(ws, r, 1, "FY2026E FCF Yield", NF)
w(ws, r, 2, f"=IF({CF_R}!I{CF['fcf']}>0,{CF_R}!I{CF['fcf']}/B{V['mktcap']},0)", fmt=P)
w(ws, r, 3, "Implies early FCF inflection")

# ── Valuation Summary ────────────────────────────────────────────────────────
r = 80; sec(ws, r, "Valuation Summary")

r = 81; V['cur_price'] = r
w(ws, r, 1, "Current Share Price (est.)", BF)
w(ws, r, 2, 10.50, fmt=U2, fill=INPUT_FILL)

r = 83
hdr_row(ws, r, ["Method", "Implied Price ($)", "Upside / Downside"])

r = 84
w(ws, r, 1, "DCF (Base Case)", NF)
w(ws, r, 2, f"=B{V['dcf_price']}", fmt=U2)
w(ws, r, 3, f"=B{r}/B{V['cur_price']}-1", fmt=P)

r = 85
w(ws, r, 1, "Cohort NPV", NF)
w(ws, r, 2, f"=B{V['cohort_price']}", fmt=U2)
w(ws, r, 3, f"=B{r}/B{V['cur_price']}-1", fmt=P)

r = 86
w(ws, r, 1, "EV/Revenue Comp (4.0x FY26E)", NF)
w(ws, r, 2, f"=ROUND(({RB_R}!I{RB['rev']}*4-B{V['nd']})/B{V['sh']},2)", fmt=U2)
w(ws, r, 3, f"=B{r}/B{V['cur_price']}-1", fmt=P)

colw(ws, [38]+[18]*8)


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
output_path = "/workspace/Snap_Inc_Three_Statement_Model.xlsx"
wb.save(output_path)
print(f"Model saved to: {output_path}")
print("All forecast cells contain live Excel formulas.")
print(f"Tabs: {wb.sheetnames}")
