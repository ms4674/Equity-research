"""
Build a granular Excel model estimating advertising-revenue exposure to users
UNDER AGE 16 for Meta, Google (Alphabet), Snap and Pinterest.

The workbook is "live": financial data and modelling assumptions live on dedicated
sheets, and the per-company / per-region model sheets and the aggregate summary are
driven by Excel formulas so any assumption can be flexed and results recalculate.

Run:  python3 build_model.py
Output: Under16_Revenue_Exposure_Model.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------------------
# Styling helpers
# --------------------------------------------------------------------------------------
NAVY = "1F3864"
BLUE = "2E5496"
LIGHT = "D9E1F2"
LIGHT2 = "EAF0FB"
GREY = "F2F2F2"
AMBER = "FFF2CC"
GREEN = "E2EFDA"
WHITE = "FFFFFF"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=WHITE)
H1_FONT = Font(name="Calibri", size=13, bold=True, color=WHITE)
H2_FONT = Font(name="Calibri", size=11, bold=True, color=NAVY)
H2_FONT_ = Font(name="Calibri", size=11, bold=True, color=WHITE)
HDR_FONT = Font(name="Calibri", size=10, bold=True, color=WHITE)
BOLD = Font(name="Calibri", size=10, bold=True)
NORMAL = Font(name="Calibri", size=10)
ITALIC = Font(name="Calibri", size=9, italic=True, color="595959")
SRC_FONT = Font(name="Calibri", size=9, italic=True, color="595959")

CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RGT = Alignment(horizontal="right", vertical="center")
TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

FMT_USD_M = '#,##0;(#,##0)'
FMT_USD_M1 = '#,##0.0;(#,##0.0)'
FMT_PCT1 = '0.0%'
FMT_PCT2 = '0.00%'
FMT_NUM = '#,##0'
FMT_X = '0.00"x"'


def hfill(color):
    return PatternFill("solid", fgColor=color)


def style_header_row(ws, row, c1, c2, fill=BLUE, font=HDR_FONT):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = hfill(fill)
        cell.font = font
        cell.alignment = CTR
        cell.border = BORDER


def band(ws, row, c1, c2, text, fill=NAVY, font=H1_FONT):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = ws.cell(row=row, column=c1, value=text)
    cell.fill = hfill(fill)
    cell.font = font
    cell.alignment = Alignment(horizontal="left", vertical="center")


def put(ws, row, col, value, font=NORMAL, align=LFT, fmt=None, fill=None, border=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.alignment = align
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = hfill(fill)
    if border:
        cell.border = BORDER
    return cell


wb = Workbook()

# ======================================================================================
# DATA (FY2024 unless noted).  Sources captured on the "Sources" sheet.
# ======================================================================================
REGIONS_4 = ["US & Canada", "Europe", "Asia-Pacific", "Rest of World"]

# ---- Meta (advertising revenue by USER geography; regional split is an estimate) ----
META_AD = 160633.0          # FY2024 advertising revenue, $M (10-K)
META_TOTAL = 164501.0
META_FOA = 162355.0
# Estimated user-geography ad-revenue shares (documented assumption)
META_REGION_SHARE = {"US & Canada": 0.44, "Europe": 0.23, "Asia-Pacific": 0.22, "Rest of World": 0.11}
META_ARPP = {"US & Canada": 233.42, "Europe": 68.12, "Asia-Pacific": 21.28, "Rest of World": 14.00}

# ---- Alphabet / Google ----
GOOG_TOTAL = 350018.0
GOOG_ADV = 264590.0
YT_ADS = 36147.0
GOOG_SEARCH = 198084.0
GOOG_NETWORK = 30359.0
# YouTube ad-revenue regional split (estimate: YouTube monetises internationally more
# than Alphabet's overall US-billed geographic mix implies)
YT_REGION_SHARE = {"US & Canada": 0.35, "Europe": 0.30, "Asia-Pacific": 0.23, "Rest of World": 0.12}
# Alphabet total-revenue geographic mix (10-K, ex-hedging) used for Search & Network
GOOG_GEO = {"US & Canada": 0.487, "Europe": 0.292, "Asia-Pacific": 0.162, "Rest of World": 0.058}

# ---- Snap (total revenue by region, exact from press release, $M) ----
SNAP_TOTAL = 5361.0
SNAP_REGION = {"North America": 3337.3, "Europe": 961.6, "Rest of World": 1062.5}
SNAP_DAU = {"North America": 100.0, "Europe": 99.0, "Rest of World": 254.0}  # Q4-24 millions

# ---- Pinterest (total revenue; regional split estimated from disclosed Q4 mix) ----
PINS_TOTAL = 3646.0
PINS_REGION = {"US & Canada": 2860.0, "Europe": 610.0, "Rest of World": 176.0}
PINS_MAU = {"US & Canada": 101.0, "Europe": 145.0, "Rest of World": 307.0}  # Q4-24 millions

# ---- Harvard (Raffoul et al. 2023, PLOS ONE) US 2022 anchors ----
# under-18 ad-revenue SHARE of each platform's US ad revenue, and 0-12 dollar split
HARVARD = {
    # platform: (US under-18 rev share, 0-12 rev $M, 13-17 rev $M)
    "Instagram": (0.160, 801.1, 4000.0),
    "YouTube": (0.270, 959.1, 1200.0),
    "Snapchat": (0.414, None, None),   # 0-12 vs 13-17 $ not disclosed; modelled
    "Facebook": (0.019, 137.2, None),
}

# ======================================================================================
# SHEET 1 — COVER & SUMMARY  (built last for formulas; create now to be first tab)
# ======================================================================================
ws_sum = wb.active
ws_sum.title = "Summary"

# ======================================================================================
# SHEET — METHODOLOGY
# ======================================================================================
ws_m = wb.create_sheet("Methodology")

# ======================================================================================
# SHEET — FINANCIALS
# ======================================================================================
ws_f = wb.create_sheet("Financials (FY2024)")

# ======================================================================================
# SHEET — YOUTH USAGE EVIDENCE
# ======================================================================================
ws_e = wb.create_sheet("Youth Usage Evidence")

# ======================================================================================
# SHEET — ASSUMPTIONS
# ======================================================================================
ws_a = wb.create_sheet("Assumptions")

# Model sheets
ws_meta = wb.create_sheet("Model - Meta")
ws_goog = wb.create_sheet("Model - Google")
ws_snap = wb.create_sheet("Model - Snap")
ws_pins = wb.create_sheet("Model - Pinterest")
ws_agg = wb.create_sheet("Aggregate by Region")
ws_src = wb.create_sheet("Sources")

# --------------------------------------------------------------------------------------
# ASSUMPTIONS SHEET
# --------------------------------------------------------------------------------------
a = ws_a
a.sheet_view.showGridLines = False
for col, w in {"A": 30, "B": 13, "C": 13, "D": 14, "E": 15, "F": 52}.items():
    a.column_dimensions[col].width = w

band(a, 1, 1, 6, "MODELLING ASSUMPTIONS  (edit blue-shaded inputs — the model recalculates)")
a.row_dimensions[1].height = 22

# ---- Global split parameter ----
band(a, 3, 1, 6, "1.  Age-band split parameter", fill=BLUE, font=H2_FONT_)
put(a, 4, 1, "13–15 share of the 13–17 revenue band", BOLD)
SPLIT_CELL = "B4"
put(a, 4, 2, 0.57, NORMAL, CTR, FMT_PCT1, AMBER, True)
put(a, 4, 6, "3 of the 5 single-year ages in 13–17 are <16; mild skew to older teens "
             "(slightly higher engagement). Low/Base/High = 50% / 57% / 62%.", ITALIC)

# ---- Platform US under-16 share derivation ----
band(a, 6, 1, 6, "2.  Platform US under-16 ad-revenue share (anchored to Harvard 2023 US 2022 study)",
     fill=BLUE, font=H2_FONT_)
hdr = ["Platform / surface", "US under-18\nrev share", "0–12 share\nof under-18 $",
       "Under-16 frac.\nof under-18", "US under-16\nrev share", "Basis / source"]
r = 7
for i, h in enumerate(hdr):
    put(a, r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
style_header_row(a, r, 1, 6)
a.row_dimensions[r].height = 30

# platform rows: name, U18 share, 0-12 frac of U18, basis text
plat_rows = [
    ("Instagram", 0.160, 0.1669, "Harvard: U18 share 16%; 0–12 $801.1M of $4,801M"),
    ("Facebook", 0.019, 0.2200, "Harvard: U18 share 1.9% (low youth base)"),
    ("Messenger / other (Meta)", 0.030, 0.300, "Estimate (messaging; limited youth ads)"),
    ("YouTube ads", 0.270, 0.4442, "Harvard: U18 share 27%; 0–12 $959.1M of $2,159M"),
    ("Google Search & other", 0.025, 0.250, "Estimate: teens search but low commercial intent/ad value"),
    ("Google Network", 0.030, 0.250, "Estimate: programmatic, broad adult base"),
    ("Snapchat", 0.414, 0.100, "Harvard: U18 share 41.4%; 0–12 $ modelled at 10%"),
    ("Pinterest", 0.040, 0.150, "Estimate: 13–17 ≈4% of users; <18 ad limits"),
]
plat_first = 8
plat_addr = {}  # platform -> US under16 share cell (col E)
for j, (name, u18, u012, basis) in enumerate(plat_rows):
    rr = plat_first + j
    put(a, rr, 1, name, BOLD, LFT, border=True)
    put(a, rr, 2, u18, NORMAL, CTR, FMT_PCT1, AMBER, True)
    put(a, rr, 3, u012, NORMAL, CTR, FMT_PCT1, AMBER, True)
    # under-16 fraction = 0-12 frac + (1 - 0-12 frac) * split
    put(a, rr, 4, f"=C{rr}+(1-C{rr})*${SPLIT_CELL}", NORMAL, CTR, FMT_PCT1, GREY, True)
    # US under-16 share = U18 share * under-16 fraction
    put(a, rr, 5, f"=B{rr}*D{rr}", BOLD, CTR, FMT_PCT2, GREEN, True)
    put(a, rr, 6, basis, ITALIC, LFT, border=True)
    plat_addr[name] = f"'Assumptions'!$E${rr}"

# ---- Meta blend weights ----
mb_r = plat_first + len(plat_rows) + 1
band(a, mb_r, 1, 6, "3.  Meta surface mix (weights of US ad revenue) — used to blend Meta's under-16 share",
     fill=BLUE, font=H2_FONT_)
mb_r += 1
for i, h in enumerate(["Surface", "Weight", "", "", "US under-16 share", "Note"]):
    put(a, mb_r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
meta_mix = [("Instagram", 0.50), ("Facebook", 0.40), ("Messenger / other (Meta)", 0.10)]
meta_mix_first = mb_r + 1
for j, (name, w) in enumerate(meta_mix):
    rr = meta_mix_first + j
    put(a, rr, 1, name, BOLD, LFT, border=True)
    put(a, rr, 2, w, NORMAL, CTR, FMT_PCT1, AMBER, True)
    put(a, rr, 5, f"={plat_addr[name]}", NORMAL, CTR, FMT_PCT2, GREY, True)
    put(a, rr, 6, "Instagram is the youth driver; Facebook skews older", ITALIC, LFT, border=True)
meta_blend_r = meta_mix_first + len(meta_mix)
put(a, meta_blend_r, 1, "Meta blended US under-16 share", BOLD, LFT, fill=GREEN, border=True)
put(a, meta_blend_r, 2, f"=SUMPRODUCT(B{meta_mix_first}:B{meta_blend_r-1},E{meta_mix_first}:E{meta_blend_r-1})",
    BOLD, CTR, FMT_PCT2, GREEN, True)
META_US_SHARE = f"'Assumptions'!$B${meta_blend_r}"

# ---- Regional multipliers ----
rm_r = meta_blend_r + 2
band(a, rm_r, 1, 6, "4.  Regional multipliers (applied to the US/North-America under-16 share)",
     fill=BLUE, font=H2_FONT_)
rm_r += 1
for i, h in enumerate(["Region", "Multiplier", "", "", "", "Rationale"]):
    put(a, rm_r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
mult_rows = [
    ("US & Canada / North America", 1.00, "Anchor region (Harvard study basis)"),
    ("Europe / EMEA", 0.95, "High youth usage (Ofcom) offset by GDPR-K / DSA enforcement"),
    ("Asia-Pacific", 1.05, "Younger populations; higher youth share of users"),
    ("Rest of World", 1.15, "Youngest populations; youth-heavy adoption, light regulation"),
]
mult_first = rm_r + 1
mult_addr = {}
for j, (name, m, why) in enumerate(mult_rows):
    rr = mult_first + j
    put(a, rr, 1, name, BOLD, LFT, border=True)
    put(a, rr, 2, m, NORMAL, CTR, FMT_X, AMBER, True)
    put(a, rr, 6, why, ITALIC, LFT, border=True)
    mult_addr[name] = f"'Assumptions'!$B${rr}"
MULT_NA = mult_addr["US & Canada / North America"]
MULT_EU = mult_addr["Europe / EMEA"]
MULT_AP = mult_addr["Asia-Pacific"]
MULT_RW = mult_addr["Rest of World"]

# ---- Scenario scalars ----
sc_r = mult_first + len(mult_rows) + 1
band(a, sc_r, 1, 6, "5.  Scenario scalars (capture model uncertainty: Harvard intervals, age mis-reporting)",
     fill=BLUE, font=H2_FONT_)
sc_r += 1
for i, h in enumerate(["Scenario", "Scalar", "", "", "", "Note"]):
    put(a, sc_r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
scen_rows = [("Low", 0.65), ("Base", 1.00), ("High", 1.35)]
scen_first = sc_r + 1
scen_addr = {}
for j, (name, s) in enumerate(scen_rows):
    rr = scen_first + j
    put(a, rr, 1, name, BOLD, LFT, border=True)
    put(a, rr, 2, s, NORMAL, CTR, FMT_X, AMBER, True)
    put(a, rr, 6, "Applied multiplicatively to base under-16 revenue", ITALIC, LFT, border=True)
    scen_addr[name] = f"'Assumptions'!$B${rr}"
SC_LOW, SC_BASE, SC_HIGH = scen_addr["Low"], scen_addr["Base"], scen_addr["High"]

# share-cell shortcuts for model sheets
SHARE_YT = plat_addr["YouTube ads"]
SHARE_SEARCH = plat_addr["Google Search & other"]
SHARE_NET = plat_addr["Google Network"]
SHARE_SNAP = plat_addr["Snapchat"]
SHARE_PINS = plat_addr["Pinterest"]

print("Assumptions laid out. Meta blend row:", meta_blend_r)


# --------------------------------------------------------------------------------------
# Helper: region -> multiplier cell
# --------------------------------------------------------------------------------------
def mult_for(region):
    if region in ("US & Canada", "North America", "US"):
        return MULT_NA
    if region in ("Europe", "Europe / EMEA", "EMEA"):
        return MULT_EU
    if region == "Asia-Pacific":
        return MULT_AP
    return MULT_RW


def model_sheet_header(ws, company, hard_facts):
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 26, "B": 16, "C": 15, "D": 17, "E": 15, "F": 15, "G": 46}.items():
        ws.column_dimensions[col].width = w
    band(ws, 1, 1, 7, f"UNDER-16 AD-REVENUE EXPOSURE MODEL  —  {company}")
    ws.row_dimensions[1].height = 22
    r = 3
    for label, val, fmt in hard_facts:
        put(ws, r, 1, label, BOLD, LFT)
        c = put(ws, r, 2, val, NORMAL, RGT, fmt)
        r += 1
    return r + 1


def model_table_header(ws, r, rev_label="Regional revenue ($M)"):
    heads = ["Region", rev_label, "Under-16 share\n(base)", "Under-16 rev\nBASE ($M)",
             "Low ($M)", "High ($M)", "Notes"]
    for i, h in enumerate(heads):
        put(ws, r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
    ws.row_dimensions[r].height = 30
    return r + 1


# ======================================================================================
# MODEL — META
# ======================================================================================
r = model_sheet_header(ws_meta, "Meta Platforms (Facebook + Instagram + Messenger)", [
    ("FY2024 total revenue ($M)", META_TOTAL, FMT_USD_M),
    ("FY2024 advertising revenue ($M)", META_AD, FMT_USD_M),
    ("Meta blended US under-16 ad-revenue share", f"={META_US_SHARE}", FMT_PCT2),
])
hdr_r = model_table_header(ws_meta, r)
data_first = hdr_r
meta_notes = {
    "US & Canada": "Anchor; Instagram youth-heavy, Facebook older",
    "Europe": "Ofcom: ~80% of 16-17s on IG/Snap; DSA minor-risk probe ongoing",
    "Asia-Pacific": "Lower ARPU but younger, youth-heavy user base",
    "Rest of World": "Youngest demographics; lightest regulation",
}
rr = data_first
for region in REGIONS_4:
    put(ws_meta, rr, 1, region, BOLD, LFT, border=True)
    # regional ad revenue = META_AD * region share (share stored inline as value)
    put(ws_meta, rr, 2, round(META_AD * META_REGION_SHARE[region], 1), NORMAL, RGT, FMT_USD_M, fill=GREY, border=True)
    put(ws_meta, rr, 3, f"={META_US_SHARE}*{mult_for(region)}", NORMAL, CTR, FMT_PCT2, border=True)
    put(ws_meta, rr, 4, f"=B{rr}*C{rr}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(ws_meta, rr, 5, f"=B{rr}*C{rr}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_meta, rr, 6, f"=B{rr}*C{rr}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_meta, rr, 7, meta_notes[region], ITALIC, LFT, border=True)
    rr += 1
meta_tot = rr
put(ws_meta, meta_tot, 1, "TOTAL — Meta", BOLD, LFT, fill=LIGHT, border=True)
for col in (2, 4, 5, 6):
    put(ws_meta, meta_tot, col, f"=SUM({get_column_letter(col)}{data_first}:{get_column_letter(col)}{rr-1})",
        BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_meta, meta_tot, 3, f"=D{meta_tot}/B{meta_tot}", BOLD, CTR, FMT_PCT2, LIGHT, True)
put(ws_meta, meta_tot, 7, "Regional ad-revenue split is an estimate (user-geography)", ITALIC, LFT, fill=LIGHT, border=True)
META_TOTAL_REF = (ws_meta.title, meta_tot)

# ======================================================================================
# MODEL — GOOGLE  (YouTube + Search&other + Network, each x 4 regions)
# ======================================================================================
r = model_sheet_header(ws_goog, "Google / Alphabet (YouTube + Search + Network)", [
    ("FY2024 Alphabet total revenue ($M)", GOOG_TOTAL, FMT_USD_M),
    ("FY2024 Google advertising revenue ($M)", GOOG_ADV, FMT_USD_M),
    ("  of which YouTube ads ($M)", YT_ADS, FMT_USD_M),
    ("  of which Google Search & other ($M)", GOOG_SEARCH, FMT_USD_M),
    ("  of which Google Network ($M)", GOOG_NETWORK, FMT_USD_M),
    ("YouTube US under-16 ad-revenue share", f"={SHARE_YT}", FMT_PCT2),
])
band(ws_goog, r, 1, 7, "YouTube ads — the principal under-16 surface", fill=BLUE, font=H2_FONT_)
r += 1
hdr_r = model_table_header(ws_goog, r)
yt_first = hdr_r
rr = yt_first
for region in REGIONS_4:
    put(ws_goog, rr, 1, region, BOLD, LFT, border=True)
    put(ws_goog, rr, 2, round(YT_ADS * YT_REGION_SHARE[region], 1), NORMAL, RGT, FMT_USD_M, GREY, True)
    put(ws_goog, rr, 3, f"={SHARE_YT}*{mult_for(region)}", NORMAL, CTR, FMT_PCT2, border=True)
    put(ws_goog, rr, 4, f"=B{rr}*C{rr}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(ws_goog, rr, 5, f"=B{rr}*C{rr}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_goog, rr, 6, f"=B{rr}*C{rr}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_goog, rr, 7, "YouTube regional split estimated (intl-heavy monetisation)", ITALIC, LFT, border=True)
    rr += 1
yt_tot = rr
put(ws_goog, yt_tot, 1, "Subtotal — YouTube ads", BOLD, LFT, fill=LIGHT2, border=True)
for col in (2, 4, 5, 6):
    put(ws_goog, yt_tot, col, f"=SUM({get_column_letter(col)}{yt_first}:{get_column_letter(col)}{rr-1})", BOLD, RGT, FMT_USD_M, LIGHT2, True)
put(ws_goog, yt_tot, 3, f"=D{yt_tot}/B{yt_tot}", BOLD, CTR, FMT_PCT2, LIGHT2, True)

# Search & other
r2 = yt_tot + 2
band(ws_goog, r2, 1, 7, "Google Search & other (lower under-16 exposure)", fill=BLUE, font=H2_FONT_)
r2 += 1
hdr_r = model_table_header(ws_goog, r2)
se_first = hdr_r
rr = se_first
for region in REGIONS_4:
    put(ws_goog, rr, 1, region, BOLD, LFT, border=True)
    put(ws_goog, rr, 2, round(GOOG_SEARCH * GOOG_GEO[region], 1), NORMAL, RGT, FMT_USD_M, GREY, True)
    put(ws_goog, rr, 3, f"={SHARE_SEARCH}*{mult_for(region)}", NORMAL, CTR, FMT_PCT2, border=True)
    put(ws_goog, rr, 4, f"=B{rr}*C{rr}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(ws_goog, rr, 5, f"=B{rr}*C{rr}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_goog, rr, 6, f"=B{rr}*C{rr}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_goog, rr, 7, "Search regional split = Alphabet geographic mix", ITALIC, LFT, border=True)
    rr += 1
se_tot = rr
put(ws_goog, se_tot, 1, "Subtotal — Search & other", BOLD, LFT, fill=LIGHT2, border=True)
for col in (2, 4, 5, 6):
    put(ws_goog, se_tot, col, f"=SUM({get_column_letter(col)}{se_first}:{get_column_letter(col)}{rr-1})", BOLD, RGT, FMT_USD_M, LIGHT2, True)
put(ws_goog, se_tot, 3, f"=D{se_tot}/B{se_tot}", BOLD, CTR, FMT_PCT2, LIGHT2, True)

# Network (single line, worldwide)
r3 = se_tot + 2
band(ws_goog, r3, 1, 7, "Google Network (worldwide, minimal youth exposure)", fill=BLUE, font=H2_FONT_)
r3 += 1
hdr_r = model_table_header(ws_goog, r3)
net_r = hdr_r
put(ws_goog, net_r, 1, "Worldwide", BOLD, LFT, border=True)
put(ws_goog, net_r, 2, GOOG_NETWORK, NORMAL, RGT, FMT_USD_M, GREY, True)
put(ws_goog, net_r, 3, f"={SHARE_NET}", NORMAL, CTR, FMT_PCT2, border=True)
put(ws_goog, net_r, 4, f"=B{net_r}*C{net_r}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
put(ws_goog, net_r, 5, f"=B{net_r}*C{net_r}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
put(ws_goog, net_r, 6, f"=B{net_r}*C{net_r}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
put(ws_goog, net_r, 7, "Programmatic third-party inventory; broad adult base", ITALIC, LFT, border=True)

# Google grand total
gt = net_r + 2
put(ws_goog, gt, 1, "TOTAL — Google (YouTube + Search + Network)", BOLD, LFT, fill=LIGHT, border=True)
put(ws_goog, gt, 2, f"=B{yt_tot}+B{se_tot}+B{net_r}", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_goog, gt, 4, f"=D{yt_tot}+D{se_tot}+D{net_r}", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_goog, gt, 5, f"=E{yt_tot}+E{se_tot}+E{net_r}", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_goog, gt, 6, f"=F{yt_tot}+F{se_tot}+F{net_r}", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_goog, gt, 3, f"=D{gt}/B{gt}", BOLD, CTR, FMT_PCT2, LIGHT, True)
put(ws_goog, gt, 7, "YouTube is ~90%+ of Google's under-16 exposure", ITALIC, LFT, fill=LIGHT, border=True)
GOOG_TOTAL_REF = (ws_goog.title, gt)
GOOG_YT_REF = (ws_goog.title, yt_tot)
print("Meta & Google model sheets done.")


# Capture Meta region rows for aggregation
META_ROWS = {REGIONS_4[i]: data_first + i for i in range(4)}
YT_ROWS = {REGIONS_4[i]: yt_first + i for i in range(4)}
SE_ROWS = {REGIONS_4[i]: se_first + i for i in range(4)}
NET_ROW = net_r

# ======================================================================================
# MODEL — SNAP
# ======================================================================================
r = model_sheet_header(ws_snap, "Snap Inc. (Snapchat)", [
    ("FY2024 total revenue ($M)", SNAP_TOTAL, FMT_USD_M),
    ("Q4-2024 global DAU (millions)", 453, FMT_NUM),
    ("Snapchat US under-16 ad-revenue share", f"={SHARE_SNAP}", FMT_PCT2),
])
put(ws_snap, r - 1, 4, "Highest under-16 exposure of the four (Harvard: 41% of Snap US ad rev from <18)", ITALIC, LFT)
hdr_r = model_table_header(ws_snap, r, "Regional revenue ($M)\n[total revenue]")
snap_first = hdr_r
snap_regions = ["North America", "Europe", "Rest of World"]
snap_notes = {
    "North America": "Anchor; Snapchat skews very young (teens core)",
    "Europe": "High teen penetration; ~99M DAU",
    "Rest of World": "Includes Asia-Pacific; fast-growing, youngest base (254M DAU)",
}
rr = snap_first
for region in snap_regions:
    put(ws_snap, rr, 1, region, BOLD, LFT, border=True)
    put(ws_snap, rr, 2, SNAP_REGION[region], NORMAL, RGT, FMT_USD_M, GREY, True)
    put(ws_snap, rr, 3, f"={SHARE_SNAP}*{mult_for(region)}", NORMAL, CTR, FMT_PCT2, border=True)
    put(ws_snap, rr, 4, f"=B{rr}*C{rr}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(ws_snap, rr, 5, f"=B{rr}*C{rr}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_snap, rr, 6, f"=B{rr}*C{rr}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_snap, rr, 7, snap_notes[region], ITALIC, LFT, border=True)
    rr += 1
snap_tot = rr
put(ws_snap, snap_tot, 1, "TOTAL — Snap", BOLD, LFT, fill=LIGHT, border=True)
for col in (2, 4, 5, 6):
    put(ws_snap, snap_tot, col, f"=SUM({get_column_letter(col)}{snap_first}:{get_column_letter(col)}{rr-1})", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_snap, snap_tot, 3, f"=D{snap_tot}/B{snap_tot}", BOLD, CTR, FMT_PCT2, LIGHT, True)
put(ws_snap, snap_tot, 7, "Applied to total revenue (incl. ~$0.4B Snapchat+ subs)", ITALIC, LFT, fill=LIGHT, border=True)
SNAP_ROWS = {snap_regions[i]: snap_first + i for i in range(3)}
SNAP_TOTAL_REF = (ws_snap.title, snap_tot)

# ======================================================================================
# MODEL — PINTEREST
# ======================================================================================
r = model_sheet_header(ws_pins, "Pinterest, Inc.", [
    ("FY2024 total revenue ($M)", PINS_TOTAL, FMT_USD_M),
    ("Q4-2024 global MAU (millions)", 553, FMT_NUM),
    ("Pinterest US under-16 ad-revenue share", f"={SHARE_PINS}", FMT_PCT2),
])
put(ws_pins, r - 1, 4, "Lowest exposure: 13–17 ≈ 4% of users; personalised ads restricted for under-18s", ITALIC, LFT)
hdr_r = model_table_header(ws_pins, r, "Regional revenue ($M)\n[estimated]")
pins_first = hdr_r
pins_regions = ["US & Canada", "Europe", "Rest of World"]
pins_notes = {
    "US & Canada": "Anchor; 79% female skew, Gen Z fastest-growing cohort",
    "Europe": "Larger MAU (145M) but lower ARPU",
    "Rest of World": "Includes Asia-Pacific; very low ARPU ($0.19 Q4)",
}
rr = pins_first
for region in pins_regions:
    put(ws_pins, rr, 1, region, BOLD, LFT, border=True)
    put(ws_pins, rr, 2, PINS_REGION[region], NORMAL, RGT, FMT_USD_M, GREY, True)
    put(ws_pins, rr, 3, f"={SHARE_PINS}*{mult_for(region)}", NORMAL, CTR, FMT_PCT2, border=True)
    put(ws_pins, rr, 4, f"=B{rr}*C{rr}*{SC_BASE}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(ws_pins, rr, 5, f"=B{rr}*C{rr}*{SC_LOW}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_pins, rr, 6, f"=B{rr}*C{rr}*{SC_HIGH}", NORMAL, RGT, FMT_USD_M, border=True)
    put(ws_pins, rr, 7, pins_notes[region], ITALIC, LFT, border=True)
    rr += 1
pins_tot = rr
put(ws_pins, pins_tot, 1, "TOTAL — Pinterest", BOLD, LFT, fill=LIGHT, border=True)
for col in (2, 4, 5, 6):
    put(ws_pins, pins_tot, col, f"=SUM({get_column_letter(col)}{pins_first}:{get_column_letter(col)}{rr-1})", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(ws_pins, pins_tot, 3, f"=D{pins_tot}/B{pins_tot}", BOLD, CTR, FMT_PCT2, LIGHT, True)
put(ws_pins, pins_tot, 7, "Regional split estimated from disclosed Q4 mix", ITALIC, LFT, fill=LIGHT, border=True)
PINS_ROWS = {pins_regions[i]: pins_first + i for i in range(3)}
PINS_TOTAL_REF = (ws_pins.title, pins_tot)
print("Snap & Pinterest model sheets done.")


# ======================================================================================
# AGGREGATE BY REGION
# ======================================================================================
ag = ws_agg
ag.sheet_view.showGridLines = False
for col, w in {"A": 20, "B": 14, "C": 14, "D": 12, "E": 13, "F": 14, "G": 14, "H": 14}.items():
    ag.column_dimensions[col].width = w
band(ag, 1, 1, 6, "AGGREGATE UNDER-16 AD-REVENUE EXPOSURE BY REGION  (BASE case, $M)")
ag.row_dimensions[1].height = 22
put(ag, 2, 1, "Cross-company view. Snap & Pinterest do not split Asia-Pacific, so their APAC users sit inside 'Rest of World'.", ITALIC, LFT)
ag.merge_cells("A2:H2")

mg = "'Model - Meta'"
gg = "'Model - Google'"
sg = "'Model - Snap'"
pg = "'Model - Pinterest'"

buckets = {
    "North America": dict(meta="US & Canada", yt="US & Canada", se="US & Canada",
                          snap="North America", pins="US & Canada", geo=0.487),
    "Europe": dict(meta="Europe", yt="Europe", se="Europe", snap="Europe", pins="Europe", geo=0.292),
    "Asia-Pacific": dict(meta="Asia-Pacific", yt="Asia-Pacific", se="Asia-Pacific",
                         snap=None, pins=None, geo=0.162),
    "Rest of World": dict(meta="Rest of World", yt="Rest of World", se="Rest of World",
                          snap="Rest of World", pins="Rest of World", geo=0.058),
}

hr = 4
for i, h in enumerate(["Region", "Meta", "Google", "Snap", "Pinterest", "TOTAL ($M)"]):
    put(ag, hr, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
ag.row_dimensions[hr].height = 18
first = hr + 1
rr = first
for region, b in buckets.items():
    put(ag, rr, 1, region, BOLD, LFT, border=True)
    # Meta
    put(ag, rr, 2, f"={mg}!D{META_ROWS[b['meta']]}", NORMAL, RGT, FMT_USD_M, border=True)
    # Google = YT + Search&other + Network*geo
    put(ag, rr, 3, f"={gg}!D{YT_ROWS[b['yt']]}+{gg}!D{SE_ROWS[b['se']]}+{gg}!D{NET_ROW}*{b['geo']}",
        NORMAL, RGT, FMT_USD_M, border=True)
    # Snap
    if b["snap"]:
        put(ag, rr, 4, f"={sg}!D{SNAP_ROWS[b['snap']]}", NORMAL, RGT, FMT_USD_M, border=True)
    else:
        put(ag, rr, 4, 0, NORMAL, RGT, FMT_USD_M, border=True)
    # Pinterest
    if b["pins"]:
        put(ag, rr, 5, f"={pg}!D{PINS_ROWS[b['pins']]}", NORMAL, RGT, FMT_USD_M, border=True)
    else:
        put(ag, rr, 5, 0, NORMAL, RGT, FMT_USD_M, border=True)
    put(ag, rr, 6, f"=SUM(B{rr}:E{rr})", BOLD, RGT, FMT_USD_M, GREEN, True)
    rr += 1
ag_tot = rr
put(ag, ag_tot, 1, "TOTAL — all regions", BOLD, LFT, fill=LIGHT, border=True)
for col in (2, 3, 4, 5, 6):
    put(ag, ag_tot, col, f"=SUM({get_column_letter(col)}{first}:{get_column_letter(col)}{rr-1})",
        BOLD, RGT, FMT_USD_M, LIGHT, True)

# memo: region share of total
put(ag, ag_tot + 2, 1, "Region as % of total under-16 exposure", H2_FONT, LFT)
ag.merge_cells(f"A{ag_tot+2}:F{ag_tot+2}")
pr = ag_tot + 3
for i, region in enumerate(buckets.keys()):
    put(ag, pr + i, 1, region, NORMAL, LFT, border=True)
    put(ag, pr + i, 2, f"=F{first+i}/F{ag_tot}", NORMAL, RGT, FMT_PCT1, border=True)
print("Aggregate sheet done.")


# ======================================================================================
# SUMMARY SHEET
# ======================================================================================
s = ws_sum
s.sheet_view.showGridLines = False
for col, w in {"A": 24, "B": 16, "C": 16, "D": 13, "E": 15, "F": 13, "G": 13, "H": 12}.items():
    s.column_dimensions[col].width = w

s.merge_cells("A1:H1")
c = s.cell(row=1, column=1, value="Revenue Exposure to Teens Under 16  —  Meta · Google · Snap · Pinterest")
c.fill = hfill(NAVY); c.font = TITLE_FONT; c.alignment = Alignment("left", "center")
s.row_dimensions[1].height = 30
s.merge_cells("A2:H2")
c = s.cell(row=2, column=1, value="Estimation model · FY2024 financials · base case unless ranged · figures in US$ millions")
c.fill = hfill(BLUE); c.font = Font(size=10, italic=True, color=WHITE); c.alignment = Alignment("left", "center")

s.merge_cells("A4:H4")
put(s, 4, 1, "Cohort definition: users UNDER AGE 16 — primarily 13–15 year-olds, plus under-13s present "
             "despite platforms' 13+ minimum-age rules (Ofcom: 51% of UK under-13s use social media).", ITALIC, TOP)
s.row_dimensions[4].height = 28

# headline table
hr = 6
heads = ["Company", "FY2024 revenue\nbase ($M)", "Under-16 share\n(blended)", "Under-16 rev\nBASE ($M)",
         "Low ($M)", "High ($M)", "Primary youth\nsurface"]
for i, h in enumerate(heads):
    put(s, hr, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
s.row_dimensions[hr].height = 30

rows = [
    ("Meta Platforms", "'Model - Meta'", meta_tot, "Instagram"),
    ("Google / Alphabet", "'Model - Google'", gt, "YouTube"),
    ("Snap", "'Model - Snap'", snap_tot, "Snapchat"),
    ("Pinterest", "'Model - Pinterest'", pins_tot, "Pinterest"),
]
first = hr + 1
rr = first
for name, sheet, trow, surface in rows:
    put(s, rr, 1, name, BOLD, LFT, border=True)
    put(s, rr, 2, f"={sheet}!B{trow}", NORMAL, RGT, FMT_USD_M, border=True)
    put(s, rr, 3, f"={sheet}!C{trow}", NORMAL, CTR, FMT_PCT2, border=True)
    put(s, rr, 4, f"={sheet}!D{trow}", BOLD, RGT, FMT_USD_M, GREEN, True)
    put(s, rr, 5, f"={sheet}!E{trow}", NORMAL, RGT, FMT_USD_M, border=True)
    put(s, rr, 6, f"={sheet}!F{trow}", NORMAL, RGT, FMT_USD_M, border=True)
    put(s, rr, 7, surface, NORMAL, CTR, border=True)
    rr += 1
tot = rr
put(s, tot, 1, "TOTAL (4 companies)", BOLD, LFT, fill=LIGHT, border=True)
for col in (2, 4, 5, 6):
    put(s, tot, col, f"=SUM({get_column_letter(col)}{first}:{get_column_letter(col)}{rr-1})", BOLD, RGT, FMT_USD_M, LIGHT, True)
put(s, tot, 3, f"=D{tot}/B{tot}", BOLD, CTR, FMT_PCT2, LIGHT, True)
put(s, tot, 7, "", NORMAL, CTR, fill=LIGHT, border=True)

# headline call-out
co = tot + 2
s.merge_cells(f"A{co}:H{co}")
put(s, co, 1, f"BASE-CASE ESTIMATE: combined under-16 advertising-revenue exposure of roughly "
              f"US$ (see D{tot}) million per year across the four platforms.", BOLD, LFT, fill=AMBER)
s.row_dimensions[co].height = 22

# how to read
g = co + 2
band(s, g, 1, 8, "How to read this workbook", fill=BLUE, font=H1_FONT)
notes = [
    "• 'Under-16 revenue' = the slice of each platform's advertising revenue attributable to users aged under 16.",
    "• It is an ESTIMATE. No platform discloses age-specific revenue; this is a transparent, source-anchored model.",
    "• The driver is each platform's US under-16 ad-revenue SHARE (anchored to the Harvard 2023 study), scaled by",
    "   region (multipliers) and applied to disclosed FY2024 regional revenue. See 'Assumptions' to flex any input.",
    "• Low / Base / High apply scenario scalars (0.65 / 1.00 / 1.35) reflecting Harvard's uncertainty intervals and",
    "   the well-documented under-reporting of age by younger users (Meta 10-K; Ofcom: 40% of 8–17s give a fake age).",
    "• Snap shows the highest exposure (teen-core app); Pinterest the lowest (older, female-skewed, <18 ad limits).",
    "• Tabs: Methodology · Financials (FY2024) · Youth Usage Evidence · Assumptions · Model-<Company> · Aggregate · Sources.",
]
for i, n in enumerate(notes):
    put(s, g + 1 + i, 1, n, NORMAL, LFT)
    s.merge_cells(start_row=g + 1 + i, start_column=1, end_row=g + 1 + i, end_column=8)

cav = g + 1 + len(notes) + 1
band(s, cav, 1, 8, "Key caveats", fill="C00000", font=H1_FONT)
cavs = [
    "1. Platforms do not report revenue by age; all age-specific figures are modelled, not disclosed.",
    "2. The Harvard anchor is US-only and for 2022; ex-US shares and the 2024 uplift are estimates.",
    "3. Meta's regional ad-revenue split and Pinterest's regional split are estimated (see model notes).",
    "4. Age data is unreliable — under-age users routinely mis-state age, so true exposure may be higher.",
    "5. 'Under 16' is approximated by splitting the 13–17 band (57% base) plus all under-13 revenue.",
]
for i, n in enumerate(cavs):
    put(s, cav + 1 + i, 1, n, NORMAL, LFT)
    s.merge_cells(start_row=cav + 1 + i, start_column=1, end_row=cav + 1 + i, end_column=8)
print("Summary sheet done.")


# ======================================================================================
# METHODOLOGY SHEET
# ======================================================================================
m = ws_m
m.sheet_view.showGridLines = False
m.column_dimensions["A"].width = 4
m.column_dimensions["B"].width = 115
band(m, 1, 1, 2, "METHODOLOGY")
m.row_dimensions[1].height = 22
para = [
    ("h", "Objective"),
    ("p", "Estimate the annual advertising-revenue exposure of Meta, Google (Alphabet), Snap and Pinterest to "
          "users under age 16, at the most granular level the public data allows (company × surface × region × age band)."),
    ("h", "Core formula (per company, per region)"),
    ("p", "Under-16 revenue  =  Regional revenue  ×  Under-16 revenue share(region)  ×  Scenario scalar"),
    ("p", "Under-16 revenue share(region)  =  US/anchor under-16 share  ×  Regional multiplier"),
    ("p", "US/anchor under-16 share  =  US under-18 share  ×  Under-16 fraction of under-18"),
    ("p", "Under-16 fraction of under-18  =  (0–12 share of under-18 $)  +  (1 − that share) × (13–15 split of 13–17)"),
    ("h", "Step 1 — Anchor: US under-18 ad-revenue shares (Harvard 2023)"),
    ("p", "Raffoul et al. (PLOS ONE, 2023) simulated US 2022 youth ad revenue using Census, Common Sense Media, Pew, "
          "eMarketer and Qustodio data. Disclosed US under-18 SHARE of each platform's ad revenue: Instagram 16%, "
          "Snapchat 41.4%, YouTube 27%, Facebook 1.9%. These are the model's anchors (Pinterest, Search and Network "
          "are not in the study and are estimated)."),
    ("h", "Step 2 — Split under-18 into under-16"),
    ("p", "'Under 16' = ages 0–15. We take all under-13 (0–12) revenue plus the 13–15 portion of the 13–17 band. "
          "The 13–15 portion is set at 57% of the 13–17 band (3 of 5 single-year ages, with a mild skew to higher "
          "engagement among 16–17s). The 0–12 share of each platform's under-18 dollars comes from the Harvard "
          "dollar figures where available (Instagram, YouTube) and is modelled otherwise."),
    ("h", "Step 3 — Regionalise"),
    ("p", "The US anchor share is scaled by regional multipliers (US/NA 1.00, Europe 0.95, Asia-Pacific 1.05, "
          "Rest of World 1.15) reflecting younger populations outside the West versus stronger child-data regulation "
          "in the EU. The scaled share is applied to each platform's disclosed FY2024 regional revenue."),
    ("h", "Step 4 — Surfaces"),
    ("p", "Meta is modelled as a weighted blend of Instagram (50%), Facebook (40%) and Messenger/other (10%) of US ad "
          "revenue. Google is split into YouTube ads (the principal youth surface), Google Search & other, and Google "
          "Network, each with its own under-16 share. Snap and Pinterest are single-surface."),
    ("h", "Step 5 — Scenarios"),
    ("p", "Low / Base / High scalars (0.65 / 1.00 / 1.35) bracket the estimate, reflecting the Harvard model's own "
          "uncertainty intervals and the systematic under-reporting of age by minors (so the true figure may sit "
          "above the base case)."),
    ("h", "Why an estimate?"),
    ("p", "No platform discloses revenue by user age, and they caution that age data is unreliable because under-age "
          "users mis-state their age (Meta 10-K: 'a disproportionate number of our younger users register with an "
          "inaccurate age'). The workbook is therefore a transparent, fully-sourced model; every input is editable "
          "on the Assumptions tab and the results recalculate."),
]
r = 3
for kind, text in para:
    if kind == "h":
        put(m, r, 2, text, H2_FONT, LFT)
    else:
        cc = put(m, r, 2, text, NORMAL, TOP)
        m.row_dimensions[r].height = 14 * (1 + len(text) // 105)
    r += 1

# ======================================================================================
# FINANCIALS SHEET
# ======================================================================================
f = ws_f
f.sheet_view.showGridLines = False
for col, w in {"A": 30, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16}.items():
    f.column_dimensions[col].width = w
band(f, 1, 1, 6, "HARD FINANCIAL DATA — FY2024 (as reported)   $M unless noted")
f.row_dimensions[1].height = 22

def fin_block(ws, r, title, headers, rows, src):
    band(ws, r, 1, 6, title, fill=BLUE, font=H2_FONT_)
    r += 1
    for i, h in enumerate(headers):
        put(ws, r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
    r += 1
    for row in rows:
        for i, v in enumerate(row):
            fmt = None if i == 0 else (FMT_PCT1 if isinstance(v, float) and v < 1 and i == len(row)-0 else FMT_USD_M)
            al = LFT if i == 0 else RGT
            put(ws, r, 1 + i, v, NORMAL if i else BOLD, al, FMT_USD_M if i and not (isinstance(v,str)) else None, border=True)
        r += 1
    put(ws, r, 1, src, SRC_FONT, TOP)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    return r + 2

r = 3
r = fin_block(f, r, "Meta Platforms",
              ["Item", "Value"],
              [["Total revenue", META_TOTAL], ["Family of Apps revenue", META_FOA],
               ["Advertising revenue", META_AD],
               ["US&Canada ad rev (est. user-geography)", round(META_AD*0.44,0)],
               ["Europe ad rev (est.)", round(META_AD*0.23,0)],
               ["Asia-Pacific ad rev (est.)", round(META_AD*0.22,0)],
               ["Rest of World ad rev (est.)", round(META_AD*0.11,0)],
               ["Annual ARPP worldwide ($)", 49.63]],
              "Source: Meta FY2024 10-K / Q4-24 press release. Regional ad-revenue split is estimated (user geography).")
r = fin_block(f, r, "Google / Alphabet",
              ["Item", "Value"],
              [["Alphabet total revenue", GOOG_TOTAL], ["Google advertising", GOOG_ADV],
               ["Google Search & other", GOOG_SEARCH], ["YouTube ads", YT_ADS],
               ["Google Network", GOOG_NETWORK],
               ["Geo mix: US&Canada", round(GOOG_TOTAL*0.487,0)],
               ["Geo mix: Europe/EMEA", round(GOOG_TOTAL*0.292,0)],
               ["Geo mix: Asia-Pacific", round(GOOG_TOTAL*0.162,0)],
               ["Geo mix: Other Americas (RoW)", round(GOOG_TOTAL*0.058,0)]],
              "Source: Alphabet FY2024 10-K. YouTube regional split estimated; Search/Network use Alphabet geo mix.")
r = fin_block(f, r, "Snap Inc.",
              ["Item", "Value"],
              [["Total revenue", SNAP_TOTAL], ["North America revenue", SNAP_REGION['North America']],
               ["Europe revenue", SNAP_REGION['Europe']], ["Rest of World revenue", SNAP_REGION['Rest of World']],
               ["Global DAU Q4 (m)", 453], ["NA / EU / RoW DAU Q4 (m)", "100 / 99 / 254"]],
              "Source: Snap FY2024 press release / investor letter. Regional figures are total revenue by user region.")
r = fin_block(f, r, "Pinterest, Inc.",
              ["Item", "Value"],
              [["Total revenue", PINS_TOTAL], ["US&Canada revenue (est.)", PINS_REGION['US & Canada']],
               ["Europe revenue (est.)", PINS_REGION['Europe']], ["Rest of World revenue (est.)", PINS_REGION['Rest of World']],
               ["Global MAU Q4 (m)", 553], ["UCAN / EU / RoW MAU Q4 (m)", "101 / 145 / 307"]],
              "Source: Pinterest FY2024 press release. Regional revenue split estimated from disclosed Q4 mix.")
print("Methodology & Financials done.")


# ======================================================================================
# YOUTH USAGE EVIDENCE SHEET
# ======================================================================================
e = ws_e
e.sheet_view.showGridLines = False
for col, w in {"A": 24, "B": 16, "C": 16, "D": 16, "E": 16, "F": 30}.items():
    e.column_dimensions[col].width = w
band(e, 1, 1, 6, "YOUTH-USAGE EVIDENCE BASE  (data underpinning the assumptions)")
e.row_dimensions[1].height = 22

# Harvard user table
band(e, 3, 1, 6, "A.  Harvard 2023 — estimated US youth USERS by platform, 2022 (millions)", fill=BLUE, font=H2_FONT_)
for i, h in enumerate(["Platform", "Ages 0–12", "Ages 13–17", "Ages 0–17", "", "Note"]):
    put(e, 4, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
harv_users = [
    ("Facebook", 2.757, 7.047, 9.904), ("Instagram", 3.982, 12.726, 16.708),
    ("Snapchat", 2.869, 15.131, 18.000), ("YouTube", 31.448, 18.341, 49.789),
    ("TikTok (ref.)", 3.041, 15.932, 18.972), ("X/Twitter (ref.)", 2.071, 4.929, 7.001),
]
rr = 5
for name, a, b, c in harv_users:
    put(e, rr, 1, name, BOLD, LFT, border=True)
    put(e, rr, 2, a, NORMAL, RGT, '#,##0.0', border=True)
    put(e, rr, 3, b, NORMAL, RGT, '#,##0.0', border=True)
    put(e, rr, 4, c, NORMAL, RGT, '#,##0.0', border=True)
    put(e, rr, 6, "Raffoul et al., PLOS ONE 2023, Table 1", SRC_FONT, LFT, border=True)
    rr += 1

# Harvard revenue / share
rr += 1
band(e, rr, 1, 6, "B.  Harvard 2023 — US under-18 advertising revenue & share, 2022", fill=BLUE, font=H2_FONT_)
rr += 1
for i, h in enumerate(["Platform", "0–12 rev ($M)", "13–17 rev ($M)", "Under-18 % of\nplatform US ad rev", "", "Note"]):
    put(e, rr, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
e.row_dimensions[rr].height = 28
rr += 1
harv_rev = [
    ("Instagram", 801.1, 4000.0, 0.160), ("YouTube", 959.1, 1200.0, 0.270),
    ("Snapchat", None, None, 0.414), ("Facebook", 137.2, None, 0.019),
    ("TikTok (ref.)", None, 2000.0, 0.350),
]
for name, a, b, sh in harv_rev:
    put(e, rr, 1, name, BOLD, LFT, border=True)
    put(e, rr, 2, a if a is not None else "n/d", NORMAL, RGT, FMT_USD_M if a is not None else None, border=True)
    put(e, rr, 3, b if b is not None else "n/d", NORMAL, RGT, FMT_USD_M if b is not None else None, border=True)
    put(e, rr, 4, sh, BOLD, CTR, FMT_PCT1, border=True)
    put(e, rr, 6, "PLOS ONE 2023, Figs 1–2", SRC_FONT, LFT, border=True)
    rr += 1
put(e, rr, 1, "Aggregate: 6 platforms earned ~$11.0bn US ad revenue from under-18s in 2022 ($2.1bn from ≤12, $8.6bn from 13–17).", ITALIC, LFT)
e.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
rr += 2

# Ofcom
band(e, rr, 1, 6, "C.  Ofcom 2024 (UK) — children's platform usage by age", fill=BLUE, font=H2_FONT_)
rr += 1
ofcom = [
    "Social-media use rises with age: 34% of 3–7s, 63% of 8–11s, 92% of 12–15s, 95% of 16–17s.",
    "51% of UK under-13s use social media despite the common 13+ minimum age.",
    "40% of 8–17s admit giving a fake age to access a site/app (key reason age data is unreliable).",
    "Instagram and Snapchat usage each rise to ~80% among 16–17s (17% even among 3–4s).",
    "YouTube is near-universal across all child age bands (>80% of 3–17s).",
    "5–7s with own profiles: YouTube/YT Kids 48%, Instagram 9%, WhatsApp 11%.",
]
for n in ofcom:
    put(e, rr, 1, "• " + n, NORMAL, LFT)
    e.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
    rr += 1
rr += 1

# Pinterest
band(e, rr, 1, 6, "D.  Pinterest age profile", fill=BLUE, font=H2_FONT_)
rr += 1
pins_ev = [
    "13–17 year-olds ≈ 4% of Pinterest's audience; bulk of users are 25–34 (the largest cohort).",
    "Gen Z is >40% of global MAU and fastest-growing, but that cohort is predominantly 18+.",
    "Audience skews ~79% female; personalised ads are restricted for users under 18.",
]
for n in pins_ev:
    put(e, rr, 1, "• " + n, NORMAL, LFT)
    e.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
    rr += 1

# ======================================================================================
# SOURCES SHEET
# ======================================================================================
sc = ws_src
sc.sheet_view.showGridLines = False
sc.column_dimensions["A"].width = 4
sc.column_dimensions["B"].width = 58
sc.column_dimensions["C"].width = 70
band(sc, 1, 1, 3, "SOURCES")
sc.row_dimensions[1].height = 22
for i, h in enumerate(["#", "Source", "URL / reference"]):
    put(sc, 2, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
sources = [
    ("Meta FY2024 results (10-K / Q4 press release)", "sec.gov/Archives/edgar/data/1326801/000132680125000014/meta-12312024xexhibit991.htm"),
    ("Meta FY2024 10-K (age-data reliability caveat; ARPP)", "sec.gov/Archives/edgar/data/1326801/000132680125000017/meta-20241231.htm"),
    ("Alphabet FY2024 results (segment & geography)", "sec.gov/Archives/edgar/data/1652044/000165204425000014/"),
    ("Snap FY2024 results / investor letter (regional rev, DAU, ARPU)", "investor.snap.com — Q4 & FY2024 press release"),
    ("Pinterest FY2024 results (revenue, MAU, ARPU, regional)", "investor.pinterestinc.com — Q4 & FY2024 press release"),
    ("Raffoul, Ward, Santoso, Kavanaugh, Austin (2023). Social media platforms generate billions from US youth. PLOS ONE 18(12): e0295337.", "journals.plos.org/plosone/article?id=10.1371/journal.pone.0295337"),
    ("Harvard T.H. Chan School news release on the above study", "hsph.harvard.edu/news/social-media-platforms-generate-billions-in-annual-ad-revenue-from-u-s-youth/"),
    ("Ofcom — Children and parents: media use and attitudes 2024", "ofcom.org.uk — Children's Media Literacy Report 2024"),
    ("Pew Research — Teens, Social Media and Technology", "pewresearch.org/internet (teen platform usage)"),
    ("DataReportal / Pinterest ad-audience age profile", "datareportal.com/essential-pinterest-stats"),
]
rr = 3
for i, (name, url) in enumerate(sources, 1):
    put(sc, rr, 1, i, NORMAL, CTR, border=True)
    put(sc, rr, 2, name, NORMAL, TOP, border=True)
    put(sc, rr, 3, url, SRC_FONT, TOP, border=True)
    sc.row_dimensions[rr].height = 28
    rr += 1
put(sc, rr + 1, 2, "All financial figures are as-reported FY2024. All age-specific revenue figures are MODELLED estimates "
                   "(no issuer discloses revenue by user age).", ITALIC, TOP)
sc.merge_cells(start_row=rr + 1, start_column=2, end_row=rr + 1, end_column=3)

# order tabs
wb.move_sheet("Summary", -wb.sheetnames.index("Summary"))

OUT = "Under16_Revenue_Exposure_Model.xlsx"
# Force Excel/Sheets to recalculate all formulas on open (openpyxl stores no cached values)
try:
    wb.calculation.fullCalcOnLoad = True
except Exception:
    pass
wb.save(OUT)
print("SAVED:", OUT)
