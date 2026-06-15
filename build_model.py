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

# The under-16 exposure model is now based on FY2025 actuals (latest full year).
# FY2024 actuals are retained for the growth comparison on the Financials tab.

# ---- Meta (advertising revenue by USER geography; regional split is an estimate) ----
META_AD = 196175.0          # FY2025 advertising revenue, $M (10-K)
META_AD_2024 = 160633.0
META_TOTAL = 200966.0
META_TOTAL_2024 = 164501.0
META_FOA = 198759.0
META_FOA_2024 = 162355.0
# Estimated user-geography ad-revenue shares (documented assumption; APAC nudged up as
# Meta cited Asia-Pacific as the fastest-growing region for ad impressions in 2025)
META_REGION_SHARE = {"US & Canada": 0.43, "Europe": 0.23, "Asia-Pacific": 0.23, "Rest of World": 0.11}
META_ARPP = {"US & Canada": 233.42, "Europe": 68.12, "Asia-Pacific": 21.28, "Rest of World": 14.00}

# ---- Alphabet / Google (FY2025) ----
GOOG_TOTAL = 402836.0
GOOG_TOTAL_2024 = 350018.0
GOOG_ADV = 294691.0
GOOG_ADV_2024 = 264590.0
YT_ADS = 40367.0
YT_ADS_2024 = 36147.0
GOOG_SEARCH = 224532.0
GOOG_SEARCH_2024 = 198084.0
GOOG_NETWORK = 29792.0
GOOG_NETWORK_2024 = 30359.0
# YouTube ad-revenue regional split (estimate: YouTube monetises internationally more
# than Alphabet's overall US-billed geographic mix implies)
YT_REGION_SHARE = {"US & Canada": 0.35, "Europe": 0.30, "Asia-Pacific": 0.23, "Rest of World": 0.12}
# Alphabet total-revenue geographic mix (FY2025 10-K) used for Search & Network
GOOG_GEO = {"US & Canada": 0.482, "Europe": 0.291, "Asia-Pacific": 0.168, "Rest of World": 0.059}

# ---- Snap (total revenue by region, exact from press release, $M) ----
SNAP_TOTAL = 5931.0
SNAP_TOTAL_2024 = 5361.0
SNAP_REGION = {"North America": 3575.6, "Europe": 1128.4, "Rest of World": 1227.4}   # FY2025 (TTM Q4-25)
SNAP_REGION_2024 = {"North America": 3337.3, "Europe": 961.6, "Rest of World": 1062.5}
SNAP_DAU = {"North America": 94.0, "Europe": 98.0, "Rest of World": 282.0}  # Q4-25 millions

# ---- Pinterest (total revenue; FY2025 full-year regional revenue, as reported) ----
PINS_TOTAL = 4222.0
PINS_TOTAL_2024 = 3646.0
PINS_REGION = {"US & Canada": 3173.0, "Europe": 775.0, "Rest of World": 274.0}     # FY2025 (reported)
PINS_REGION_2024 = {"US & Canada": 2884.0, "Europe": 593.0, "Rest of World": 169.0}  # FY2024 (reported)
PINS_MAU = {"US & Canada": 105.0, "Europe": 158.0, "Rest of World": 356.0}  # Q4-25 millions

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
ws_f = wb.create_sheet("Financials")

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
ws_reg = wb.create_sheet("Regulatory Headwind")
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
    ("FY2025 total revenue ($M)", META_TOTAL, FMT_USD_M),
    ("FY2025 advertising revenue ($M)", META_AD, FMT_USD_M),
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
    ("FY2025 Alphabet total revenue ($M)", GOOG_TOTAL, FMT_USD_M),
    ("FY2025 Google advertising revenue ($M)", GOOG_ADV, FMT_USD_M),
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
    ("FY2025 total revenue ($M)", SNAP_TOTAL, FMT_USD_M),
    ("Q4-2025 global DAU (millions)", 474, FMT_NUM),
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
    ("FY2025 total revenue ($M)", PINS_TOTAL, FMT_USD_M),
    ("Q4-2025 global MAU (millions)", 619, FMT_NUM),
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
c = s.cell(row=2, column=1, value="Estimation model · FY2025 financials · base case unless ranged · figures in US$ millions")
c.fill = hfill(BLUE); c.font = Font(size=10, italic=True, color=WHITE); c.alignment = Alignment("left", "center")

s.merge_cells("A4:H4")
put(s, 4, 1, "Cohort definition: users UNDER AGE 16 — primarily 13–15 year-olds, plus under-13s present "
             "despite platforms' 13+ minimum-age rules (Ofcom: 51% of UK under-13s use social media).", ITALIC, TOP)
s.row_dimensions[4].height = 28

# headline table
hr = 6
heads = ["Company", "FY2025 revenue\nbase ($M)", "Under-16 share\n(blended)", "Under-16 rev\nBASE ($M)",
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
    "   region (multipliers) and applied to disclosed FY2025 regional revenue. See 'Assumptions' to flex any input.",
    "• Low / Base / High apply scenario scalars (0.65 / 1.00 / 1.35) reflecting Harvard's uncertainty intervals and",
    "   the well-documented under-reporting of age by younger users (Meta 10-K; Ofcom: 40% of 8–17s give a fake age).",
    "• Snap shows the highest exposure (teen-core app); Pinterest the lowest (older, female-skewed, <18 ad limits).",
    "• Tabs: Methodology · Financials · Youth Usage Evidence · Assumptions · Model-<Company> · Aggregate · Regulatory Headwind · Sources.",
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
          "in the EU. The scaled share is applied to each platform's disclosed FY2025 regional revenue."),
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
band(f, 1, 1, 4, "HARD FINANCIAL DATA — FY2024 vs FY2025 (as reported)   $M unless noted")
f.row_dimensions[1].height = 22

def fin_block(ws, r, title, rows, src):
    band(ws, r, 1, 4, title, fill=BLUE, font=H2_FONT_)
    r += 1
    for i, h in enumerate(["Item", "FY2024", "FY2025", "YoY %"]):
        put(ws, r, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
    r += 1
    for label, v24, v25 in rows:
        put(ws, r, 1, label, BOLD, LFT, border=True)
        num = not (isinstance(v24, str) or isinstance(v25, str))
        put(ws, r, 2, v24, NORMAL, RGT, FMT_USD_M if num else None, border=True)
        put(ws, r, 3, v25, NORMAL, RGT, FMT_USD_M if num else None, border=True)
        if num and v24:
            put(ws, r, 4, f"=C{r}/B{r}-1", NORMAL, CTR, FMT_PCT1, border=True)
        else:
            put(ws, r, 4, "", NORMAL, CTR, border=True)
        r += 1
    put(ws, r, 1, src, SRC_FONT, TOP)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    return r + 2

r = 3
r = fin_block(f, r, "Meta Platforms",
              [["Total revenue", META_TOTAL_2024, META_TOTAL],
               ["Family of Apps revenue", META_FOA_2024, META_FOA],
               ["Advertising revenue", META_AD_2024, META_AD],
               ["US & Canada ad rev (est. user-geography)", round(META_AD_2024*0.44), round(META_AD*0.43)],
               ["Europe ad rev (est.)", round(META_AD_2024*0.23), round(META_AD*0.23)],
               ["Asia-Pacific ad rev (est.)", round(META_AD_2024*0.22), round(META_AD*0.23)],
               ["Rest of World ad rev (est.)", round(META_AD_2024*0.11), round(META_AD*0.11)]],
              "Source: Meta FY2024 & FY2025 10-Ks. Regional ad-revenue split is estimated (user geography).")
r = fin_block(f, r, "Google / Alphabet",
              [["Alphabet total revenue", GOOG_TOTAL_2024, GOOG_TOTAL],
               ["Google advertising", GOOG_ADV_2024, GOOG_ADV],
               ["Google Search & other", GOOG_SEARCH_2024, GOOG_SEARCH],
               ["YouTube ads", YT_ADS_2024, YT_ADS],
               ["Google Network", GOOG_NETWORK_2024, GOOG_NETWORK],
               ["Geo: United States", round(GOOG_TOTAL_2024*0.487), round(GOOG_TOTAL*0.482)],
               ["Geo: EMEA", round(GOOG_TOTAL_2024*0.292), round(GOOG_TOTAL*0.291)],
               ["Geo: APAC", round(GOOG_TOTAL_2024*0.162), round(GOOG_TOTAL*0.168)],
               ["Geo: Other Americas", round(GOOG_TOTAL_2024*0.058), round(GOOG_TOTAL*0.059)]],
              "Source: Alphabet FY2024 & FY2025 10-Ks. YouTube regional split estimated; Search/Network use Alphabet geo mix.")
r = fin_block(f, r, "Snap Inc.",
              [["Total revenue", SNAP_TOTAL_2024, SNAP_TOTAL],
               ["North America revenue", SNAP_REGION_2024['North America'], SNAP_REGION['North America']],
               ["Europe revenue", SNAP_REGION_2024['Europe'], SNAP_REGION['Europe']],
               ["Rest of World revenue", SNAP_REGION_2024['Rest of World'], SNAP_REGION['Rest of World']],
               ["Global DAU Q4 (m)", 453, 474],
               ["NA / EU / RoW DAU Q4 (m)", "100 / 99 / 254", "94 / 98 / 282"]],
              "Source: Snap FY2024 & FY2025 press releases. Regional figures are total revenue by user region.")
r = fin_block(f, r, "Pinterest, Inc.",
              [["Total revenue", PINS_TOTAL_2024, PINS_TOTAL],
               ["US & Canada revenue", PINS_REGION_2024['US & Canada'], PINS_REGION['US & Canada']],
               ["Europe revenue", PINS_REGION_2024['Europe'], PINS_REGION['Europe']],
               ["Rest of World revenue", PINS_REGION_2024['Rest of World'], PINS_REGION['Rest of World']],
               ["Global MAU Q4 (m)", 553, 619],
               ["UCAN / EU / RoW MAU Q4 (m)", "101 / 145 / 307", "105 / 158 / 356"]],
              "Source: Pinterest FY2024 & FY2025 press releases (full-year regional revenue, as reported).")
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
    ("Meta FY2025 results (10-K / Q4-25 press release, Jan 2026)", "sec.gov/Archives/edgar/data/1326801/000162828026003832/meta-12312025xexhibit991.htm"),
    ("Meta FY2024 results (prior-year comparison)", "sec.gov/Archives/edgar/data/1326801/000132680125000014/meta-12312024xexhibit991.htm"),
    ("Alphabet FY2025 results (segment & geography, Feb 2026)", "sec.gov/Archives/edgar/data/1652044/000165204426000018/"),
    ("Snap FY2025 results (regional rev, DAU, ARPU, Feb 2026)", "investor.snap.com — Q4 & FY2025 press release"),
    ("Pinterest FY2025 results (full-year regional revenue, MAU, Feb 2026)", "investor.pinterestinc.com — Q4 & FY2025 press release"),
    ("Raffoul, Ward, Santoso, Kavanaugh, Austin (2023). Social media platforms generate billions from US youth. PLOS ONE 18(12): e0295337.", "journals.plos.org/plosone/article?id=10.1371/journal.pone.0295337"),
    ("Harvard T.H. Chan School news release on the above study", "hsph.harvard.edu/news/social-media-platforms-generate-billions-in-annual-ad-revenue-from-u-s-youth/"),
    ("Ofcom — Children and parents: media use and attitudes 2024", "ofcom.org.uk — Children's Media Literacy Report 2024"),
    ("eSafety Commissioner (AU) — Social Media Minimum Age: which platforms are age-restricted", "esafety.gov.au/.../social-media-age-restrictions/which-platforms-are-age-restricted"),
    ("eSafety Commissioner (AU) — Social Media Minimum Age FAQ (penalties, excluded services)", "esafety.gov.au/.../social-media-age-restrictions/faqs"),
    ("AVPA — EU countries / US state age-assurance trackers (2026)", "avpassociation.com — EU & US state age-assurance laws"),
    ("TechCrunch / Wired-Parents — global under-16 social-media ban trackers (2026)", "techcrunch.com/2026/04/23/social-media-ban-children-countries-list/"),
    ("DataReportal / Pinterest ad-audience age profile", "datareportal.com/essential-pinterest-stats"),
]
rr = 3
for i, (name, url) in enumerate(sources, 1):
    put(sc, rr, 1, i, NORMAL, CTR, border=True)
    put(sc, rr, 2, name, NORMAL, TOP, border=True)
    put(sc, rr, 3, url, SRC_FONT, TOP, border=True)
    sc.row_dimensions[rr].height = 28
    rr += 1
put(sc, rr + 1, 2, "All financial figures are as-reported FY2024/FY2025. All age-specific revenue and ban-headwind "
                   "figures are MODELLED estimates (no issuer discloses revenue by user age or by jurisdiction at this granularity).", ITALIC, TOP)
sc.merge_cells(start_row=rr + 1, start_column=2, end_row=rr + 1, end_column=3)

# ======================================================================================
# REGULATORY HEADWIND SHEET
# ======================================================================================
rg = ws_reg
rg.sheet_view.showGridLines = False
for col, w in {"A": 26, "B": 34, "C": 11, "D": 10, "E": 10, "F": 9, "G": 13, "H": 13}.items():
    rg.column_dimensions[col].width = w
band(rg, 1, 1, 8, "REGULATORY HEADWIND — UNDER-16 SOCIAL-MEDIA AGE BANS")
rg.row_dimensions[1].height = 22
rg.merge_cells("A2:H2")
put(rg, 2, 1, "Quantifies the revenue / growth headwind from under-16 bans. 'At-risk' revenue = the slice of "
              "under-16 revenue a ban is estimated to remove. All inputs (blue) are editable.", ITALIC, LFT)
rg.row_dimensions[2].height = 26

# ---- 1. Australia coverage ----
band(rg, 4, 1, 8, "1.  Australia — Social Media Minimum Age Act (in force 10 Dec 2025; minimum age 16; fines up to A$49.5m)",
     fill=BLUE, font=H2_FONT_)
for i, h in enumerate(["Issuer", "In scope?", "Covered surface(s)", "Excluded / note"]):
    put(rg, 5, 1 + i if i < 3 else 4, h, HDR_FONT, CTR, fill=BLUE, border=True)
rg.merge_cells("D5:H5")
cov = [
    ("Meta", "YES", "Facebook, Instagram, Threads", "Messenger & WhatsApp are excluded (messaging)"),
    ("Google", "YES", "YouTube", "YouTube Kids, Search and Network are out of scope"),
    ("Snap", "YES", "Snapchat", "Entire app in scope — most exposed"),
    ("Pinterest", "NO", "— (excluded)", "eSafety: not an 'age-restricted social media platform'"),
]
rr = 6
for name, scope, surf, note in cov:
    put(rg, rr, 1, name, BOLD, LFT, border=True)
    put(rg, rr, 2, scope, BOLD, CTR, fill=(GREEN if scope == "YES" else GREY), border=True)
    put(rg, rr, 3, surf, NORMAL, LFT, border=True)
    put(rg, rr, 4, note, ITALIC, LFT, border=True)
    rg.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=8)
    rr += 1

# ---- 2. Inputs ----
ir = rr + 1
band(rg, ir, 1, 8, "2.  Headwind model inputs (editable)", fill=BLUE, font=H2_FONT_)
ir += 1
inputs = [
    ("Developed-market under-16 multiplier (vs US anchor)", 0.97, "Ban markets (AU/UK/EU/US) skew developed"),
    ("Cohort factor — under-16 regime (AU, UK, ES, PT)", 1.00, "Full under-16 cohort"),
    ("Cohort factor — under-15 regime (DK, FR, GR, NL...)", 0.82, "Removes only 0–14 (keeps 15s)"),
    ("Cohort factor — US (under-14 / parental-consent)", 0.55, "Narrower cohort, consent not a full ban"),
    ("Enforcement effectiveness — Australia (in force)", 0.60, "Age assurance imperfect; VPN/evasion leakage"),
    ("Enforcement effectiveness — EU/UK pipeline", 0.55, "Similar assurance regime once enacted"),
    ("Enforcement effectiveness — US (incl. litigation discount)", 0.25, "Many state laws enjoined on 1st-Amendment grounds"),
]
input_cell = {}
keys = ["devmult", "c16", "c15", "cUS", "eAU", "ePIPE", "eUS"]
for j, (label, val, note) in enumerate(inputs):
    rrr = ir + j
    put(rg, rrr, 1, label, BOLD, LFT, border=True)
    rg.merge_cells(start_row=rrr, start_column=1, end_row=rrr, end_column=2)
    put(rg, rrr, 3, val, NORMAL, CTR, FMT_X, AMBER, True)
    put(rg, rrr, 4, note, ITALIC, LFT, border=True)
    rg.merge_cells(start_row=rrr, start_column=4, end_row=rrr, end_column=8)
    input_cell[keys[j]] = f"$C${rrr}"
DEV = input_cell["devmult"]

# ---- 3 & 4. Per-company at-risk blocks ----
baskets = [
    # name, status, revshare{company}, cohort_key, enf_key, pinterest_covered
    ("Australia", "In force (10 Dec 2025)",
     {"Meta": 0.018, "Google": 0.020, "Snap": 0.025, "Pinterest": 0.020}, "c16", "eAU", 0),
    ("United Kingdom", "Consulting; target ~2027",
     {"Meta": 0.055, "Google": 0.060, "Snap": 0.040, "Pinterest": 0.040}, "c16", "ePIPE", 1),
    ("EU pipeline", "DK/FR/GR/PT passed; ES/NO/NL/IT advancing",
     {"Meta": 0.060, "Google": 0.060, "Snap": 0.050, "Pinterest": 0.050}, "c15", "ePIPE", 1),
    ("US states", "FL/UT/TN enforceable; many enjoined",
     {"Meta": 0.040, "Google": 0.040, "Snap": 0.050, "Pinterest": 0.030}, "cUS", "eUS", 1),
]

companies = [
    # name, exposure_base_ref, exposure_label, total_rev, u16_ref
    ("Meta", f"'Model - Meta'!B{meta_tot}", "ad revenue (FB+IG+Messenger)", META_TOTAL, META_US_SHARE),
    ("Google", f"'Model - Google'!B{yt_tot}", "YouTube ads (only covered surface)", GOOG_TOTAL, SHARE_YT),
    ("Snap", f"'Model - Snap'!B{snap_tot}", "total revenue", SNAP_TOTAL, SHARE_SNAP),
    ("Pinterest", f"'Model - Pinterest'!B{pins_tot}", "total revenue", PINS_TOTAL, SHARE_PINS),
]

cur = ir + len(inputs) + 1
atrisk = {c[0]: {} for c in companies}  # company -> basket -> cell
totalrev_cell = {}
for name, exp_ref, exp_label, total_rev, u16_ref in companies:
    band(rg, cur, 1, 8, f"3.  {name} — exposure base = {exp_label}; FY2025 total revenue ${total_rev:,.0f}m",
         fill=BLUE, font=H2_FONT_)
    cur += 1
    for i, h in enumerate(["Jurisdiction basket", "Status", "Rev share", "Cohort", "Enforce",
                           "Covered", "At-risk ($M)", "bps of rev"]):
        put(rg, cur, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
    rg.row_dimensions[cur].height = 26
    cur += 1
    block_first = cur
    for bname, status, shares, ck, ek, pcov in baskets:
        covered = pcov if name == "Pinterest" else 1
        put(rg, cur, 1, bname, BOLD, LFT, border=True)
        put(rg, cur, 2, status, ITALIC, LFT, border=True)
        put(rg, cur, 3, shares[name], NORMAL, CTR, FMT_PCT1, AMBER, True)
        put(rg, cur, 4, f"={input_cell[ck]}", NORMAL, CTR, FMT_X, border=True)
        put(rg, cur, 5, f"={input_cell[ek]}", NORMAL, CTR, FMT_X, border=True)
        put(rg, cur, 6, covered, NORMAL, CTR, FMT_X, (GREY if covered else "F4CCCC"), True)
        # at-risk = expbase * revshare * u16 * devmult * cohort * enforce * covered
        put(rg, cur, 7, f"={exp_ref}*C{cur}*{u16_ref}*{DEV}*D{cur}*E{cur}*F{cur}",
            BOLD, RGT, FMT_USD_M1, GREEN, True)
        put(rg, cur, 8, f"=G{cur}/{total_rev}*10000", NORMAL, CTR, '#,##0"bp"', border=True)
        atrisk[name][bname] = cur
        cur += 1
    # subtotal across all baskets (= broad scenario)
    put(rg, cur, 1, "All baskets (broad-adoption scenario)", BOLD, LFT, fill=LIGHT, border=True)
    rg.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=6)
    put(rg, cur, 7, f"=SUM(G{block_first}:G{cur-1})", BOLD, RGT, FMT_USD_M1, LIGHT, True)
    put(rg, cur, 8, f"=G{cur}/{total_rev}*10000", BOLD, CTR, '#,##0"bp"', LIGHT, True)
    totalrev_cell[name] = total_rev
    cur += 2

# ---- 5. Scenario summary ----
band(rg, cur, 1, 8, "4.  Scenario summary — under-16-ban headwind by company (At-risk $M  /  bps of FY2025 revenue)",
     fill=NAVY, font=H1_FONT)
cur += 1
scen_defs = [
    ("A. In force today", ["Australia"], "Australia ban only"),
    ("B. Legislated + advancing", ["Australia", "United Kingdom", "EU pipeline"], "+ UK & EU pipeline (~2026–27)"),
    ("C. Broad global adoption", ["Australia", "United Kingdom", "EU pipeline", "US states"], "+ US states"),
]
hdr_cols = ["Scenario", "Meta", "Google", "Snap", "Pinterest", "TOTAL $M", "", "Note"]
for i, h in enumerate(hdr_cols):
    put(rg, cur, 1 + i, h, HDR_FONT, CTR, fill=BLUE, border=True)
cur += 1
scen_first = cur
for sname, blist, note in scen_defs:
    put(rg, cur, 1, sname, BOLD, LFT, border=True)
    comp_cols = {"Meta": 2, "Google": 3, "Snap": 4, "Pinterest": 5}
    for cname, col in comp_cols.items():
        terms = "+".join(f"G{atrisk[cname][b]}" for b in blist)
        put(rg, cur, col, f"={terms}", NORMAL, RGT, FMT_USD_M1, border=True)
    put(rg, cur, 6, f"=SUM(B{cur}:E{cur})", BOLD, RGT, FMT_USD_M1, GREEN, True)
    put(rg, cur, 8, note, ITALIC, LFT, border=True)
    rg.merge_cells(start_row=cur, start_column=8, end_row=cur, end_column=8)
    cur += 1
# bps row for each scenario (vs each company's total revenue)
put(rg, cur, 1, "Headwind in bps of FY2025 revenue", H2_FONT, LFT)
rg.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=8)
cur += 1
for i, (sname, blist, note) in enumerate(scen_defs):
    srow = scen_first + i
    put(rg, cur, 1, sname, NORMAL, LFT, border=True)
    for cname, col, tr in [("Meta", 2, META_TOTAL), ("Google", 3, GOOG_TOTAL),
                           ("Snap", 4, SNAP_TOTAL), ("Pinterest", 5, PINS_TOTAL)]:
        put(rg, cur, col, f"={get_column_letter(col)}{srow}/{tr}*10000", NORMAL, CTR, '#,##0"bp"', border=True)
    put(rg, cur, 8, "Google bps vs Alphabet total; YouTube-only exposure", ITALIC, LFT, border=True)
    cur += 1

# ---- 6. Interpretation ----
cur += 1
band(rg, cur, 1, 8, "5.  Interpretation", fill=BLUE, font=H2_FONT_)
cur += 1
interp = [
    "• Direct Australia impact is small for the giants (Meta low-single-digit bps; Google ~2–3bps of Alphabet) and",
    "   ZERO for Pinterest (excluded). It is most material for Snap (~30–40bps of revenue), whose user base is teen-core.",
    "• Even under broad global adoption (Australia + UK + EU + US states), the headwind is ~tens of bps for Meta/Google",
    "   but can reach ~150bps+ for Snap — set against FY2025 revenue growth of Meta +22%, Google +15%, Snap +11%, Pins +16%.",
    "• The larger risk is not the immediate revenue line but (a) regulatory CONTAGION (40+ countries now active) and",
    "   (b) erosion of the youth USER FUNNEL: losing under-16s removes the top of the future-ARPU pipeline, a multi-year drag.",
    "• Mitigants: bans target accounts, not viewing (logged-out YouTube continues); ad budgets/engagement partly redistribute",
    "   to older users; messaging surfaces (WhatsApp/Messenger) and Pinterest are largely carved out.",
]
for n in interp:
    put(rg, cur, 1, n, NORMAL, LFT)
    rg.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=8)
    cur += 1
REG_SCEN_FIRST = scen_first  # for Summary reference (rows: A,B,C at scen_first..+2; TOTAL col F)

# ======================================================================================
# SUMMARY — append regulatory-headwind callout
# ======================================================================================
gr = cav + 1 + len(cavs) + 2
band(s, gr, 1, 8, "Regulatory headwind — under-16 social-media bans (see 'Regulatory Headwind' tab)",
     fill=BLUE, font=H1_FONT)
gr += 1
reg_tab = "'Regulatory Headwind'"
reg_notes = [
    ("Australia's under-16 ban (in force Dec 2025) covers Facebook, Instagram, Snapchat and YouTube; "
     "Pinterest, Messenger and WhatsApp are EXCLUDED.", None),
    ("Estimated at-risk revenue — Scenario A (Australia, in force today):",
     f"={reg_tab}!F{scen_first}"),
    ("Estimated at-risk revenue — Scenario B (+ UK & EU pipeline, ~2026–27):",
     f"={reg_tab}!F{scen_first+1}"),
    ("Estimated at-risk revenue — Scenario C (+ US states, broad adoption):",
     f"={reg_tab}!F{scen_first+2}"),
]
for i, (txt, ref) in enumerate(reg_notes):
    put(s, gr + i, 1, txt, NORMAL, LFT)
    s.merge_cells(start_row=gr + i, start_column=1, end_row=gr + i, end_column=6)
    if ref:
        put(s, gr + i, 7, ref, BOLD, RGT, FMT_USD_M1, AMBER)
        put(s, gr + i, 8, "$M / yr", ITALIC, LFT)
gr += len(reg_notes)
put(s, gr, 1, "Most exposed: Snap (teen-core). Least: Pinterest (excluded in AU, older audience). "
              "FY2025 growth for context: Meta +22%, Google +15%, Snap +11%, Pinterest +16%.", ITALIC, LFT)
s.merge_cells(start_row=gr, start_column=1, end_row=gr, end_column=8)
print("Regulatory headwind sheet done.")

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
