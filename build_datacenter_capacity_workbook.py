#!/usr/bin/env python3
"""Build datacenter_gw_capacity_us_vs_international.xlsx.

Data-center IT-load capacity (GW), United States vs international, split into
live, under construction, and planned pipeline, with a country political-risk
indicator and a HIGH-risk flag.

Capacity: Cushman & Wakefield H1 2026 regional updates (Americas, EMEA, APAC)
and the 2026 Global Data Center Market Comparison. China is cross-checked
against SemiAnalysis and Rystad. Political-risk evidence is public reporting
as of 7 November 2026 (see the Sources sheet).

Run: python3 build_datacenter_capacity_workbook.py
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "datacenter_gw_capacity_us_vs_international.xlsx"
COMPILED = "7 November 2026"

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(color="FFFFFF", bold=True, size=11)
ALT_FILL = PatternFill("solid", fgColor="F3F6FB")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
US_FILL = PatternFill("solid", fgColor="E2EFDA")
HIGH_FILL = PatternFill("solid", fgColor="F8CBAD")
ELEV_FILL = PatternFill("solid", fgColor="FCE4D6")
MOD_FILL = PatternFill("solid", fgColor="FFF2CC")
LOW_FILL = PatternFill("solid", fgColor="E2EFDA")
TITLE_FONT = Font(bold=True, size=15, color=NAVY)
SUB_FONT = Font(italic=True, size=10, color="595959")
SECTION_FONT = Font(bold=True, size=12, color=NAVY)
BOLD = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

GW = "#,##0.00"
GW1 = "#,##0.0"
PCT = "0.0%"
MULT = '0.0"x"'
SCORE = "0.00"


def style_header(ws, row, ncols, start_col=1, height=None):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    if height:
        ws.row_dimensions[row].height = height


def paint_alt(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = WRAP
            if (r - r0) % 2 == 1:
                cell.fill = ALT_FILL


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()

# ---------------------------------------------------------------- Inputs
inp = wb.active
inp.title = "Inputs"
inp.sheet_view.showGridLines = False
inp["B2"] = "Inputs and assumptions — yellow cells are editable; the rest of the workbook recalculates"
inp["B2"].font = TITLE_FONT
inp["B3"] = (
    f"Compiled {COMPILED}. Capacity is GW of IT (critical) load. Regional headlines are "
    "Cushman & Wakefield (C&W) H1 2026. Under construction is the near-committed book; "
    "planned is announced, site-identified capacity and is not fully committed."
)
inp["B3"].font = SUB_FONT

inp["B5"] = "Regional headline capacity (C&W H1 2026)"
inp["B5"].font = SECTION_FONT
for i, h in enumerate(["Region", "Live (GW)", "Under construction (GW)", "Planned (GW)", "Source", "Note"]):
    inp.cell(row=6, column=2 + i, value=h)
style_header(inp, 6, 6, start_col=2, height=32)
regions = [
    ("Americas", 50.3, 37.7, 285.0, "S1",
     "C&W Americas Data Center Update H1 2026 (14 Sep 2026). 94.1% of live capacity is in the U.S.; 91.7% of capacity under construction was precommitted."),
    ("EMEA", 12.1, 3.8, 14.5, "S4",
     "C&W EMEA Data Centre Update H1 2026. Committed pipeline of 18.3 GW = 3.8 under construction + 14.5 planned."),
    ("APAC", 15.135, 4.763, 21.692, "S7",
     "Sum of the 14 APAC markets on the C&W H1 2026 interactive map. C&W rounds under construction and planned to 4.8 GW and 21.7 GW."),
]
for i, row in enumerate(regions):
    for j, v in enumerate(row):
        cell = inp.cell(row=7 + i, column=2 + j, value=v)
        if 1 <= j <= 3:
            cell.number_format = GW
            cell.fill = INPUT_FILL
            cell.border = BORDER
        else:
            cell.border = BORDER
            cell.alignment = WRAP
AMER_ROW, EMEA_ROW, APAC_ROW = 7, 8, 9

inp["B11"] = "Derivation and estimate assumptions"
inp["B11"].font = SECTION_FONT
for i, h in enumerate(["Assumption", "Value", "Unit", "Basis", "Used in"]):
    inp.cell(row=12, column=2 + i, value=h)
style_header(inp, 12, 5, start_col=2, height=30)
assumptions = [
    ("U.S. share of Americas live capacity", 0.941, "share",
     "C&W Americas H1 2026: 50.3 GW operational, 94.1% located in the U.S. (S1)", "Countries — United States live"),
    ("Alberta live capacity", 0.05, "GW",
     "Estimate: eStruxture CAL-1/CAL-2 about 35 MW plus the AWS Canada West region. Not published separately by C&W. (S12)", "Countries — Canada"),
    ("Alberta under construction", 1.19, "GW",
     "Meta Sturgeon County about 1.0 GW (ground broken July 2026) + eStruxture CAL-3 90 MW + Bitdeer Fox Creek about 100 MW. (S12, S13)", "Countries — Canada"),
    ("Alberta planned pipeline, site-identified", 8.0, "GW",
     "Judgement. Wonder Valley (headline up to 7.5 GW), Beacon AI campuses, Keephills 230 MW and peers. The AESO queue is about 19.6 GW, but only 1.2 GW is allocated through 2028. (S12, S13)", "Countries — Canada"),
    ("Saudi Arabia live", 0.30, "GW",
     "Residual of the C&W Middle East 972 MW live total after the UAE (413 MW), Israel (121 MW) and Türkiye (92 MW). (S4)", "Countries — Saudi Arabia"),
    ("Saudi Arabia under construction", 0.25, "GW",
     "Estimate: HUMAIN Riyadh and Dammam initial phases, DataVolt/Oxagon phase 1, AWS Riyadh. (S16, S17)", "Countries — Saudi Arabia"),
    ("Saudi Arabia planned pipeline", 1.50, "GW",
     "Estimate, conservative and site-identified. HUMAIN targets about 1.9 GW by 2030; Together AI 250 MW (Aug 2026); AWS $5.3B Riyadh hub. (S16, S17)", "Countries — Saudi Arabia"),
    ("Türkiye under construction", 0.10, "GW",
     "C&W national pipeline is 284 MW combined. The split between under construction and planned is estimated. (S4)", "Countries — Türkiye"),
    ("Türkiye planned", 0.184, "GW", "Residual of the 284 MW national pipeline. (S4)", "Countries — Türkiye"),
    ("South Africa live", 0.30, "GW",
     "Johannesburg 222 MW live; Africa total 469 MW. Cape Town is a developing market. (S4)", "Countries — South Africa"),
    ("South Africa under construction", 0.08, "GW", "Estimate from the Johannesburg 175 MW pipeline. (S4)", "Countries — South Africa"),
    ("South Africa planned", 0.15, "GW", "Estimate covering Johannesburg and Cape Town. (S4)", "Countries — South Africa"),
    ("China live, full-fleet (adjusted view)", 24.0, "GW",
     "SemiAnalysis China Datacenter Model, 25 Sep 2026: fleet above 24 GW. Rystad: 32 GW installed at end-2025 across all data-center types, 40 GW by end-2026. (S9, S10)", "Summary — adjusted view"),
    ("China under construction, full-fleet", 20.0, "GW",
     "SemiAnalysis: about 20 GW of dated pipeline, excluded from the 24 GW operating fleet. (S9)", "Summary — adjusted view"),
    ("China planned, full-fleet", 30.0, "GW",
     "SemiAnalysis: about 30 GW announced. Rystad: 28 GW of new projects online by 2030. (S9, S10)", "Summary — adjusted view"),
]
for i, row in enumerate(assumptions):
    r = 13 + i
    for j, v in enumerate(row):
        cell = inp.cell(row=r, column=2 + j, value=v)
        cell.border = BORDER
        cell.alignment = WRAP
        if j == 1:
            cell.fill = INPUT_FILL
            cell.number_format = PCT if row[2] == "share" else GW
A = {name: f"Inputs!$C${13 + i}" for i, (name, *_) in enumerate(assumptions)}
set_widths(inp, [3, 52, 14, 14, 16, 110, 28])

# -------------------------------------------------------------- Countries
co = wb.create_sheet("Countries")
co.sheet_view.showGridLines = False
co["A1"] = "Data center capacity by country — live, under construction, planned pipeline, political risk"
co["A1"].font = TITLE_FONT
co["A2"] = (
    f"GW of IT load, C&W H1 2026 unless the notes say otherwise. Residual rows reconcile listed markets to the "
    f"regional headlines on Inputs. Risk columns look up the Political Risk sheet. Compiled {COMPILED}."
)
co["A2"].font = SUB_FONT
co.row_dimensions[2].height = 32

CO_HDR = [
    "Country / market", "Region", "Sub-region", "US / International",
    "Live (GW)", "Under construction (GW)", "Planned pipeline (GW)",
    "Pipeline, U/C + planned (GW)", "Pipeline / live",
    "Political risk score (1-5)", "Risk tier", "HIGH political risk",
    "Data basis", "Source",
]
CO_HDR_ROW = 4
for i, h in enumerate(CO_HDR, 1):
    co.cell(row=CO_HDR_ROW, column=i, value=h)
style_header(co, CO_HDR_ROW, len(CO_HDR), height=36)

# name, region, sub, us/intl, live, uc, planned, note, source
# live/uc/planned may be a number or an Excel formula. "US_*" and "RESID" are filled below.
countries = [
    ("United States", "Americas", "North America", "United States", "US_LIVE", "US_UC", "US_PL",
     "Live = Americas headline times the 94.1% U.S. share. Under construction and planned = Americas headline minus the identified non-U.S. markets below, so small unidentified non-U.S. markets stay inside the U.S. figure (upward bias well under 1 GW on under construction). Virginia alone is 11.3 GW live and 6.4 GW under construction; West Texas has 2.9 GW under construction.",
     "S1, S2, S3"),
    ("Canada", "Americas", "North America", "International",
     f"=0.438+0.469+0.031+{A['Alberta live capacity']}",
     f"=0.234+0.185+0.016+{A['Alberta under construction']}",
     f"=0.365+0.563+0.013+{A['Alberta planned pipeline, site-identified']}",
     "Toronto 438/234/365 MW, Montreal-Quebec 469/185/563 MW and Vancouver 31/16/13 MW are the latest published C&W market detail (H2 2025 edition), plus the Alberta estimates on Inputs. The AESO large-load queue is about 19.6 GW against 1.2 GW allocated through 2028.",
     "S1, S11, S12, S13"),
    ("Brazil", "Americas", "Latin America", "International", 0.661, 0.250, 2.066,
     "Sao Paulo-Campinas, H1 2026: 661 MW live, 250 MW under construction, 2,066 MW planned. Vacancy 4.1%.", "S1, S14"),
    ("Mexico", "Americas", "Latin America", "International", 0.336, 0.081, 0.080,
     "Queretaro live capacity is 336 MW (H1 2026). Under construction and planned are the latest published detail, 81 MW and 80 MW (H1 2025).", "S1, S14, S15"),
    ("Chile", "Americas", "Latin America", "International", 0.272, 0.095, 0.079,
     "Santiago live capacity is 272 MW (H1 2026). Under construction and planned are the latest published detail, 95 MW and 79 MW (H1 2025).", "S1, S14, S15"),
    ("Colombia", "Americas", "Latin America", "International", 0.058, 0.056, 0.300,
     "Bogota, H1 2026: 58 MW live, 56 MW under construction, 300 MW planned, including an ODATA 120 MW project.", "S14"),
    ("Rest of Americas", "Americas", "Other", "International", "RESID_AM_L", 0.0, 0.0,
     "Live residual versus the C&W Americas headline (other Latin American and Canadian markets). The under-construction and planned residual is embedded in the U.S. derivation.", "S1"),
    ("United Kingdom", "EMEA", "Europe", "International", 1.376, 0.278, 1.496,
     "London only, Q2 2026: 1,376 MW live, 278 MW under construction, 1,496 MW planned, colocation vacancy 5.6%. Other UK markets sit in Rest of EMEA.", "S4, S5"),
    ("Germany", "EMEA", "Europe", "International", "=0.981+0.106", "=0.411+0.046", "=1.267+0.448",
     "Frankfurt 981/411/1,267 MW plus Berlin 106/46/448 MW (Q2 2026). Berlin has a further 510 MW at an early stage, which is excluded.", "S4, S5"),
    ("Ireland", "EMEA", "Europe", "International", 1.253, 0.104, 0.588,
     "Dublin, Q2 2026: 1,253 MW live, 104 MW under construction, 588 MW planned. C&W describes growth as heavily constrained by regulation.", "S4, S5"),
    ("Netherlands", "EMEA", "Europe", "International", 0.927, 0.102, 0.251,
     "Amsterdam, Q2 2026: 927 MW live, 102 MW under construction, 251 MW planned. The pipeline is down 20% year on year. Groningen/Eemshaven sits in Rest of EMEA.", "S4, S5"),
    ("France", "EMEA", "Europe", "International", 0.665, 0.278, 1.010,
     "Paris, Q2 2026: 665 MW live, 278 MW under construction, more than 1 GW planned. RTE counts about 18 GW of national connection requests.", "S4, S5"),
    ("Spain", "EMEA", "Europe", "International", "=0.261+0.111+0.068", "=0.094+0+0.015", "=0.377+0.414+0.075",
     "Madrid 261/94/377 MW, Zaragoza 111/0/414 MW and Barcelona 68/15/75 MW (Q2 2026). Zaragoza has a further 2.4 GW at an early stage, which is excluded.", "S4, S5"),
    ("Finland", "EMEA", "Europe", "International", 0.329, 0.083, 0.737,
     "Helsinki, Q2 2026. New EMEA powerhouse market.", "S4, S5"),
    ("Italy", "EMEA", "Europe", "International", 0.212, 0.113, 0.764,
     "Milan, Q2 2026.", "S4, S5"),
    ("Norway", "EMEA", "Europe", "International", 0.208, 0.120, 0.216,
     "Oslo, Q2 2026.", "S4, S5"),
    ("Sweden", "EMEA", "Europe", "International", 0.198, 0.080, 0.148,
     "Stockholm, Q2 2026. Gavle sits in Rest of EMEA.", "S4, S5"),
    ("Poland", "EMEA", "Europe", "International", 0.159, 0.005, 0.156,
     "Warsaw, Q2 2026.", "S4, S5"),
    ("Denmark", "EMEA", "Europe", "International", 0.143, 0.017, 0.106,
     "Copenhagen, Q2 2026.", "S4, S5"),
    ("Austria", "EMEA", "Europe", "International", 0.079, 0.009, 0.104,
     "Vienna, Q2 2026.", "S4, S5"),
    ("Portugal", "EMEA", "Europe", "International", 0.068, 0.050, 0.374,
     "National, Q2 2026. A further 1.24 GW early-stage pipeline at Sines is excluded.", "S4, S5"),
    ("United Arab Emirates", "EMEA", "Middle East", "International", "=0.221+0.192", "=0.232+0.072", "=0.159+0.112",
     "Abu Dhabi 221/232/159 MW plus Dubai 192/72/112 MW (Q2 2026). Abu Dhabi has a further 800 MW early-stage. The 5 GW Stargate UAE campus is being replanned as a dispersed network after the March 2026 Iranian strikes (Reuters, Sep 2026).",
     "S4, S5, S18"),
    ("Saudi Arabia", "EMEA", "Middle East", "International",
     f"={A['Saudi Arabia live']}", f"={A['Saudi Arabia under construction']}", f"={A['Saudi Arabia planned pipeline']}",
     "Estimate; see Inputs. C&W puts the Middle East at 972 MW live. Riyadh is Established, Jeddah Emerging, Dammam Developing.", "S4, S16, S17"),
    ("Israel", "EMEA", "Middle East", "International", 0.121, 0.180, 0.437,
     "National, Q2 2026: 121 MW live, 180 MW under construction, 437 MW planned. Tel Aviv is 99 MW live and 109 MW under construction.", "S4, S5"),
    ("Türkiye", "EMEA", "Middle East", "International", 0.092,
     f"={A['Türkiye under construction']}", f"={A['Türkiye planned']}",
     "National live capacity 92 MW (Q2 2026). The 284 MW combined pipeline is split on Inputs.", "S4, S5"),
    ("South Africa", "EMEA", "Africa", "International",
     f"={A['South Africa live']}", f"={A['South Africa under construction']}", f"={A['South Africa planned']}",
     "Estimate; see Inputs. Johannesburg is 222 MW live with a 175 MW pipeline. Africa is 469 MW live.", "S4, S5"),
    ("Rest of EMEA", "EMEA", "Other", "International", "RESID_EM_L", "RESID_EM_U", "RESID_EM_P",
     "Residual versus the C&W EMEA headline: the rest of the UK and Germany, Switzerland, Belgium, other Nordics (the Nordics total 2.1 GW live), the Gulf outside the UAE and Saudi Arabia, and the rest of Africa.", "S4"),
    ("Mainland China", "APAC", "North Asia", "International", 4.905, 0.471, 1.420,
     "C&W tracks colocation and hyperscale cloud in major cities (32 operators, 255 sites). Full-fleet estimates are much larger: SemiAnalysis above 24 GW operating, about 20 GW dated pipeline and about 30 GW announced; Rystad 32 GW installed at end-2025. See the adjusted view on Summary.",
     "S7, S9, S10"),
    ("Japan", "APAC", "North Asia", "International", 1.824, 0.526, 2.971,
     "37 operators, 181 sites. Tokyo is 1,397 MW live. Second-largest tracked APAC market.", "S7, S8"),
    ("India", "APAC", "South Asia", "International", 1.789, 0.509, 3.351,
     "33 operators, 147 sites. Mumbai is 890 MW live.", "S7, S8"),
    ("Australia", "APAC", "Oceania", "International", 1.607, 0.311, 4.242,
     "44 operators, 157 sites. Sydney is 917 MW live and 2,018 MW planned.", "S7, S8"),
    ("Malaysia", "APAC", "Southeast Asia", "International", 1.318, 1.039, 4.399,
     "26 operators, 67 sites. Johor is 1,110 MW live and 602 MW under construction. Largest under-construction book in APAC.", "S7, S8"),
    ("Singapore", "APAC", "Southeast Asia", "International", 1.068, 0.025, 0.207,
     "26 operators, 51 sites. New capacity is released through government calls. DC-CFA2 offers at least 200 MW.", "S7, S8"),
    ("South Korea", "APAC", "North Asia", "International", 0.848, 0.314, 1.019,
     "30 operators, 81 sites. Seoul is 663 MW live.", "S7, S8"),
    ("Hong Kong", "APAC", "North Asia", "International", 0.592, 0.185, 0.675,
     "21 operators, 49 sites.", "S7, S8"),
    ("Indonesia", "APAC", "Southeast Asia", "International", 0.468, 0.434, 1.523,
     "39 operators, 125 sites. Jakarta is 356 MW live and 395 MW under construction.", "S7, S8"),
    ("Taiwan", "APAC", "North Asia", "International", 0.270, 0.061, 0.134,
     "12 operators, 42 sites.", "S7"),
    ("Thailand", "APAC", "Southeast Asia", "International", 0.139, 0.859, 1.227,
     "28 operators, 64 sites. Bangkok under construction rose 148% to 859 MW.", "S7, S8"),
    ("New Zealand", "APAC", "Oceania", "International", 0.127, 0.005, 0.322,
     "22 operators, 54 sites.", "S7"),
    ("Philippines", "APAC", "Southeast Asia", "International", 0.107, 0.005, 0.102,
     "16 operators, 40 sites.", "S7"),
    ("Vietnam", "APAC", "Southeast Asia", "International", 0.073, 0.019, 0.100,
     "12 operators, 30 sites. Ho Chi Minh City's pipeline is 68 MW.", "S7, S8"),
]

CO_FIRST = CO_HDR_ROW + 1
CO_LAST = CO_FIRST + len(countries) - 1
row_of = {c[0]: CO_FIRST + i for i, c in enumerate(countries)}


def rows_where(pred):
    return [CO_FIRST + i for i, c in enumerate(countries) if pred(c)]


def sum_cells(col, rows):
    return "+".join(f"{col}{r}" for r in rows)


amer_non_us = rows_where(lambda c: c[1] == "Americas" and c[3] == "International" and not c[0].startswith("Rest"))
amer_listed_live = rows_where(lambda c: c[1] == "Americas" and not c[0].startswith("Rest"))
emea_listed = rows_where(lambda c: c[1] == "EMEA" and not c[0].startswith("Rest"))

for i, (name, region, sub, usflag, live, uc, pl, note, src) in enumerate(countries):
    r = CO_FIRST + i
    if name == "United States":
        live = f"=Inputs!$C${AMER_ROW}*{A['U.S. share of Americas live capacity']}"
        uc = f"=Inputs!$D${AMER_ROW}-({sum_cells('F', amer_non_us)})"
        pl = f"=Inputs!$E${AMER_ROW}-({sum_cells('G', amer_non_us)})"
    elif name == "Rest of Americas":
        live = f"=Inputs!$C${AMER_ROW}-({sum_cells('E', amer_listed_live)})"
        uc, pl = 0.0, 0.0
    elif name == "Rest of EMEA":
        live = f"=Inputs!$C${EMEA_ROW}-({sum_cells('E', emea_listed)})"
        uc = f"=Inputs!$D${EMEA_ROW}-({sum_cells('F', emea_listed)})"
        pl = f"=Inputs!$E${EMEA_ROW}-({sum_cells('G', emea_listed)})"
    co.cell(row=r, column=1, value=name)
    co.cell(row=r, column=2, value=region)
    co.cell(row=r, column=3, value=sub)
    co.cell(row=r, column=4, value=usflag)
    for col, v in ((5, live), (6, uc), (7, pl)):
        cell = co.cell(row=r, column=col, value=v)
        cell.number_format = GW
    co.cell(row=r, column=8, value=f"=F{r}+G{r}").number_format = GW
    co.cell(row=r, column=9, value=f'=IF(E{r}>0,H{r}/E{r},"n/a")').number_format = MULT
    if name.startswith("Rest of"):
        for col, v in ((10, "n/a"), (11, "n/a"), (12, "n/a")):
            co.cell(row=r, column=col, value=v)
    else:
        co.cell(row=r, column=10, value=f"=IFERROR(VLOOKUP(A{r},'Political Risk'!$A$6:$K$80,9,FALSE),\"n/a\")").number_format = SCORE
        co.cell(row=r, column=11, value=f"=IFERROR(VLOOKUP(A{r},'Political Risk'!$A$6:$K$80,10,FALSE),\"n/a\")")
        co.cell(row=r, column=12, value=f"=IFERROR(VLOOKUP(A{r},'Political Risk'!$A$6:$K$80,11,FALSE),\"n/a\")")
    co.cell(row=r, column=13, value=note)
    co.cell(row=r, column=14, value=src)

paint_alt(co, CO_FIRST, CO_LAST, 1, len(CO_HDR))
us_row = row_of["United States"]
for col in range(1, len(CO_HDR) + 1):
    co.cell(row=us_row, column=col).fill = US_FILL
for r in range(CO_FIRST, CO_LAST + 1):
    for col in (4, 10, 11, 12):
        co.cell(row=r, column=col).alignment = CENTER

# Independent arithmetic, used to write the narrative and to assert the formulas reconcile.
US_SHARE = 0.941
AMER = (50.3, 37.7, 285.0)
EMEA = (12.1, 3.8, 14.5)
APAC = (15.135, 4.763, 21.692)
canada_py = (0.438 + 0.469 + 0.031 + 0.05, 0.234 + 0.185 + 0.016 + 1.19, 0.365 + 0.563 + 0.013 + 8.0)
non_us_amer_py = [canada_py, (0.661, 0.250, 2.066), (0.336, 0.081, 0.080), (0.272, 0.095, 0.079), (0.058, 0.056, 0.300)]
nus = tuple(sum(x[i] for x in non_us_amer_py) for i in range(3))
us_py = (AMER[0] * US_SHARE, AMER[1] - nus[1], AMER[2] - nus[2])
rest_am_live = AMER[0] - us_py[0] - nus[0]

def add(*rows):
    return tuple(sum(r[i] for r in rows) for i in range(3))

emea_named = [
    (1.376, 0.278, 1.496), (0.981 + 0.106, 0.411 + 0.046, 1.267 + 0.448), (1.253, 0.104, 0.588),
    (0.927, 0.102, 0.251), (0.665, 0.278, 1.010),
    (0.261 + 0.111 + 0.068, 0.094 + 0.015, 0.377 + 0.414 + 0.075),
    (0.329, 0.083, 0.737), (0.212, 0.113, 0.764), (0.208, 0.120, 0.216), (0.198, 0.080, 0.148),
    (0.159, 0.005, 0.156), (0.143, 0.017, 0.106), (0.079, 0.009, 0.104), (0.068, 0.050, 0.374),
    (0.221 + 0.192, 0.232 + 0.072, 0.159 + 0.112), (0.30, 0.25, 1.50), (0.121, 0.180, 0.437),
    (0.092, 0.10, 0.184), (0.30, 0.08, 0.15),
]
emea_sum = add(*emea_named)
rest_emea = tuple(EMEA[i] - emea_sum[i] for i in range(3))
apac_named = [
    (4.905, 0.471, 1.420), (1.824, 0.526, 2.971), (1.789, 0.509, 3.351), (1.607, 0.311, 4.242),
    (1.318, 1.039, 4.399), (1.068, 0.025, 0.207), (0.848, 0.314, 1.019), (0.592, 0.185, 0.675),
    (0.468, 0.434, 1.523), (0.270, 0.061, 0.134), (0.139, 0.859, 1.227), (0.127, 0.005, 0.322),
    (0.107, 0.005, 0.102), (0.073, 0.019, 0.100),
]
apac_sum = add(*apac_named)
assert all(abs(apac_sum[i] - APAC[i]) < 1e-9 for i in range(3)), apac_sum
assert rest_am_live > -0.05 and all(x > -0.05 for x in rest_emea), (rest_am_live, rest_emea)

intl_live = (AMER[0] - us_py[0]) + EMEA[0] + APAC[0]
intl_uc = (AMER[1] - us_py[1]) + EMEA[1] + APAC[1]
intl_pl = (AMER[2] - us_py[2]) + EMEA[2] + APAC[2]
glob = (us_py[0] + intl_live, us_py[1] + intl_uc, us_py[2] + intl_pl)
# HIGH-flag countries and their (live, uc, planned), C&W basis.
high_rows = {
    "Mainland China": (4.905, 0.471, 1.420),
    "Hong Kong": (0.592, 0.185, 0.675),
    "Taiwan": (0.270, 0.061, 0.134),
    "United Arab Emirates": (0.221 + 0.192, 0.232 + 0.072, 0.159 + 0.112),
    "Saudi Arabia": (0.30, 0.25, 1.50),
    "Israel": (0.121, 0.180, 0.437),
    "Ireland": (1.253, 0.104, 0.588),
    "Netherlands": (0.927, 0.102, 0.251),
}
high_live = sum(v[0] for v in high_rows.values())
high_pipe = sum(v[1] + v[2] for v in high_rows.values())
china_adj = (24.0 - 4.905, 20.0 - 0.471, 30.0 - 1.420)
adj_intl = (intl_live + china_adj[0], intl_uc + china_adj[1], intl_pl + china_adj[2])
adj_glob = tuple(us_py[i] + adj_intl[i] for i in range(3))

us_fml = co.cell(row=us_row, column=6).value
us_pl_fml = co.cell(row=us_row, column=7).value
assert "Inputs!$D$" in us_fml and "Inputs!$E$" in us_pl_fml, (us_fml, us_pl_fml)
assert "Inputs!$C$" not in us_fml

TOT = CO_LAST + 2
co.cell(row=TOT, column=1, value="Totals (these reconcile to the Inputs headlines)").font = SECTION_FONT
tot_rows = [
    ("United States", f'=SUMIF($D${CO_FIRST}:$D${CO_LAST},"United States",{{c}}{CO_FIRST}:{{c}}{CO_LAST})'),
    ("International", f'=SUMIF($D${CO_FIRST}:$D${CO_LAST},"International",{{c}}{CO_FIRST}:{{c}}{CO_LAST})'),
    ("Americas", f'=SUMIF($B${CO_FIRST}:$B${CO_LAST},"Americas",{{c}}{CO_FIRST}:{{c}}{CO_LAST})'),
    ("EMEA", f'=SUMIF($B${CO_FIRST}:$B${CO_LAST},"EMEA",{{c}}{CO_FIRST}:{{c}}{CO_LAST})'),
    ("APAC", f'=SUMIF($B${CO_FIRST}:$B${CO_LAST},"APAC",{{c}}{CO_FIRST}:{{c}}{CO_LAST})'),
    ("Global", f"=SUM({{c}}{CO_FIRST}:{{c}}{CO_LAST})"),
]
for i, (label, fml) in enumerate(tot_rows):
    r = TOT + 1 + i
    co.cell(row=r, column=1, value=label).font = BOLD
    for col_i, col in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
        cell = co.cell(row=r, column=col_i, value=fml.format(c=col))
        cell.number_format = GW
        cell.font = BOLD
        cell.fill = TOTAL_FILL
        cell.border = BORDER
    cell = co.cell(row=r, column=9, value=f"=IF(E{r}>0,H{r}/E{r},\"n/a\")")
    cell.number_format = MULT
    cell.border = BORDER
    co.cell(row=r, column=1).border = BORDER
    co.cell(row=r, column=1).fill = TOTAL_FILL

set_widths(co, [26, 12, 16, 18, 14, 18, 16, 18, 14, 16, 14, 16, 78, 16])
co.freeze_panes = "B5"
co.auto_filter.ref = f"A{CO_HDR_ROW}:N{CO_LAST}"
co.auto_filter.add_sort_condition("E5:E80")

# ---------------------------------------------------------- Political Risk
pr = wb.create_sheet("Political Risk")
pr.sheet_view.showGridLines = False
pr["A1"] = "Political-risk indicator by country (1 = low, 5 = high)"
pr["A1"].font = TITLE_FONT
pr["A2"] = (
    f"Analyst scores on evidence compiled {COMPILED}. The composite is the weighted average of five dimensions; "
    "weights and tier thresholds in row 4 are editable. HIGH = composite at or above the High threshold, or any single dimension scored 5."
)
pr["A2"].font = SUB_FONT
pr.row_dimensions[2].height = 32
pr["A4"] = "Weights"
pr["A4"].font = BOLD
for i, w in enumerate((0.30, 0.20, 0.25, 0.15, 0.10)):
    cell = pr.cell(row=4, column=3 + i, value=w)
    cell.number_format = "0%"
    cell.fill = INPUT_FILL
    cell.border = BORDER
    cell.alignment = CENTER
pr["H4"] = "Low below"
pr["I4"] = 2.0
pr["J4"] = "Moderate below"
pr["K4"] = 2.75
pr["L4"] = "Elevated below"
pr["M4"] = 3.25
pr["N4"] = "At or above this last threshold is High"
for addr in ("I4", "K4", "M4"):
    pr[addr].fill = INPUT_FILL
    pr[addr].number_format = "0.00"
    pr[addr].border = BORDER
for addr in ("H4", "J4", "L4"):
    pr[addr].alignment = Alignment(horizontal="right", vertical="center")

PR_HDR = [
    "Country / market", "US / International",
    "D1 Domestic policy and community opposition",
    "D2 Grid and energy-policy constraint",
    "D3 Geopolitical and physical security",
    "D4 U.S. export-control and chip access",
    "D5 Regulation, sovereignty and tax incentives",
    "Highest dimension", "Composite", "Risk tier", "HIGH political risk", "12-month trend",
    "Evidence, 2025-26", "Source",
]
for i, h in enumerate(PR_HDR, 1):
    pr.cell(row=5, column=i, value=h)
style_header(pr, 5, len(PR_HDR), height=48)

# name, us/intl, d1..d5, trend, evidence, source
risk = [
    ("United States", "United States", 4, 4, 1, 1, 3, "Rising",
     "November 2026 midterms: voters backed restrictions in 44 of 45 counted local data-center referendums, 21 of them in Ohio. Thirty-two of 34 winning gubernatorial candidates support tighter rules, and 12 support bans or moratoria. Governorships flipped in Texas, Ohio, Georgia, Nevada and Iowa. Sixteen states considered moratorium bills in 2026; New York's one-year moratorium on sites above 20 MW passed the legislature, and Virginia HB 1515 was continued. More than 300 state bills were filed in six weeks, and at least 18 states created large-load rate classes. AEP Ohio requires large loads to pay for at least 85% of requested capacity for 12 years. C&W puts the average American power-delivery timeline near five years. Offsetting this, there is no federal restriction, most moratorium bills have failed or been vetoed, and 91.7% of capacity under construction is already precommitted.",
     "S19, S20, S21, S22, S23, S3"),
    ("Canada", "International", 2, 4, 1, 1, 2, "Stable",
     "Alberta's AESO interim cap allows about 1.2 GW of new large load through 2028 against roughly 20 GW requested, which pushes projects to on-site gas. Quebec allocates power by ministerial decision. Federal sovereign-compute policy is supportive. A 2026 U.S.-Canada trade dispute is a secondary risk.",
     "S12, S13"),
    ("Brazil", "International", 2, 2, 2, 2, 3, "Easing",
     "The federal REDATA incentive regime and a power surplus support growth, and Alibaba Cloud is entering. Election-year fiscal noise and LGPD compliance are the main frictions.",
     "S14"),
    ("Mexico", "International", 3, 3, 2, 2, 3, "Stable",
     "Queretaro is power- and water-constrained. Power policy runs through CFE. Judicial reform and the 2026 USMCA review add regulatory uncertainty.",
     "S14, S15"),
    ("Chile", "International", 3, 2, 1, 1, 2, "Stable",
     "Community and water opposition has delayed hyperscale projects around Santiago. A permitting reform is in progress. Institutions are otherwise stable and the grid is renewable-rich.",
     "S14, S15, S24"),
    ("Colombia", "International", 3, 3, 3, 1, 3, "Stable",
     "The grid is hydro-heavy and exposed to El Nino. The 2026 election cycle and internal security are the other risks. Free-trade-zone incentives support data centers.",
     "S14"),
    ("United Kingdom", "International", 3, 4, 1, 1, 2, "Stable",
     "New connections in west London have been pushed into the mid-2030s. In September 2026 Scotland held planning decisions on data centers above 50 MW pending new guidance. Offsetting this, data centers are designated critical national infrastructure, and AI Growth Zones plus planning reform are supportive.",
     "S24, S26"),
    ("Germany", "International", 3, 3, 1, 1, 3, "Stable",
     "The Energy Efficiency Act sets PUE, waste-heat and renewable-electricity duties, stricter for sites starting from 1 July 2026. Frankfurt zones data centers into designated areas. Power prices are high. EU Energy Efficiency Directive reporting is in force.",
     "S25, S26"),
    ("Ireland", "International", 5, 5, 1, 1, 3, "Stable",
     "The December 2025 CRU connection policy requires new data centers to bring matching dispatchable generation or storage and to cover 80% of annual demand with additional Irish renewables, on a six-year glide path. Connections around Dublin were effectively frozen from 2021 to 2025. C&W calls growth heavily constrained. A political debate on a moratorium continues.",
     "S4, S24, S27"),
    ("Netherlands", "International", 5, 4, 1, 1, 3, "Stable",
     "A national decree bars hyperscale data centers outside two designated municipalities. Amsterdam bans new data centers and expansions until at least 2030. Grid congestion is national. The Amsterdam pipeline fell 20% year on year.",
     "S4, S24, S25"),
    ("France", "International", 2, 3, 1, 1, 2, "Easing",
     "From May 2026, strategic data centers can be designated projects of major national interest, with simplified permitting. The sovereignty agenda is supportive. The binding constraint is RTE's connection queue of about 18 GW.",
     "S4, S6"),
    ("Italy", "International", 2, 3, 1, 1, 2, "Easing",
     "Law 49/2026 creates a single authorisation for data centers. Grid connection timelines remain long.",
     "S25"),
    ("Spain", "International", 2, 3, 1, 1, 3, "Stable",
     "A draft royal decree on data-center sustainability, resilience and digital sovereignty closed its hearing in September 2026. Grid-access allocation is contested. The April 2025 Iberian blackout raised resilience scrutiny.",
     "S25"),
    ("Finland", "International", 2, 2, 2, 1, 2, "Stable",
     "Permitting is supportive and clean power is cheap. NATO's border with Russia, and hybrid threats to subsea cables, are the security premium.",
     "S4"),
    ("Norway", "International", 3, 3, 2, 1, 3, "Rising",
     "A June 2026 law removes the data-center exemption from the ban on building where surplus heat is not reused, and requires a cost-benefit analysis before construction. Crypto-mining restrictions and the grid queue add friction.",
     "S25"),
    ("Sweden", "International", 2, 3, 2, 1, 2, "Stable",
     "Stockholm is grid-constrained and an earlier tax relief was rolled back. Other nodes, such as Gavle, have cleaner power.",
     "S4"),
    ("Denmark", "International", 3, 4, 2, 1, 2, "Rising",
     "Energinet paused new transmission connection agreements from March to June 2026 and now ranks data centers behind households, health, industry and renewables when capacity is tight.",
     "S24, S25"),
    ("Poland", "International", 2, 3, 3, 1, 2, "Stable",
     "The government is supportive. The grid is still coal-heavy, and the country sits next to the Russia-Ukraine war. Microsoft and Google continue to invest.",
     "S4"),
    ("Austria", "International", 2, 2, 1, 1, 2, "Stable",
     "A small market. The binding overlay is the EU Energy Efficiency Directive.", "S4, S25"),
    ("Portugal", "International", 2, 2, 1, 1, 2, "Stable",
     "The Sines campus is politically supported. The grid is renewable-rich. EU reporting duties apply.", "S4, S25"),
    ("United Arab Emirates", "International", 1, 1, 5, 3, 2, "Rising",
     "Iranian drones struck two AWS data centers in the UAE on 1 March 2026 and damaged a third in Bahrain the same morning. The IRGC published a target list that included AWS, Google, Microsoft, Oracle and Nvidia, and later named Stargate UAE. In September 2026 Reuters reported that the 5 GW campus is being replanned as a dispersed network, with air defense and underground construction under consideration. Chip access improved in July 2026 when BIS moved the UAE to Country Group A:5 and whitelisted G42 and Core42, but advanced computing items still need a licence except for approved entities.",
     "S17, S18, S29"),
    ("Saudi Arabia", "International", 1, 1, 5, 4, 2, "Rising",
     "Saudi territory was struck during the 2026 Iran war. HUMAIN's purchase of 18,000 accelerators remains contingent on a U.S. licence, security and reporting conditions, and divestment from Huawei. Saudi Arabia has not received the UAE's A:5 upgrade. The government itself is strongly pro-build.",
     "S16, S17, S28"),
    ("Israel", "International", 2, 2, 5, 1, 2, "Stable",
     "An active war theatre in 2026, with missile and drone exposure for critical infrastructure. Domestic policy and technology demand are supportive, and chip access is not the constraint.",
     "S4, S17"),
    ("Türkiye", "International", 3, 3, 3, 2, 4, "Stable",
     "Macro and institutional volatility, data-localisation rules, and proximity to regional conflict. Newly added to the C&W index.",
     "S4"),
    ("South Africa", "International", 3, 4, 2, 1, 3, "Stable",
     "A history of Eskom load-shedding forces on-site generation. Policy is uncertain and POPIA adds a localisation overlay. Johannesburg and Cape Town are still attracting hyperscalers.",
     "S4"),
    ("Mainland China", "International", 2, 1, 4, 5, 4, "Stable",
     "U.S. export controls restrict advanced AI chips. Nvidia has guided to no China data-center compute revenue, and licensed H200 shipments were under 1% of data-center revenue. China has also restricted Nvidia purchases. Strategic rivalry and a Taiwan contingency are the geopolitical risk, alongside foreign-ownership limits and data laws. Domestic build policy is the opposite: East Data West Computing hubs, quota-steered siting, three-to-six-month permitting and abundant power. The buildout is chip-gated, not power-gated.",
     "S9, S10, S30"),
    ("Hong Kong", "International", 2, 2, 4, 5, 3, "Stable",
     "The same U.S. advanced-chip licensing applies to entities whose parent is headquartered in China or Macau, wherever the chips sit. The National Security Law and the U.S. sanctions regime are additional overlays.",
     "S30, S31"),
    ("Taiwan", "International", 2, 3, 5, 1, 2, "Stable",
     "A cross-strait conflict or blockade is the largest geopolitical tail risk in the region for data-center capacity. Power adequacy is tight after the nuclear phase-out and TSMC's load growth. Policy is otherwise supportive.",
     "S7"),
    ("Japan", "International", 1, 2, 2, 1, 1, "Stable",
     "National AI and semiconductor policy is strongly supportive. The constraint is the Tokyo grid, which is pushing new campuses toward Osaka and Hokkaido.",
     "S7, S8"),
    ("South Korea", "International", 2, 3, 3, 1, 2, "Stable",
     "The Distributed Energy Act and grid-impact reviews push new builds out of the Seoul metro. North Korea is the security exposure. AI-compute policy is supportive.",
     "S7"),
    ("Singapore", "International", 4, 3, 1, 2, 2, "Stable",
     "Capacity is rationed. A moratorium ran from 2019 to 2022, a later roadmap unlocked about 300 MW, and the second call for applications offers at least 200 MW, tied to efficiency conditions. C&W ranks the grid the most stable in its sample. Export-control scrutiny falls on Chinese-owned tenants.",
     "S8, S24, S30"),
    ("Malaysia", "International", 3, 3, 2, 4, 3, "Rising",
     "Johor has stopped approving water-intensive Tier 1 and Tier 2 facilities and set a water-usage ceiling. BIS imposed extra AI-chip licensing on Malaysia because of diversion risk, and May 2026 guidance blocks sales to subsidiaries of China-headquartered parents. Malaysia now requires permits for U.S.-origin AI chips and has seized servers. The government is otherwise pro-build.",
     "S8, S30, S31, S32"),
    ("Thailand", "International", 3, 2, 2, 4, 3, "Rising",
     "BIS imposed the same extra licensing on Thailand. In July 2026 a White House official accused Moonshot AI of using Nvidia GB300s through a Thai facility. Bangkok's under-construction book rose 148%. Political instability is a standing overlay.",
     "S30, S33"),
    ("Indonesia", "International", 3, 3, 2, 2, 4, "Stable",
     "Data-localisation rules, slow permitting and uneven grid reliability. Batam and Jakarta are growing quickly anyway.",
     "S7, S8"),
    ("Vietnam", "International", 3, 3, 2, 2, 4, "Easing",
     "A data law and localisation rules still apply, but 2026 reforms opened more foreign ownership of data centers. G42/FPT and a proposed $2 billion AI campus have been announced.",
     "S8"),
    ("Philippines", "International", 3, 3, 3, 2, 3, "Stable",
     "South China Sea tensions, high power cost and permitting friction. CREATE MORE incentives are supportive.",
     "S7"),
    ("India", "International", 2, 3, 3, 2, 4, "Stable",
     "Maharashtra, Telangana and Andhra Pradesh are actively recruiting hyperscale campuses. Grid reliability varies by state. The May 2025 India-Pakistan confrontation is the security overlay, and the DPDP Act adds a localisation track. Large GPU clusters still need a U.S. licence, but India is treated as an eligible partner.",
     "S7, S8"),
    ("Australia", "International", 2, 3, 1, 1, 2, "Stable",
     "Grid-connection delays, energy prices and planning scrutiny in Sydney. Federal policy and AUKUS / Pax Silica alignment are supportive.",
     "S7, S8"),
    ("New Zealand", "International", 1, 2, 1, 1, 1, "Stable",
     "Supportive policy, a renewable grid and a small market.", "S7"),
]
# Confirm every scored country exists on the Countries sheet.
country_names = {c[0] for c in countries}
missing = [r[0] for r in risk if r[0] not in country_names]
assert not missing, missing

PR_FIRST = 6
for i, (name, flag, d1, d2, d3, d4, d5, trend, evidence, src) in enumerate(risk):
    r = PR_FIRST + i
    pr.cell(row=r, column=1, value=name)
    pr.cell(row=r, column=2, value=flag)
    for j, d in enumerate((d1, d2, d3, d4, d5)):
        cell = pr.cell(row=r, column=3 + j, value=d)
        cell.alignment = CENTER
    pr.cell(row=r, column=8, value=f"=MAX(C{r}:G{r})").alignment = CENTER
    comp = pr.cell(row=r, column=9, value=f"=SUMPRODUCT(C{r}:G{r},$C$4:$G$4)")
    comp.number_format = SCORE
    comp.alignment = CENTER
    comp.font = BOLD
    pr.cell(row=r, column=10, value=f'=IF(I{r}<$I$4,"Low",IF(I{r}<$K$4,"Moderate",IF(I{r}<$M$4,"Elevated","High")))').alignment = CENTER
    pr.cell(row=r, column=11, value=f'=IF(OR(I{r}>=$M$4,H{r}=5),"YES","no")').alignment = CENTER
    pr.cell(row=r, column=12, value=trend).alignment = CENTER
    pr.cell(row=r, column=13, value=evidence)
    pr.cell(row=r, column=14, value=src)
PR_LAST = PR_FIRST + len(risk) - 1
paint_alt(pr, PR_FIRST, PR_LAST, 1, len(PR_HDR))
for col in range(1, len(PR_HDR) + 1):
    pr.cell(row=PR_FIRST, column=col).fill = US_FILL
for r in range(PR_FIRST, PR_LAST + 1):
    for col in range(3, 13):
        pr.cell(row=r, column=col).alignment = CENTER

pr.conditional_formatting.add(f"K{PR_FIRST}:K{PR_LAST}", CellIsRule(operator="equal", formula=['"YES"'], fill=HIGH_FILL, font=Font(bold=True, color="9C0006")))
pr.conditional_formatting.add(f"J{PR_FIRST}:J{PR_LAST}", CellIsRule(operator="equal", formula=['"High"'], fill=HIGH_FILL))
pr.conditional_formatting.add(f"J{PR_FIRST}:J{PR_LAST}", CellIsRule(operator="equal", formula=['"Elevated"'], fill=ELEV_FILL))
pr.conditional_formatting.add(f"J{PR_FIRST}:J{PR_LAST}", CellIsRule(operator="equal", formula=['"Moderate"'], fill=MOD_FILL))
pr.conditional_formatting.add(f"J{PR_FIRST}:J{PR_LAST}", CellIsRule(operator="equal", formula=['"Low"'], fill=LOW_FILL))
pr.conditional_formatting.add(f"C{PR_FIRST}:G{PR_LAST}", CellIsRule(operator="equal", formula=["5"], fill=HIGH_FILL, font=Font(bold=True)))
pr.conditional_formatting.add(f"C{PR_FIRST}:G{PR_LAST}", CellIsRule(operator="equal", formula=["4"], fill=ELEV_FILL))

# Python-side scores, so the HIGH set used in the narrative matches the sheet.
W = (0.30, 0.20, 0.25, 0.15, 0.10)
scored = {}
for name, flag, d1, d2, d3, d4, d5, trend, evidence, src in risk:
    dims = (d1, d2, d3, d4, d5)
    composite = sum(d * w for d, w in zip(dims, W))
    high = composite >= 3.25 or max(dims) == 5
    scored[name] = (composite, high)
assert abs(scored["United States"][0] - 2.70) < 1e-9
flagged = {n for n, (_, h) in scored.items() if h}
assert flagged == set(high_rows), flagged.symmetric_difference(set(high_rows))

rub_row = PR_LAST + 2
pr.cell(row=rub_row, column=1, value="Scoring rubric, applied to each dimension").font = SECTION_FONT
rubric = [
    (1, "Negligible. Policy is supportive, or there is no material constraint."),
    (2, "Low. Ordinary regulatory friction. Nothing that binds the pipeline."),
    (3, "Moderate. Constraints that slow or reroute projects: longer timelines, local opposition, conditional licensing."),
    (4, "Significant. Policy that caps, rations or materially reprices new capacity: connection caps, extra export licensing, live moratorium bills."),
    (5, "Binding. A formal ban or moratorium, a mandatory self-generation rule, a kinetic conflict, or a hard export-control exclusion."),
]
for i, (score, text) in enumerate(rubric):
    r = rub_row + 1 + i
    pr.cell(row=r, column=1, value=score).alignment = CENTER
    pr.cell(row=r, column=1).border = BORDER
    pr.cell(row=r, column=2, value=text).alignment = WRAP
    pr.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    pr.cell(row=r, column=2).border = BORDER
set_widths(pr, [24, 16, 16, 16, 16, 16, 16, 12, 12, 12, 14, 14, 88, 18])
pr.freeze_panes = "C6"
pr.auto_filter.ref = f"A5:N{PR_LAST}"

# -------------------------------------------------------------- US States
us = wb.create_sheet("US States")
us.sheet_view.showGridLines = False
us["A1"] = "United States — large markets and state political risk after the November 2026 midterms"
us["A1"].font = TITLE_FONT
us["A2"] = (
    "Market GW figures are C&W, mostly the H2 2025 edition (the latest public market detail) plus the May 2026 global comparison. "
    "n/a means the market total was not disclosed. State scores use the same 1-to-5 rubric. The national score on Political Risk is Moderate; several states are High."
)
us["A2"].font = SUB_FONT
us.row_dimensions[2].height = 32
US_HDR = ["State", "Markets", "Live (GW)", "Under construction (GW)", "Planned (GW)", "Vintage",
          "What changed in 2025-26", "State score (1-5)", "State tier", "Trend", "Source"]
for i, h in enumerate(US_HDR, 1):
    us.cell(row=4, column=i, value=h)
style_header(us, 4, len(US_HDR), height=32)
states = [
    ("Virginia", "Northern Virginia, the largest market in the world", 11.3, 6.4, None, "C&W global comparison, 2026",
     "HB 1515, which would block final approvals until interconnection requests are met or until July 2028, was continued into the 2027 session. The sales-tax exemption is under review. Loudoun and Prince William have active local opposition. Dominion has a large-load tariff, and PJM capacity prices have been blamed on data centers.",
     3.5, "Rising", "S3, S21, S22"),
    ("Texas", "Dallas (ranked the world's top primary market), West Texas, Austin-San Antonio", None, None, None, "C&W H1 2026",
     "SB 6 (2025) sets large-load interconnection and ERCOT curtailment rules. West Texas has 2.9 GW under construction. The Dallas planned pipeline nearly tripled. The 2026 governor's race was won by a Democrat who campaigned on pausing or regulating new projects. On-site gas generation is now common.",
     3.0, "Rising", "S2, S3, S19, S20"),
    ("Georgia", "Atlanta, the world's number-two primary market", 1.757, 1.322, 13.037, "C&W H2 2025",
     "HB 1059, which would stop local permitting until December 2028, has been introduced, alongside proposals to cut the data-center sales-tax credit. Georgia Power has large-load rules. The governorship flipped to a Democrat who supports tighter regulation.",
     3.5, "Rising", "S2, S19, S20, S21"),
    ("Ohio", "Columbus, the world's number-four primary market", 3.150, 1.150, 9.709, "C&W H2 2025",
     "Twenty-one local anti-data-center measures passed in November 2026. Governor-elect Acton backs a conditional moratorium: projects must cover 100% of utility costs and use union labor. AEP's tariff makes large loads pay for at least 85% of requested capacity for 12 years. Senator Husted lost after being cast as the face of data centers.",
     4.5, "Rising", "S2, S19, S20"),
    ("Arizona", "Phoenix", 2.341, 1.006, 7.492, "C&W H2 2025; under construction near 1.6 GW by H1 2026",
     "Water use is under scrutiny and both APS and SRP have large-load tariffs. The planned pipeline was flat in the first half of 2026: developers are advancing queued projects rather than adding commitments.",
     2.5, "Stable", "S2, S34"),
    ("Oregon", "Portland and eastern Oregon", 3.317, 0.716, 4.138, "C&W H2 2025",
     "The 2025 POWER Act creates a separate rate class for large loads. A PGE-GridCARE partnership is expected to unlock 80 MW in 2026 and more than 320 MW by 2029.",
     2.5, "Stable", "S2"),
    ("Illinois", "Chicago", None, None, None, "C&W H2 2025",
     "The planned pipeline nearly tripled. Seven local restriction referendums passed in November 2026. ComEd has a large-load tariff proceeding.",
     3.0, "Rising", "S2, S20"),
    ("Pennsylvania", "Homer City and Amazon's $20 billion program", None, None, None, "C&W H1 2026",
     "SB 1359 and HB 2533, a three-year moratorium or a municipal opt-out, have been introduced. The governor signed an August 2026 order on siting and ratepayer protection. PJM cost allocation is contested.",
     3.0, "Rising", "S21, S22"),
    ("Wisconsin", "Port Washington and Microsoft's Fairwater campus", None, None, None, "Project announcements",
     "A bill that would have blocked data centers until 14 conditions were met failed. The governorship went to a Democrat running on regulation, and local referendums passed.",
     3.0, "Rising", "S20, S21"),
    ("New York", "New York City and northern New Jersey", None, None, None, "Legislative record, 2026",
     "A 10141 / S 9144, a one-year moratorium on data centers above 20 MW, passed the legislature and awaits the governor. DEC and the Public Service Commission are writing rules on ratepayer impacts.",
     4.0, "Rising", "S21, S22"),
    ("Iowa", "Meta, Google and Microsoft campuses", None, None, None, "2026 election results",
     "The governorship flipped to a Democrat who supports tighter regulation. MidAmerican's large-load tariff is under review.",
     2.5, "Rising", "S20"),
    ("Nevada", "Reno and Las Vegas", None, None, None, "2026 election results",
     "The governorship flipped to a Democrat. NV Energy has a large-load tariff, and water is a constraint.",
     2.5, "Rising", "S20"),
    ("Michigan", "Saline Township (Stargate 'The Barn')", None, None, None, "Legislative record, 2026",
     "HB 5594 and HB 5595 would bar new approvals until April 2027. They have been introduced, not passed. Local zoning litigation is active.",
     3.0, "Rising", "S21"),
    ("Other states with failed moratorium bills", "Oklahoma, Maine, South Dakota, Minnesota, New Hampshire", None, None, None, "NCSL tracker, 2026",
     "Moratorium bills failed or were vetoed in Oklahoma, Maine, South Dakota, Minnesota and New Hampshire. Most attempts to stop data centers have not become law.",
     2.0, "Stable", "S21"),
]
for i, row in enumerate(states):
    r = 5 + i
    for j, v in enumerate(row):
        if j == 8:
            us.cell(row=r, column=9, value=f'=IF(H{r}<2,"Low",IF(H{r}<2.75,"Moderate",IF(H{r}<3.25,"Elevated","High")))')
            us.cell(row=r, column=10, value=v)
        elif j == 9:
            us.cell(row=r, column=11, value=v)
        else:
            cell = us.cell(row=r, column=j + 1, value="n/a" if v is None else v)
            if 2 <= j <= 4 and v is not None:
                cell.number_format = GW
US_LAST = 4 + len(states)
paint_alt(us, 5, US_LAST, 1, len(US_HDR))
for r in range(5, US_LAST + 1):
    for col in (3, 4, 5, 8, 9, 10):
        us.cell(row=r, column=col).alignment = CENTER
us.conditional_formatting.add(f"I5:I{US_LAST}", CellIsRule(operator="equal", formula=['"High"'], fill=HIGH_FILL))
us.conditional_formatting.add(f"I5:I{US_LAST}", CellIsRule(operator="equal", formula=['"Elevated"'], fill=ELEV_FILL))
us.conditional_formatting.add(f"I5:I{US_LAST}", CellIsRule(operator="equal", formula=['"Moderate"'], fill=MOD_FILL))
us.conditional_formatting.add(f"I5:I{US_LAST}", CellIsRule(operator="equal", formula=['"Low"'], fill=LOW_FILL))
set_widths(us, [28, 42, 12, 16, 14, 28, 78, 14, 12, 12, 16])
us.freeze_panes = "B5"

# ---------------------------------------------------------------- Summary
sm = wb.create_sheet("Summary", 0)
sm.sheet_view.showGridLines = False
sm["B2"] = "Data center power capacity: United States vs international"
sm["B2"].font = TITLE_FONT
sm["B3"] = (
    f"Live, under construction, and planned pipeline, in GW of IT load, with a HIGH political-risk flag. "
    f"Compiled {COMPILED}. Under construction is the near-committed book (91.7% of the Americas book, and nearly all of the APAC book, is pre-leased). "
    "Planned pipeline is announced and site-identified; it is not fully committed. Yellow cells on Inputs and Political Risk drive every figure below."
)
sm["B3"].font = SUB_FONT
sm.row_dimensions[3].height = 36

sm["B5"] = "A. Base case — C&W-tracked colocation and hyperscale self-build, H1 2026"
sm["B5"].font = SECTION_FONT
SA = ["Geography", "Live (GW)", "Under construction (GW)", "Planned pipeline (GW)",
      "Total pipeline (GW)", "Pipeline / live", "Share of global live", "Share of global under construction",
      "Share of global planned", "Live GW flagged HIGH risk", "Pipeline GW flagged HIGH risk",
      "Share of pipeline flagged HIGH risk"]
for i, h in enumerate(SA):
    sm.cell(row=6, column=2 + i, value=h)
style_header(sm, 6, len(SA), start_col=2, height=36)

CR = f"Countries!$D${CO_FIRST}:$D${CO_LAST}"
FLAG = f"Countries!$L${CO_FIRST}:$L${CO_LAST}"

def ccol(letter):
    return f"Countries!${letter}${CO_FIRST}:${letter}${CO_LAST}"

for i, (label, key) in enumerate((("United States", "United States"), ("International", "International"))):
    r = 7 + i
    sm.cell(row=r, column=2, value=label).font = BOLD
    sm.cell(row=r, column=3, value=f'=SUMIF({CR},"{key}",{ccol("E")})')
    sm.cell(row=r, column=4, value=f'=SUMIF({CR},"{key}",{ccol("F")})')
    sm.cell(row=r, column=5, value=f'=SUMIF({CR},"{key}",{ccol("G")})')
    sm.cell(row=r, column=6, value=f"=D{r}+E{r}")
    sm.cell(row=r, column=7, value=f"=IF(C{r}=0,\"n/a\",F{r}/C{r})")
    sm.cell(row=r, column=8, value=f"=C{r}/$C$9")
    sm.cell(row=r, column=9, value=f"=D{r}/$D$9")
    sm.cell(row=r, column=10, value=f"=E{r}/$E$9")
    sm.cell(row=r, column=11, value=f'=SUMIFS({ccol("E")},{CR},"{key}",{FLAG},"YES")')
    sm.cell(row=r, column=12, value=f'=SUMIFS({ccol("H")},{CR},"{key}",{FLAG},"YES")')
    sm.cell(row=r, column=13, value=f"=IF(F{r}=0,\"n/a\",L{r}/F{r})")
sm.cell(row=9, column=2, value="Global").font = BOLD
for col in range(3, 7):
    letter = get_column_letter(col)
    sm.cell(row=9, column=col, value=f"={letter}7+{letter}8")
sm.cell(row=9, column=7, value="=F9/C9")
for col in (8, 9, 10):
    sm.cell(row=9, column=col, value=1)
sm.cell(row=9, column=11, value="=K7+K8")
sm.cell(row=9, column=12, value="=L7+L8")
sm.cell(row=9, column=13, value="=L9/F9")
for rr, fill in ((7, US_FILL), (8, None), (9, TOTAL_FILL)):
    for col in range(2, 14):
        cell = sm.cell(row=rr, column=col)
        cell.border = BORDER
        cell.alignment = CENTER
        if col in (3, 4, 5, 6, 11, 12):
            cell.number_format = GW1
        elif col == 7:
            cell.number_format = MULT
        elif col in (8, 9, 10, 13):
            cell.number_format = PCT
        if fill:
            cell.fill = fill
            cell.font = BOLD
sm["B10"] = (
    "HIGH means the Political Risk flag is YES: Mainland China, Hong Kong, Taiwan, the UAE, Saudi Arabia, Israel, Ireland and the Netherlands. "
    f"The United States scores {scored['United States'][0]:.2f}, which is Moderate, with the risk rising after the November 2026 midterms. Ohio and New York score High at state level; see US States."
)
sm["B10"].font = SUB_FONT
sm.row_dimensions[10].height = 32

sm["B12"] = "B. Adjusted view — replaces C&W's tracked China figure with the SemiAnalysis full-fleet estimate"
sm["B12"].font = SECTION_FONT
for i, h in enumerate(SA[:10]):
    sm.cell(row=13, column=2 + i, value=h)
style_header(sm, 13, 10, start_col=2, height=36)
china_r = row_of["Mainland China"]
adj = {
    "live": f"({A['China live, full-fleet (adjusted view)']}-Countries!E{china_r})",
    "uc": f"({A['China under construction, full-fleet']}-Countries!F{china_r})",
    "pl": f"({A['China planned, full-fleet']}-Countries!G{china_r})",
}
sm["B14"] = "United States"
sm["C14"] = "=C7"
sm["D14"] = "=D7"
sm["E14"] = "=E7"
sm["B15"] = "International, China on a full-fleet basis"
sm["C15"] = f"=C8+{adj['live']}"
sm["D15"] = f"=D8+{adj['uc']}"
sm["E15"] = f"=E8+{adj['pl']}"
sm["B16"] = "Global, adjusted"
for col in range(3, 6):
    letter = get_column_letter(col)
    sm.cell(row=16, column=col, value=f"={letter}14+{letter}15")
for rr in (14, 15, 16):
    sm.cell(row=rr, column=6, value=f"=D{rr}+E{rr}")
    sm.cell(row=rr, column=7, value=f"=F{rr}/C{rr}")
    sm.cell(row=rr, column=8, value=f"=C{rr}/$C$16")
    sm.cell(row=rr, column=9, value=f"=D{rr}/$D$16")
    sm.cell(row=rr, column=10, value=f"=E{rr}/$E$16")
sm["K14"] = "=K7"
sm["K15"] = f"=K8+{adj['live']}"
sm["K16"] = "=K14+K15"
for rr, fill in ((14, US_FILL), (15, None), (16, TOTAL_FILL)):
    for col in range(2, 12):
        cell = sm.cell(row=rr, column=col)
        cell.border = BORDER
        cell.alignment = CENTER
        cell.font = BOLD
        if col in (3, 4, 5, 6, 11):
            cell.number_format = GW1
        elif col == 7:
            cell.number_format = MULT
        elif col in (8, 9, 10):
            cell.number_format = PCT
        if fill:
            cell.fill = fill
sm["B17"] = (
    "China is flagged HIGH because of U.S. export controls, so the China uplift adds to high-risk live capacity. "
    "Rystad, using a broader definition that includes enterprise and telecom sites, puts China at 32 GW installed at the end of 2025 and 40 GW by the end of 2026. Switch the China cells on Inputs to use that view."
)
sm["B17"].font = SUB_FONT
sm.row_dimensions[17].height = 32

sm["B19"] = "C. Regions, base case — these rows reconcile to the C&W headlines on Inputs"
sm["B19"].font = SECTION_FONT
SC = ["Region", "Live (GW)", "Under construction (GW)", "Planned (GW)", "Total pipeline (GW)", "Pipeline / live",
      "Live GW flagged HIGH", "Pipeline GW flagged HIGH", "C&W live", "C&W under construction", "C&W planned", "Reconciles"]
for i, h in enumerate(SC):
    sm.cell(row=20, column=2 + i, value=h)
style_header(sm, 20, len(SC), start_col=2, height=36)
REG = f"Countries!$B${CO_FIRST}:$B${CO_LAST}"
for i, (reg, inp_row) in enumerate((("Americas", AMER_ROW), ("EMEA", EMEA_ROW), ("APAC", APAC_ROW))):
    r = 21 + i
    sm.cell(row=r, column=2, value=reg).font = BOLD
    sm.cell(row=r, column=3, value=f'=SUMIF({REG},"{reg}",{ccol("E")})')
    sm.cell(row=r, column=4, value=f'=SUMIF({REG},"{reg}",{ccol("F")})')
    sm.cell(row=r, column=5, value=f'=SUMIF({REG},"{reg}",{ccol("G")})')
    sm.cell(row=r, column=6, value=f"=D{r}+E{r}")
    sm.cell(row=r, column=7, value=f"=F{r}/C{r}")
    sm.cell(row=r, column=8, value=f'=SUMIFS({ccol("E")},{REG},"{reg}",{FLAG},"YES")')
    sm.cell(row=r, column=9, value=f'=SUMIFS({ccol("H")},{REG},"{reg}",{FLAG},"YES")')
    sm.cell(row=r, column=10, value=f"=Inputs!C{inp_row}")
    sm.cell(row=r, column=11, value=f"=Inputs!D{inp_row}")
    sm.cell(row=r, column=12, value=f"=Inputs!E{inp_row}")
    sm.cell(row=r, column=13, value=f'=IF(AND(ABS(C{r}-J{r})<0.01,ABS(D{r}-K{r})<0.01,ABS(E{r}-L{r})<0.01),"Yes","CHECK")')
sm["B24"] = "Global"
sm["B24"].font = BOLD
for col, letter in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (8, "H"), (9, "I"), (10, "J"), (11, "K"), (12, "L")):
    sm.cell(row=24, column=col, value=f"={letter}21+{letter}22+{letter}23")
sm["G24"] = "=F24/C24"
sm["M24"] = '=IF(AND(ABS(C24-C9)<0.01,ABS(D24-D9)<0.01,ABS(E24-E9)<0.01),"Yes","CHECK")'
for rr, fill in ((21, None), (22, None), (23, None), (24, TOTAL_FILL)):
    for col in range(2, 14):
        cell = sm.cell(row=rr, column=col)
        cell.border = BORDER
        cell.alignment = CENTER
        if col in (3, 4, 5, 6, 8, 9, 10, 11, 12):
            cell.number_format = GW1
        elif col == 7:
            cell.number_format = MULT
        if fill:
            cell.fill = fill
            cell.font = BOLD

sm["B26"] = "What the numbers say"
sm["B26"].font = SECTION_FONT


def g(x):
    return f"{x:.1f}"


takeaways = [
    f"The United States has {g(us_py[0])} GW live, of {g(glob[0])} GW tracked worldwide ({us_py[0] / glob[0]:.0%}). International live capacity is {g(intl_live)} GW, led by Europe, then the tracked China figure, Japan, India, Australia and Malaysia.",
    f"Under construction, the near-committed book, is {g(us_py[1])} GW in the United States of {g(glob[1])} GW worldwide ({us_py[1] / glob[1]:.0%}). Virginia alone (6.4 GW) and West Texas (2.9 GW) each rival or exceed all of EMEA, which has {EMEA[1]:.1f} GW under construction. APAC has {APAC[1]:.1f} GW, about half of it in Southeast Asia.",
    f"The planned pipeline is {g(us_py[2])} GW in the United States of {g(glob[2])} GW worldwide ({us_py[2] / glob[2]:.0%}). This is announced capacity, not a commitment. Americas capacity under construction is 91.7% pre-leased; APAC capacity under construction is nearly all pre-leased; EMEA's planned book is gated by power, with a 5.2-year average connection time.",
    f"The U.S. pipeline is {(us_py[1] + us_py[2]) / us_py[0]:.1f} times live capacity. The international pipeline is {(intl_uc + intl_pl) / intl_live:.1f} times live capacity. The U.S. buildout depends far more on converting announcements, which is where the new domestic political risk sits: 44 of 45 local referendums in November 2026 restricted data centers, and 12 incoming governors back moratoria.",
    f"On the C&W basis, {g(high_live)} GW of international live capacity and {g(high_pipe)} GW of international pipeline sit in HIGH-risk jurisdictions: China, Hong Kong, Taiwan, the UAE, Saudi Arabia, Israel, Ireland and the Netherlands. Using the full-fleet China estimate, high-risk live capacity rises to about {g(high_live + china_adj[0])} GW.",
    "The risks are different in kind. U.S. risk is domestic: ratepayer politics, moratorium bills and large-load tariffs. Gulf risk is kinetic — Iranian strikes hit AWS sites in the UAE and Bahrain on 1 March 2026 — and it is also licence-dependent. China and Hong Kong are inside the U.S. export-control wall. Taiwan's risk is a conflict or blockade. Ireland and the Netherlands have binding rules on new connections.",
    f"China is the swing factor. On the tracked basis the United States is {us_py[0] / glob[0]:.0%} of global live capacity. On the SemiAnalysis full-fleet basis that falls to {us_py[0] / adj_glob[0]:.0%}. No other assumption moves the comparison as much. Both views are in this sheet; the switch is the China block on Inputs.",
]
for i, text in enumerate(takeaways):
    cell = sm.cell(row=27 + i, column=2, value=text)
    cell.alignment = WRAP
    sm.merge_cells(start_row=27 + i, start_column=2, end_row=27 + i, end_column=13)
    sm.row_dimensions[27 + i].height = 34

chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.title = "United States vs international: live, under construction, planned (GW)"
chart.y_axis.title = "GW of IT load"
data = Reference(sm, min_col=3, max_col=5, min_row=6, max_row=8)
cats = Reference(sm, min_col=2, min_row=7, max_row=8)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.height = 8
chart.width = 18
sm.add_chart(chart, "B35")

chart2 = BarChart()
chart2.type = "col"
chart2.grouping = "clustered"
chart2.title = "Capacity sitting in HIGH political-risk jurisdictions (GW)"
chart2.y_axis.title = "GW of IT load"
data2 = Reference(sm, min_col=11, max_col=12, min_row=6, max_row=8)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
chart2.shape = 4
chart2.height = 8
chart2.width = 15
sm.add_chart(chart2, "B52")
set_widths(sm, [3, 46, 16, 18, 18, 16, 14, 16, 18, 16, 16, 16, 16])

# ----------------------------------------------------------------- Sources
so = wb.create_sheet("Sources")
so.sheet_view.showGridLines = False
so["A1"] = "Sources"
so["A1"].font = TITLE_FONT
so["A2"] = f"Public sources, read on or before {COMPILED}. Reference codes in the other sheets point here."
so["A2"].font = SUB_FONT
for i, h in enumerate(["Ref", "Publisher and title", "Date", "URL", "What it supports"], 1):
    so.cell(row=4, column=i, value=h)
style_header(so, 4, 5, height=22)
sources = [
    ("S1", "Cushman & Wakefield, Americas Data Center Update H1 2026", "14 Sep 2026",
     "https://www.cushmanwakefield.com/en/insights/americas-data-center-update",
     "Americas: 50.3 GW live, 94.1% of it in the U.S.; 37.7 GW under construction, 91.7% precommitted; 285 GW planned. Vacancy 3.8%."),
    ("S2", "Conversational Geek Stats, summaries of the C&W Americas Data Center Market Overview", "1 Aug 2026",
     "https://stats.conversationalgeek.com/analysis/cushman-wakefield-2026-data-centers",
     "Market detail for Atlanta, Phoenix, Columbus, Portland/eastern Oregon, Toronto, Montreal and Vancouver. AEP's 85% tariff. PGE-GridCARE."),
    ("S3", "Cushman & Wakefield, 2026 Global Data Center Market Comparison, press release", "21 May 2026",
     "https://ir.cushmanwakefield.com/news/press-release-details/2026/Dallas-Texas-Ranked-No--1-Primary-Data-Market-in-the-World-as-AI-Demand-Power-Constraints-and-Regulation-Reshape-CRE-Strategy/default.aspx",
     "Global under construction 31.7 GW in the prior edition. Americas about 80% of global under construction. Virginia 11.3 GW live. West Texas 2.9 GW under construction. Dallas ranked first."),
    ("S4", "Cushman & Wakefield, EMEA Data Centre Update H1 2026 (report)", "Sep 2026",
     "https://www.galivel.com/media/files/emea_data_centre_marketbeat_h1_2026_avec_compression.pdf",
     "EMEA 12.1 / 3.8 / 14.5 GW. Market indicators for London, Frankfurt, Dublin, Paris, Amsterdam, Milan, Helsinki, Madrid, Oslo, Stockholm, Warsaw, Copenhagen, Vienna, Abu Dhabi, Dubai and the spotlight markets. Nordics 2.1 GW. Middle East 972 MW. Africa 469 MW."),
    ("S5", "Cushman & Wakefield, EMEA Data Centre Update H1 2026 (summary page)", "Sep 2026",
     "https://www.cushmanwakefield.com/en/insights/emea-data-centre-update",
     "Headline EMEA figures. FLAPD holds 43% of live capacity and 32% of the pipeline."),
    ("S6", "Cushman & Wakefield, Global Data Center Market Comparison 2026, Germany release", "29 May 2026",
     "https://www.cushmanwakefield.com/en/germany/news/2026/05/global-data-center-market-comparison",
     "Prior-vintage regional totals and grid connection times: 4.4 years globally, 5.2 in EMEA, 2.7 in APAC, 5.0 in the Americas."),
    ("S7", "Cushman & Wakefield, APAC Data Centres interactive map, H1 2026", "Aug 2026",
     "https://view.ceros.com/cushmanwakefield-apac/apac-data-centres-group-interactive-map-2",
     "Live, under construction and planned MW for 14 APAC markets."),
    ("S8", "Cushman & Wakefield, Asia Pacific pipeline reaches 26.5 GW in H1 2026", "5 Aug 2026",
     "https://www.cushmanwakefield.com/en/singapore/news/2026/08/apac-dc-h1-2026",
     "APAC 4.8 GW under construction, 21.7 GW planned, 1.4 GW delivered. Southeast Asia about half of under construction. Johor, Sydney, Bangkok, Mumbai, Jakarta."),
    ("S9", "SemiAnalysis, The Chinese AI Infrastructure Boom", "25 Sep 2026",
     "https://newsletter.semianalysis.com/p/the-chinese-ai-infrastructure-boom",
     "China fleet above 24 GW, about 20 GW dated pipeline, about 30 GW announced. United States 56 GW at year-end 2026 on SemiAnalysis's own basis. China is chip-gated, not power-gated."),
    ("S10", "Rystad Energy, China's data center capacity set to top 60 GW by 2030", "2026",
     "https://www.rystadenergy.com/news/chinas-data-center-capacity-doubling-of-power",
     "China 32 GW installed at end-2025, 40 GW by end-2026, 28 GW of new projects by 2030. Broader than the C&W colo-plus-hyperscale scope."),
    ("S11", "Conversational Geek Stats, C&W Americas report on Canada", "2026",
     "https://stats.conversationalgeek.com/analysis/cushman-wakefield-2026-canada",
     "Toronto 438/234/365 MW, Montreal-Quebec 469/185/563 MW, Vancouver 31/16/13 MW."),
    ("S12", "Zeever Research, Mapping Canada's Data-Centre Boom", "15 Jul 2026",
     "https://www.zeever.ca/research/canadian-data-centres/the-map",
     "Alberta: Meta Sturgeon County 1 GW under construction, eStruxture CAL-3 90 MW, Bitdeer Fox Creek about 100 MW, Wonder Valley up to 7.5 GW proposed. AESO requests versus the 1,200 MW cap."),
    ("S13", "Alberta Post, Alberta data centres: operating and proposed projects", "Jul 2026",
     "https://abpost.ca/alberta-data-centres-operating-proposed-projects/",
     "AESO queue of 19,565 MW as of 30 July 2026. 1,200 MW allocated: 970 MW to the Meta project and 230 MW to Keephills."),
    ("S14", "MastekHw, reporting the C&W Americas H1 2026 update for Latin America", "Sep 2026",
     "https://www.mastekhw.com/noticias/cushman-wakefield-presento-el-estudio-americas-data-center-h1-2026-update/",
     "Sao Paulo 661 MW, Queretaro 336 MW, Santiago 272 MW, Bogota 58 MW live. Bogota 56 MW under construction and 300 MW planned. Sao Paulo 250 MW under construction and 2,066 MW planned."),
    ("S15", "ebizlatam, reporting an earlier C&W Americas update", "2025",
     "https://ebizlatam.com/santiago-de-chile-es-el-segundo-mercado-de-data-centers-mas-relevante-de-la-region/",
     "Latest published under-construction and planned detail for Queretaro (81/80 MW) and Santiago (95/79 MW)."),
    ("S16", "TechScoop, Export approval is the gating item for Saudi AI compute", "2026",
     "https://techscoop.io/ai-data/saudi-nvidia-chips-export-approval-constraint",
     "HUMAIN's 18,000 accelerators depend on a U.S. licence. Riyadh and Dammam initial 100 MW sites. Stargate UAE described as a 1 GW cluster inside a 5 GW campus."),
    ("S17", "Middle East Institute, Gulf AI and the Iran war", "28 Sep 2026",
     "https://mei.edu/publication/gulf-ai-the-war-as-a-catalyst-for-deepening-regional-us-strategic-tech-cooperation/",
     "Iranian drones hit two AWS sites in the UAE and one in Bahrain on 1 March 2026. IRGC target list. UAE moved to Country Group A:5 on 10 July 2026. Gulf states planning 8-10 GW. Three of 233 Gulf data centers were hit."),
    ("S18", "Reuters, UAE revises AI data center plan after Iranian attacks", "11 Sep 2026",
     "https://www.reuters.com/world/middle-east/uae-revises-ai-data-center-plan-after-iranian-attacks-sources-say-2026-09-11/",
     "The 5 GW campus is likely to become a network of sites. Air defense and underground construction are being considered."),
    ("S19", "CNBC, Voters rebuff data centers in every ballot measure except one", "4 Nov 2026",
     "https://www.cnbc.com/2026/11/04/data-center-ballot-measures-2026-midterm.html",
     "Forty-one measures in 10 states. Restrictions won everywhere except Harwood, North Dakota. Ohio's governor-elect backs a conditional moratorium."),
    ("S20", "PBS News / AP, Midterm opposition foreshadows tougher data-center rules", "Nov 2026",
     "https://www.pbs.org/newshour/nation/opposition-to-ai-data-centers-in-the-midterms-foreshadows-tougher-regulations-ahead",
     "Thirty-two of 34 winning governors back tighter rules; 12 back bans or moratoria. Governorships flipped in Texas, Ohio, Georgia, Nevada and Iowa. About three-quarters of voters oppose a local data center."),
    ("S21", "National Conference of State Legislatures, Which states are banning data centers?", "2026",
     "https://www.ncsl.org/fiscal/which-states-are-banning-data-centers",
     "Sixteen-state moratorium tracker, with status. New York passed the legislature. Virginia was continued. Maine was vetoed. Several failed."),
    ("S22", "MultiState, State data center legislation in 2026", "20 Feb 2026",
     "https://www.multistate.us/insider/2026/2/20/state-data-center-legislation-in-2026-tackles-energy-and-tax-issues",
     "More than 300 bills in 30-plus states. Eighteen states with large-load rate classes. Tax-incentive rollbacks proposed in Virginia, Georgia and Oklahoma."),
    ("S23", "Oxford Economics, US data centers are on the ballot", "2026",
     "https://www.oxfordeconomics.com/resource/us-data-centers-are-on-the-ballot/",
     "Governor races in Texas, Georgia and Ohio matter most. Local opposition mostly relocates projects. National-security framing limits federal action."),
    ("S24", "Reuters, Where authorities are restricting data centres amid the AI boom", "18 Aug 2026",
     "https://www.reuters.com/legal/litigation/where-authorities-are-restricting-data-centres-amid-ai-boom-2026-08-18/",
     "Amsterdam's ban through 2030. The Dutch hyperscale siting decree. Dublin's freeze, ended December 2025 with a self-generation condition. Restrictions also noted for Singapore, Johor, India and Chile."),
    ("S25", "DC Comply, European data centre regulation tracker, status as of 15 Oct 2026", "15 Oct 2026",
     "https://dccomply.com/insights/european-data-centre-regulation-tracker",
     "EU Energy Efficiency Directive, the rating scheme, the Cloud and AI Development Act proposal, Germany's Energy Efficiency Act, Ireland's CRU decision, the Dutch decree, Italy's Law 49/2026, Spain's draft royal decree, Norway's June 2026 law, Denmark's Energinet pause."),
    ("S26", "Coloprice, compiled data-center restrictions dataset", "2026",
     "https://coloprice.com/data/data-center-restrictions.csv",
     "Scotland's hold on planning decisions above 50 MW (16 Sep 2026). Frankfurt zoning. West London grid limits. Singapore's second capacity call. Dates for the Danish pause."),
    ("S27", "Commission for Regulation of Utilities (Ireland), data-centre connection decision", "Dec 2025",
     "https://www.cru.ie/about-us/news/the-cru-publishes-its-decision-on-new-electricity-connection-policy-for-data-centres/",
     "Matching on-site or nearby generation or storage. Eighty percent additional Irish renewables, with a six-year glide path."),
    ("S28", "Alhurra, The Iran war and the Gulf AI industry", "2026",
     "https://alhurra.com/en/28623",
     "U.S. authorisation of advanced-chip exports to HUMAIN, conditional on security and reporting."),
    ("S29", "Centre for Technology and Strategy, Oxford, Can money buy AI autonomy?", "2026",
     "https://www.cts-oxford.org/insights/can-money-buy-ai-autonomy/",
     "The July 2026 BIS rule moving the UAE from country groups D:3 and D:4 to A:5."),
    ("S30", "Asia Times, Nvidia GPU crackdown hits China-linked Southeast Asia data centers", "Jun 2026",
     "https://asiatimes.com/2026/06/nvidia-gpu-crackdown-hits-china-linked-southeast-asia-data-centers/",
     "BIS guidance of 31 May 2026: a licence is required for advanced chips destined for any entity whose ultimate parent is headquartered in China or Macau. Chips already installed may keep running."),
    ("S31", "U.S. Bureau of Industry and Security, rescission of the AI Diffusion Rule and tighter chip controls", "13 May 2025",
     "https://media.bis.gov/news-updates/press-releases/commerce-expands-export-controls-advanced-chips-artificial-intelligence-0",
     "Additional licensing for AI chips to Malaysia and Thailand, citing diversion. Guidance on Huawei Ascend."),
    ("S32", "Associated Press, Malaysia seizes AI chip servers at Kuala Lumpur airport", "2026",
     "https://apnews.com/article/malaysia-us-ai-chips-export-transit-technology-79a912de57b21a4e59b27c9a86340202",
     "Malaysia's permit regime for U.S.-origin AI chips. Thirty servers seized in a suspected re-export."),
    ("S33", "CNBC, China AI firms tap Nvidia power overseas as the U.S. weighs a crackdown", "19 Aug 2026",
     "https://www.cnbc.com/2026/08/19/china-ai-nvidia-chips-us-export-controls.html",
     "Remote cloud access is outside the physical-chip rules. Allegation that Moonshot used GB300s via Thailand. JLL: global capacity could reach about 200 GW by 2030."),
    ("S34", "Roger Williams (Cushman & Wakefield), Phoenix in the H1 2026 update", "17 Sep 2026",
     "https://www.linkedin.com/posts/roger-williams-0098103_americas-data-center-update-h1-2026-cushman-activity-7506330956274999296-kRvI",
     "Phoenix added 151 MW in the first half of 2026. Under construction near 1.6 GW. The planned pipeline was flat."),
    ("S35", "Synergy Research Group, U.S. capacity will keep growing despite headwinds", "23 Jul 2026",
     "https://www.srgresearch.com/articles/synergy-reports-that-us-data-center-capacity-will-continue-to-balloon-despite-increasing-headwinds",
     "Independent cross-check. U.S. capacity expected to double in three years. About 45 GW of U.S. IT capacity in the known pipeline. The U.S. stays well above half of global operational capacity for five years."),
    ("S36", "Synergy Research Group, hyperscale capacity growth", "Q3 2025 data",
     "https://www.srgresearch.com/articles/hyperscale-spending-spree-is-driving-dramatic-growth-in-data-center-capacity",
     "The U.S. held 55% of worldwide hyperscale operational capacity, up from 52% three years earlier."),
]
for i, row in enumerate(sources):
    r = 5 + i
    for j, v in enumerate(row, 1):
        cell = so.cell(row=r, column=j, value=v)
        cell.border = BORDER
        cell.alignment = WRAP
        if j == 4:
            cell.hyperlink = v
            cell.font = Font(color="0563C1", underline="single")
        elif (r - 5) % 2 == 1:
            cell.fill = ALT_FILL
set_widths(so, [6, 62, 16, 55, 78])
so.freeze_panes = "A5"
so.auto_filter.ref = f"A4:E{4 + len(sources)}"

# ------------------------------------------------------------- Methodology
me = wb.create_sheet("Methodology")
me.sheet_view.showGridLines = False
me["B2"] = "Methodology, definitions and caveats"
me["B2"].font = TITLE_FONT
notes = [
    ("What is counted", None),
    ("Live", "Commissioned IT load. C&W's 'in operation'. Captive enterprise and telecom sites are outside this scope in most markets."),
    ("Under construction", "Construction has started. This is the near-committed book: 91.7% of Americas capacity under construction was precommitted, and nearly all APAC capacity under construction is pre-leased. There is no published global figure for capacity that is contracted but not yet under construction, so this workbook does not invent one."),
    ("Planned pipeline", "Announced projects with a site and a development plan. C&W's 'planned'. Early-stage capacity is excluded (Zaragoza 2.4 GW, Abu Dhabi 0.8 GW, Portugal 1.24 GW, and similar). A large share of planned capacity will not be delivered on the announced schedule. The Americas planned book is several times the under-construction book."),
    ("Committed, in the question's sense", "Read under construction as committed. Read planned as announced. The commitment-quality notes above are the bridge between the two words."),
    ("Gigawatts", "IT load, not grid draw. Grid draw runs about 1.2 to 1.3 times IT load."),
    ("International", "Everything outside the United States, including Canada and Latin America."),
    ("", None),
    ("How the U.S. and the residuals are derived", None),
    ("U.S. live", "Americas headline times 94.1%. Both numbers are C&W's (S1)."),
    ("U.S. under construction and planned", "Americas headline minus Canada, Brazil, Mexico, Chile and Colombia. Unidentified non-U.S. markets therefore sit inside the U.S. figure. The bias is well under 1 GW for under construction."),
    ("Residual rows", "Rest of Americas and Rest of EMEA close the list to the regional headline. They carry no risk flag, so they are left out of the HIGH-risk totals."),
    ("Estimates", "Canada's Alberta figures, and the Saudi, Turkish and South African splits, are estimates. They are the yellow cells on Inputs, and every dependent cell recalculates."),
    ("APAC", "The 14 markets on the C&W map sum to the APAC headline on Inputs. There is no residual."),
    ("China", "C&W tracks 4.9 GW in major cities. SemiAnalysis puts the fleet above 24 GW, with about 20 GW dated and 30 GW announced. Rystad, counting all data-center types, says 32 GW at the end of 2025 and 40 GW by the end of 2026. The adjusted view uses SemiAnalysis. This gap dominates any U.S.-versus-international comparison."),
    ("Vintage", "Regional headlines are H1 2026, published August to September 2026. Some city detail (Toronto, Montreal, Vancouver, Atlanta, Phoenix, Columbus, Oregon, and the Queretaro and Santiago forward books) is from the latest public market page, which is the H2 2025 edition, and is labelled as such."),
    ("Cross-check", "Synergy, on its own basis, has the U.S. above half of global operational capacity and holding about 45 GW of future U.S. IT capacity (S35, S36). That is the same shape as this workbook, not the same number. Do not add the two sources together."),
    ("", None),
    ("Political-risk indicator", None),
    ("Dimensions and weights", "Domestic policy and community opposition, 30%. Grid and energy-policy constraints, 20%. Geopolitical and physical security, 25%. U.S. export-control and chip access, 15%. Regulation, sovereignty and tax incentives, 10%. Weights are editable."),
    ("Score", "1 is negligible and 5 is binding. The rubric is at the bottom of Political Risk. Scores are judgements on the cited evidence. They are not a published index."),
    ("Tier", "Low below 2.00, Moderate below 2.75, Elevated below 3.25, High at or above 3.25. Thresholds are editable."),
    ("HIGH flag", "The flag is YES when the composite is in the High tier, or when any one dimension is a 5. The override is deliberate: the UAE scores 2.40 overall and still has a kinetic-risk score of 5; Ireland's connection rule is a binding cap that an average would soften."),
    ("Why the United States is Moderate", "There is no federal ban, and most state moratorium bills failed or were vetoed. The November 2026 elections change the direction, which is why domestic policy and grid policy both score 4 and the trend is Rising. Ohio and New York are High on the US States sheet."),
    ("Not scored", "Financing risk, equipment tariffs, and water stress, except where water has already become a permitting rule (Johor, Chile, Arizona)."),
    ("", None),
    ("Using the file", None),
    ("Edit", "Change a yellow cell on Inputs or a score or weight on Political Risk. Summary, Countries, the totals and both charts recalculate in Excel."),
    ("Rebuild", "python3 build_datacenter_capacity_workbook.py regenerates the workbook from the data in this script."),
]
r = 4
for key, val in notes:
    if val is None:
        if key:
            me.cell(row=r, column=2, value=key).font = SECTION_FONT
    else:
        me.cell(row=r, column=2, value=key).font = BOLD
        me.cell(row=r, column=2).alignment = WRAP
        me.cell(row=r, column=3, value=val).alignment = WRAP
        me.row_dimensions[r].height = 32
    r += 1
set_widths(me, [3, 42, 120])

order = ["Summary", "Countries", "Political Risk", "US States", "Inputs", "Sources", "Methodology"]
wb._sheets = [wb[name] for name in order]
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "Data center GW capacity — US vs international"
    ws.oddFooter.left.text = f"Compiled {COMPILED}. Capacity is IT load. Planned pipeline is not fully committed."
    ws.oddFooter.right.text = "Page &P of &N"
    ws.print_title_rows = "1:4"
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.zoomScale = 100
wb.active = 0
wb.save(OUT)

print(f"wrote {OUT}")
print(f"US live/uc/planned {us_py[0]:.2f} / {us_py[1]:.2f} / {us_py[2]:.2f}")
print(f"Intl live/uc/planned {intl_live:.2f} / {intl_uc:.2f} / {intl_pl:.2f}")
print(f"Global {glob[0]:.2f} / {glob[1]:.2f} / {glob[2]:.2f}")
print(f"US share live/uc/planned {us_py[0]/glob[0]:.1%} / {us_py[1]/glob[1]:.1%} / {us_py[2]/glob[2]:.1%}")
print(f"HIGH intl live {high_live:.2f} pipeline {high_pipe:.2f}; adjusted high live {high_live + china_adj[0]:.2f}")
print(f"Adjusted US live share {us_py[0]/adj_glob[0]:.1%}")
print(f"EMEA residual {rest_emea[0]:.2f} / {rest_emea[1]:.2f} / {rest_emea[2]:.2f}; Americas live residual {rest_am_live:.2f}")
print(f"Countries {len(countries)} scored {len(risk)} sources {len(sources)}")
