"""Build the AI Capex Cut Scenario Model workbook.

Generates an Excel file that models the impact of AI capex spending cuts on
2026E sales and EPS growth across AI supply-chain sub-segments, including each
sub-segment's exposure to training vs. inference capex.

The workbook is formula-driven: edit the yellow input cells on the
'Scenario Assumptions' and 'Sub-Segment Exposures' sheets and every downstream
sheet recalculates.

Usage:  python3 build_ai_capex_model.py
Output: AI_Capex_Cut_Scenario_Model.xlsx (same directory)
"""

import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
YELLOW = "FFF2CC"   # input cells
GREY = "F2F2F2"
WHITE = "FFFFFF"

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=11, italic=True, color="595959")
HDR_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
BOLD = Font(name="Calibri", size=11, bold=True)
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color=NAVY)

HDR_FILL = PatternFill("solid", fgColor=NAVY)
BAND_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
INPUT_FILL = PatternFill("solid", fgColor=YELLOW)
GREY_FILL = PatternFill("solid", fgColor=GREY)

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

PCT = "0.0%"
PCT_SIGNED = "+0.0%;-0.0%;0.0%"
LEV = '0.0"x"'
USD_BN = '$#,##0"bn"'


def style_header_row(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BOX


def box_range(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = BOX


def band_rows(ws, r1, r2, c1, c2):
    """Alternate-row banding for readability."""
    for r in range(r1, r2 + 1):
        if (r - r1) % 2 == 1:
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = GREY_FILL


# ---------------------------------------------------------------------------
# Model inputs
# ---------------------------------------------------------------------------

# Global AI capex assumptions (2025E base year)
CAPEX_BASE_2025_BN = 475          # hyperscalers + neoclouds + enterprise AI infra
TRAINING_SHARE_2025 = 0.55        # share of 2025 AI capex tied to training buildouts

# Scenarios: (name, training capex growth 26E y/y, inference capex growth 26E y/y, narrative)
SCENARIOS = [
    ("Base Case",
     0.20, 0.45,
     "Hyperscaler capex guidance holds. Frontier training race continues; "
     "inference scales with copilot/agent adoption and token growth."),
    ("Mild Cut — Digestion",
     0.00, 0.30,
     "Budget digestion after three years of hypergrowth. Training efficiency "
     "gains (distillation, MoE, longer depreciation) trim frontier spend; "
     "inference growth stays intact."),
    ("Moderate Cut — ROI Reset",
     -0.20, 0.15,
     "ROI scrutiny bites. Training capex declines as fewer frontier labs "
     "compete and clusters are re-used; inference growth slows on a "
     "monetization lag."),
    ("Severe Cut — AI Winter",
     -0.45, -0.10,
     "Funding pullback across labs and neoclouds. Training capex cut sharply; "
     "inference also contracts while the capacity overhang is absorbed."),
]

# Sub-segments:
# (name, representative names, AI-capex revenue exposure (% of total sales),
#  training capex exposure (% of AI-linked revenue), EPS leverage (x) to a
#  1pp change in sales growth, baseline 26E sales growth, baseline 26E EPS
#  growth, notes)
SEGMENTS = [
    ("AI Accelerators (GPUs)", "NVIDIA, AMD",
     0.85, 0.55, 1.4, 0.35, 0.38,
     "Data center is ~85-90% of NVDA revenue. Inference share of GPU demand "
     "rising but frontier training still sets the marginal order."),
    ("Custom AI ASICs (XPUs)", "Broadcom (AI), Marvell, Alchip",
     0.45, 0.45, 1.5, 0.32, 0.35,
     "TPU / Trainium / MTIA programs skew to inference serving at scale; "
     "multi-year design wins add some backlog protection."),
    ("HBM & AI DRAM", "Micron, SK Hynix, Samsung",
     0.40, 0.60, 2.0, 0.28, 0.45,
     "HBM demand attaches to accelerator units (training-heavy). High "
     "fixed-cost model gives memory the largest EPS beta."),
    ("AI Networking (Switching/NICs)", "Arista, NVIDIA Networking, Broadcom, Cisco",
     0.50, 0.65, 1.4, 0.25, 0.28,
     "Training clusters drive east-west fabric density (higher switch/optics "
     "content per GPU than inference pods)."),
    ("Optical Transceivers & Interconnect", "Coherent, Lumentum, Fabrinet, Credo",
     0.55, 0.60, 1.8, 0.30, 0.35,
     "800G/1.6T ramps tied to large training fabrics; short lead times mean "
     "fast order response when capex changes."),
    ("Foundry & Advanced Packaging", "TSMC (CoWoS/SoIC)",
     0.30, 0.55, 1.3, 0.18, 0.20,
     "AI accelerators ~30% of revenue and rising; CoWoS capacity is the "
     "bottleneck. Long-term agreements dampen the first-year hit."),
    ("Semicap Equipment (WFE)", "ASML, Applied Materials, Lam, KLA",
     0.25, 0.55, 1.5, 0.12, 0.15,
     "Second-order exposure: tool orders lag chip capex by 2-4 quarters, so "
     "cuts land with a delay but persist longer."),
    ("AI Servers (ODM/OEM)", "Supermicro, Dell ISG, HPE, Quanta, Wiwynn",
     0.50, 0.55, 1.1, 0.22, 0.18,
     "Low-margin integration/pass-through: big revenue swings, smaller EPS "
     "leverage than component makers."),
    ("Power, Cooling & Electricals", "Vertiv, Eaton, Schneider, Modine",
     0.35, 0.50, 1.3, 0.18, 0.22,
     "Facility infrastructure is largely training/inference agnostic; "
     "multi-quarter backlog cushions the first year of a cut."),
    ("Data Center Colo & REITs", "Digital Realty, Equinix, GDS",
     0.25, 0.45, 0.9, 0.12, 0.10,
     "Long leases limit downside to in-place revenue; the hit shows up in "
     "bookings/pre-leasing and development yields."),
    ("Storage (Nearline HDD/eSSD)", "Seagate, Western Digital",
     0.20, 0.30, 1.6, 0.15, 0.25,
     "Skews to inference/data-lake buildouts (retention of generated data); "
     "least training-dependent hardware segment."),
    ("EDA & Semi IP", "Synopsys, Cadence, Arm",
     0.20, 0.60, 1.1, 0.12, 0.14,
     "Design activity for training silicon drives bookings, but ratable "
     "revenue and backlog make near-term EPS resilient."),
    ("Neoclouds (GPU Cloud)", "CoreWeave, Nebius, Crusoe",
     0.95, 0.60, 2.2, 0.60, 0.80,
     "Most sensitive segment: leveraged balance sheets plus utilization and "
     "GPU-price risk amplify any capex/demand cut."),
    ("Hyperscalers (AI Cloud Revenue)", "Microsoft, Alphabet, Amazon, Meta, Oracle",
     0.12, 0.30, 0.4, 0.13, 0.14,
     "They are the spenders: a cut lowers depreciation and lifts near-term "
     "FCF/EPS, largely offsetting slower AI cloud revenue — low net EPS beta."),
]

N_SEG = len(SEGMENTS)

# Sheet layout anchors (keep segment rows aligned across sheets)
SEG_FIRST_ROW = 4
SEG_LAST_ROW = SEG_FIRST_ROW + N_SEG - 1
SCEN_FIRST_ROW = 9           # first scenario row on 'Scenario Assumptions'

SH_ASSUM = "Scenario Assumptions"
SH_EXPO = "Sub-Segment Exposures"
SH_SALES = "Sales Growth Impact"
SH_EPS = "EPS Growth Impact"


def build():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True

    # ------------------------------------------------------------------
    # Sheet 1: Read Me
    # ------------------------------------------------------------------
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 110

    ws["B2"] = "AI Capex Spending Cut — Scenario Analysis"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Impact on 2026E sales and EPS growth across AI supply-chain "
                "sub-segments, with training vs. inference capex exposure")
    ws["B3"].font = SUBTITLE_FONT

    lines = [
        ("", None),
        ("HOW TO USE", SECTION_FONT),
        ("• Yellow cells are inputs. Edit scenario capex growth rates on 'Scenario Assumptions' and segment "
         "exposures/baselines on 'Sub-Segment Exposures'; all downstream sheets recalculate automatically.", None),
        ("", None),
        ("SHEET GUIDE", SECTION_FONT),
        ("• Scenario Assumptions — global 2025E AI capex base, training/inference split, and four 2026E scenarios "
         "(Base, Mild Cut/Digestion, Moderate Cut/ROI Reset, Severe Cut/AI Winter) with separate training and "
         "inference capex growth paths.", None),
        ("• Sub-Segment Exposures — 14 sub-segments with: AI-capex revenue exposure (% of sales), training vs. "
         "inference capex exposure mix, EPS operating leverage, and baseline 2026E sales/EPS growth.", None),
        ("• Sales Growth Impact — per scenario: each segment's blended AI capex growth (weighted by its "
         "training/inference mix), the change in sales growth vs. Base (pp), and the resulting scenario sales growth.", None),
        ("• EPS Growth Impact — applies each segment's EPS leverage factor to the sales-growth delta to get "
         "scenario EPS growth.", None),
        ("• Summary Matrix — one-page view of sales and EPS growth by segment across all four scenarios "
         "(color-scaled).", None),
        ("", None),
        ("METHODOLOGY", SECTION_FONT),
        ("1. Each scenario specifies separate 2026E growth rates for TRAINING capex and INFERENCE capex. Cuts hit "
         "training first (frontier-model ROI scrutiny, efficiency gains such as distillation); inference is more "
         "resilient because it is tied to usage and revenue.", None),
        ("2. A segment's AI-linked revenue growth = (training exposure × training capex growth) + (inference "
         "exposure × inference capex growth).", None),
        ("3. Scenario sales growth = baseline sales growth + AI revenue exposure × (segment AI-linked growth in "
         "scenario − segment AI-linked growth in Base). Baseline growth is assumed to embed the Base capex case.", None),
        ("4. Scenario EPS growth = baseline EPS growth + (sales-growth delta × EPS leverage). Leverage reflects "
         "gross-margin/fixed-cost structure (memory ~2x, ODM servers ~1x, hyperscalers ~0.4x since lower capex "
         "cuts depreciation).", None),
        ("", None),
        ("CAVEATS", SECTION_FONT),
        ("• All figures are illustrative, directional estimates for scenario analysis — not consensus data and not "
         "investment advice. Exposures and baselines are approximations as of September 2026; replace with live "
         "estimates before relying on outputs.", None),
        ("• First-year sensitivities only: backlogs (semicap, power) and long leases (REITs) delay — not remove — "
         "the impact; multi-year effects would differ.", None),
    ]
    r = 5
    for text, font in lines:
        cell = ws.cell(row=r, column=2, value=text)
        cell.alignment = WRAP
        if font:
            cell.font = font
        r += 1

    # ------------------------------------------------------------------
    # Sheet 2: Scenario Assumptions
    # ------------------------------------------------------------------
    ws = wb.create_sheet(SH_ASSUM)
    ws.sheet_view.showGridLines = False
    widths = {"A": 30, "B": 20, "C": 20, "D": 20, "E": 22, "F": 20, "G": 78}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws["A1"] = "Scenario Assumptions — 2026E AI Capex"
    ws["A1"].font = TITLE_FONT

    ws["A3"] = "Global AI capex (hyperscalers + neoclouds + enterprise)"
    ws["A3"].font = SECTION_FONT
    ws["A4"] = "2025E AI capex base ($bn)"
    ws["B4"] = CAPEX_BASE_2025_BN
    ws["B4"].number_format = USD_BN
    ws["A5"] = "Training share of 2025E AI capex"
    ws["B5"] = TRAINING_SHARE_2025
    ws["B5"].number_format = PCT
    ws["A6"] = "Inference share of 2025E AI capex"
    ws["B6"] = "=1-B5"
    ws["B6"].number_format = PCT
    for ref in ("B4", "B5"):
        ws[ref].fill = INPUT_FILL
        ws[ref].border = BOX
    ws["B6"].border = BOX

    hdr_row = SCEN_FIRST_ROW - 1  # 8
    headers = ["Scenario", "Training capex growth (26E y/y)",
               "Inference capex growth (26E y/y)", "Blended AI capex growth",
               "Implied 2026E AI capex ($bn)", "Δ vs Base Case ($bn)",
               "Narrative"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=hdr_row, column=i, value=h)
    style_header_row(ws, hdr_row, 1, len(headers))

    for i, (name, g_train, g_inf, narrative) in enumerate(SCENARIOS):
        r = SCEN_FIRST_ROW + i
        ws.cell(row=r, column=1, value=name).font = BOLD
        c_t = ws.cell(row=r, column=2, value=g_train)
        c_i = ws.cell(row=r, column=3, value=g_inf)
        for c in (c_t, c_i):
            c.number_format = PCT_SIGNED
            c.fill = INPUT_FILL
        ws.cell(row=r, column=4, value=f"=$B$5*B{r}+$B$6*C{r}").number_format = PCT_SIGNED
        ws.cell(row=r, column=5, value=f"=$B$4*(1+D{r})").number_format = USD_BN
        ws.cell(row=r, column=6,
                value=f"=E{r}-$E${SCEN_FIRST_ROW}").number_format = '+$#,##0"bn";-$#,##0"bn";$0"bn"'
        ws.cell(row=r, column=7, value=narrative).alignment = WRAP
        ws.row_dimensions[r].height = 42
    box_range(ws, hdr_row, SCEN_FIRST_ROW + len(SCENARIOS) - 1, 1, len(headers))

    note_r = SCEN_FIRST_ROW + len(SCENARIOS) + 1
    ws.cell(row=note_r, column=1,
            value="Design principle: capex cuts hit TRAINING budgets first (frontier ROI scrutiny, model-efficiency "
                  "gains), while INFERENCE capex is defended because it maps to live usage and revenue. Training-"
                  "exposed segments therefore underperform in every cut scenario.").alignment = WRAP
    ws.cell(row=note_r, column=1).font = SUBTITLE_FONT
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=7)
    ws.row_dimensions[note_r].height = 30

    # ------------------------------------------------------------------
    # Sheet 3: Sub-Segment Exposures
    # ------------------------------------------------------------------
    ws = wb.create_sheet(SH_EXPO)
    ws.sheet_view.showGridLines = False
    widths = {"A": 32, "B": 36, "C": 15, "D": 14, "E": 14, "F": 12, "G": 14,
              "H": 14, "I": 70}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws["A1"] = "Sub-Segment Exposures & Baselines"
    ws["A1"].font = TITLE_FONT

    hdr = ["Sub-segment", "Representative names",
           "AI capex revenue exposure (% of sales)",
           "Training capex exposure (% of AI rev)",
           "Inference capex exposure (% of AI rev)",
           "EPS leverage (x)",
           "Baseline 26E sales growth",
           "Baseline 26E EPS growth",
           "Notes"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header_row(ws, 3, 1, len(hdr))
    ws.row_dimensions[3].height = 58

    for i, seg in enumerate(SEGMENTS):
        name, names, expo, train, lev, g_sales, g_eps, notes = seg
        r = SEG_FIRST_ROW + i
        ws.cell(row=r, column=1, value=name).font = BOLD
        ws.cell(row=r, column=2, value=names).alignment = WRAP
        c = ws.cell(row=r, column=3, value=expo); c.number_format = PCT; c.fill = INPUT_FILL
        c = ws.cell(row=r, column=4, value=train); c.number_format = PCT; c.fill = INPUT_FILL
        ws.cell(row=r, column=5, value=f"=1-D{r}").number_format = PCT
        c = ws.cell(row=r, column=6, value=lev); c.number_format = LEV; c.fill = INPUT_FILL
        c = ws.cell(row=r, column=7, value=g_sales); c.number_format = PCT_SIGNED; c.fill = INPUT_FILL
        c = ws.cell(row=r, column=8, value=g_eps); c.number_format = PCT_SIGNED; c.fill = INPUT_FILL
        ws.cell(row=r, column=9, value=notes).alignment = WRAP
        ws.row_dimensions[r].height = 40
    box_range(ws, 3, SEG_LAST_ROW, 1, len(hdr))
    ws.freeze_panes = "B4"

    # Stacked bar chart: training vs inference exposure by segment
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Training vs Inference Capex Exposure by Sub-Segment"
    chart.height = 12
    chart.width = 24
    data = Reference(ws, min_col=4, max_col=5, min_row=3, max_row=SEG_LAST_ROW)
    cats = Reference(ws, min_col=1, min_row=SEG_FIRST_ROW, max_row=SEG_LAST_ROW)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    ws.add_chart(chart, f"A{SEG_LAST_ROW + 3}")

    # ------------------------------------------------------------------
    # Sheet 4: Sales Growth Impact
    # ------------------------------------------------------------------
    ws = wb.create_sheet(SH_SALES)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    for c in range(2, 14):
        ws.column_dimensions[get_column_letter(c)].width = 15

    ws["A1"] = "Sales Growth Impact by Scenario (2026E)"
    ws["A1"].font = TITLE_FONT

    sub_hdrs = ["Segment AI-linked capex growth", "Δ sales growth vs Base (pp)",
                "Scenario sales growth (26E)"]
    for s_idx, (scen_name, *_rest) in enumerate(SCENARIOS):
        c1 = 2 + s_idx * 3
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c1 + 2)
        top = ws.cell(row=2, column=c1, value=scen_name)
        top.font = HDR_FONT; top.fill = HDR_FILL; top.alignment = CENTER
        for j, sh in enumerate(sub_hdrs):
            ws.cell(row=3, column=c1 + j, value=sh)
    ws.cell(row=3, column=1, value="Sub-segment")
    style_header_row(ws, 3, 1, 13)
    ws.row_dimensions[3].height = 45

    base_blend_col = "B"  # Base scenario blended-growth column
    for i in range(N_SEG):
        r = SEG_FIRST_ROW + i
        ws.cell(row=r, column=1, value=f"='{SH_EXPO}'!A{r}").font = BOLD
        for s_idx in range(len(SCENARIOS)):
            c1 = 2 + s_idx * 3
            scen_r = SCEN_FIRST_ROW + s_idx
            blend_col = get_column_letter(c1)
            # blended AI-linked growth = train_expo * scen_train + inf_expo * scen_inf
            ws.cell(row=r, column=c1,
                    value=(f"='{SH_EXPO}'!$D{r}*'{SH_ASSUM}'!B${scen_r}"
                           f"+'{SH_EXPO}'!$E{r}*'{SH_ASSUM}'!C${scen_r}")
                    ).number_format = PCT_SIGNED
            # delta vs Base = exposure * (scenario blended - base blended)
            ws.cell(row=r, column=c1 + 1,
                    value=f"='{SH_EXPO}'!$C{r}*({blend_col}{r}-${base_blend_col}{r})"
                    ).number_format = PCT_SIGNED
            # scenario sales growth = baseline + delta
            ws.cell(row=r, column=c1 + 2,
                    value=f"='{SH_EXPO}'!$G{r}+{get_column_letter(c1 + 1)}{r}"
                    ).number_format = PCT_SIGNED
    box_range(ws, 3, SEG_LAST_ROW, 1, 13)
    band_rows(ws, SEG_FIRST_ROW, SEG_LAST_ROW, 1, 13)
    ws.freeze_panes = "B4"

    # ------------------------------------------------------------------
    # Sheet 5: EPS Growth Impact
    # ------------------------------------------------------------------
    ws = wb.create_sheet(SH_EPS)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 16

    ws["A1"] = "EPS Growth Impact by Scenario (2026E)"
    ws["A1"].font = TITLE_FONT

    sub_hdrs = ["Δ EPS growth vs Base (pp)", "Scenario EPS growth (26E)"]
    for s_idx, (scen_name, *_rest) in enumerate(SCENARIOS):
        c1 = 2 + s_idx * 2
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c1 + 1)
        top = ws.cell(row=2, column=c1, value=scen_name)
        top.font = HDR_FONT; top.fill = HDR_FILL; top.alignment = CENTER
        for j, sh in enumerate(sub_hdrs):
            ws.cell(row=3, column=c1 + j, value=sh)
    ws.cell(row=3, column=1, value="Sub-segment")
    style_header_row(ws, 3, 1, 9)
    ws.row_dimensions[3].height = 45

    # Sales sheet delta columns per scenario: C, F, I, L
    sales_delta_cols = ["C", "F", "I", "L"]
    for i in range(N_SEG):
        r = SEG_FIRST_ROW + i
        ws.cell(row=r, column=1, value=f"='{SH_EXPO}'!A{r}").font = BOLD
        for s_idx in range(len(SCENARIOS)):
            c1 = 2 + s_idx * 2
            dcol = sales_delta_cols[s_idx]
            ws.cell(row=r, column=c1,
                    value=f"='{SH_EXPO}'!$F{r}*'{SH_SALES}'!{dcol}{r}"
                    ).number_format = PCT_SIGNED
            ws.cell(row=r, column=c1 + 1,
                    value=f"='{SH_EXPO}'!$H{r}+{get_column_letter(c1)}{r}"
                    ).number_format = PCT_SIGNED
    box_range(ws, 3, SEG_LAST_ROW, 1, 9)
    band_rows(ws, SEG_FIRST_ROW, SEG_LAST_ROW, 1, 9)
    ws.freeze_panes = "B4"

    # ------------------------------------------------------------------
    # Sheet 6: Summary Matrix
    # ------------------------------------------------------------------
    ws = wb.create_sheet("Summary Matrix")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    for c in range(2, 6):
        ws.column_dimensions[get_column_letter(c)].width = 17

    ws["A1"] = "Summary — 2026E Growth by Scenario"
    ws["A1"].font = TITLE_FONT

    scen_names = [s[0] for s in SCENARIOS]

    def summary_block(title, top_row, src_sheet, src_cols):
        ws.cell(row=top_row, column=1, value=title).font = SECTION_FONT
        hdr_r = top_row + 1
        ws.cell(row=hdr_r, column=1, value="Sub-segment")
        for j, nm in enumerate(scen_names):
            ws.cell(row=hdr_r, column=2 + j, value=nm)
        style_header_row(ws, hdr_r, 1, 5)
        ws.row_dimensions[hdr_r].height = 30
        for i in range(N_SEG):
            src_r = SEG_FIRST_ROW + i
            r = hdr_r + 1 + i
            ws.cell(row=r, column=1, value=f"='{SH_EXPO}'!A{src_r}").font = BOLD
            for j, col in enumerate(src_cols):
                ws.cell(row=r, column=2 + j,
                        value=f"='{src_sheet}'!{col}{src_r}"
                        ).number_format = PCT_SIGNED
        box_range(ws, hdr_r, hdr_r + N_SEG, 1, 5)
        rng = f"B{hdr_r + 1}:E{hdr_r + N_SEG}"
        ws.conditional_formatting.add(
            rng,
            ColorScaleRule(start_type="num", start_value=-0.4, start_color="F8696B",
                           mid_type="num", mid_value=0, mid_color="FFEB84",
                           end_type="num", end_value=0.4, end_color="63BE7B"))
        return hdr_r + N_SEG

    end1 = summary_block("2026E revenue growth by scenario", 3, SH_SALES,
                         ["D", "G", "J", "M"])
    summary_block("2026E EPS growth by scenario", end1 + 3, SH_EPS,
                  ["C", "E", "G", "I"])
    ws.freeze_panes = "B2"

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "AI_Capex_Cut_Scenario_Model.xlsx")
    wb.save(out)
    print(f"Saved {out}")


def print_expected_values():
    """Recompute the model in pure Python and print the summary matrices.

    Used to sanity-check the Excel formulas against an independent
    implementation of the same methodology.
    """
    t_share = TRAINING_SHARE_2025
    base_t, base_i = SCENARIOS[0][1], SCENARIOS[0][2]

    print("\nBlended AI capex growth by scenario:")
    for name, g_t, g_i, _ in SCENARIOS:
        blended = t_share * g_t + (1 - t_share) * g_i
        print(f"  {name:<28} train {g_t:+.0%}  inf {g_i:+.0%}  blended {blended:+.1%}"
              f"  implied capex ${CAPEX_BASE_2025_BN * (1 + blended):,.0f}bn")

    print(f"\n{'Sub-segment':<38}" + "".join(f"{s[0][:22]:>24}" for s in SCENARIOS))
    print("  -- 2026E sales growth / EPS growth --")
    for seg in SEGMENTS:
        name, _, expo, train, lev, g_sales, g_eps, _ = seg
        cells = []
        base_blend = train * base_t + (1 - train) * base_i
        for _, g_t, g_i, _n in SCENARIOS:
            blend = train * g_t + (1 - train) * g_i
            d_sales = expo * (blend - base_blend)
            sales = g_sales + d_sales
            eps = g_eps + lev * d_sales
            cells.append(f"{sales:+.1%} / {eps:+.1%}")
        print(f"{name:<38}" + "".join(f"{c:>24}" for c in cells))


if __name__ == "__main__":
    build()
    print_expected_values()
