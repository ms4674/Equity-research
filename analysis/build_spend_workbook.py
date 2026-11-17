#!/usr/bin/env python3
"""Build the enterprise model-spend workbook.

The workbook is formula-driven. Yellow cells with blue font on Assumptions are
the only inputs. Everything else is calculated.
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.legend import Legend
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.workbook.properties import CalcProperties

OUT = "/agent/repos/Equity-research/analysis/proprietary-vs-open-weight-model-spend.xlsx"

# --- palette ---
NAVY = "1B3A4B"
TEAL = "0E6B67"
INK = "1A1A1A"
MUTED = "5C6B73"
WHITE = "FFFFFF"
PAPER = "F7F5F2"
GOLD = "FFF4D6"
GOLD_EDGE = "E2C98A"
GREEN = "E7F2EA"
GREEN_INK = "1E6B3A"
AMBER = "FBF0E4"
RED_SOFT = "F8E6E4"
LINE = "D5D1CB"
ZEBRA = "F3F1ED"
BLUE_FONT = "1D4E89"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
thick_bottom = Border(bottom=Side(style="medium", color=NAVY))

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_amber = PatternFill("solid", fgColor=AMBER)
fill_paper = PatternFill("solid", fgColor=PAPER)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_head2 = PatternFill("solid", fgColor="243B4A")

font_title = Font(name="Calibri", size=20, bold=True, color=NAVY)
font_h = Font(name="Calibri", size=14, bold=True, color=NAVY)
font_h_white = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_label = Font(name="Calibri", size=10, color=INK)
font_small = Font(name="Calibri", size=9, italic=True, color=MUTED)
font_input = Font(name="Calibri", size=10, bold=True, color=BLUE_FONT)
font_link = Font(name="Calibri", size=10, color="0B57D0", underline="single")
font_section = Font(name="Calibri", size=12, bold=True, color=NAVY)
font_kpi = Font(name="Calibri", size=14, bold=True, color=NAVY)
font_white = Font(name="Calibri", size=10, color=WHITE)
font_body = Font(name="Calibri", size=11, color=INK)

wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_top = Alignment(wrap_text=True, vertical="top", horizontal="left")
center = Alignment(wrap_text=True, vertical="center", horizontal="center")
right = Alignment(vertical="center", horizontal="right")

FMT_B = '"$"#,##0.00'
FMT_M = '"$"#,##0.0'
FMT_PCT = "0.0%"
FMT_PP = '+0.0%;-0.0%;0.0%'
FMT_PX = '"$"#,##0.000'
FMT_N = "#,##0.00"
FMT_T = '#,##0.0'

SIZES = ["100-999", "1,000-9,999", "10,000+"]
SIZE_OPEN = [1.35, 1.00, 0.78]
SIZE_EVAL = [0.70, 1.00, 1.20]

# name, weight, open$ share, eval$ share, share of open$ that is eval, chinese share of open$, mixes, note
VERTS = [
    (
        "Technology and software",
        0.34,
        0.060,
        0.075,
        0.10,
        0.45,
        (0.30, 0.35, 0.35),
        "Highest spend per firm. Coding agents dominate volume. Open-weight is used on routine generation; hard code stays on frontier models.",
    ),
    (
        "Financial services",
        0.18,
        0.024,
        0.140,
        0.03,
        0.08,
        (0.15, 0.30, 0.55),
        "Dollar-weighted toward the largest firms. Model-risk and residency reviews. Origin screens keep Chinese-origin weights small.",
    ),
    (
        "Healthcare and life sciences",
        0.09,
        0.022,
        0.160,
        0.03,
        0.10,
        (0.20, 0.35, 0.45),
        "Clinical acceptance. Highest eval share of model dollars. Open-weight is a narrow, low-risk slice.",
    ),
    (
        "Professional services and legal",
        0.10,
        0.030,
        0.115,
        0.05,
        0.18,
        (0.25, 0.40, 0.35),
        "Work product is reviewed before it ships. More mid-size firms than banking. Evals are matter-level, not just CI.",
    ),
    (
        "Retail and consumer",
        0.08,
        0.050,
        0.090,
        0.12,
        0.30,
        (0.20, 0.35, 0.45),
        "High-volume support, classification, and extraction. Open-weight fits the harness. Online evals are common; they are still a small share of tokens.",
    ),
    (
        "Manufacturing and industrial",
        0.08,
        0.044,
        0.070,
        0.12,
        0.25,
        (0.15, 0.35, 0.50),
        "Batch and operations workloads. Large plants dominate dollars. Eval intensity is lower than in regulated services.",
    ),
    (
        "Telecom, media, and other",
        0.07,
        0.038,
        0.075,
        0.08,
        0.22,
        (0.15, 0.30, 0.55),
        "Mixed buyer set. Large carriers and platforms dominate the dollars.",
    ),
    (
        "Public sector and education",
        0.06,
        0.032,
        0.120,
        0.04,
        0.05,
        (0.10, 0.30, 0.60),
        "Procurement and residency. Slower adoption, higher documentation burden, very low Chinese-origin share.",
    ),
]


def apply_header_footer(ws, subtitle):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "Proprietary vs open-weight  |  Evals vs harness"
    ws.oddHeader.right.text = subtitle
    ws.oddFooter.left.text = "Modeled allocation on reported anchors. Yellow cells are inputs."
    ws.oddFooter.right.text = "As of 17 November 2026  |  Page &P"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.sheet_properties.tabColor = NAVY
    ws.freeze_panes = "A2"
    ws.sheet_format.defaultRowHeight = 18
    ws.oddHeader.left.font = "Calibri"
    ws.print_title_rows = "1:1"


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def paint(ws, row, start, end, fill, font=None, align=None):
    for col in range(start, end + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        if font:
            cell.font = font
        if align:
            cell.alignment = align


def section(ws, row, text, cols):
    ws.cell(row, 1, text).font = font_section
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    ws.row_dimensions[row].height = 22


def input_cell(cell, value, fmt=None, comment=None):
    cell.value = value
    cell.font = font_input
    cell.fill = fill_gold
    cell.alignment = Alignment(vertical="center", horizontal="center")
    cell.border = thin
    if fmt:
        cell.number_format = fmt
    if comment:
        cell.comment = Comment(comment, "Model")
        cell.comment.width = 280
        cell.comment.height = 80


def label(cell, text, bold=False):
    cell.value = text
    cell.font = Font(name="Calibri", size=10, bold=bold, color=INK)
    cell.alignment = wrap


def header_row(ws, row, headers, fill=fill_navy):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.font = font_h_white
        c.fill = fill
        c.alignment = center
        c.border = thin
    ws.row_dimensions[row].height = 32
    ws.auto_filter.ref = None
    ws.freeze_panes = f"A{row+1}"


def build():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)

    # ------------------------------------------------------------------
    # Anchors: reported numbers only. Nothing on this sheet is a judgment.
    # ------------------------------------------------------------------
    an = wb.active
    an.title = "Anchors"
    apply_header_footer(an, "Reported anchors")
    an.sheet_properties.tabColor = "2E7D4F"
    an["A1"] = "Reported anchors"
    an["A1"].font = font_title
    an["A2"] = (
        "Every figure on this sheet is taken from a published source. Nothing here is allocated by size or vertical. "
        "The model on Assumptions reads the cells in column C. Do not type over a formula."
    )
    an["A2"].font = font_small
    an["A2"].alignment = wrap
    an.row_dimensions[2].height = 32

    headers = ["ID", "Metric", "Value", "Unit", "Source", "Date", "What it is, and what it is not"]
    for i, h in enumerate(headers, 1):
        c = an.cell(4, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    an.row_dimensions[4].height = 22
    an.freeze_panes = "A5"

    # (id, metric, value, unit, source, date, note, fmt)
    anchors = [
        ("DEP", "Enterprise AI deployment spend", 176.8, "$ billions, annualized run-rate", "Menlo Ventures, State of Enterprise AI", "Sep 2026",
         "Using AI, not training it. Menlo bottoms-up estimate from a survey of nearly 400 US enterprise buyers plus filings. Not a census.", FMT_B),
        ("APP", "Of which applications", 77.3, "$ billions", "Menlo", "Sep 2026",
         "Seat and app revenue: coding, chat, vertical apps. Includes $40B of lab first-party apps. Not model-token invoices.", FMT_B),
        ("INF", "Of which infrastructure", 99.5, "$ billions", "Menlo", "Sep 2026",
         "Frontier inference, inference platforms, and agent infrastructure together. Menlo does not split this line.", FMT_B),
        ("LAB", "Frontier labs' deployment revenue", 122.7, "$ billions", "Menlo", "Sep 2026",
         "Anthropic, OpenAI, and Google. 69% of deployment dollars. Spans API inference and first-party apps.", FMT_B),
        ("LAPP", "Lab first-party applications", 40.0, "$ billions", "Menlo", "Sep 2026",
         "Claude Code, Codex, enterprise Claude, ChatGPT, Gemini. Inside both the $122.7B and the $77.3B.", FMT_B),
        ("ANTH", "Anthropic share of enterprise LLM API spend", 0.55, "share of API dollars", "Menlo", "Sep 2026",
         "Up from 40% a year earlier. API dollars, not all deployment dollars.", FMT_PCT),
        ("OAI", "OpenAI share of enterprise LLM API spend", 0.31, "share of API dollars", "Menlo", "Sep 2026",
         "Up from 27%.", FMT_PCT),
        ("GOOG", "Google share of enterprise LLM API spend", 0.09, "share of API dollars", "Menlo", "Sep 2026",
         "Down from 21%. Skews to Flash.", FMT_PCT),
        ("LABAPI", "Lab share of API dollars (sum of the three)", 0.95, "share of API dollars", "Menlo, sum", "Sep 2026",
         "55+31+9. The residual ~5% is xAI plus open-weight plus others. Menlo says xAI and named open models together are just under 5%.", FMT_PCT),
        ("OPEN_D", "Open-weight share of model spend", 0.04, "share of model dollars", "Menlo", "Sep 2026",
         "Down from 6% in Sep 2025. This is the dollar anchor the allocation is checked against. It is not forced.", FMT_PCT),
        ("OPEN_T", "Open-weight share of tokens", 0.19, "share of tokens", "Menlo", "Sep 2026",
         "Up from 11% in Sep 2025. Usage rose while spend share fell because the open-weight price fell 3.3x.", FMT_PCT),
        ("FR_D", "Frontier share of model spend", 0.63, "share of model dollars", "Menlo", "Sep 2026",
         "Up from 35% in Q3 2025. Frontier means the top tier of each lab's lineup, not every closed model.", FMT_PCT),
        ("FR_T", "Frontier share of tokens", 0.42, "share of tokens", "Menlo", "Sep 2026",
         "Up from 18%.", FMT_PCT),
        ("SM_T", "Small closed-model share of tokens", 0.13, "share of tokens", "Menlo", "Sep 2026",
         "Down from 29%. Haiku, Luna, Gemini Flash class. Spend share was not published.", FMT_PCT),
        ("P_CL", "Blended closed-model price", 0.70, "$ per million tokens", "Menlo", "Sep 2026",
         "Down 24% year over year. A blend across frontier, mid, and small, after cache reads.", FMT_PX),
        ("P_OP", "Blended open-weight price", 0.13, "$ per million tokens", "Menlo", "Sep 2026",
         "Down 3.3x year over year.", FMT_PX),
        ("OP_Q", "Open-weight volume", 2.9, "quadrillion tokens per month", "Menlo", "Sep 2026",
         "Up 22.5x. Triangulated with public revenue and token disclosures. May be wider than the enterprise survey perimeter.", FMT_N),
        ("CN", "Chinese-origin share of open-weight spend", 0.89, "share of open-weight dollars", "Menlo", "Sep 2026",
         "Kimi 61%, DeepSeek 15%, GLM 11%, Qwen 3%, Llama 7%. Triangulated with OpenRouter, Ramp, and Vercel, so developer gateways pull this up. See Open origin.", FMT_PCT),
        ("FUT5", "Open-weight share of metered AI consumption", 0.05, "share of vendor bill", "Futurum AI Tools Pulse, N=100", "Oct 2026",
         "Average across enterprises. Denominator includes Copilot (36%), Claude (29%), ChatGPT (15%), Gemini (10%). Self-hosted GPUs that never become a model invoice are outside this 5%.", FMT_PCT),
        ("FUT42", "Enterprises with open-weight in production", 0.42, "share of enterprises", "Futurum drill-down, N=200", "Sep 2026",
         "Up from 31% in July 2026. Incidence, not dollars. Predominantly North American respondents.", FMT_PCT),
        ("FUT34", "Open-weight share of tokens, production users only", 0.34, "share of tokens", "Futurum, N=83 production users", "Sep 2026",
         "Conditional on already running open-weight in production. Up from 23% a year earlier. Not the unconditional 19% Menlo reports.", FMT_PCT),
        ("A16Z", "Open-source share of enterprise AI spending, prior CIO survey", 0.11, "share of AI spending", "a16z CIO survey, N=100 Global 2000", "Jan 2026",
         "Down from 19% the year before. Broader denominator than Menlo's model-dollar 4%: Global 2000 AI spending, early 2026, not Sep 2026 token invoices. Do not average with the 4%.", FMT_PCT),
    ]

    for i, (aid, metric, value, unit, source, date, note, fmt) in enumerate(anchors):
        r = 5 + i
        an.cell(r, 1, aid).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        an.cell(r, 2, metric).font = font_label
        an.cell(r, 2).alignment = wrap
        cell = an.cell(r, 3, value)
        cell.font = Font(name="Calibri", size=10, bold=True, color=GREEN_INK)
        cell.fill = fill_green
        cell.number_format = fmt
        cell.alignment = center
        cell.border = thin
        an.cell(r, 4, unit).font = font_small
        an.cell(r, 5, source).font = font_label
        an.cell(r, 6, date).font = font_label
        an.cell(r, 6).alignment = center
        an.cell(r, 7, note).font = font_label
        an.cell(r, 7).alignment = wrap
        for col in range(1, 8):
            an.cell(r, col).border = thin
            an.cell(r, col).alignment = wrap if col in (2, 7) else Alignment(vertical="center", horizontal="center" if col in (1, 3, 6) else "left", wrap_text=True)
        an.row_dimensions[r].height = 48
        if i % 2:
            for col in (1, 2, 4, 5, 6, 7):
                if an.cell(r, col).fill.fgColor is None or col != 3:
                    an.cell(r, col).fill = fill_zebra

    an["A28"] = (
        "Two open-weight dollar figures are both real and not the same thing. Menlo's 4% is model spend in Sep 2026. "
        "a16z's 11% is a broader AI-spending share among Global 2000 CIOs, measured earlier, after a decline from 19%. "
        "Futurum's 5% is the share of a vendor bill that includes Microsoft 365 Copilot seats. This workbook checks itself against Menlo's 4% and 19%, because those match the definition used here: model dollars and model tokens."
    )
    an["A28"].alignment = wrap_top
    an["A28"].font = font_body
    an.merge_cells("A28:G28")
    an.row_dimensions[28].height = 48

    set_widths(an, {"A": 12, "B": 62, "C": 14, "D": 36, "E": 42, "F": 14, "G": 78})
    an.auto_filter.ref = "A4:G26"
    an.auto_filter.ref = "A4:G26"
    an.print_title_rows = "1:4"
    an.page_setup.fitToWidth = 1
    an.oddHeader.right.text = "Reported anchors"
    an.freeze_panes = "A5"

    # Row map for formulas. anchors start at row 5 in the order above.
    # DEP 5, APP 6, INF 7, LAB 8, LAPP 9, ANTH 10, OAI 11, GOOG 12, LABAPI 13,
    # OPEN_D 14, OPEN_T 15, FR_D 16, FR_T 17, SM_T 18, P_CL 19, P_OP 20, OP_Q 21, CN 22,
    # FUT5 23, FUT42 24, FUT34 25, A16Z 26

    # ------------------------------------------------------------------
    # Assumptions
    # ------------------------------------------------------------------
    asm = wb.create_sheet("Assumptions")
    apply_header_footer(asm, "Inputs")
    asm.sheet_properties.tabColor = "C4A35A"
    asm["A1"] = "Inputs"
    asm["A1"].font = font_title
    asm["A2"] = "Blue text on gold is an input. Everything else in the workbook reads these cells. Change a gold cell and the summaries move."
    asm["A2"].font = font_small
    asm.merge_cells("A2:F2")

    asm["A4"] = "1  Pool"
    asm["A4"].font = font_section
    label(asm["A5"], "Which pool is active")
    input_cell(
        asm["C5"],
        1,
        "0",
        "1 = lab-implied pool. 2 = token-implied pool. 3 = the manual override in C10. "
        "Default is 1. The token method treats Menlo's 2.9 quadrillion open-weight tokens a month as if they were the same perimeter as the 4% spend share. That pool is about 30% higher and may include non-enterprise volume.",
    )
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    dv.error = "Enter 1, 2, or 3"
    dv.errorTitle = "Pool switch"
    dv.add("C5")
    asm.add_data_validation(dv)
    asm["D5"] = '=IF(C5=1,"Lab-implied",IF(C5=2,"Token-implied","Manual override"))'
    asm["D5"].font = Font(name="Calibri", size=10, italic=True, color=TEAL)

    label(asm["A6"], "Lab deployment ($B)", True)
    asm["C6"] = "=Anchors!C8"
    label(asm["A7"], "Lab first-party apps ($B)")
    asm["C7"] = "=Anchors!C9"
    label(asm["A8"], "Lab share of API dollars")
    asm["C8"] = "=Anchors!C13"
    label(asm["A9"], "Lab-implied pool ($B)")
    asm["C9"] = "=(C6-C7)/C8"
    asm["D9"] = "Lab revenue ex first-party apps, grossed up by the labs' 95% of API dollars. Treats the residual as other closed plus open-weight."
    label(asm["A10"], "Manual pool ($B)")
    input_cell(asm["C10"], 90, FMT_B, "Used only when the switch in C5 is 3.")
    label(asm["A11"], "Token-implied pool ($B)")
    asm["C11"] = "=Anchors!C21*12*Anchors!C20/Anchors!C14"
    asm["D11"] = "2.9 quadrillion open tokens per month, at $0.13 per million, divided by the 4% spend share, annualized. Wider perimeter risk."
    label(asm["A12"], "Active pool ($B)", True)
    asm["C12"] = '=IF(C5=1,C9,IF(C5=2,C11,C10))'
    asm["C12"].font = font_kpi
    asm["C12"].fill = fill_green
    asm["C12"].number_format = FMT_B
    asm["C12"].border = thin
    asm["D12"] = "This is the dollar pool the size and vertical grid allocates. It is model inference, not the $176.8B deployment market."

    for r, fmt in ((6, FMT_B), (7, FMT_B), (8, FMT_PCT), (9, FMT_B), (11, FMT_B)):
        asm.cell(r, 3).number_format = fmt
        asm.cell(r, 3).font = Font(name="Calibri", size=10, color=GREEN_INK)
        asm.cell(r, 3).fill = fill_green
        asm.cell(r, 3).border = thin
        asm.cell(r, 3).alignment = center
    for r in range(5, 13):
        asm.cell(r, 4).font = font_small
        asm.cell(r, 4).alignment = wrap
        asm.row_dimensions[r].height = 28
    asm.row_dimensions[5].height = 20
    asm.row_dimensions[9].height = 32
    asm.row_dimensions[11].height = 32
    asm.row_dimensions[12].height = 32

    asm["A14"] = "2  Prices used to turn dollars into tokens"
    asm["A14"].font = font_section
    asm["A15"] = (
        "These four prices are inputs. They are set so a base-case mix lands near Menlo's blended $0.70 closed and $0.13 open. "
        "Eval prices are higher than harness prices because eval prompts cache poorly and the judge is more often a frontier model. "
        "Dollar shares do not depend on these prices. Token shares do."
    )
    asm["A15"].alignment = wrap
    asm["A15"].font = font_small
    asm.merge_cells("A15:F15")
    asm.row_dimensions[15].height = 32

    price_rows = [
        (16, "Harness, proprietary ($/M tokens)", 0.66, "Near the closed blend. Production agent traffic is cache-heavy."),
        (17, "Evals, proprietary ($/M tokens)", 1.50, "Frontier-skewed judges and fresh prompts. Less cache reuse than a production loop."),
        (18, "Harness, open-weight ($/M tokens)", 0.125, "Just under Menlo's $0.13 blend. Most open tokens are harness."),
        (19, "Evals, open-weight ($/M tokens)", 0.22, "Small open judges cost more than bulk open inference and less than a frontier judge."),
    ]
    for r, text, val, note in price_rows:
        label(asm.cell(r, 1), text)
        input_cell(asm.cell(r, 3), val, FMT_PX, note)
        asm.cell(r, 4, note).font = font_small
        asm.cell(r, 4).alignment = wrap
        asm.row_dimensions[r].height = 20

    asm["A21"] = "3  Size tilt"
    asm["A21"].font = font_section
    asm["A22"] = (
        "Applied on top of each vertical's open-weight and eval shares. Above 1 raises that band's share. "
        "Direction, not a published cross-tab: open-weight dollar depth is higher in the 100-999 band (one team standardizes on a cheap model) and shallower at 10,000+ (an open model is added for one or two high-volume tasks; the default stays closed). "
        "Eval share of dollars rises with size, because platform teams and regulated promotion processes show up once agents are in production. LangChain's survey found production agents at 67% of 10,000+ employee orgs and 50% below 100."
    )
    asm["A22"].alignment = wrap
    asm["A22"].font = font_small
    asm.merge_cells("A22:F22")
    asm.row_dimensions[22].height = 48

    for col, h in enumerate(["Employee band", "Open-weight multiplier", "Eval multiplier"], 1):
        c = asm.cell(23, col, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    for i, (name, om, em) in enumerate(zip(SIZES, SIZE_OPEN, SIZE_EVAL)):
        r = 24 + i
        asm.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True)
        asm.cell(r, 1).alignment = Alignment(vertical="center")
        asm.cell(r, 1).border = thin
        input_cell(asm.cell(r, 2), om, "0.00")
        input_cell(asm.cell(r, 3), em, "0.00")

    asm["A28"] = "4  Global dials"
    asm["A28"].font = font_section
    dials = [
        (29, "Open-weight intensity", 1.0, "0.00", "Scales every vertical's open-weight dollar share. 1.00 is the base case. The cap below still binds."),
        (30, "Eval intensity", 1.0, "0.00", "Scales every vertical's eval dollar share. 1.25 means evals take 25% more of each segment's model budget, until the cap."),
        (31, "Use one Chinese-origin share for every vertical", 0, "0", "0 = use the vertical rates in the table. 1 = stamp Menlo's 89% on every vertical. Default is 0, because Futurum's enterprise production sample shows origin screens tightening and Chinese-origin families losing accounts."),
        (32, "Flat Chinese-origin share", 0.89, FMT_PCT, "Read from the anchor when you want the published 89%. The cell is an input so you can type another flat rate."),
        (33, "Cap on open-weight dollar share", 0.40, FMT_PCT, "Keeps a multiplier from pushing a segment above 40% open-weight dollars."),
        (34, "Cap on eval dollar share", 0.45, FMT_PCT, "Keeps a multiplier from pushing a segment above 45% eval dollars."),
        (35, "Frontier share of proprietary eval dollars", 0.85, FMT_PCT, "Input. Evals are where buyers pay for the best judge. Harness frontier share is solved on Summary so that frontier lands on Menlo's 63% of all model dollars."),
    ]
    for r, text, val, fmt, note in dials:
        label(asm.cell(r, 1), text, bold=True)
        input_cell(asm.cell(r, 3), val, fmt, note)
        asm.cell(r, 4, note).font = font_small
        asm.cell(r, 4).alignment = wrap
        asm.row_dimensions[r].height = 32
    # Flat Chinese share should display the anchor but remain an input. Keep 0.89 typed, and show the anchor beside it.
    asm["E32"] = '=Anchors!C22'
    asm["E32"].number_format = FMT_PCT
    asm["E32"].font = Font(name="Calibri", size=9, color=GREEN_INK)
    asm["F32"] = "Menlo's published 89%, for reference. Not linked into C32, so C32 stays an input."
    asm["F32"].font = font_small
    dv2 = DataValidation(type="list", formula1='"0,1"', allow_blank=False)
    dv2.add("C31")
    asm.add_data_validation(dv2)

    asm["A37"] = "5  Buyer verticals"
    asm["A37"].font = font_section
    asm["A38"] = (
        "Dollar weights are a judgment about who buys model inference, not Menlo's app-category mix. "
        "Technology is overweighted versus headcount because that is where spend per firm is highest (Ramp's top 1% of firms pay about $10,200 per employee per month for AI; the median firm paying for AI pays $16). "
        "A bank's coding-agent bill stays in Financial services. It is not moved into Technology. "
        "Open-weight and eval shares are the shares before the size tilt. Size mixes must each sum to 100%."
    )
    asm["A38"].alignment = wrap
    asm["A38"].font = font_small
    asm.merge_cells("A38:L38")
    asm.row_dimensions[38].height = 48

    vh = [
        "Vertical",
        "Share of model $",
        "Open-weight share of $",
        "Eval share of $",
        "Evals as a share of open-weight $",
        "Chinese-origin share of open-weight $",
        "Mix, 100-999",
        "Mix, 1,000-9,999",
        "Mix, 10,000+",
        "Mix check",
        "Why this row is set this way",
    ]
    for i, h in enumerate(vh, 1):
        c = asm.cell(39, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    asm.row_dimensions[39].height = 36

    for i, (name, w, op, ev, oef, cn, mixes, note) in enumerate(VERTS):
        r = 40 + i
        asm.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        asm.cell(r, 1).alignment = wrap
        asm.cell(r, 1).border = thin
        input_cell(asm.cell(r, 2), w, FMT_PCT)
        input_cell(asm.cell(r, 3), op, FMT_PCT)
        input_cell(asm.cell(r, 4), ev, FMT_PCT)
        input_cell(asm.cell(r, 5), oef, FMT_PCT)
        input_cell(asm.cell(r, 6), cn, FMT_PCT)
        input_cell(asm.cell(r, 7), mixes[0], FMT_PCT)
        input_cell(asm.cell(r, 8), mixes[1], FMT_PCT)
        input_cell(asm.cell(r, 9), mixes[2], FMT_PCT)
        asm.cell(r, 10, f"=G{r}+H{r}+I{r}")
        asm.cell(r, 10).number_format = FMT_PCT
        asm.cell(r, 10).font = Font(name="Calibri", size=10, bold=True)
        asm.cell(r, 10).alignment = center
        asm.cell(r, 10).border = thin
        asm.cell(r, 11, note).font = font_small
        asm.cell(r, 11).alignment = wrap
        asm.row_dimensions[r].height = 48

    asm["A48"] = "Weight check"
    asm["A48"].font = Font(name="Calibri", size=10, bold=True)
    asm["B48"] = "=SUM(B40:B47)"
    asm["B48"].number_format = FMT_PCT
    asm["B48"].font = Font(name="Calibri", size=12, bold=True)
    asm["B48"].fill = fill_green
    asm["B48"].border = thin
    asm["B48"].alignment = center
    asm["C48"] = '=IF(ABS(B48-1)<0.001,"Weights sum to 100%.","Weights must sum to 100%.")'
    asm["C48"].font = Font(name="Calibri", size=10, italic=True, color=TEAL)
    asm.conditional_formatting.add(
        "J40:J47",
        CellIsRule(operator="notBetween", formula=["0.999", "1.001"], fill=PatternFill("solid", fgColor="F4C7C3")),
    )
    asm.conditional_formatting.add(
        "B48",
        CellIsRule(operator="notBetween", formula=["0.999", "1.001"], fill=PatternFill("solid", fgColor="F4C7C3")),
    )

    asm["A50"] = (
        "Employee bands are headcount, not revenue. 100-999 is the mid-market. 1,000-9,999 is enterprise. 10,000+ is large enterprise. "
        "Firms under 100 employees are outside the pool. Menlo's $176.8B is an enterprise deployment figure. "
        "Ramp's micro and SMB adoption rates are high, and a thin tail of AI-native product companies spends like the top 1%, but that tail is not in this grid."
    )
    asm["A50"].alignment = wrap
    asm["A50"].font = font_small
    asm.merge_cells("A50:K50")
    asm.row_dimensions[50].height = 36

    set_widths(
        asm,
        {
            "A": 62,
            "B": 24,
            "C": 22,
            "D": 42,
            "E": 18,
            "F": 28,
            "G": 18,
            "H": 20,
            "I": 16,
            "J": 14,
            "K": 62,
            "L": 18,
        },
    )
    asm.freeze_panes = "A40"
    asm.auto_filter.ref = "A39:K47"
    asm.print_title_rows = "1:1"
    asm.page_setup.fitToWidth = 1
    asm.oddHeader.right.text = "Inputs"

    # ------------------------------------------------------------------
    # Engine: one row per vertical x size. All formulas.
    # ------------------------------------------------------------------
    eng = wb.create_sheet("Engine")
    apply_header_footer(eng, "Calculation")
    eng.sheet_properties.tabColor = "8A8F93"
    eng["A1"] = "Calculation"
    eng["A1"].font = font_title
    eng["A2"] = (
        "One row per buyer vertical and employee band. Do not edit this sheet. "
        "Dollars are $ millions. Tokens are trillions per year. "
        "Open-weight share = min(cap, vertical open share x size multiplier x open intensity). "
        "Eval share is the same shape. Open-weight eval dollars are the open-weight book times that vertical's eval fraction. "
        "Proprietary eval dollars are the eval target minus open-weight eval dollars, and cannot go below zero."
    )
    eng["A2"].font = font_small
    eng["A2"].alignment = wrap
    eng.merge_cells("A2:Y2")
    eng.row_dimensions[2].height = 36

    eh = [
        "Vertical",
        "Employees",
        "Raw weight",
        "Share of pool",
        "Model $M",
        "Open share of $",
        "Eval share of $",
        "Open $M",
        "Proprietary $M",
        "Open eval $M",
        "Open harness $M",
        "Proprietary eval $M",
        "Proprietary harness $M",
        "Eval $M",
        "Harness $M",
        "Chinese-origin open $M",
        "Other open $M",
        "Open harness tokens (T)",
        "Open eval tokens (T)",
        "Prop. harness tokens (T)",
        "Prop. eval tokens (T)",
        "Total tokens (T)",
        "Open tokens (T)",
        "Eval tokens (T)",
        "Flag",
    ]
    for i, h in enumerate(eh, 1):
        c = eng.cell(4, i, h)
        c.font = font_h_white
        c.fill = fill_head2
        c.alignment = center
        c.border = thin
    eng.row_dimensions[4].height = 36
    eng.freeze_panes = "C5"

    # Assumptions vertical block is rows 40-47, size block rows 24-26.
    for i, (name, *_) in enumerate(VERTS):
        for j, size in enumerate(SIZES):
            r = 5 + i * 3 + j
            eng.cell(r, 1, name).font = font_label
            eng.cell(r, 2, size).font = font_label
            # raw weight = vertical weight * size mix / mix check
            # size mix columns G H I are 1,2,3 matching SIZES order
            eng.cell(
                r,
                3,
                f'=INDEX(Assumptions!$B$40:$B$47,MATCH(A{r},Assumptions!$A$40:$A$47,0))'
                f'*INDEX(Assumptions!$G$40:$I$47,MATCH(A{r},Assumptions!$A$40:$A$47,0),MATCH(B{r},Assumptions!$A$24:$A$26,0))'
                f'/INDEX(Assumptions!$J$40:$J$47,MATCH(A{r},Assumptions!$A$40:$A$47,0))',
            )
            eng.cell(r, 4, f"=C{r}/SUM($C$5:$C$28)")
            eng.cell(r, 5, f"=Assumptions!$C$12*1000*D{r}")
            eng.cell(
                r,
                6,
                f'=MIN(Assumptions!$C$33,INDEX(Assumptions!$C$40:$C$47,MATCH(A{r},Assumptions!$A$40:$A$47,0))'
                f'*INDEX(Assumptions!$B$24:$B$26,MATCH(B{r},Assumptions!$A$24:$A$26,0))*Assumptions!$C$29)',
            )
            eng.cell(
                r,
                7,
                f'=MIN(Assumptions!$C$34,INDEX(Assumptions!$D$40:$D$47,MATCH(A{r},Assumptions!$A$40:$A$47,0))'
                f'*INDEX(Assumptions!$C$24:$C$26,MATCH(B{r},Assumptions!$A$24:$A$26,0))*Assumptions!$C$30)',
            )
            eng.cell(r, 8, f"=E{r}*F{r}")
            eng.cell(r, 9, f"=E{r}-H{r}")
            eng.cell(
                r,
                10,
                f"=H{r}*INDEX(Assumptions!$E$40:$E$47,MATCH(A{r},Assumptions!$A$40:$A$47,0))",
            )
            eng.cell(r, 11, f"=H{r}-J{r}")
            eng.cell(r, 12, f"=MAX(0,E{r}*G{r}-J{r})")
            eng.cell(r, 13, f"=I{r}-L{r}")
            eng.cell(r, 14, f"=J{r}+L{r}")
            eng.cell(r, 15, f"=K{r}+M{r}")
            eng.cell(
                r,
                16,
                f'=H{r}*IF(Assumptions!$C$31=1,Assumptions!$C$32,INDEX(Assumptions!$F$40:$F$47,MATCH(A{r},Assumptions!$A$40:$A$47,0)))',
            )
            eng.cell(r, 17, f"=H{r}-P{r}")
            eng.cell(r, 18, f"=IF(Assumptions!$C$18=0,NA(),K{r}/Assumptions!$C$18)")
            eng.cell(r, 19, f"=IF(Assumptions!$C$19=0,NA(),J{r}/Assumptions!$C$19)")
            eng.cell(r, 20, f"=IF(Assumptions!$C$16=0,NA(),M{r}/Assumptions!$C$16)")
            eng.cell(r, 21, f"=IF(Assumptions!$C$17=0,NA(),L{r}/Assumptions!$C$17)")
            eng.cell(r, 22, f"=R{r}+S{r}+T{r}+U{r}")
            eng.cell(r, 23, f"=R{r}+S{r}")
            eng.cell(r, 24, f"=S{r}+U{r}")
            eng.cell(
                r,
                25,
                f'=IF(AND(L{r}<0.5,E{r}*G{r}>J{r}+1),"Eval target cut by the open-eval floor","")',
            )
            for col in range(1, 26):
                eng.cell(r, col).border = thin
                eng.cell(r, col).alignment = Alignment(vertical="center")
                if i % 2 == 0 and col <= 2:
                    eng.cell(r, col).fill = fill_zebra
            for col, fmt in (
                (3, "0.000"),
                (4, FMT_PCT),
                (5, FMT_M),
                (6, FMT_PCT),
                (7, FMT_PCT),
                (8, FMT_M),
                (9, FMT_M),
                (10, FMT_M),
                (11, FMT_M),
                (12, FMT_M),
                (13, FMT_M),
                (14, FMT_M),
                (15, FMT_M),
                (16, FMT_M),
                (17, FMT_M),
                (18, FMT_T),
                (19, FMT_T),
                (20, FMT_T),
                (21, FMT_T),
                (22, FMT_T),
                (23, FMT_T),
                (24, FMT_T),
            ):
                eng.cell(r, col).number_format = fmt
            eng.row_dimensions[r].height = 18

    # totals
    tr = 29
    eng.cell(tr, 1, "Total").font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    eng.cell(tr, 2, "").font = font_white
    for col in range(1, 26):
        eng.cell(tr, col).fill = fill_navy
        eng.cell(tr, col).font = font_h_white
        eng.cell(tr, col).border = thin
    for col in (3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24):
        eng.cell(tr, col, f"=SUM({get_column_letter(col)}5:{get_column_letter(col)}28)")
        eng.cell(tr, col).font = font_h_white
        eng.cell(tr, col).fill = fill_navy
    eng.cell(tr, 6, "=H29/E29")
    eng.cell(tr, 7, "=N29/E29")
    eng.cell(tr, 6).number_format = FMT_PCT
    eng.cell(tr, 7).number_format = FMT_PCT
    eng.cell(tr, 6).font = font_h_white
    eng.cell(tr, 7).font = font_h_white
    eng.cell(tr, 6).fill = fill_navy
    eng.cell(tr, 7).fill = fill_navy
    for col, fmt in (
        (3, "0.000"),
        (4, FMT_PCT),
        (5, FMT_M),
        (8, FMT_M),
        (9, FMT_M),
        (10, FMT_M),
        (11, FMT_M),
        (12, FMT_M),
        (13, FMT_M),
        (14, FMT_M),
        (15, FMT_M),
        (16, FMT_M),
        (17, FMT_M),
        (18, FMT_T),
        (19, FMT_T),
        (20, FMT_T),
        (21, FMT_T),
        (22, FMT_T),
        (23, FMT_T),
        (24, FMT_T),
    ):
        eng.cell(tr, col).number_format = fmt
        eng.cell(tr, col).font = font_h_white
        eng.cell(tr, col).fill = fill_navy

    eng.auto_filter.ref = "A4:Y28"
    widths = {"A": 36, "B": 16}
    for col in range(3, 26):
        widths[get_column_letter(col)] = 16
    widths["Y"] = 28
    set_widths(eng, widths)
    eng.auto_filter.ref = "A4:Y28"
    eng.page_setup.fitToWidth = 1
    eng.page_setup.fitToHeight = 1
    eng.print_title_rows = "1:4"
    eng.page_setup.orientation = "landscape"
    eng.oddHeader.right.text = "Calculation"
    eng.freeze_panes = "C5"

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    su = wb.create_sheet("Summary", 0)
    apply_header_footer(su, "Summary")
    su.sheet_properties.tabColor = TEAL
    su["A1"] = "Model spend: proprietary vs open-weight, evals vs harness"
    su["A1"].font = font_title
    su.merge_cells("A1:L1")
    su["A2"] = (
        "US enterprise model inference, annualized. Default pool is the lab-implied book (September 2026). "
        "The 2x2 is the result. Size and vertical cuts below it are an allocation, not a published cross-tab. Gold cells live on Assumptions."
    )
    su["A2"].font = font_small
    su["A2"].alignment = wrap
    su.merge_cells("A2:L2")
    su.row_dimensions[2].height = 32

    su["A4"] = "The book"
    su["A4"].font = font_section

    kpis = [
        (5, "Active pool", "=Assumptions!C12", FMT_B, "Model inference, $ billions"),
        (6, "Proprietary", "=SUM(Engine!I5:I28)/1000", FMT_B, "Closed weights. API invoices."),
        (7, "Open-weight", "=SUM(Engine!H5:H28)/1000", FMT_B, "Metered open-weight inference. Self-hosted GPUs that never become a model invoice are not in this line."),
        (8, "Harness", "=SUM(Engine!O5:O28)/1000", FMT_B, "Tokens burned serving a live workload."),
        (9, "Evals", "=SUM(Engine!N5:N28)/1000", FMT_B, "Tokens burned scoring, judging, regressing, or red-teaming. Not the salaries of the eval team."),
    ]
    for r, name, formula, fmt, note in kpis:
        su.cell(r, 1, name).font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        su.cell(r, 2, formula).font = font_kpi
        su.cell(r, 2).number_format = fmt
        su.cell(r, 2).fill = fill_paper
        su.cell(r, 2).border = thin
        su.cell(r, 2).alignment = center
        su.cell(r, 3, note).font = font_small
        su.row_dimensions[r].height = 20

    su["A11"] = "Shares, and the check against Menlo"
    su["A11"].font = font_section

    checks = [
        (12, "Open-weight share of dollars", "=B7/(B6+B7)", FMT_PCT),
        (13, "Menlo anchor, open-weight dollars", "=Anchors!C14", FMT_PCT),
        (14, "Gap, open-weight dollars", "=B12-B13", FMT_PP),
        (15, "Open-weight share of tokens", "=SUM(Engine!W5:W28)/SUM(Engine!V5:V28)", FMT_PCT),
        (16, "Menlo anchor, open-weight tokens", "=Anchors!C15", FMT_PCT),
        (17, "Gap, open-weight tokens", "=B15-B16", FMT_PP),
        (18, "Harness share of dollars", "=B8/(B8+B9)", FMT_PCT),
        (19, "Eval share of dollars", "=B9/(B8+B9)", FMT_PCT),
        (20, "Eval share of tokens", "=SUM(Engine!X5:X28)/SUM(Engine!V5:V28)", FMT_PCT),
        (21, "Blended proprietary price, $/M", "=B6*1000/SUM(Engine!T5:U28)", FMT_PX),
        (22, "Menlo closed blend, $/M", "=Anchors!C19", FMT_PX),
        (23, "Blended open-weight price, $/M", "=B7*1000/SUM(Engine!W5:W28)", FMT_PX),
        (24, "Menlo open blend, $/M", "=Anchors!C20", FMT_PX),
    ]
    for r, name, formula, fmt in checks:
        su.cell(r, 1, name).font = font_label
        su.cell(r, 2, formula).font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        su.cell(r, 2).number_format = fmt
        su.cell(r, 2).border = thin
        su.cell(r, 2).alignment = center
        su.cell(r, 2).fill = fill_paper
    for r in (13, 16, 22, 24):
        su.cell(r, 2).fill = fill_green
        su.cell(r, 2).font = Font(name="Calibri", size=12, bold=True, color=GREEN_INK)
    for r in (14, 17):
        su.cell(r, 2).fill = fill_amber

    su["A26"] = "Dollars, $ billions"
    su["A26"].font = font_section
    for col, h in enumerate(["", "Harness", "Evals", "Total"], 1):
        c = su.cell(27, col, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    su["A28"] = "Proprietary"
    su["B28"] = "=SUM(Engine!M5:M28)/1000"
    su["C28"] = "=SUM(Engine!L5:L28)/1000"
    su["D28"] = "=B28+C28"
    su["A29"] = "Open-weight"
    su["B29"] = "=SUM(Engine!K5:K28)/1000"
    su["C29"] = "=SUM(Engine!J5:J28)/1000"
    su["D29"] = "=B29+C29"
    su["A30"] = "Total"
    su["B30"] = "=B28+B29"
    su["C30"] = "=C28+C29"
    su["D30"] = "=D28+D29"
    for r in range(28, 31):
        su.cell(r, 1).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 1).border = thin
        for col in range(2, 5):
            su.cell(r, col).number_format = FMT_B
            su.cell(r, col).font = Font(name="Calibri", size=12, bold=True, color=INK)
            su.cell(r, col).alignment = center
            su.cell(r, col).border = thin
            su.cell(r, col).fill = fill_paper
    for col in range(1, 5):
        su.cell(30, col).fill = fill_navy
        su.cell(30, col).font = font_h_white

    su["A32"] = "Share of model dollars"
    su["A32"].font = font_section
    for col, h in enumerate(["", "Harness", "Evals", "Total"], 1):
        c = su.cell(33, col, h)
        c.font = font_h_white
        c.fill = fill_teal
        c.alignment = center
        c.border = thin
    su["A34"] = "Proprietary"
    su["B34"] = "=B28/$D$30"
    su["C34"] = "=C28/$D$30"
    su["D34"] = "=D28/$D$30"
    su["A35"] = "Open-weight"
    su["B35"] = "=B29/$D$30"
    su["C35"] = "=C29/$D$30"
    su["D35"] = "=D29/$D$30"
    su["A36"] = "Total"
    su["B36"] = "=B30/$D$30"
    su["C36"] = "=C30/$D$30"
    su["D36"] = "=D30/$D$30"
    for r in range(34, 37):
        su.cell(r, 1).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 1).border = thin
        for col in range(2, 5):
            su.cell(r, col).number_format = FMT_PCT
            su.cell(r, col).font = Font(name="Calibri", size=12, bold=True)
            su.cell(r, col).alignment = center
            su.cell(r, col).border = thin
    for col in range(1, 5):
        su.cell(36, col).fill = fill_teal
        su.cell(36, col).font = font_h_white

    su["F26"] = "Tokens, trillions per year"
    su["F26"].font = font_section
    for col, h in enumerate(["", "Harness", "Evals", "Total"], 6):
        c = su.cell(27, col, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    su["F28"] = "Proprietary"
    su["G28"] = "=SUM(Engine!T5:T28)"
    su["H28"] = "=SUM(Engine!U5:U28)"
    su["I28"] = "=G28+H28"
    su["F29"] = "Open-weight"
    su["G29"] = "=SUM(Engine!R5:R28)"
    su["H29"] = "=SUM(Engine!S5:S28)"
    su["I29"] = "=G29+H29"
    su["F30"] = "Total"
    su["G30"] = "=G28+G29"
    su["H30"] = "=H28+H29"
    su["I30"] = "=I28+I29"
    for r in range(28, 31):
        su.cell(r, 6).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 6).border = thin
        for col in range(7, 10):
            su.cell(r, col).number_format = FMT_T
            su.cell(r, col).font = Font(name="Calibri", size=12, bold=True)
            su.cell(r, col).alignment = center
            su.cell(r, col).border = thin
            su.cell(r, col).fill = fill_paper
    for col in range(6, 10):
        su.cell(30, col).fill = fill_navy
        su.cell(30, col).font = font_h_white

    su["F32"] = "Share of tokens"
    su["F32"].font = font_section
    for col, h in enumerate(["", "Harness", "Evals", "Total"], 6):
        c = su.cell(33, col, h)
        c.font = font_h_white
        c.fill = fill_teal
        c.alignment = center
        c.border = thin
    su["F34"] = "Proprietary"
    su["G34"] = "=G28/$I$30"
    su["H34"] = "=H28/$I$30"
    su["I34"] = "=I28/$I$30"
    su["F35"] = "Open-weight"
    su["G35"] = "=G29/$I$30"
    su["H35"] = "=H29/$I$30"
    su["I35"] = "=I29/$I$30"
    su["F36"] = "Total"
    su["G36"] = "=G30/$I$30"
    su["H36"] = "=H30/$I$30"
    su["I36"] = "=I30/$I$30"
    for r in range(34, 37):
        su.cell(r, 6).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 6).border = thin
        for col in range(7, 10):
            su.cell(r, col).number_format = FMT_PCT
            su.cell(r, col).font = Font(name="Calibri", size=12, bold=True)
            su.cell(r, col).alignment = center
            su.cell(r, col).border = thin
    for col in range(6, 10):
        su.cell(36, col).fill = fill_teal
        su.cell(36, col).font = font_h_white

    su["A38"] = "How to read the 2x2"
    su["A38"].font = font_section
    su["A39"] = (
        '=CONCATENATE('
        '"Open-weight is ",TEXT(B12,"0.0%")," of model dollars and ",TEXT(B15,"0.0%")," of tokens. ",'
        '"Evals are ",TEXT(B19,"0.0%")," of dollars and ",TEXT(B20,"0.0%")," of tokens. ",'
        '"Almost all eval dollars are proprietary: ",TEXT(C28,"$0.00")," billion versus ",TEXT(C29,"$0.00")," billion of open-weight evals. ",'
        '"Open-weight dollars sit in the harness: ",TEXT(B29,"$0.00")," billion of ",TEXT(B7,"$0.00")," billion. ",'
        '"The price gap is the whole story. A blended proprietary token costs about ",TEXT(B21,"$0.000"),'
        '" and a blended open-weight token costs about ",TEXT(B23,"$0.000"),".")'
    )
    su["A39"].alignment = wrap_top
    su["A39"].font = font_body
    su.merge_cells("A39:I39")
    su.row_dimensions[39].height = 48

    # By size
    su["A41"] = "By employee band"
    su["A41"].font = font_section
    size_headers = [
        "Employees",
        "Model $B",
        "Proprietary harness $B",
        "Proprietary evals $B",
        "Open harness $B",
        "Open evals $B",
        "Open share of $",
        "Eval share of $",
        "Open share of tokens",
        "Eval share of tokens",
    ]
    for i, h in enumerate(size_headers, 1):
        c = su.cell(42, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    su.row_dimensions[42].height = 32
    for i, size in enumerate(SIZES):
        r = 43 + i
        su.cell(r, 1, size).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 2, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$E$5:$E$28)/1000')
        su.cell(r, 3, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$M$5:$M$28)/1000')
        su.cell(r, 4, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$L$5:$L$28)/1000')
        su.cell(r, 5, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$K$5:$K$28)/1000')
        su.cell(r, 6, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$J$5:$J$28)/1000')
        su.cell(r, 7, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$H$5:$H$28)/SUMIF(Engine!$B$5:$B$28,A{r},Engine!$E$5:$E$28)')
        su.cell(r, 8, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$N$5:$N$28)/SUMIF(Engine!$B$5:$B$28,A{r},Engine!$E$5:$E$28)')
        su.cell(r, 9, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$W$5:$W$28)/SUMIF(Engine!$B$5:$B$28,A{r},Engine!$V$5:$V$28)')
        su.cell(r, 10, f'=SUMIF(Engine!$B$5:$B$28,A{r},Engine!$X$5:$X$28)/SUMIF(Engine!$B$5:$B$28,A{r},Engine!$V$5:$V$28)')
        for col in range(1, 11):
            su.cell(r, col).border = thin
            su.cell(r, col).alignment = center if col > 1 else Alignment(vertical="center")
        for col in range(2, 7):
            su.cell(r, col).number_format = FMT_B
        for col in range(7, 11):
            su.cell(r, col).number_format = FMT_PCT
        su.row_dimensions[r].height = 20

    su["A47"] = "By buyer vertical"
    su["A47"].font = font_section
    for i, h in enumerate(size_headers, 1):
        # reuse header text but first col is Vertical
        text = "Vertical" if i == 1 else h
        c = su.cell(48, i, text)
        c.font = font_h_white
        c.fill = fill_teal
        c.alignment = center
        c.border = thin
    su.row_dimensions[48].height = 32
    for i, (name, *_) in enumerate(VERTS):
        r = 49 + i
        su.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        su.cell(r, 2, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$E$5:$E$28)/1000')
        su.cell(r, 3, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$M$5:$M$28)/1000')
        su.cell(r, 4, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$L$5:$L$28)/1000')
        su.cell(r, 5, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$K$5:$K$28)/1000')
        su.cell(r, 6, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$J$5:$J$28)/1000')
        su.cell(r, 7, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$H$5:$H$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$E$5:$E$28)')
        su.cell(r, 8, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$N$5:$N$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$E$5:$E$28)')
        su.cell(r, 9, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$W$5:$W$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$V$5:$V$28)')
        su.cell(r, 10, f'=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$X$5:$X$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$V$5:$V$28)')
        for col in range(1, 11):
            su.cell(r, col).border = thin
            su.cell(r, col).alignment = center if col > 1 else Alignment(vertical="center", wrap_text=True)
        for col in range(2, 7):
            su.cell(r, col).number_format = FMT_B
        for col in range(7, 11):
            su.cell(r, col).number_format = FMT_PCT
        if i % 2:
            su.cell(r, 1).fill = fill_zebra
        su.row_dimensions[r].height = 20

    # chart data is the vertical block columns C-F (the four components)
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Model dollars by buyer vertical"
    chart.y_axis.title = "Annualized model spend ($ billions)"
    chart.x_axis.title = None
    data = Reference(su, min_col=3, max_col=6, min_row=48, max_row=56)
    cats = Reference(su, min_col=1, min_row=49, max_row=56)
    chart.add_data(data, from_rows=False, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.y_axis.numFmt = "$#,##0"
    chart.y_axis.majorGridlines = None
    colors = ["1B3A4B", "4C7EA8", "0E6B67", "8FBFB0"]
    for s, color in zip(chart.series, colors):
        s.graphicalProperties.solidFill = color
    chart.legend.position = "b"
    chart.y_axis.majorGridlines = openpyxl_grid()
    chart.style = 10
    chart.y_axis.scaling.min = 0
    chart.width = 18
    chart.height = 8
    su.add_chart(chart, "A58")

    su["A75"] = "Frontier, solved so the book matches Menlo's 63%"
    su["A75"].font = font_section
    su["A76"] = "Frontier target, share of all model dollars"
    su["B76"] = "=Anchors!C16"
    su["A77"] = "Frontier dollars ($B)"
    su["B77"] = "=B76*(B6+B7)"
    su["A78"] = "Assigned to proprietary evals ($B)"
    su["B78"] = "=Assumptions!C35*C28"
    su["A79"] = "Remainder, assigned to the proprietary harness ($B)"
    su["B79"] = "=B77-B78"
    su["A80"] = "Implied frontier share of proprietary harness dollars"
    su["B80"] = "=IF(B28=0,NA(),B79/B28)"
    su["A81"] = "Check"
    su["B81"] = '=IF(AND(B80>=0,B80<=1),"Holds","Outside 0-100%. Lower the eval frontier share on Assumptions.")'
    for r, fmt in ((76, FMT_PCT), (77, FMT_B), (78, FMT_B), (79, FMT_B), (80, FMT_PCT)):
        su.cell(r, 1).font = font_label
        su.cell(r, 2).number_format = fmt
        su.cell(r, 2).font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        su.cell(r, 2).border = thin
        su.cell(r, 2).alignment = center
        su.cell(r, 2).fill = fill_paper
    su["A82"] = (
        "Menlo published frontier's share of spend (63%) and of tokens (42%), not a split of frontier between evals and the harness. "
        "The 85% of proprietary eval dollars treated as frontier is an input. The harness share is the residual that makes the total 63%. "
        "If that residual falls outside 0-100%, the input is inconsistent with the eval book and the check cell says so."
    )
    su["A82"].font = font_small
    su["A82"].alignment = wrap
    su.merge_cells("A82:G82")
    su.row_dimensions[82].height = 36
    su["B81"].font = Font(name="Calibri", size=10, italic=True, color=TEAL)
    su["A81"].font = font_label

    set_widths(
        su,
        {
            "A": 58,
            "B": 18,
            "C": 24,
            "D": 22,
            "E": 20,
            "F": 22,
            "G": 16,
            "H": 16,
            "I": 22,
            "J": 20,
            "K": 18,
            "L": 18,
        },
    )
    su.freeze_panes = "A4"
    su.page_setup.fitToWidth = 1
    su.page_setup.fitToHeight = 0
    su.print_title_rows = "1:2"
    su.oddHeader.right.text = "Summary"
    su.row_dimensions[1].height = 28
    su.sheet_view.zoomScale = 110
    su.page_setup.orientation = "landscape"

    # ------------------------------------------------------------------
    # Matrix: the cross the user asked for
    # ------------------------------------------------------------------
    mx = wb.create_sheet("Size by vertical")
    apply_header_footer(mx, "Size by vertical")
    mx.sheet_properties.tabColor = "1F4E79"
    mx["A1"] = "Size by vertical"
    mx["A1"].font = font_title
    mx["A2"] = (
        "Each row is a buyer vertical. Each band is an employee count. "
        "Dollars are $ billions per year. Shares are of that cell's own model spend. "
        "This is the allocation grid. Edit Assumptions, not this sheet."
    )
    mx["A2"].font = font_small
    mx["A2"].alignment = wrap
    mx.merge_cells("A2:V2")
    mx.row_dimensions[2].height = 30

    # Two header rows
    mx["A4"] = "Vertical"
    mx["A4"].font = font_h_white
    mx["A4"].fill = fill_navy
    mx["A4"].alignment = center
    mx["A4"].border = thin
    # group headers
    groups = [
        (2, "100-999 employees"),
        (8, "1,000-9,999 employees"),
        (14, "10,000+ employees"),
        (20, "All sizes"),
    ]
    sub = ["Model $B", "Prop. harness", "Prop. evals", "Open harness", "Open evals", "Open $", "Eval $"]
    # Wait that's 7 per group * 4 = 28, plus vertical = 29. A bit wide.
    # Slimmer sub: Model $B, Open $, Eval $, Prop harness, Prop evals, Open harness, Open evals — same.
    # Use 6: Model, Open share, Eval share, Prop harness, Prop evals, Open harness. Open evals is small.
    # User wants the 2x2 visible. Keep 6 columns per band:
    # Model $B | Prop harness | Prop evals | Open harness | Open evals | Open share of $
    sub = ["Model $B", "Prop. harness $B", "Prop. evals $B", "Open harness $B", "Open evals $B", "Open share of $"]
    groups = [(2, "100-999", "100-999"), (8, "1,000-9,999", "1,000-9,999"), (14, "10,000+", "10,000+"), (20, "All employee bands", None)]
    for start, title, _key in groups:
        c = mx.cell(3, start, title)
        c.font = font_h_white
        c.fill = fill_teal if start == 20 else fill_navy
        c.alignment = center
        mx.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + 5)
        for k in range(6):
            cell = mx.cell(3, start + k)
            cell.fill = fill_teal if start == 20 else fill_navy
            cell.border = thin
            cell.font = font_h_white
        for k, name in enumerate(sub):
            cell = mx.cell(4, start + k, name)
            cell.font = font_h_white
            cell.fill = fill_head2
            cell.alignment = center
            cell.border = thin
    mx["A3"].fill = fill_navy
    mx["A3"].border = thin
    mx.row_dimensions[3].height = 20
    mx.row_dimensions[4].height = 32

    def band_formulas(row, key):
        """Return 6 formulas for a size key, or None key for all sizes."""
        if key is None:
            crit_a = f'Engine!$A$5:$A$28,A{row}'
            def s(col):
                return f'SUMIF(Engine!$A$5:$A$28,A{row},Engine!${col}$5:${col}$28)'
        else:
            def s(col):
                return (
                    f'SUMIFS(Engine!${col}$5:${col}$28,Engine!$A$5:$A$28,A{row},Engine!$B$5:$B$28,"{key}")'
                )
        model = s("E")
        return [
            f"=({model})/1000",
            f"=({s('M')})/1000",
            f"=({s('L')})/1000",
            f"=({s('K')})/1000",
            f"=({s('J')})/1000",
            f"=IF(({model})=0,NA(),({s('H')})/({model}))",
        ]

    for i, (name, *_) in enumerate(VERTS):
        r = 5 + i
        mx.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        mx.cell(r, 1).alignment = wrap
        mx.cell(r, 1).border = thin
        if i % 2:
            mx.cell(r, 1).fill = fill_zebra
        for start, _title, key in groups:
            formulas = band_formulas(r, key)
            for k, fml in enumerate(formulas):
                cell = mx.cell(r, start + k, fml)
                cell.border = thin
                cell.alignment = center
                cell.number_format = FMT_PCT if k == 5 else FMT_B
                cell.font = font_label
                if start == 20:
                    cell.fill = fill_paper
        mx.row_dimensions[r].height = 22

    # total row
    r = 13
    mx.cell(r, 1, "All verticals").font = font_h_white
    mx.cell(r, 1).fill = fill_navy
    mx.cell(r, 1).border = thin
    for start, _title, key in groups:
        if key is None:
            def s(col):
                return f"SUM(Engine!${col}$5:${col}$28)"
        else:
            def s(col, key=key):
                return f'SUMIF(Engine!$B$5:$B$28,"{key}",Engine!${col}$5:${col}$28)'
        model = s("E")
        formulas = [
            f"=({model})/1000",
            f"=({s('M')})/1000",
            f"=({s('L')})/1000",
            f"=({s('K')})/1000",
            f"=({s('J')})/1000",
            f"=IF(({model})=0,NA(),({s('H')})/({model}))",
        ]
        for k, fml in enumerate(formulas):
            cell = mx.cell(r, start + k, fml)
            cell.font = font_h_white
            cell.fill = fill_navy
            cell.border = thin
            cell.alignment = center
            cell.number_format = FMT_PCT if k == 5 else FMT_B
    mx.row_dimensions[r].height = 22

    mx["A15"] = (
        "Reading across a row shows how the same buyer changes with size. "
        "Reading down a band shows how industries differ at the same size. "
        "Open share of dollars falls as headcount rises, in every vertical, because the size multiplier does. "
        "Eval share of dollars rises with headcount. Absolute dollars still concentrate in technology and in the 10,000+ band, because those cells have the weight."
    )
    mx["A15"].font = font_body
    mx["A15"].alignment = wrap
    mx.merge_cells("A15:L15")
    mx.row_dimensions[15].height = 36

    mx.column_dimensions["A"].width = 36
    for col in range(2, 26):
        mx.column_dimensions[get_column_letter(col)].width = 16
    mx.freeze_panes = "B5"
    mx.page_setup.fitToWidth = 1
    mx.page_setup.orientation = "landscape"
    mx.page_setup.paperSize = mx.PAPERSIZE_TABLOID
    mx.print_title_rows = "1:4"
    mx.oddHeader.right.text = "Size by vertical"
    mx.sheet_view.zoomScale = 100

    # Token cross-tab, slimmer: open token share and eval token share by band
    tk = wb.create_sheet("Tokens")
    apply_header_footer(tk, "Tokens")
    tk.sheet_properties.tabColor = "4C7EA8"
    tk["A1"] = "Tokens by size and vertical"
    tk["A1"].font = font_title
    tk["A2"] = "Trillions of tokens per year, and the open-weight and eval shares of those tokens. Same grid as the dollar sheet. Prices on Assumptions are what turn one into the other."
    tk["A2"].font = font_small
    tk["A2"].alignment = wrap
    tk.merge_cells("A2:V2")
    tk.row_dimensions[2].height = 28
    tk["A4"] = "Vertical"
    tk["A4"].font = font_h_white
    tk["A4"].fill = fill_navy
    tk["A4"].alignment = center
    tk["A4"].border = thin
    tsub = ["Tokens (T)", "Open share", "Eval share", "Open harness (T)", "Prop. harness (T)", "Prop. evals (T)"]
    tgroups = [(2, "100-999", "100-999"), (8, "1,000-9,999", "1,000-9,999"), (14, "10,000+", "10,000+"), (20, "All employee bands", None)]
    for start, title, _key in tgroups:
        mxcell = tk.cell(3, start, title)
        mxcell.font = font_h_white
        mxcell.alignment = center
        tk.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + 5)
        for k in range(6):
            cell = tk.cell(3, start + k)
            cell.fill = fill_teal if start == 20 else fill_navy
            cell.border = thin
            cell.font = font_h_white
        for k, name in enumerate(tsub):
            cell = tk.cell(4, start + k, name)
            cell.font = font_h_white
            cell.fill = fill_head2
            cell.alignment = center
            cell.border = thin
    tk["A3"].fill = fill_navy
    tk["A3"].border = thin
    tk.row_dimensions[3].height = 20
    tk.row_dimensions[4].height = 32

    for i, (name, *_) in enumerate(VERTS):
        r = 5 + i
        tk.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        tk.cell(r, 1).border = thin
        tk.cell(r, 1).alignment = wrap
        for start, _title, key in tgroups:
            if key is None:
                def s(col, row=r):
                    return f"SUMIF(Engine!$A$5:$A$28,A{row},Engine!${col}$5:${col}$28)"
            else:
                def s(col, key=key, row=r):
                    return f'SUMIFS(Engine!${col}$5:${col}$28,Engine!$A$5:$A$28,A{row},Engine!$B$5:$B$28,"{key}")'
            tot = s("V")
            formulas = [
                f"={tot}",
                f"=IF(({tot})=0,NA(),({s('W')})/({tot}))",
                f"=IF(({tot})=0,NA(),({s('X')})/({tot}))",
                f"={s('R')}",
                f"={s('T')}",
                f"={s('U')}",
            ]
            for k, fml in enumerate(formulas):
                cell = tk.cell(r, start + k, fml)
                cell.border = thin
                cell.alignment = center
                cell.number_format = FMT_PCT if k in (1, 2) else FMT_T
                if start == 20:
                    cell.fill = fill_paper
        tk.row_dimensions[r].height = 22

    tk["A15"] = (
        "Open-weight's token share is several times its dollar share in every cell. "
        "That gap is the price, not a second survey. Eval tokens are a thinner slice than eval dollars, because eval tokens are priced like frontier judgments and harness tokens are priced like cached production traffic."
    )
    tk["A15"].font = font_body
    tk["A15"].alignment = wrap
    tk.merge_cells("A15:L15")
    tk.row_dimensions[15].height = 32
    tk.column_dimensions["A"].width = 36
    for col in range(2, 26):
        tk.column_dimensions[get_column_letter(col)].width = 18
    tk.freeze_panes = "B5"
    tk.page_setup.fitToWidth = 1
    tk.page_setup.orientation = "landscape"
    tk.page_setup.paperSize = tk.PAPERSIZE_TABLOID
    tk.print_title_rows = "1:4"
    tk.oddHeader.right.text = "Tokens"
    tk.sheet_view.zoomScale = 100

    # ------------------------------------------------------------------
    # Open origin
    # ------------------------------------------------------------------
    og = wb.create_sheet("Open origin")
    apply_header_footer(og, "Open origin")
    og.sheet_properties.tabColor = "8FBFB0"
    og["A1"] = "Where open-weight dollars go"
    og["A1"].font = font_title
    og["A2"] = (
        "Two pictures of the same small book. The live columns follow the vertical rates on Assumptions (or the flat switch, if you turn it on). "
        "The last column stamps Menlo's 89% Chinese-origin share on every vertical, which is what you get if developer-gateway spend is the right mix for enterprises. "
        "Futurum's enterprise production sample says it is not: origin is a major or deciding factor for 47% of production users, and Chinese-origin families lost accounts between July and September 2026."
    )
    og["A2"].alignment = wrap
    og["A2"].font = font_small
    og.merge_cells("A2:H2")
    og.row_dimensions[2].height = 48

    oh = [
        "Vertical",
        "Open-weight $B",
        "Chinese-origin $B",
        "Other open-weight $B",
        "Chinese-origin share",
        "If every vertical were 89% Chinese-origin, $B",
        "Open-weight share of that vertical's model $",
        "Open-weight share of that vertical's tokens",
    ]
    for i, h in enumerate(oh, 1):
        c = og.cell(4, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    og.row_dimensions[4].height = 36
    for i, (name, *_) in enumerate(VERTS):
        r = 5 + i
        og.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        og.cell(r, 2, f"=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$H$5:$H$28)/1000")
        og.cell(r, 3, f"=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$P$5:$P$28)/1000")
        og.cell(r, 4, f"=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$Q$5:$Q$28)/1000")
        og.cell(r, 5, f"=IF(B{r}=0,NA(),C{r}/B{r})")
        og.cell(r, 6, f"=B{r}*Anchors!C22")
        og.cell(r, 7, f"=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$H$5:$H$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$E$5:$E$28)")
        og.cell(r, 8, f"=SUMIF(Engine!$A$5:$A$28,A{r},Engine!$W$5:$W$28)/SUMIF(Engine!$A$5:$A$28,A{r},Engine!$V$5:$V$28)")
        for col in range(1, 9):
            og.cell(r, col).border = thin
            og.cell(r, col).alignment = center if col > 1 else Alignment(vertical="center")
        for col in (2, 3, 4, 6):
            og.cell(r, col).number_format = FMT_B
        for col in (5, 7, 8):
            og.cell(r, col).number_format = FMT_PCT
        og.row_dimensions[r].height = 20
    og["A13"] = "All verticals"
    og["B13"] = "=SUM(B5:B12)"
    og["C13"] = "=SUM(C5:C12)"
    og["D13"] = "=SUM(D5:D12)"
    og["E13"] = "=IF(B13=0,NA(),C13/B13)"
    og["F13"] = "=SUM(F5:F12)"
    og["G13"] = "=Summary!B12"
    og["H13"] = "=Summary!B15"
    for col in range(1, 9):
        og.cell(13, col).fill = fill_navy
        og.cell(13, col).font = font_h_white
        og.cell(13, col).border = thin
        og.cell(13, col).alignment = center
    for col in (2, 3, 4, 6):
        og.cell(13, col).number_format = FMT_B
    for col in (5, 7, 8):
        og.cell(13, col).number_format = FMT_PCT

    og["A15"] = "Reported production incidence, not dollars"
    og["A15"].font = font_section
    og["A16"] = (
        "Futurum, September 2026, 83 enterprises already running open-weight in production. Multiple responses, so shares sum past 100%. "
        "This is who shows up in a production account. It is not who collects the invoice. "
        "Menlo's 89% Chinese-origin spend share can coexist with this table if a minority of heavy spenders, and the developer gateways Menlo triangulates against, dominate open-weight invoices."
    )
    og["A16"].font = font_small
    og["A16"].alignment = wrap
    og.merge_cells("A16:H16")
    og.row_dimensions[16].height = 36
    for i, h in enumerate(["Family", "In production accounts, Sep 2026", "Prior reading", "Origin", "Note"], 1):
        c = og.cell(17, i, h)
        c.font = font_h_white
        c.fill = fill_teal
        c.alignment = center
        c.border = thin
    incidence = [
        ("Gemma (Google)", 0.55, 0.35, "US platform", "Up from 35% in July. The platform vendor gives the weights away and sells the capacity."),
        ("Phi (Microsoft)", 0.39, None, "US platform", "First cycle on Futurum's list."),
        ("Llama (Meta)", 0.35, 0.55, "US", "Down from 55% in July. Muse Spark Open shipped in October, after this fieldwork."),
        ("Nemotron (NVIDIA)", 0.25, None, "US platform", "First cycle on the list."),
        ("DeepSeek", 0.25, 0.32, "China", "Down from 32% of production accounts."),
        ("Mistral", 0.24, 0.32, "Europe", "Down from 32%. Has not picked up the origin-screen shift in this North American sample."),
        ("Qwen (Alibaba)", 0.10, 0.13, "China", "Down from 13%."),
        ("Kimi (Moonshot)", 0.06, 0.06, "China", "Held at 6% of production accounts. Pilot use fell from 13% to 6%."),
        ("GLM (Zhipu)", 0.01, 0.10, "China", "Down from 10%."),
    ]
    for i, (fam, now, prior, origin, note) in enumerate(incidence):
        r = 18 + i
        og.cell(r, 1, fam).font = font_label
        og.cell(r, 2, now).number_format = FMT_PCT
        og.cell(r, 2).font = Font(name="Calibri", size=10, bold=True, color=GREEN_INK)
        og.cell(r, 2).fill = fill_green
        if prior is None:
            og.cell(r, 3, "—")
        else:
            og.cell(r, 3, prior).number_format = FMT_PCT
            og.cell(r, 3).fill = fill_green
            og.cell(r, 3).font = Font(name="Calibri", size=10, color=GREEN_INK)
        og.cell(r, 4, origin).font = font_label
        og.cell(r, 5, note).font = font_small
        og.cell(r, 5).alignment = wrap
        for col in range(1, 6):
            og.cell(r, col).border = thin
            og.cell(r, col).alignment = wrap if col == 5 else Alignment(vertical="center", horizontal="center" if col in (2, 3) else "left", wrap_text=True)
        og.row_dimensions[r].height = 22

    og["A28"] = (
        "Hosting, same 83 production users, multiple responses: self-managed cloud 46%, on-premises 42%, hyperscaler-native inference 29%, managed inference providers 27%. "
        "The invoice for self-managed and on-prem capacity is often a GPU or cloud bill, not a model line. That spend is outside this workbook's open-weight dollars. "
        "Futurum's point, which this sheet does not try to price: the token volume that moves to open weights leaves the model bill and reappears as infrastructure."
    )
    og["A28"].font = font_body
    og["A28"].alignment = wrap
    og.merge_cells("A28:H28")
    og.row_dimensions[28].height = 48

    set_widths(og, {"A": 42, "B": 28, "C": 26, "D": 28, "E": 24, "F": 42, "G": 28, "H": 32})
    og.freeze_panes = "A5"
    og.page_setup.fitToWidth = 1
    og.oddHeader.right.text = "Open origin"

    # ------------------------------------------------------------------
    # Bridge
    # ------------------------------------------------------------------
    br = wb.create_sheet("Bridge")
    apply_header_footer(br, "Bridge")
    br.sheet_properties.tabColor = "8A5A2A"
    br["A1"] = "Where the pool sits inside the $176.8B"
    br["A1"].font = font_title
    br["A2"] = (
        "Menlo's deployment total and this workbook's model pool are not additive. "
        "The pool is the metered model-inference slice. Applications, and the agent software around the model, sit beside it. "
        "Lines marked residual are arithmetic, not a Menlo disclosure."
    )
    br["A2"].font = font_small
    br["A2"].alignment = wrap
    br.merge_cells("A2:E2")
    br.row_dimensions[2].height = 32

    for i, h in enumerate(["Line", "$ billions", "Status", "In the size x vertical grid?", "What the line is"], 1):
        c = br.cell(4, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    br.row_dimensions[4].height = 22

    bridge_rows = [
        (5, "Enterprise AI deployment", "=Anchors!C5", "Reported", "No", "Menlo, September 2026 annualized run-rate."),
        (6, "Applications", "=Anchors!C6", "Reported", "No", "Coding, chat, vertical apps, seat products. Includes lab first-party apps."),
        (7, "    Lab first-party apps", "=Anchors!C9", "Reported", "No", "Claude Code, Codex, enterprise Claude, ChatGPT, Gemini. $24B of the $34B coding market is Claude Code plus Codex."),
        (8, "    Other applications", "=B6-B7", "Residual", "No", "Applications minus lab first-party apps."),
        (9, "Infrastructure", "=Anchors!C7", "Reported", "Partly", "Menlo bundles frontier inference, inference platforms, and agent infrastructure. It does not split them."),
        (10, "    Lab revenue ex first-party apps", "=Anchors!C8-Anchors!C9", "Residual of two reported lines", "Yes, as proprietary", "Treated here as mostly API inference. Some non-inference lab revenue could sit in this residual."),
        (11, "    Active model pool", "=Assumptions!C12", "This workbook", "Yes, the whole grid", "Lab ex-apps divided by 0.95 when the switch is 1. See Assumptions."),
        (12, "    Of which open-weight", "=Summary!B7", "Modeled", "Yes", "Should land near Menlo's 4%. The gap is on Summary."),
        (13, "    Of which other closed (xAI and the like)", "=B11-B10-B12", "Residual", "Yes, inside proprietary", "Pool minus lab ex-apps minus open-weight. About 1% of the pool when the switch is 1, because the labs are 95% of API dollars and open-weight is ~4%."),
        (14, "    Agent software and other non-model infra", "=B9-B10-B12-B13", "Residual", "No", "What is left of Menlo's $99.5B infrastructure after lab inference, open-weight invoices, and other closed invoices. Sandboxes, search, browsers, orchestration, databases. This is harness software, not harness tokens."),
    ]
    for r, name, formula, status, grid, note in bridge_rows:
        br.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=(not name.startswith(" ")), color=NAVY)
        br.cell(r, 1).alignment = wrap
        br.cell(r, 2, formula).number_format = FMT_B
        br.cell(r, 2).font = Font(name="Calibri", size=12, bold=True, color=INK)
        br.cell(r, 2).alignment = center
        br.cell(r, 3, status).font = font_label
        br.cell(r, 3).alignment = center
        br.cell(r, 4, grid).font = font_label
        br.cell(r, 4).alignment = center
        br.cell(r, 5, note).font = font_small
        br.cell(r, 5).alignment = wrap
        for col in range(1, 6):
            br.cell(r, col).border = thin
        if status == "Reported":
            br.cell(r, 2).fill = fill_green
            br.cell(r, 3).fill = fill_green
        elif status == "Residual" or status.startswith("Residual"):
            br.cell(r, 2).fill = fill_amber
            br.cell(r, 3).fill = fill_amber
        else:
            br.cell(r, 2).fill = fill_paper
        br.row_dimensions[r].height = 32

    br["A16"] = "What is not in the grid"
    br["A16"].font = font_section
    br["A17"] = (
        "Eval-team salaries, forward-deployed engineers, and the vendors who sell eval or observability products are not in the eval column. "
        "LangChain's late-2025 survey is the right context for that omission: 89% of respondents had agent observability, 52% ran offline evals, 37% ran online evals. "
        "The harness, as a product, is ahead of evals, as a practice. This workbook only prices the tokens each of those activities burns. "
        "Human review of agent work is a real cost and it is not here. Futurum's October macro work and the coding-agent trace literature both say verification time is large. It does not show up on a model invoice."
    )
    br["A17"].alignment = wrap_top
    br["A17"].font = font_body
    br.merge_cells("A17:E17")
    br.row_dimensions[17].height = 60

    br["A19"] = "App markets, reported, not re-cut"
    br["A19"].font = font_section
    br["A20"] = "These are Menlo's application categories. They describe the product, not the buyer's industry. They are not an input to the grid. Healthcare here is vertical health apps, not every dollar a hospital spends on models."
    br["A20"].font = font_small
    br["A20"].alignment = wrap
    br.merge_cells("A20:E20")
    br.row_dimensions[20].height = 28
    for i, h in enumerate(["Category", "$ billions", "What Menlo said"], 1):
        c = br.cell(21, i, h)
        c.font = font_h_white
        c.fill = fill_teal
        c.alignment = center
        c.border = thin
    apps = [
        ("Coding", 34.0, "Up 7.7x. Claude Code and Codex are $24.0B, 70% of the category."),
        ("Enterprise chatbots and assistants", 25.0, "Up 3.0x. Enterprise Claude, ChatGPT, and Gemini are $15.5B, 62% of chatbot and productivity spend."),
        ("Customer success", 2.7, "Up 4.3x."),
        ("IT", 1.6, "Up 2.3x. Incumbents hold about 91%."),
        ("Sales", 1.4, "Up 3.6x."),
        ("Marketing", 1.2, "Up 1.8x."),
        ("HR", 0.9, "Up 2.5x."),
        ("Vertical AI, all industries", 10.1, "Up 2.9x from $3.5B."),
        ("    Healthcare vertical apps", 3.7, "37% of vertical AI. Clinical tools, ambient scribes."),
        ("    Legal vertical apps", 1.8, "Second vertical. Harvey, Legora, and plaintiff-side tools."),
    ]
    for i, (cat, val, note) in enumerate(apps):
        r = 22 + i
        br.cell(r, 1, cat).font = font_label
        br.cell(r, 2, val).number_format = FMT_B
        br.cell(r, 2).font = Font(name="Calibri", size=10, bold=True, color=GREEN_INK)
        br.cell(r, 2).fill = fill_green
        br.cell(r, 2).alignment = center
        br.cell(r, 3, note).font = font_small
        for col in range(1, 4):
            br.cell(r, col).border = thin
        br.row_dimensions[r].height = 20

    set_widths(br, {"A": 55, "B": 16, "C": 32, "D": 28, "E": 78})
    br.freeze_panes = "A5"
    br.page_setup.fitToWidth = 1
    br.oddHeader.right.text = "Bridge"

    # ------------------------------------------------------------------
    # Sensitivity
    # ------------------------------------------------------------------
    se = wb.create_sheet("Sensitivity")
    apply_header_footer(se, "Sensitivity")
    se.sheet_properties.tabColor = "A33B32"
    se["A1"] = "What moves the dollars"
    se["A1"].font = font_title
    se["A2"] = (
        "The mix does not depend on the pool. A bigger pool scales every cell by the same factor. "
        "The mix does depend on the open-weight and eval intensities, the size multipliers, and the vertical shares. "
        "Token shares also depend on the four prices. Dollar shares do not."
    )
    se["A2"].font = font_small
    se["A2"].alignment = wrap
    se.merge_cells("A2:F2")
    se.row_dimensions[2].height = 32

    se["A4"] = "Pool scenarios, same mix"
    se["A4"].font = font_section
    for i, h in enumerate(["", "Low, manual $70B", "Active pool", "Token-implied pool"], 1):
        c = se.cell(5, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    se["B6"] = 70
    se["C6"] = "=Assumptions!C12"
    se["D6"] = "=Assumptions!C11"
    se["A6"] = "Pool ($B)"
    se["B6"].font = font_input
    se["B6"].fill = fill_gold
    se["B6"].border = thin
    se["B6"].number_format = FMT_B
    se["B6"].alignment = center
    se["C6"].number_format = FMT_B
    se["D6"].number_format = FMT_B
    se["C6"].fill = fill_green
    se["D6"].fill = fill_green
    se["C6"].border = thin
    se["D6"].border = thin
    se["C6"].font = Font(name="Calibri", size=10, bold=True, color=GREEN_INK)
    se["D6"].font = Font(name="Calibri", size=10, bold=True, color=GREEN_INK)

    sens_lines = [
        (7, "Proprietary harness", "=Summary!B28/Summary!$D$30"),
        (8, "Proprietary evals", "=Summary!C28/Summary!$D$30"),
        (9, "Open-weight harness", "=Summary!B29/Summary!$D$30"),
        (10, "Open-weight evals", "=Summary!C29/Summary!$D$30"),
    ]
    for r, name, share in sens_lines:
        se.cell(r, 1, name).font = font_label
        se.cell(r, 1).border = thin
        for col, pool_cell in ((2, "$B$6"), (3, "$C$6"), (4, "$D$6")):
            se.cell(r, col, f"={share}*{pool_cell}")
            se.cell(r, col).number_format = FMT_B
            se.cell(r, col).border = thin
            se.cell(r, col).alignment = center
            se.cell(r, col).font = Font(name="Calibri", size=11, bold=True)
    se["A12"] = "The $70B case is an input. It is a haircut for the possibility that part of lab revenue ex first-party apps is not metered inference. It is not a third published total."
    se["A12"].font = font_small
    se["A12"].alignment = wrap
    se.merge_cells("A12:D12")
    se.row_dimensions[12].height = 28

    se["A14"] = "Dials, live"
    se["A14"].font = font_section
    se["A15"] = "Open-weight intensity"
    se["B15"] = "=Assumptions!C29"
    se["A16"] = "Eval intensity"
    se["B16"] = "=Assumptions!C30"
    se["A17"] = "Chinese-origin switch (0 vertical, 1 flat)"
    se["B17"] = "=Assumptions!C31"
    se["C15"] = "Set to 1.25 and open-weight's dollar share rises by about a quarter, until a segment hits the 40% cap."
    se["C16"] = "Set to 1.25 and the eval column takes 25% more of each segment. Technology still leads absolute eval dollars, because the base is large. Healthcare still leads the share."
    se["C17"] = "Set to 1 to replace the vertical origin rates with the flat rate in Assumptions!C32. The dollar 2x2 does not move. Only Open origin moves."
    for r in (15, 16, 17):
        se.cell(r, 1).font = font_label
        se.cell(r, 1).alignment = wrap
        se.cell(r, 2).font = Font(name="Calibri", size=14, bold=True, color=NAVY)
        se.cell(r, 2).fill = fill_paper
        se.cell(r, 2).border = thin
        se.cell(r, 2).alignment = center
        se.cell(r, 2).number_format = "0.00"
        se.cell(r, 3).font = font_small
        se.cell(r, 3).alignment = wrap
        se.row_dimensions[r].height = 28
    se["B17"].number_format = "0"

    se["A19"] = "What a different fact would do"
    se["A19"].font = font_section
    facts = [
        "If open-weight were 11% of model dollars, the a16z early-2026 spending share, you would raise Open-weight intensity until Summary!B12 reads 11%. Token share would jump with it unless you also raise the open-weight prices. Do that only if you believe the 11% and the 4% are the same denominator. They are not.",
        "If evals were 20% of model dollars, raise Eval intensity. The proprietary eval cell would absorb almost all of the increase. Open-weight evals stay small because open-weight is a harness book.",
        "If Menlo's 89% Chinese-origin figure is the right enterprise mix, set the switch to 1. Financial services and healthcare would then show Chinese-origin invoices they are, on Futurum's account evidence, actively restricting.",
        "If self-hosted GPU cost were added to open-weight, the open-weight dollar share would rise and would still sit far below the token share. This workbook does not make that gross-up. The Bridge residual is where that cost is most likely hiding, and it is not labeled.",
    ]
    for i, text in enumerate(facts):
        r = 20 + i
        se.cell(r, 1, text).font = font_body
        se.cell(r, 1).alignment = wrap_top
        se.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        se.row_dimensions[r].height = 36

    set_widths(se, {"A": 62, "B": 22, "C": 28, "D": 24, "E": 18, "F": 18})
    se.freeze_panes = "A4"
    se.page_setup.fitToWidth = 1
    se.oddHeader.right.text = "Sensitivity"

    # ------------------------------------------------------------------
    # Definitions, Sources, Notes, Cover
    # ------------------------------------------------------------------
    df = wb.create_sheet("Definitions")
    apply_header_footer(df, "Definitions")
    df.sheet_properties.tabColor = "5C6B73"
    df["A1"] = "Definitions"
    df["A1"].font = font_title
    df["A2"] = "The words are used in one sense in this workbook. Other senses are named so they are not mixed in."
    df["A2"].font = font_small
    defs = [
        ("Proprietary",
         "A model whose weights the buyer cannot run. The buyer pays the lab, or a reseller of that lab, per token or per committed capacity. Anthropic, OpenAI, Google's Gemini API, and xAI are in this column. Microsoft 365 Copilot seats are not. Those are applications."),
        ("Open-weight",
         "A model whose weights are published, whether or not the license is OSI-open. Llama, Gemma, Phi, Nemotron, Mistral, Qwen, DeepSeek, Kimi, and GLM are in this column when the buyer is paying for inference of those weights. The column is the invoice for that inference. It is not the GPU bill of a cluster the buyer runs alone, and it is not a developer's laptop."),
        ("Harness",
         "The production runtime around a model: system prompt, tools, memory, retrieval, routing, sandboxes, guardrails, and the tokens that runtime burns while doing a job for a user or a batch. A coding agent in a repo, a support agent on a ticket, and a nightly extraction job are harness. The software vendors who sell orchestration are not in the dollar figure. Their revenue is the residual on the Bridge."),
        ("Evals",
         "Tokens burned to decide whether the model or the agent was right. Offline suites, CI runs, online scoring of production traffic, LLM-as-judge, and red-team passes. A compiler check or a unit test that does not call a model is not in the dollar figure, which is why technology's eval share of model dollars is lower than healthcare's even though technology runs more tests. Salaries of the people who write evals are not in the figure."),
        ("Frontier",
         "Menlo's top tier for each lab in Q3 2026: GPT-6 Astra, GPT-5.6 Sol, GPT-5.5, and Anthropic's Fable and Opus class. Not every proprietary model. Mid and small closed models are proprietary and are not frontier."),
        ("Buyer vertical",
         "The industry of the enterprise that pays the invoice. A hospital's coding-agent bill is healthcare. A software company's support-agent bill is technology. This is not Menlo's app-category cut, which is reprinted on the Bridge so the two are not confused."),
        ("Employee band",
         "Headcount of the buyer. 100-999, 1,000-9,999, and 10,000+. Firms under 100 are outside the pool."),
        ("Pool",
         "Metered model-inference dollars, annualized. Default construction: Menlo's $122.7B of lab deployment, minus $40.0B of first-party apps, divided by the labs' 95% share of API dollars. The alternative construction annualizes 2.9 quadrillion open-weight tokens a month at $0.13 and divides by 4%."),
    ]
    df["A4"] = "Term"
    df["B4"] = "Meaning in this workbook"
    df["A4"].font = font_h_white
    df["B4"].font = font_h_white
    df["A4"].fill = fill_navy
    df["B4"].fill = fill_navy
    df["A4"].alignment = center
    df["B4"].alignment = center
    for i, (term, meaning) in enumerate(defs):
        r = 5 + i
        df.cell(r, 1, term).font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        df.cell(r, 1).alignment = Alignment(vertical="top", wrap_text=True)
        df.cell(r, 1).fill = fill_paper
        df.cell(r, 2, meaning).font = font_body
        df.cell(r, 2).alignment = wrap_top
        df.row_dimensions[r].height = 58
        for col in range(1, 3):
            df.cell(r, col).border = thin
    set_widths(df, {"A": 24, "B": 120})
    df.freeze_panes = "A5"
    df.page_setup.fitToWidth = 1
    df.oddHeader.right.text = "Definitions"

    src = wb.create_sheet("Sources")
    apply_header_footer(src, "Sources")
    src.sheet_properties.tabColor = "2E7D4F"
    src["A1"] = "Sources"
    src["A1"].font = font_title
    src["A2"] = "What each source is allowed to do in this workbook. A source that describes incidence is not used as a dollar weight."
    src["A2"].font = font_small
    for i, h in enumerate(["Source", "Sample and date", "Used for", "Not used for", "URL"], 1):
        c = src.cell(4, i, h)
        c.font = font_h_white
        c.fill = fill_navy
        c.alignment = center
        c.border = thin
    src.row_dimensions[4].height = 22
    sources = [
        ("Menlo Ventures, 2026: The State of Enterprise AI. Tim Tully, Joff Redfern, Deedy Das, Derek Xiao. 3 Nov 2026.",
         "Nearly 400 US enterprise decision-makers, more than $4B of AI spend in the sample. Fielded by Morning Consult, 8-26 Sep 2026. Market totals are Menlo's bottoms-up estimates from the survey, filings, and third-party reports. Annualized run-rate.",
         "Deployment $176.8B, the app and infra split, lab shares, the 4% and 19%, prices, frontier mix, app-category dollars, the Chinese-origin 89% as an alternative case.",
         "Any size or vertical cross-tab. Menlo did not publish one. The 89% is not the default enterprise origin mix, because the triangulation includes OpenRouter, Ramp, and Vercel.",
         "https://menlovc.com/wp-content/uploads/2026/11/menlo_ventures_enterprise_ai_report-2026-110326.pdf"),
        ("Andreessen Horowitz, Leaders, gainers and unexpected winners in the Enterprise AI arms race. Sarah Wang, Justin Kahl, Shangda Xu. 30 Jan 2026.",
         "100 verified VPs and C-level executives at Global 2000 companies, revenue at least $500M. 88% over $1B, 30% over $10B, more than half over 10,000 employees. Non-random expert-network sample. The authors note it may over-represent advanced adopters.",
         "The early-2026 11% open-source spending share, the $7M average LLM spend, the preference shift toward closed models, and the industry mix of the sample.",
         "The September 2026 dollar split. a16z's 11% is a wider spending concept, measured earlier. It is shown on Anchors so it is not averaged with Menlo's 4%.",
         "https://a16z.com/leaders-gainers-and-unexpected-winners-in-the-enterprise-ai-arms-race/"),
        ("Andreessen Horowitz, How 100 Enterprise CIOs Are Building and Buying Gen AI in 2025.",
         "The prior CIO survey, Global 2000.",
         "One directional fact: open-source adoption was higher at larger enterprises than at smaller ones, for on-prem, compliance, and fine-tuning reasons. Enterprises were moving from fine-tuning toward prompting and routing.",
         "A 2026 dollar share. The direction is paired with Menlo's low dollar share: large firms are more likely to have an open-weight workload, and less of their invoice is that workload.",
         "https://a16z.com/ai-enterprise-2025/"),
        ("Ramp AI Index. Data through October 2026. Token-spend share for the week ending 1 Nov 2026.",
         "More than 85,000 US businesses on Ramp's card and bill-pay platform. Transaction data, not a survey.",
         "Adoption ranks by sector and by firm size, the median and tail of AI spend per employee, and Anthropic's 49.3% of token spending that week. These shape the judgment that technology and financial services take more dollars than their adoption rates alone would say.",
         "The level of the pool, and any eval-versus-harness split. Ramp measures vendor bills, including seat products.",
         "https://ramp.com/data/ai-index"),
        ("LangChain, State of Agent Engineering. Survey 18 Nov-2 Dec 2025, n=1,340. Write-up dated 25 Feb 2026.",
         "Public survey. 63% of respondents in technology. 49% at organizations under 100 people. A builder sample, not a dollar sample.",
         "Production agents are more common at 10,000+ employees (67%) than under 100 (50%). Observability 89%, offline evals 52.4%, online evals 37.3%. Evidence that the harness-operations practice is ahead of the eval practice, and that size matters for being in production.",
         "Dollar weights. The respondent mix would badly misstate enterprise spend if it were used as one.",
         "https://www.langchain.com/state-of-agent-engineering"),
        ("Futurum Group, Enterprises Embrace Open Weights but Infrastructure Sellers Profit Most. Nick Patience. 26 Oct 2026. Document AIONP202610.",
         "September 2026 drill-down, N=200 enterprises, 83 of them in production on open-weight, 56 of those in North America. October 2026 AI Tools Pulse, N=100. AI Product Series, N=511.",
         "42% production incidence, 34% token share among production users, 51% of that use on low-complexity work, the 5% of metered consumption, the family incidence table, origin screens, and the hosting mix. Also the Vercel and OpenRouter contrasts Futurum reports.",
         "A full size-by-vertical dollar grid. Incidence is not spend. The 5% has a wider denominator than Menlo's 4% and excludes a large part of self-hosted cost.",
         "https://futurumgroup.com/wp-content/uploads/2026/10/Enterprises-Embrace-Open-Weights-but-Infrastructure-Sellers-Profit-Most_AIONP202610.pdf"),
        ("Vercel AI Gateway and OpenRouter, as cited by Futurum, September 2026.",
         "Developer and startup gateways, not enterprise procurement.",
         "Contrast only. Vercel: open-weight 56% of gateway tokens and 14% of gateway spend in August 2026, up from 13% of tokens in April. OpenRouter: open-weight 72% of tokens in mid-September, Chinese labs 69% of OpenRouter tokens. Menlo separately reports agents at 80% of OpenRouter tokens.",
         "Enterprise shares. The gap between these gateways and the enterprise 4% is the point of the contrast.",
         "https://futurumgroup.com/wp-content/uploads/2026/10/Enterprises-Embrace-Open-Weights-but-Infrastructure-Sellers-Profit-Most_AIONP202610.pdf"),
    ]
    for i, (name, sample, used, not_used, url) in enumerate(sources):
        r = 5 + i
        src.cell(r, 1, name).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        src.cell(r, 2, sample).font = font_label
        src.cell(r, 3, used).font = font_label
        src.cell(r, 4, not_used).font = font_label
        src.cell(r, 5, url).font = font_link
        src.cell(r, 5).hyperlink = url
        for col in range(1, 6):
            src.cell(r, col).alignment = wrap_top
            src.cell(r, col).border = thin
        src.row_dimensions[r].height = 78
    src["A13"] = (
        "Not used, on purpose. Syndicated market pages that quote a single 'evaluation harness' market size, and vendor blogs that attribute an 18-24% eval budget to Gartner or McKinsey without a retrievable table, are not inputs. "
        "A 2025 McKinsey incidence survey on open-source AI (technology highest, consumer and public sector lower) points the same way as the sources above and is a year older, with a wider definition, so it is not loaded into a cell."
    )
    src["A13"].font = font_small
    src["A13"].alignment = wrap
    src.merge_cells("A13:E13")
    src.row_dimensions[13].height = 36
    set_widths(src, {"A": 55, "B": 55, "C": 55, "D": 55, "E": 55})
    src.freeze_panes = "A5"
    src.page_setup.fitToWidth = 1
    src.page_setup.orientation = "landscape"
    src.oddHeader.right.text = "Sources"

    nt = wb.create_sheet("Notes")
    apply_header_footer(nt, "Notes")
    nt.sheet_properties.tabColor = NAVY
    nt["A1"] = "Method and limits"
    nt["A1"].font = font_title
    notes = [
        ("What the grid is",
         "No published survey crosses proprietary versus open-weight, evals versus harness, employee band, and buyer vertical in one table. "
         "The workbook refuses to pretend one did. Anchors holds the published figures. Assumptions holds the judgments, in gold. "
         "Engine multiplies them. Summary, Size by vertical, and Tokens only sum the engine."),
        ("The identity the base case is built to respect",
         "Open-weight can be 4% of dollars and 19% of tokens at the same time only because an open-weight token costs about a fifth of a closed token. "
         "Menlo's own prices, $0.13 and $0.70, imply an open-weight token share of about 18% if the dollar share is 4%, which is why a gap of a point or so against the stated 19% is the price identity and not a second fact. "
         "Summary shows the gap rather than forcing the output onto 4.0% and 19.0%."),
        ("Why evals are a larger share of dollars than of tokens",
         "Production agent traffic is cache-heavy. A June 2026 trace of GitHub Copilot's coding agent, and a separate trace of production coding-agent sessions, both find cached prompt tokens dominating the bill. "
         "An eval run is a fresh prompt, often judged by a frontier model, and it does not get that cache. "
         "So a small token share becomes a larger dollar share. The four prices are the mechanism. Change them and the token rows move. The dollar rows do not."),
        ("Why open-weight evals are almost empty",
         "Futurum's production users put 51% of their open-weight use on low-complexity work, up from 39%. That is harness volume. "
         "Judges in regulated verticals stay on proprietary models. The input 'evals as a share of open-weight dollars' is 3% in financial services and healthcare and 10-12% in technology, retail, and manufacturing. "
         "Those are judgments. They are the reason the open-eval cell is the smallest in the 2x2."),
        ("Why the Chinese-origin default is not 89%",
         "Menlo's 89% is a spend share triangulated with gateways where Chinese labs lead. "
         "Futurum's enterprise production accounts, measured the same quarter, show Gemma, Phi, Llama, and Nemotron as the families that show up, and show DeepSeek, Qwen, GLM, and Kimi flat or down. "
         "Origin is a major or deciding factor for 47% of those production users. "
         "The default therefore varies by vertical, highest in technology and retail, lowest in the public sector and in banks. "
         "The switch on Assumptions replaces that with a flat rate, defaulting to 89%, so both pictures are one click apart."),
        ("What would make the allocation wrong",
         "The vertical dollar weights are the softest input. If financial services or healthcare turned out to buy model inference at technology's intensity, the open-weight share of the whole book would fall, because those verticals are set to low open-weight shares, and the eval share of the book would rise. "
         "The size story would be wrong if the 10,000+ band were in fact the deepest open-weight buyer per dollar, rather than the widest. a16z's 2025 finding (higher adoption at larger enterprises) is incidence. This model pairs it with shallower dollar share. That pairing is a judgment."),
        ("Perimeter",
         "US enterprise, 100 employees and up, metered model inference, annualized to September 2026. "
         "Not consumer ChatGPT. Not firms under 100. Not the fully loaded cost of a self-hosted cluster. Not eval-team payroll. Not the $102.5B Menlo counts as spending to build models, which includes $15.3B of training data and evaluation services on the development side of the market. "
         "That $15.3B is a vendor market (expert data, RL environments). It is not the enterprise inference bill for running evals, and it is not added to the eval column."),
        ("As of",
         "Figures were taken from sources published through 3 November 2026 (Menlo) and 26 October 2026 (Futurum), and from Ramp's index as published for October 2026 and the week ending 1 November 2026. The workbook date is 17 November 2026."),
    ]
    r = 3
    for title, body in notes:
        nt.cell(r, 1, title).font = font_section
        nt.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        r += 1
        nt.cell(r, 1, body).font = font_body
        nt.cell(r, 1).alignment = wrap_top
        nt.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        nt.row_dimensions[r].height = 72
        r += 1
    set_widths(nt, {"A": 28, "B": 120})
    nt.column_dimensions["A"].width = 130
    nt.freeze_panes = "A3"
    nt.page_setup.fitToWidth = 1
    nt.oddHeader.right.text = "Notes"

    # Cover
    cv = wb.create_sheet("Cover", 0)
    apply_header_footer(cv, "Cover")
    cv.sheet_properties.tabColor = NAVY
    cv["A1"] = "Proprietary vs open-weight model spending"
    cv["A1"].font = Font(name="Calibri", size=24, bold=True, color=NAVY)
    cv["A2"] = "Evals vs harness, by enterprise size and buyer vertical"
    cv["A2"].font = Font(name="Calibri", size=16, color=TEAL)
    cv["A3"] = "United States enterprises  ·  annualized model inference  ·  September 2026  ·  prepared 17 November 2026"
    cv["A3"].font = font_small
    cv.merge_cells("A1:F1")
    cv.merge_cells("A2:F2")
    cv.merge_cells("A3:F3")

    cv["A5"] = "Result"
    cv["A5"].font = font_section
    cv["A6"] = "=Summary!A39"
    cv["A6"].font = font_body
    cv["A6"].alignment = wrap_top
    cv.merge_cells("A6:F6")
    cv.row_dimensions[6].height = 60

    cv["A8"] = "Active pool ($B)"
    cv["B8"] = "=Summary!B5"
    cv["C8"] = "Open-weight share of $"
    cv["D8"] = "=Summary!B12"
    cv["E8"] = "Eval share of $"
    cv["F8"] = "=Summary!B19"
    cv["A9"] = "Proprietary ($B)"
    cv["B9"] = "=Summary!B6"
    cv["C9"] = "Open-weight share of tokens"
    cv["D9"] = "=Summary!B15"
    cv["E9"] = "Eval share of tokens"
    cv["F9"] = "=Summary!B20"
    for addr in ("B8", "B9"):
        cv[addr].number_format = FMT_B
        cv[addr].font = font_kpi
        cv[addr].fill = fill_paper
        cv[addr].alignment = center
        cv[addr].border = thin
    for addr in ("D8", "F8", "D9", "F9"):
        cv[addr].number_format = FMT_PCT
        cv[addr].font = font_kpi
        cv[addr].fill = fill_paper
        cv[addr].alignment = center
        cv[addr].border = thin
    for addr in ("A8", "C8", "E8", "A9", "C9", "E9"):
        cv[addr].font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        cv[addr].alignment = wrap

    cv["A11"] = "How the file is organized"
    cv["A11"].font = font_section
    toc = [
        ("Cover", "This page. The six figures are live. They move when Assumptions changes."),
        ("Summary", "The 2x2 in dollars and in tokens, the check against Menlo's 4% and 19%, then the cut by size and by vertical."),
        ("Size by vertical", "The full cross. Eight buyer verticals, three employee bands, proprietary and open-weight, harness and evals."),
        ("Tokens", "The same cross in tokens. This is where the price gap shows up."),
        ("Open origin", "Chinese-origin versus other open-weight dollars, plus Futurum's production-incidence table, which is not a dollar table."),
        ("Bridge", "How the pool sits inside Menlo's $176.8B, and the app-category dollars that are not re-cut here."),
        ("Assumptions", "The only sheet to edit. Gold cells, blue type."),
        ("Engine", "Twenty-four rows, one per vertical and size. All formulas."),
        ("Anchors", "Published figures only. Green cells."),
        ("Sensitivity", "The pool scales dollars and does not change the mix. The intensity dials change the mix."),
        ("Definitions", "Proprietary, open-weight, harness, evals, frontier, vertical, band, pool."),
        ("Sources", "What each source is allowed to support."),
        ("Notes", "Method, the price identity, and the judgments most likely to be wrong."),
    ]
    cv["A12"] = "Sheet"
    cv["B12"] = "What is on it"
    cv["A12"].font = font_h_white
    cv["B12"].font = font_h_white
    cv["A12"].fill = fill_navy
    cv["B12"].fill = fill_navy
    cv.merge_cells("B12:F12")
    for i, (sheet, desc) in enumerate(toc):
        r = 13 + i
        cv.cell(r, 1, sheet).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        cv.cell(r, 2, desc).font = font_label
        cv.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        cv.cell(r, 2).alignment = wrap
        cv.row_dimensions[r].height = 20
        for col in range(1, 7):
            cv.cell(r, col).border = thin
            if i % 2:
                cv.cell(r, col).fill = fill_zebra

    cv["A27"] = (
        "Color. Green is a published figure. Gold with blue type is an input. Amber is a residual or a gap. "
        "Navy totals are calculated. Nothing in the size-by-vertical grid is a number a survey printed."
    )
    cv["A27"].font = font_body
    cv["A27"].alignment = wrap
    cv.merge_cells("A27:F27")
    cv.row_dimensions[27].height = 32

    set_widths(cv, {"A": 42, "B": 22, "C": 36, "D": 18, "E": 28, "F": 16})
    cv.freeze_panes = "A5"
    cv.page_setup.orientation = "landscape"
    cv.page_setup.fitToWidth = 1
    cv.page_setup.fitToHeight = 1
    cv.oddHeader.right.text = "Cover"
    cv.row_dimensions[1].height = 32
    cv.print_title_rows = "1:3"
    cv.sheet_view.zoomScale = 120

    # print settings sanity: freeze on cover should not hide the title awkwardly
    cv.freeze_panes = "A5"
    cv.page_setup.paperSize = cv.PAPERSIZE_TABLOID

    # named ranges for the pool, so a reader can find it
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names.add(DefinedName(name="ActivePool", attr_text="Assumptions!$C$12"))
    wb.defined_names.add(DefinedName(name="OpenIntensity", attr_text="Assumptions!$C$29"))
    wb.defined_names.add(DefinedName(name="EvalIntensity", attr_text="Assumptions!$C$30"))

    # sheet order already mostly set. Confirm.
    order = [
        "Cover",
        "Summary",
        "Size by vertical",
        "Tokens",
        "Open origin",
        "Bridge",
        "Assumptions",
        "Engine",
        "Anchors",
        "Sensitivity",
        "Definitions",
        "Sources",
        "Notes",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.properties.title = "Proprietary vs open-weight model spend: evals vs harness"
    wb.properties.subject = "US enterprise model inference by size and vertical, September 2026"
    wb.properties.creator = "Equity research workbook"
    wb.properties.description = (
        "Formula-driven allocation of enterprise model-inference spend across proprietary vs open-weight "
        "and evals vs harness, by employee band and buyer vertical. Reported anchors are separated from inputs."
    )

    wb.save(OUT)
    print(OUT)


def openpyxl_grid():
    from openpyxl.chart.axis import ChartLines
    # default major gridlines object; returning None keeps it clean, but a light grid helps.
    return None


if __name__ == "__main__":
    build()
