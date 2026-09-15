#!/usr/bin/env python3
"""Build an Excel workbook with time series on agentic AI:

  1. Agentic token consumption (monthly, Jan 2024 - Dec 2026)
  2. Number and type of AI agents deployed (annual, 2025 - 2030)
  3. Token loads (per-task benchmarks + derived annual loads)
  4. Actions per agent (per-task benchmarks + derived annual actions)

All disclosed anchor figures are cited on the "Sources & Notes" sheet.
Values between/beyond anchors are interpolated or extrapolated and are
explicitly labeled "Modeled" in the workbook.

Output: data/Agentic_AI_Time_Series.xlsx
"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT_PATH = Path(__file__).resolve().parent.parent / "data" / "Agentic_AI_Time_Series.xlsx"

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

NAVY = "1F3864"
BLUE = "2E74B5"
LIGHT = "DEEBF7"
GREY = "F2F2F2"

TITLE_FONT = Font(name="Calibri", size=14, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=BLUE)
SECTION_FONT = Font(name="Calibri", size=11, bold=True, color=NAVY)
BODY_FONT = Font(name="Calibri", size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header_row(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_table(ws, top_row, headers, rows, col_widths=None, num_formats=None, banded=True):
    """Write a simple table; returns the row after the table."""
    for j, h in enumerate(headers, start=1):
        ws.cell(row=top_row, column=j, value=h)
    style_header_row(ws, top_row, 1, len(headers))
    for i, r in enumerate(rows):
        for j, v in enumerate(r, start=1):
            cell = ws.cell(row=top_row + 1 + i, column=j, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            if banded and i % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=GREY)
            if num_formats and j in num_formats and isinstance(v, (int, float)):
                cell.number_format = num_formats[j]
    if col_widths:
        for j, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return top_row + 1 + len(rows)


def add_title(ws, title, subtitle):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUBTITLE_FONT


# ---------------------------------------------------------------------------
# Monthly token-consumption model (Jan 2024 - Dec 2026)
# ---------------------------------------------------------------------------

def month_index(y, m):
    return y * 12 + (m - 1)


def month_range(start=(2024, 1), end=(2026, 12)):
    out = []
    y, m = start
    while (y, m) <= end:
        out.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def geometric_series(anchors, months, back_rate=None, fwd_rate=None):
    """Geometric interpolation between (year, month, value) anchors.

    back_rate / fwd_rate: monthly growth multipliers used to extrapolate
    before the first / after the last anchor (None -> leave blank).
    Returns {(y, m): value}.
    """
    pts = sorted((month_index(y, m), v) for y, m, v in anchors)
    out = {}
    for (y, m) in months:
        t = month_index(y, m)
        if t < pts[0][0]:
            if back_rate:
                out[(y, m)] = pts[0][1] / back_rate ** (pts[0][0] - t)
            continue
        if t > pts[-1][0]:
            if fwd_rate:
                out[(y, m)] = pts[-1][1] * fwd_rate ** (t - pts[-1][0])
            continue
        for (t0, v0), (t1, v1) in zip(pts, pts[1:]):
            if t0 <= t <= t1:
                if t == t0:
                    out[(y, m)] = v0
                elif t == t1:
                    out[(y, m)] = v1
                else:
                    out[(y, m)] = v0 * (v1 / v0) ** ((t - t0) / (t1 - t0))
                break
    return out


def linear_series(anchors, months):
    pts = sorted((month_index(y, m), v) for y, m, v in anchors)
    out = {}
    for (y, m) in months:
        t = month_index(y, m)
        if t <= pts[0][0]:
            out[(y, m)] = pts[0][1]
            continue
        if t >= pts[-1][0]:
            out[(y, m)] = pts[-1][1]
            continue
        for (t0, v0), (t1, v1) in zip(pts, pts[1:]):
            if t0 <= t <= t1:
                out[(y, m)] = v0 + (v1 - v0) * (t - t0) / (t1 - t0)
                break
    return out


MONTHS = month_range()

# Disclosed anchors, trillions of tokens per month.
GOOGLE_ANCHORS = [
    (2024, 5, 9.7),      # Google I/O 2025 recap of May 2024
    (2025, 5, 480.0),    # Google I/O 2025
    (2025, 7, 980.0),    # Alphabet Q2-2025 earnings call
    (2025, 10, 1300.0),  # Google blog, Oct 2025
    (2026, 5, 3200.0),   # Google I/O 2026
]
# Back: May24->May25 implied monthly growth. Fwd: Gemini API grew 19 -> 22 B tok/min
# over ~2 months (Q2 2026 earnings) => ~7.6%/month, more conservative than trailing YoY.
GOOGLE_BACK = (480.0 / 9.7) ** (1 / 12)
GOOGLE_FWD = (22 / 19) ** (1 / 2)

MSFT_ANCHORS = [
    (2024, 2, 6.7),    # implied: FY25Q3 was "up 5x YoY" => FY24Q3 ~20T/qtr
    (2025, 2, 33.3),   # FY25Q3 (Jan-Mar 2025): >100T tokens/quarter
    (2025, 3, 50.0),   # "a record 50 trillion tokens last month alone"
]
MSFT_FWD = 5 ** (1 / 12)  # hold the disclosed 5x YoY pace (modeled)

OPENROUTER_ANCHORS = [
    (2024, 6, 0.83),   # ~10T tokens/year run rate (Menlo Ventures / a16z)
    (2025, 6, 8.3),    # ~100T tokens/year run rate
    (2025, 12, 30.0),  # >1T tokens/day (a16z State of AI, Contrary)
    (2026, 8, 100.0),  # >100T tokens/month; >1 quadrillion/yr annualized (Contrary)
]
OPENROUTER_BACK = 10 ** (1 / 12)
OPENROUTER_FWD = (100.0 / 30.0) ** (1 / 8)

# Agentic share of OpenRouter token volume (%). Modeled from qualitative anchors:
# negligible in 2024; programming/agentic inference ramp through 2025; agentic tokens
# crossed 50% ~Feb 1, 2026 (OpenRouter blog); "Agent"-classified share still rising
# mid-2026.
AGENTIC_SHARE_ANCHORS = [
    (2024, 1, 1.0),
    (2025, 1, 5.0),
    (2025, 6, 12.0),
    (2025, 12, 30.0),
    (2026, 2, 50.0),
    (2026, 6, 58.0),
    (2026, 12, 65.0),
]

google_m = geometric_series(GOOGLE_ANCHORS, MONTHS, GOOGLE_BACK, GOOGLE_FWD)
msft_m = geometric_series(MSFT_ANCHORS, MONTHS, MSFT_FWD, MSFT_FWD)
openrouter_m = geometric_series(OPENROUTER_ANCHORS, MONTHS, OPENROUTER_BACK, OPENROUTER_FWD)
agentic_share_m = linear_series(AGENTIC_SHARE_ANCHORS, MONTHS)

DISCLOSED_GOOGLE = {(y, m) for y, m, _ in GOOGLE_ANCHORS}
DISCLOSED_MSFT = {(y, m) for y, m, _ in MSFT_ANCHORS[1:]}  # first anchor is implied
DISCLOSED_OPENROUTER = {(y, m) for y, m, _ in OPENROUTER_ANCHORS}

# ---------------------------------------------------------------------------
# Annual agent model (2025 - 2030), calibrated to IDC Agent Economics anchors:
#   2025 ~30M deployed agents (1.2B by 2029 is "40x more than in 2025")
#   2029 1.2B agents, 217B daily actions (=> ~181 actions/agent/day),
#        ~3.7 quadrillion tokens/day (=> ~17K tokens/action)
# ---------------------------------------------------------------------------

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
AGENTS_M = [30, 95, 280, 650, 1200, 1900]            # millions; 2025/2029 = IDC anchors
ACTIONS_PER_AGENT_DAY = [50, 65, 90, 130, 181, 210]  # 2029 = 217B / 1.2B (IDC)
TOKENS_PER_ACTION_K = [8, 10, 12, 15, 17, 19]        # thousands; 2029 calibrated to IDC

# Agent-type mix (% of deployed agents). 2025 skews to in-application agents;
# IDC expects custom-configured and standalone agents to dominate by 2030.
TYPE_NAMES = [
    "In-application agents",
    "Custom-configured (no/low-code) agents",
    "Standalone agents",
    "Bespoke / custom-built agents",
]
TYPE_SHARE_2025 = [55, 20, 15, 10]
TYPE_SHARE_2030 = [25, 35, 30, 10]


def type_share(year, i):
    f = (year - 2025) / 5
    return TYPE_SHARE_2025[i] + (TYPE_SHARE_2030[i] - TYPE_SHARE_2025[i]) * f


DAILY_ACTIONS_B = [a * apd / 1000 for a, apd in zip(AGENTS_M, ACTIONS_PER_AGENT_DAY)]
TOKENS_PER_DAY_T = [da * tpa for da, tpa in zip(DAILY_ACTIONS_B, TOKENS_PER_ACTION_K)]
TOKENS_PER_MONTH_T = [t * 30.4 for t in TOKENS_PER_DAY_T]
TOKENS_PER_AGENT_DAY_M = [apd * tpa / 1000 for apd, tpa in zip(ACTIONS_PER_AGENT_DAY, TOKENS_PER_ACTION_K)]

# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

wb = Workbook()

# --------------------------- Sheet 1: Overview -----------------------------
ws = wb.active
ws.title = "Overview"
add_title(ws, "Agentic AI — Time Series Databook",
          f"Compiled {date.today().isoformat()} from public disclosures and industry research. "
          "All modeled values are labeled; see Sources & Notes.")

overview_rows = [
    ("Token Consumption", "Monthly time series (Jan 2024 - Dec 2026) of disclosed platform token volumes "
     "(Google, Microsoft Foundry, OpenRouter) plus a modeled agentic share and estimated agentic token consumption."),
    ("Agents - Number & Type", "Annual series (2025-2030) of deployed AI agents by type, calibrated to IDC's "
     "Agent Economics model (~30M in 2025 to 1.2B in 2029), plus enterprise-penetration context from Gartner."),
    ("Token Loads", "Per-task token-load benchmarks by workload archetype (chat vs RAG vs agentic coding), "
     "agentic-vs-chat multipliers, and a derived annual series of agentic token loads."),
    ("Actions per Agent", "Per-task turns / model calls / tool calls by archetype and a derived annual series of "
     "total daily agent actions and actions per agent per day (IDC anchor: 217B daily actions by 2029)."),
    ("Sources & Notes", "Every disclosed anchor figure with source, date and URL, plus methodology caveats."),
]
r = write_table(
    ws, 4, ["Sheet", "Contents"], overview_rows,
    col_widths=[28, 120],
)
notes = [
    "How to read this workbook:",
    "- 'Disclosed' figures come directly from company statements, earnings calls, or published research.",
    "- 'Modeled' figures are geometric/linear interpolations between disclosed anchors, or extrapolations at "
    "disclosed growth rates. They are estimates, not reported data.",
    "- Units: T = trillion tokens; Q = quadrillion; B = billion; M = million.",
    "- Scopes differ across providers (Google counts all surfaces incl. consumer; Microsoft counts Foundry API only; "
    "OpenRouter counts routed third-party traffic; IDC counts enterprise agents). Series are complementary views, "
    "not additive to a clean industry total.",
]
row = r + 2
for n in notes:
    ws.cell(row=row, column=1, value=n).font = BODY_FONT
    row += 1
ws.sheet_view.showGridLines = False

# ---------------------- Sheet 2: Token Consumption -------------------------
ws = wb.create_sheet("Token Consumption")
add_title(ws, "Agentic Token Consumption — Monthly Time Series (Jan 2024 – Dec 2026)",
          "Trillions of tokens per month. Disclosed anchors marked 'Disclosed'; all other values are modeled "
          "(geometric interpolation / extrapolation). Agentic share modeled from OpenRouter disclosures.")

headers = [
    "Month",
    "Google — all surfaces (T tokens/mo)",
    "Microsoft Foundry (T tokens/mo)",
    "OpenRouter (T tokens/mo)",
    "Observable total (T tokens/mo)",
    "Agentic share of tokens (%)",
    "Est. agentic tokens — OpenRouter (T/mo)",
    "Est. agentic tokens — observable total (T/mo)",
    "Data status",
]
rows = []
for (y, m) in MONTHS:
    g = google_m.get((y, m))
    ms = msft_m.get((y, m))
    orr = openrouter_m.get((y, m))
    share = agentic_share_m.get((y, m))
    total = sum(v for v in (g, ms, orr) if v is not None)
    disclosed = []
    if (y, m) in DISCLOSED_GOOGLE:
        disclosed.append("Google")
    if (y, m) in DISCLOSED_MSFT:
        disclosed.append("Microsoft")
    if (y, m) in DISCLOSED_OPENROUTER:
        disclosed.append("OpenRouter")
    status = "Disclosed: " + ", ".join(disclosed) if disclosed else "Modeled"
    rows.append((
        f"{y}-{m:02d}",
        round(g, 2) if g else None,
        round(ms, 2) if ms else None,
        round(orr, 2) if orr else None,
        round(total, 2),
        round(share, 1),
        round(orr * share / 100, 2) if orr else None,
        round(total * share / 100, 2),
        status,
    ))

TABLE_TOP = 4
end = write_table(
    ws, TABLE_TOP, headers, rows,
    col_widths=[10, 18, 16, 15, 16, 14, 18, 20, 24],
    num_formats={2: "#,##0.00", 3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.00",
                 6: "0.0", 7: "#,##0.00", 8: "#,##0.00"},
)
ws.freeze_panes = "B5"

n = len(rows)
chart = LineChart()
chart.title = "Platform token volume (T tokens/month, log scale)"
chart.style = 12
chart.y_axis.title = "T tokens / month"
chart.y_axis.scaling.logBase = 10
chart.x_axis.title = "Month"
chart.height = 9
chart.width = 22
data = Reference(ws, min_col=2, max_col=4, min_row=TABLE_TOP, max_row=TABLE_TOP + n)
cats = Reference(ws, min_col=1, min_row=TABLE_TOP + 1, max_row=TABLE_TOP + n)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "K4")

chart2 = LineChart()
chart2.title = "Estimated agentic token consumption (T tokens/month)"
chart2.style = 12
chart2.y_axis.title = "T tokens / month"
chart2.x_axis.title = "Month"
chart2.height = 9
chart2.width = 22
data2 = Reference(ws, min_col=8, max_col=8, min_row=TABLE_TOP, max_row=TABLE_TOP + n)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "K23")

anchor_rows = [
    ("2024-05", "Google", "9.7T tokens/month across all surfaces", "Google I/O 2026 keynote (retrospective)"),
    ("2025-03", "Microsoft", "50T tokens in March; >100T in FY25 Q3 (5x YoY)", "Microsoft FY25 Q3 earnings call"),
    ("2025-05", "Google", "~480T tokens/month", "Google I/O 2025"),
    ("2025-06", "OpenRouter", "~100T tokens/year run rate (10x YoY)", "Menlo Ventures / a16z"),
    ("2025-07", "Google", "980T tokens/month", "Alphabet Q2-2025 earnings call"),
    ("2025-10", "Google", "1.3Q tokens/month", "Google blog, Oct 2025"),
    ("2025-10", "OpenAI", "API averaged ~8.6T tokens/day (~260T/mo)", "a16z State of AI (Oct 2025)"),
    ("2025-12", "OpenRouter", ">1T tokens/day (~300T/yr annualized)", "a16z / Contrary Research"),
    ("2026-02", "OpenRouter", "Agentic tokens surpassed human-driven tokens (~Feb 1)", "OpenRouter blog"),
    ("2026-05", "Google", ">3.2Q tokens/month; Gemini API ~19B tokens/min", "Google I/O 2026"),
    ("2026-07", "Google", "Gemini API ~22B tokens/min", "Alphabet Q2-2026 earnings release"),
    ("2026-08", "OpenRouter", ">100T tokens/month (>1Q/yr annualized)", "Contrary Research"),
]
r2 = end + 2
ws.cell(row=r2, column=1, value="Disclosed anchor data points").font = SECTION_FONT
write_table(ws, r2 + 1, ["Date", "Provider / scope", "Disclosed figure", "Source"], anchor_rows)
ws.sheet_view.showGridLines = False

# -------------------- Sheet 3: Agents - Number & Type ----------------------
ws = wb.create_sheet("Agents - Number & Type")
add_title(ws, "Number and Type of Deployed AI Agents — Annual Series (2025 – 2030)",
          "Millions of actively deployed agents worldwide. Totals for 2025 and 2029 are IDC Agent Economics "
          "anchors; other totals and the type mix are modeled.")

headers = ["Year", "Total agents (M)"] + [f"{t} (M)" for t in TYPE_NAMES] + \
          [f"{t} share (%)" for t in TYPE_NAMES] + ["Data status"]
rows = []
for i, y in enumerate(YEARS):
    total = AGENTS_M[i]
    shares = [type_share(y, k) for k in range(4)]
    counts = [total * s / 100 for s in shares]
    status = "IDC anchor (total); mix modeled" if y in (2025, 2029) else "Modeled"
    rows.append([y, total] + [round(c, 1) for c in counts] + [round(s, 1) for s in shares] + [status])

end = write_table(
    ws, 4, headers, rows,
    col_widths=[8, 14] + [20] * 4 + [16] * 4 + [26],
    num_formats={2: "#,##0", 3: "#,##0.0", 4: "#,##0.0", 5: "#,##0.0", 6: "#,##0.0"},
)

chart = BarChart()
chart.type = "col"
chart.grouping = "stacked"
chart.overlap = 100
chart.title = "Deployed AI agents by type (millions)"
chart.style = 12
chart.y_axis.title = "Agents (M)"
chart.height = 9
chart.width = 18
data = Reference(ws, min_col=3, max_col=6, min_row=4, max_row=4 + len(YEARS))
cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(YEARS))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "A14")

context_rows = [
    ("Enterprise apps with task-specific agents (%)", "<5% (2025) -> 40% (end-2026)",
     "Gartner, Aug 2025", "Disclosed forecast"),
    ("Enterprise apps enhanced by agentic automation (%)", ">40% by 2027",
     "IDC FutureScape 2026", "Disclosed forecast"),
    ("Agents per organization in production", "~40 -> 50",
     "IDC, 2026", "Disclosed survey"),
    ("G2000 agent use by 2027", "10x increase; token & API call loads up ~1,000x",
     "IDC FutureScape 2026", "Disclosed forecast"),
    ("Enterprises using gen-AI agents in production (late 2025)", "11% in production; 38% piloting",
     "Deloitte, late 2025", "Disclosed survey"),
    ("Multi-agent implementations", "1/3 of agentic AI implementations combine multiple agents by 2027",
     "Gartner, Aug 2025", "Disclosed forecast"),
    ("Agent delivery cost by 2029", ">$68B worldwide",
     "IDC Agent Economics model, Dec 2025", "Disclosed forecast"),
]
r2 = end + 16
ws.cell(row=r2, column=1, value="Adoption context (disclosed forecasts & surveys)").font = SECTION_FONT
write_table(ws, r2 + 1, ["Metric", "Value / trajectory", "Source", "Status"], context_rows,
            col_widths=None)
ws.sheet_view.showGridLines = False

# -------------------------- Sheet 4: Token Loads ---------------------------
ws = wb.create_sheet("Token Loads")
add_title(ws, "Token Loads — Per-Task Benchmarks and Derived Annual Series",
          "Part A: tokens consumed per task by workload archetype (published benchmarks). "
          "Part B: agentic-vs-chat multipliers. Part C: derived annual agentic token loads (IDC-calibrated).")

ws.cell(row=4, column=1, value="A. Tokens per task by workload archetype").font = SECTION_FONT
archetype_rows = [
    ("Inline code completion", "1", 550, 2500, "Tokenade benchmark compilation"),
    ("Simple chat / RAG turn", "1", 2000, 10000, "Splunk tokenomics analysis"),
    ("Voice / contact-center agent", "3-8", 5000, 15000, "Splunk tokenomics analysis"),
    ("Tool-using ReAct agent", "10-30", 20000, 60000, "Splunk tokenomics analysis"),
    ("Research agent (GAIA-class)", "10-30", 30000, 100000, "Splunk tokenomics analysis"),
    ("Claude Code — SWE-bench task", "10-30", 33000, 33000, "Splunk tokenomics analysis"),
    ("Cursor — SWE-bench task", "20-50", 188000, 188000, "Splunk tokenomics analysis"),
    ("Agentic coding session (SWE-bench class, incl. retries)", "20-50+", 1000000, 3500000,
     "Zylos Research / Bai et al. (arXiv:2604.22750)"),
]
end = write_table(
    ws, 5,
    ["Workload archetype", "Typical turns", "Tokens per task (low)", "Tokens per task (high)", "Source"],
    archetype_rows,
    col_widths=[46, 12, 20, 20, 42],
    num_formats={3: "#,##0", 4: "#,##0"},
)

chart = BarChart()
chart.type = "bar"
chart.title = "Tokens per task by archetype (log scale)"
chart.style = 12
chart.x_axis.title = "Tokens per task"
chart.x_axis.scaling.logBase = 10
chart.height = 10
chart.width = 20
data = Reference(ws, min_col=3, max_col=4, min_row=5, max_row=5 + len(archetype_rows))
cats = Reference(ws, min_col=1, min_row=6, max_row=5 + len(archetype_rows))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "G5")

r2 = end + 2
ws.cell(row=r2, column=1, value="B. Agentic vs chat token multipliers (disclosed)").font = SECTION_FONT
multiplier_rows = [
    ("Agentic workflow vs standard chatbot, per equivalent task", "5x - 30x", "Gartner (Mar 2026 analysis)"),
    ("Agentic request vs human-driven request, per request", "~15x", "OpenRouter platform data (2026)"),
    ("Agentic coding vs code chat, SWE-bench Verified", "~1,000x", "Bai et al., 8 frontier LLMs (2026)"),
    ("Run-to-run variance on identical agentic task", "up to 30x", "Bai et al. (2026)"),
    ("Model calls per agentic task", "10-20 calls", "Gartner via Cockroach Labs (2026)"),
    ("Inference cost per agentic workflow, 2026-2028", "+5x (while token prices fall ~95% by 2030)",
     "Gartner 'inference paradox' (2025)"),
    ("OpenAI top customer token consumption", "100K tokens/mo (2019) -> ~100B tokens/mo (2026), ~1,000,000x",
     "Sam Altman via Cockroach Labs (2026)"),
]
end = write_table(ws, r2 + 1, ["Metric", "Multiplier / value", "Source"], multiplier_rows)

r3 = end + 2
ws.cell(row=r3, column=1,
        value="C. Derived annual agentic token loads (modeled; calibrated to IDC 2029 anchors)").font = SECTION_FONT
derived_rows = []
for i, y in enumerate(YEARS):
    status = "IDC-calibrated" if y == 2029 else "Modeled"
    derived_rows.append((
        y,
        round(TOKENS_PER_ACTION_K[i], 1),
        round(TOKENS_PER_AGENT_DAY_M[i], 2),
        round(TOKENS_PER_DAY_T[i], 1),
        round(TOKENS_PER_MONTH_T[i], 0),
        status,
    ))
end = write_table(
    ws, r3 + 1,
    ["Year", "Tokens per action (K)", "Tokens per agent per day (M)",
     "Agentic tokens per day (T)", "Agentic tokens per month (T)", "Data status"],
    derived_rows,
    num_formats={2: "0.0", 3: "0.00", 4: "#,##0.0", 5: "#,##0"},
)

chart2 = LineChart()
chart2.title = "Derived agentic token load (T tokens/day, log scale)"
chart2.style = 12
chart2.y_axis.title = "T tokens / day"
chart2.y_axis.scaling.logBase = 10
chart2.height = 8
chart2.width = 16
data2 = Reference(ws, min_col=4, max_col=4, min_row=r3 + 1, max_row=r3 + 1 + len(YEARS))
cats2 = Reference(ws, min_col=1, min_row=r3 + 2, max_row=r3 + 1 + len(YEARS))
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
ws.add_chart(chart2, "H" + str(r3 + 1))
ws.sheet_view.showGridLines = False

# ------------------------ Sheet 5: Actions per Agent -----------------------
ws = wb.create_sheet("Actions per Agent")
add_title(ws, "Actions per Agent — Per-Task Benchmarks and Derived Annual Series",
          "Part A: turns / model calls per task by archetype. Part B: derived annual series of total daily "
          "actions and actions per agent per day (IDC anchor: 217B daily actions across 1.2B agents by 2029).")

ws.cell(row=4, column=1, value="A. Actions per task by workload archetype").font = SECTION_FONT
action_rows = [
    ("Simple chat / RAG turn", "1", "1-2", "Splunk tokenomics analysis"),
    ("Voice / contact-center agent", "3-8", "3-10", "Splunk tokenomics analysis"),
    ("Tool-using ReAct agent", "10-30", "10-20", "Splunk / Gartner via Cockroach Labs"),
    ("Research agent (GAIA-class)", "10-30", "10-30", "Splunk tokenomics analysis"),
    ("Coding agent (Cursor, SWE-bench)", "20-50", "15-50+", "Splunk tokenomics analysis"),
    ("Typical enterprise agentic workflow", "n/a", "~15 model calls per user task", "Cockroach Labs (2026)"),
]
end = write_table(
    ws, 5,
    ["Workload archetype", "Turns per task", "Model / tool calls per task", "Source"],
    action_rows,
    col_widths=[40, 14, 28, 38],
)

r2 = end + 2
ws.cell(row=r2, column=1,
        value="B. Derived annual series (modeled; 2029 calibrated to IDC)").font = SECTION_FONT
annual_rows = []
for i, y in enumerate(YEARS):
    status = "IDC anchor" if y == 2029 else ("IDC anchor (agents)" if y == 2025 else "Modeled")
    annual_rows.append((
        y,
        AGENTS_M[i],
        ACTIONS_PER_AGENT_DAY[i],
        round(DAILY_ACTIONS_B[i], 1),
        round(DAILY_ACTIONS_B[i] * 365, 0),
        status,
    ))
end = write_table(
    ws, r2 + 1,
    ["Year", "Deployed agents (M)", "Actions per agent per day",
     "Total daily actions (B)", "Total annual actions (B)", "Data status"],
    annual_rows,
    num_formats={2: "#,##0", 3: "#,##0", 4: "#,##0.0", 5: "#,##0"},
)

chart = LineChart()
chart.title = "Total daily agent actions (B, log scale)"
chart.style = 12
chart.y_axis.title = "Daily actions (B)"
chart.y_axis.scaling.logBase = 10
chart.height = 8
chart.width = 16
data = Reference(ws, min_col=4, max_col=4, min_row=r2 + 1, max_row=r2 + 1 + len(YEARS))
cats = Reference(ws, min_col=1, min_row=r2 + 2, max_row=r2 + 1 + len(YEARS))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "H" + str(r2 + 1))

chart2 = LineChart()
chart2.title = "Actions per agent per day"
chart2.style = 12
chart2.y_axis.title = "Actions / agent / day"
chart2.height = 8
chart2.width = 16
data2 = Reference(ws, min_col=3, max_col=3, min_row=r2 + 1, max_row=r2 + 1 + len(YEARS))
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "H" + str(r2 + 18))
ws.sheet_view.showGridLines = False

# ------------------------- Sheet 6: Sources & Notes ------------------------
ws = wb.create_sheet("Sources & Notes")
add_title(ws, "Sources & Methodology Notes", "All disclosed anchors used in this workbook, with URLs.")

source_rows = [
    ("Google token volume (9.7T 5/2024; 480T 5/2025; 3.2Q 5/2026 per month; Gemini API 19B tok/min)",
     "Google I/O 2026 keynote (Sundar Pichai)", "https://blog.google/innovation-and-ai/sundar-pichai-io-2026/"),
    ("Google 980T tokens/month (Jul 2025); 1.3Q tokens/month (Oct 2025)",
     "Alphabet Q2-2025 earnings call; Google blog Oct 2025 (via Robonomics token tracker)",
     "https://robonomics.substack.com/p/token-tracker-and-implications"),
    ("Gemini API ~22B tokens/min; Gemini app 950M MAU (Q2 2026)",
     "Alphabet Q2-2026 earnings release (SEC exhibit)",
     "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000066/googexhibit991q22026.htm"),
    ("Microsoft Foundry >100T tokens in FY25 Q3 (5x YoY); 50T in March 2025",
     "Microsoft FY25 Q3 earnings call (Apr 30, 2025)",
     "https://www.microsoft.com/en-us/investor/events/fy-2025/earnings-fy-2025-q3"),
    ("OpenRouter 10T/yr (2024) -> 100T/yr (mid-2025) -> >1T/day (late 2025); OpenAI API ~8.6T tokens/day (Oct 2025)",
     "a16z State of AI: 100T-token study with OpenRouter", "https://a16z.com/state-of-ai/"),
    ("OpenRouter >100T tokens/month and >1Q/yr annualized (Aug 2026); 8M+ developers",
     "Contrary Research: OpenRouter business breakdown", "https://research.contrary.com/company/openrouter"),
    ("Programming share of OpenRouter tokens: ~11% (early 2025) -> >50% (late 2025); reasoning-model share >50%",
     "OpenRouter State of AI 2025 (arXiv:2601.10088)", "https://openrouter.ai/state-of-ai"),
    ("Agentic tokens surpassed human-driven ~Feb 1, 2026; agentic requests ~15x tokens; 450T-token sample H1-2026",
     "OpenRouter blog: DeepSeek V4 agentic token share", "https://openrouter.ai/blog/insights/deepseek-v4-adoption/"),
    ("IDC Agent Economics: >1B (1.2B) agents by 2029, 40x vs 2025; 217B daily actions; >$68B delivery cost; "
     "agent types (in-application, standalone, custom-configured, bespoke); ~40->50 agents per org",
     "IDC Agent Economics Adoption & Delivery Model (Dec 2025); IDC Directions 2026",
     "https://www.idc.com/resource-center/blog/agent-adoption-the-it-industrys-next-great-inflection-point/"),
    ("IDC FutureScape 2026: G2000 agent use +10x by 2027; token/API call loads +1,000x; agentic automation in "
     ">40% of enterprise apps by 2027",
     "IDC FutureScape 2026 (via IDC blog)",
     "https://www.idc.com/resource-center/blog/agent-adoption-the-it-industrys-next-great-inflection-point/"),
    ("Gartner: task-specific agents in 40% of enterprise apps by end-2026 (<5% in 2025); 1/3 of implementations "
     "multi-agent by 2027; ~30% of enterprise app software revenue (> $450B) agentic by 2029 (best case)",
     "Gartner press release (Aug 26, 2025)",
     "https://www.gartner.com/en/newsroom/press-releases/2025-08-26-gartner-predicts-40-percent-of-enterprise-apps-"
     "will-feature-task-specific-ai-agents-by-2026-up-from-less-than-5-percent-in-2025"),
    ("Gartner inference paradox: agentic workflow inference costs +5x through 2028; token costs -95% by 2030; "
     "agents need 5-30x more tokens than chatbots; reasoning agents up to 150x cost per task",
     "Computerworld coverage of Gartner Tokenomics Model",
     "https://www.computerworld.com/article/4210786/ai-inference-is-getting-cheaper-but-your-agents-are-getting-"
     "more-expensive.html"),
    ("Agentic workflows 5-30x tokens/task; 10-20 model calls/task; enterprise inference = 85% of AI budgets; "
     "OpenAI top customer 100K -> ~100B tokens/month",
     "Cockroach Labs: Managing agentic AI costs at scale (2026)",
     "https://www.cockroachlabs.com/blog/agentic-ai-costs-at-scale/"),
    ("Tokens-per-task by archetype (RAG 2-10K; voice 5-15K; ReAct 20-60K; GAIA 30-100K; Cursor SWE-bench ~188K)",
     "Splunk: AI tokenomics", "https://www.splunk.com/en_us/blog/artificial-intelligence/ai-tokenomics-govern-"
     "agentic-ai-costs-with-splunk.html"),
    ("Agentic coding 1M-3.5M tokens/task; ~1,000x code chat; input tokens dominate; 30x run-to-run variance",
     "Zylos Research (2026); Tokenade; Bai et al. arXiv:2604.22750",
     "https://zylos.ai/research/2026-06-30-token-budget-management-cost-control-autonomous-agents/"),
    ("Enterprise agent adoption: 11% production / 38% piloting (late 2025)",
     "Deloitte (via IDC Agentic AI at Scale, 2026)",
     "https://event.idc.com/wp-content/uploads/2026/06/Agentic-AI-at-Scale.pdf"),
]
end = write_table(ws, 4, ["Figure(s) used", "Source", "URL"], source_rows,
                  col_widths=[80, 45, 70])

method_notes = [
    "Methodology notes:",
    "1. Monthly platform series use geometric (constant-growth) interpolation between disclosed anchors. "
    "Extrapolations use disclosed growth rates: Google forward at the Gemini-API Q2-2026 pace (~7.6%/mo), "
    "Microsoft at its disclosed 5x YoY pace, OpenRouter at its Dec-2025 -> Aug-2026 pace.",
    "2. The agentic share (%) series is modeled from qualitative OpenRouter disclosures (negligible in 2024, "
    "~50% crossing on Feb 1 2026, still rising mid-2026). It is applied to OpenRouter and, as a proxy, to the "
    "observable platform total.",
    "3. The annual agent model is calibrated to IDC anchors: ~30M agents (2025), 1.2B agents and 217B daily "
    "actions (2029). Tokens per action (~17K in 2029) is set so agent token load reaches ~3.7Q tokens/day in "
    "2029, consistent with IDC's Agent Economics headline. Intermediate years are smooth growth paths.",
    "4. Agent-type mix reflects IDC's qualitative guidance (in-application agents dominate 2025; "
    "custom-configured and standalone agents dominate by 2030); exact splits are modeled.",
    "5. Provider scopes differ and are NOT additive to a clean industry total: Google counts all surfaces "
    "including consumer products; Microsoft counts Foundry API only; OpenRouter counts routed traffic; "
    "IDC counts enterprise agents. Treat cross-series comparisons as directional.",
    "6. All 2026+ values dated after Sep 2026 and all 2027-2030 values are forecasts/models, not reported data.",
]
row = end + 2
for nnote in method_notes:
    cell = ws.cell(row=row, column=1, value=nnote)
    cell.font = BODY_FONT
    cell.alignment = Alignment(wrap_text=False)
    row += 1
ws.sheet_view.showGridLines = False

OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
