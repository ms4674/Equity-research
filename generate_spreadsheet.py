#!/usr/bin/env python3
"""
Generate a comprehensive Excel spreadsheet comparing token consumption
across five AI agent verticals: Coding Agents, Video Editing, Customer Service,
Contract Review, and Legal.

Data sources:
  - arxiv.org/abs/2604.22750 (agentic coding token analysis)
  - AgentMarketCap.ai (enterprise agent cost data, Apr 2026)
  - iternal.ai/token-usage-guide (10 use-case cost profiles, 2026)
  - RelayPlane.com/benchmarks/agent-costs (production agent benchmarks, Mar 2026)
  - AICostCheck.com (agent cost breakdowns, Mar 2026)
  - TokenMix.ai (chatbot & API cost calculators, Apr 2026)
  - Zylos.ai (AI agent cost optimization, Feb 2026)
  - yemhub.com (legal document AI cost calculator, 2026)
  - aicostcheck.com (document summarization costs, 2026)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from copy import copy

wb = openpyxl.Workbook()

# ── Style palette ──────────────────────────────────────────────────────────────
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", bold=True, size=11)
LIGHT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
ACCENT_FILLS = {
    "coding":   PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "video":    PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "cs":       PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
    "contract": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "legal":    PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid"),
}
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
NUM_FMT_COMMA = '#,##0'
NUM_FMT_DOLLAR = '$#,##0.00'
NUM_FMT_DOLLAR4 = '$#,##0.0000'
NUM_FMT_PERCENT = '0.0%'


def style_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_subheader_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def apply_borders(ws, min_row, max_row, max_col):
    for r in range(min_row, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER


def auto_width(ws, max_col, min_width=12, max_width=28):
    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        best = min_width
        for cell in ws[col_letter]:
            if cell.value:
                best = max(best, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[col_letter].width = best


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — Per-Task Token Consumption Comparison
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Per-Task Comparison"

ws1.merge_cells("A1:H1")
title_cell = ws1["A1"]
title_cell.value = "Token Consumption Per Task — Five AI Agent Verticals (2026)"
title_cell.font = Font(name="Calibri", bold=True, size=14, color="2F5496")
title_cell.alignment = Alignment(horizontal="center")

headers = [
    "Vertical", "Task Type",
    "Avg Input\nTokens", "Avg Output\nTokens", "Avg Total\nTokens",
    "Cost @ Frontier\n($/task)", "Cost @ Mid-Tier\n($/task)", "Cost @ Budget\n($/task)",
]
for c, h in enumerate(headers, 1):
    ws1.cell(row=3, column=c, value=h)
style_header_row(ws1, 3, len(headers))

data_rows = [
    # (vertical, task, input, output, total, cost_frontier, cost_mid, cost_budget)
    # ── Coding Agents ──
    ("Coding Agents", "Single-file code edit", 4200, 850, 5050, 0.031, 0.019, 0.005),
    ("Coding Agents", "Multi-file refactor", 25000, 5000, 30000, 0.17, 0.10, 0.026),
    ("Coding Agents", "Code review (PR)", 15000, 3000, 18000, 0.062, 0.038, 0.010),
    ("Coding Agents", "New feature (end-to-end)", 42000, 8500, 50500, 0.53, 0.33, 0.084),
    ("Coding Agents", "Bug investigation", 20000, 4000, 24000, 0.089, 0.055, 0.014),
    ("Coding Agents", "Agentic SWE-bench task", 2500000, 500000, 3000000, 5.00, 3.00, 0.80),

    # ── Video Editing ──
    ("Video Editing", "Scene description / script (text-only)", 2000, 500, 2500, 0.04, 0.02, 0.003),
    ("Video Editing", "Video analysis (1 min clip)", 946800, 50000, 996800, 16.70, 8.50, 1.20),
    ("Video Editing", "Edit instruction generation", 5000, 1500, 6500, 0.10, 0.05, 0.008),
    ("Video Editing", "Multi-step workflow (5-10 actions)", 50000, 3500, 53500, 0.85, 0.42, 0.06),
    ("Video Editing", "Video generation (15s clip)", 308880, 0, 308880, 2.18, 1.00, 0.30),

    # ── Customer Service ──
    ("Customer Service", "Simple chatbot query", 800, 200, 1000, 0.012, 0.005, 0.001),
    ("Customer Service", "FAQ / KB lookup", 2500, 400, 2900, 0.045, 0.018, 0.002),
    ("Customer Service", "Multi-turn support chat", 6200, 1500, 7700, 0.12, 0.05, 0.008),
    ("Customer Service", "Complex agent (tool use)", 14000, 800, 14800, 0.15, 0.047, 0.005),
    ("Customer Service", "Escalation w/ sentiment analysis", 10000, 600, 10600, 0.10, 0.040, 0.006),

    # ── Contract Review ──
    ("Contract Review", "Short contract (10 pp)", 12000, 1200, 13200, 0.054, 0.026, 0.005),
    ("Contract Review", "Standard contract (30 pp)", 15000, 1200, 16200, 0.063, 0.030, 0.005),
    ("Contract Review", "Long contract (50 pp)", 25000, 1500, 26500, 0.105, 0.056, 0.009),
    ("Contract Review", "Due diligence (multi-doc)", 275000, 10000, 285000, 3.85, 1.50, 0.25),
    ("Contract Review", "Clause extraction (batch, per doc)", 15000, 800, 15800, 0.063, 0.030, 0.002),

    # ── Legal (General) ──
    ("Legal", "Legal research query (RAG)", 12500, 2000, 14500, 0.18, 0.07, 0.01),
    ("Legal", "Brief drafting", 6000, 5000, 11000, 0.13, 0.06, 0.008),
    ("Legal", "Case law analysis", 50000, 5000, 55000, 0.65, 0.28, 0.04),
    ("Legal", "Regulatory compliance check", 30000, 3000, 33000, 0.38, 0.16, 0.02),
    ("Legal", "Full litigation support session", 200000, 20000, 220000, 3.50, 1.40, 0.20),
]

vertical_colors = {
    "Coding Agents": ACCENT_FILLS["coding"],
    "Video Editing": ACCENT_FILLS["video"],
    "Customer Service": ACCENT_FILLS["cs"],
    "Contract Review": ACCENT_FILLS["contract"],
    "Legal": ACCENT_FILLS["legal"],
}

for i, row in enumerate(data_rows):
    r = 4 + i
    vertical, task, inp, out, total, cf, cm, cb = row
    ws1.cell(row=r, column=1, value=vertical)
    ws1.cell(row=r, column=2, value=task)
    ws1.cell(row=r, column=3, value=inp).number_format = NUM_FMT_COMMA
    ws1.cell(row=r, column=4, value=out).number_format = NUM_FMT_COMMA
    ws1.cell(row=r, column=5, value=total).number_format = NUM_FMT_COMMA
    ws1.cell(row=r, column=6, value=cf).number_format = NUM_FMT_DOLLAR
    ws1.cell(row=r, column=7, value=cm).number_format = NUM_FMT_DOLLAR
    ws1.cell(row=r, column=8, value=cb).number_format = NUM_FMT_DOLLAR4

    fill = vertical_colors.get(vertical, LIGHT_FILL)
    for c in range(1, 9):
        ws1.cell(row=r, column=c).fill = fill

apply_borders(ws1, 3, 3 + len(data_rows), 8)
auto_width(ws1, 8, min_width=14)


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — Quarterly Aggregate Time Series (Q1 2024 – Q1 2026)
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Quarterly Time Series")

ws2.merge_cells("A1:G1")
t2 = ws2["A1"]
t2.value = "Estimated Aggregate Token Consumption — Quarterly Time Series (Billions of Tokens / Quarter)"
t2.font = Font(name="Calibri", bold=True, size=13, color="2F5496")
t2.alignment = Alignment(horizontal="center")

ws2.merge_cells("A2:G2")
note = ws2["A2"]
note.value = (
    "Estimates synthesised from OpenRouter traffic data, Anthropic/OpenAI disclosures, "
    "Gartner projections (40% enterprise agent adoption by 2026), and arxiv:2604.22750. "
    "Video editing represents LLM text tokens consumed for planning/editing workflows, not raw video generation compute."
)
note.font = Font(name="Calibri", italic=True, size=9, color="808080")
note.alignment = Alignment(wrap_text=True)

ts_headers = [
    "Quarter",
    "Coding Agents\n(B tokens)",
    "Video Editing\n(B tokens)",
    "Customer Service\n(B tokens)",
    "Contract Review\n(B tokens)",
    "Legal\n(B tokens)",
    "Total\n(B tokens)",
]
for c, h in enumerate(ts_headers, 1):
    ws2.cell(row=4, column=c, value=h)
style_header_row(ws2, 4, len(ts_headers))

# Time-series estimates (billions of tokens per quarter, global aggregate)
# Rationale:
#   - Programming rose from 11% to >50% of all LLM token usage by late 2025 (OpenRouter).
#   - Weekly token volume grew >3,000% over 18 months through mid-2025.
#   - Enterprise LLM spending hit $8.4B in H1 2025.
#   - Gartner: 40% enterprise apps integrate AI agents by 2026 (up from <5% in 2025).
#   - Coding agent tasks consume 1000x more tokens than chat (arxiv:2604.22750).
#   - Customer service is highest-volume but low tokens/task.
#   - Legal/contract adoption lags coding but is accelerating.
ts_data = [
    # (quarter, coding, video, cs, contract, legal)
    ("Q1 2024",   80,    5,   120,   8,   6),
    ("Q2 2024",  140,   10,   180,  14,  10),
    ("Q3 2024",  250,   18,   260,  22,  16),
    ("Q4 2024",  450,   30,   380,  35,  25),
    ("Q1 2025",  800,   55,   550,  55,  40),
    ("Q2 2025", 1400,  100,   800,  90,  65),
    ("Q3 2025", 2400,  180,  1150, 140, 100),
    ("Q4 2025", 4000,  320,  1650, 220, 160),
    ("Q1 2026", 6500,  550,  2300, 350, 260),
]

for i, (qtr, coding, video, cs, contract, legal) in enumerate(ts_data):
    r = 5 + i
    total = coding + video + cs + contract + legal
    ws2.cell(row=r, column=1, value=qtr)
    ws2.cell(row=r, column=2, value=coding).number_format = NUM_FMT_COMMA
    ws2.cell(row=r, column=3, value=video).number_format = NUM_FMT_COMMA
    ws2.cell(row=r, column=4, value=cs).number_format = NUM_FMT_COMMA
    ws2.cell(row=r, column=5, value=contract).number_format = NUM_FMT_COMMA
    ws2.cell(row=r, column=6, value=legal).number_format = NUM_FMT_COMMA
    ws2.cell(row=r, column=7, value=total).number_format = NUM_FMT_COMMA

    if i % 2 == 0:
        for c in range(1, 8):
            ws2.cell(row=r, column=c).fill = LIGHT_FILL

apply_borders(ws2, 4, 4 + len(ts_data), 7)
auto_width(ws2, 7, min_width=16)

# QoQ growth row
grow_row = 5 + len(ts_data) + 1
ws2.cell(row=grow_row, column=1, value="Q1-2024 → Q1-2026 Growth")
ws2.cell(row=grow_row, column=1).font = Font(bold=True)
multipliers = [
    6500 / 80,   # coding
    550 / 5,     # video
    2300 / 120,  # cs
    350 / 8,     # contract
    260 / 6,     # legal
    (6500+550+2300+350+260) / (80+5+120+8+6),  # total
]
for c, m in enumerate(multipliers, 2):
    cell = ws2.cell(row=grow_row, column=c, value=f"{m:.0f}x")
    cell.font = Font(bold=True, color="C00000")
    cell.alignment = Alignment(horizontal="center")

# Market share row
share_row = grow_row + 1
ws2.cell(row=share_row, column=1, value="Q1 2026 Share of Total")
ws2.cell(row=share_row, column=1).font = Font(bold=True)
q1_total = 6500 + 550 + 2300 + 350 + 260
for c, val in enumerate([6500, 550, 2300, 350, 260, q1_total], 2):
    cell = ws2.cell(row=share_row, column=c, value=val / q1_total)
    cell.number_format = NUM_FMT_PERCENT
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

# ── Line chart ──
chart = LineChart()
chart.title = "Quarterly Token Consumption by Vertical (B tokens)"
chart.y_axis.title = "Billion Tokens"
chart.x_axis.title = "Quarter"
chart.style = 10
chart.height = 15
chart.width = 24

cats = Reference(ws2, min_col=1, min_row=5, max_row=5 + len(ts_data) - 1)
for col_idx in range(2, 7):
    vals = Reference(ws2, min_col=col_idx, min_row=4, max_row=5 + len(ts_data) - 1)
    chart.add_data(vals, titles_from_data=True)
chart.set_categories(cats)

colors = ["548235", "ED7D31", "4472C4", "FFC000", "7030A0"]
for idx, series in enumerate(chart.series):
    series.graphicalProperties.line.width = 25000
    if idx < len(colors):
        series.graphicalProperties.line.solidFill = colors[idx]

ws2.add_chart(chart, "A" + str(share_row + 3))


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — Monthly Cost Projections at Scale
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Monthly Cost @ Scale")

ws3.merge_cells("A1:H1")
t3 = ws3["A1"]
t3.value = "Monthly Cost Projections at Production Scale — by Vertical & Model Tier (March 2026 Pricing)"
t3.font = Font(name="Calibri", bold=True, size=13, color="2F5496")
t3.alignment = Alignment(horizontal="center")

cost_headers = [
    "Vertical", "Tasks/Month",
    "Budget Model\n($/mo)", "Mid-Tier\n($/mo)", "Frontier\n($/mo)",
    "Budget-to-Frontier\nMultiplier",
    "Optimized (Routed)\n($/mo)", "Savings vs\nFrontier",
]
for c, h in enumerate(cost_headers, 1):
    ws3.cell(row=3, column=c, value=h)
style_header_row(ws3, 3, len(cost_headers))

cost_data = [
    # (vertical, tasks_mo, budget, mid, frontier, multiplier, optimized, savings_pct)
    ("Coding Agents",    1500,    60,    788,   1313,  22, 250,   0.81),
    ("Video Editing",     500,    15,    210,    425,  28,  85,   0.80),
    ("Customer Service",10000,    10,     50,    900,  90,  22,   0.976),
    ("Contract Review",  2000,    10,     60,    126,  13,  28,   0.78),
    ("Legal",            1000,    10,     70,    180,  18,  35,   0.81),
]

for i, (vert, tasks, bud, mid, front, mult, opt, sav) in enumerate(cost_data):
    r = 4 + i
    ws3.cell(row=r, column=1, value=vert)
    ws3.cell(row=r, column=2, value=tasks).number_format = NUM_FMT_COMMA
    ws3.cell(row=r, column=3, value=bud).number_format = NUM_FMT_DOLLAR
    ws3.cell(row=r, column=4, value=mid).number_format = NUM_FMT_DOLLAR
    ws3.cell(row=r, column=5, value=front).number_format = NUM_FMT_DOLLAR
    ws3.cell(row=r, column=6, value=f"{mult}x").alignment = Alignment(horizontal="center")
    ws3.cell(row=r, column=7, value=opt).number_format = NUM_FMT_DOLLAR
    ws3.cell(row=r, column=8, value=sav).number_format = NUM_FMT_PERCENT

    fill = vertical_colors.get(vert, LIGHT_FILL)
    for c in range(1, 9):
        ws3.cell(row=r, column=c).fill = fill

apply_borders(ws3, 3, 3 + len(cost_data), 8)
auto_width(ws3, 8, min_width=14)

# Bar chart
bar = BarChart()
bar.type = "col"
bar.title = "Monthly Cost by Vertical & Model Tier"
bar.y_axis.title = "USD / Month"
bar.style = 10
bar.height = 14
bar.width = 22

cats = Reference(ws3, min_col=1, min_row=4, max_row=4 + len(cost_data) - 1)
for col_idx in [3, 4, 5, 7]:
    vals = Reference(ws3, min_col=col_idx, min_row=3, max_row=3 + len(cost_data))
    bar.add_data(vals, titles_from_data=True)
bar.set_categories(cats)
ws3.add_chart(bar, "A" + str(4 + len(cost_data) + 2))


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — Token Economics Detail
# ═══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Token Economics Detail")

ws4.merge_cells("A1:G1")
t4 = ws4["A1"]
t4.value = "Token Economics — Structural Characteristics by Vertical"
t4.font = Font(name="Calibri", bold=True, size=13, color="2F5496")
t4.alignment = Alignment(horizontal="center")

econ_headers = [
    "Characteristic",
    "Coding Agents", "Video Editing", "Customer Service",
    "Contract Review", "Legal",
]
for c, h in enumerate(econ_headers, 1):
    ws4.cell(row=3, column=c, value=h)
style_header_row(ws4, 3, len(econ_headers))

econ_data = [
    ("Avg tokens/task (typical)", "18,000–3,500,000", "2,500–1,000,000", "1,000–15,000", "13,000–285,000", "11,000–220,000"),
    ("Input:Output ratio", "~5:1 to 10:1", "~19:1 (video heavy)", "~4:1 to 18:1", "~12:1 to 15:1", "~3:1 to 10:1"),
    ("Cost driver", "Context accumulation,\ntool call overhead", "Multimodal (vision)\ntoken volume", "Volume of\ninteractions", "Document length,\nmulti-page ingest", "RAG retrieval,\nlong-context windows"),
    ("LLM calls/task", "8–15+", "3–10", "2–5", "1–3", "2–8"),
    ("Context window utilisation", "High (50–80%)", "Very high (vision)", "Low–Medium (10–30%)", "Medium (30–60%)", "Medium–High (40–70%)"),
    ("Prompt caching savings", "40–60%", "20–35%", "30–50%", "40–60%", "30–50%"),
    ("Model routing savings", "60–80%", "50–70%", "80–95%", "70–85%", "60–80%"),
    ("Variability (same task)", "Up to 30x", "Moderate (3–5x)", "Low (1.5–3x)", "Low (1.5–2x)", "Moderate (2–5x)"),
    ("Primary model tier", "Frontier", "Frontier/Multimodal", "Budget/Mid-tier", "Mid-tier", "Frontier/Mid-tier"),
    ("Share of global LLM tokens\n(Q1 2026 est.)", "~65%", "~6%", "~23%", "~3.5%", "~2.5%"),
    ("Growth trajectory", "Fastest (81x since Q1-24)", "Very fast (110x)", "Steady (19x)", "Fast (44x)", "Fast (43x)"),
]

for i, row_data in enumerate(econ_data):
    r = 4 + i
    for c, val in enumerate(row_data):
        cell = ws4.cell(row=r, column=c + 1, value=val)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        if c == 0:
            cell.font = Font(bold=True)
    if i % 2 == 0:
        for c in range(1, 7):
            ws4.cell(row=r, column=c).fill = LIGHT_FILL

apply_borders(ws4, 3, 3 + len(econ_data), 6)
auto_width(ws4, 6, min_width=18, max_width=32)


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — Model Pricing Reference
# ═══════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Model Pricing (Mar 2026)")

ws5.merge_cells("A1:F1")
t5 = ws5["A1"]
t5.value = "LLM API Pricing Reference — March 2026"
t5.font = Font(name="Calibri", bold=True, size=13, color="2F5496")
t5.alignment = Alignment(horizontal="center")

price_headers = ["Model", "Provider", "Input $/1M", "Output $/1M", "Context Window", "Best For"]
for c, h in enumerate(price_headers, 1):
    ws5.cell(row=3, column=c, value=h)
style_header_row(ws5, 3, len(price_headers))

pricing = [
    ("Claude Opus 4.6", "Anthropic", 15.00, 75.00, "200K", "Hardest tasks, legal, max capability"),
    ("Claude Sonnet 4.6", "Anthropic", 3.00, 15.00, "200K", "Complex reasoning, coding, large codebases"),
    ("Claude Haiku 4.5", "Anthropic", 0.80, 4.00, "200K", "Fast, cheap tasks, high volume"),
    ("GPT-5.4", "OpenAI", 2.50, 15.00, "128K", "General tasks, vision, broad compatibility"),
    ("GPT-5 mini", "OpenAI", 0.25, 2.00, "128K", "Low-cost OpenAI-compatible workloads"),
    ("GPT-5 nano", "OpenAI", 0.05, 0.40, "128K", "Classification, routing, simple tasks"),
    ("Gemini 3 Pro", "Google", 4.00, 18.00, "2M", "Long context, video analysis, multimodal"),
    ("Gemini 2.5 Flash", "Google", 0.15, 0.60, "1M", "Lowest cost, massive context"),
    ("DeepSeek V3.2", "DeepSeek", 0.14, 0.28, "128K", "Budget agent workloads"),
    ("Qwen3-Coder-480B", "Alibaba", 1.00, 5.00, "256K", "Coding-specific tasks"),
]

for i, (model, provider, inp_p, out_p, ctx, best) in enumerate(pricing):
    r = 4 + i
    ws5.cell(row=r, column=1, value=model)
    ws5.cell(row=r, column=2, value=provider)
    ws5.cell(row=r, column=3, value=inp_p).number_format = NUM_FMT_DOLLAR
    ws5.cell(row=r, column=4, value=out_p).number_format = NUM_FMT_DOLLAR
    ws5.cell(row=r, column=5, value=ctx).alignment = Alignment(horizontal="center")
    ws5.cell(row=r, column=6, value=best)
    if i % 2 == 0:
        for c in range(1, 7):
            ws5.cell(row=r, column=c).fill = LIGHT_FILL

apply_borders(ws5, 3, 3 + len(pricing), 6)
auto_width(ws5, 6, min_width=14, max_width=40)


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 6 — Sources & Methodology
# ═══════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Sources & Methodology")

ws6.merge_cells("A1:B1")
ws6["A1"].value = "Data Sources & Methodology"
ws6["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")

sources = [
    ("arxiv:2604.22750", "How Do AI Agents Spend Your Money? — First systematic study of token consumption in agentic coding (Apr 2026)"),
    ("AgentMarketCap.ai", "The AI Agent Token Consumption Gap — Enterprise agentic workload cost analysis (Apr 2026)"),
    ("iternal.ai", "AI Token Usage Guide 2026 — 10 use-case cost profiles with token breakdowns"),
    ("RelayPlane.com", "Agent Cost Benchmarks 2026 — Real production median token counts (Mar 2026)"),
    ("AICostCheck.com", "AI Agent Costs 2026 — Real math + 7 optimisation strategies; Document summarisation costs (Mar 2026)"),
    ("TokenMix.ai", "AI Video Generation API 2026 — Provider cost comparison; Chatbot cost calculator (Apr 2026)"),
    ("Zylos.ai", "AI Agent Cost Optimization — Token economics & FinOps in production (Feb 2026)"),
    ("yemhub.com", "Legal Document AI Cost Calculator — Contract review & analysis pricing (2026)"),
    ("DigitalApplied.com", "LLM API Pricing Index — Agent deployment cost tracking & patterns (Q1 2026)"),
    ("OpenRouter data", "Programming share rose from 11% to >50% of all LLM token usage by late 2025"),
    ("Gartner", "40% of enterprise apps to integrate AI agents by 2026, up from <5% in 2025"),
]

ws6.cell(row=3, column=1, value="Source").font = Font(bold=True)
ws6.cell(row=3, column=2, value="Description").font = Font(bold=True)
style_header_row(ws6, 3, 2)

for i, (src, desc) in enumerate(sources):
    r = 4 + i
    ws6.cell(row=r, column=1, value=src).font = Font(bold=True)
    ws6.cell(row=r, column=2, value=desc).alignment = Alignment(wrap_text=True)
    if i % 2 == 0:
        ws6.cell(row=r, column=1).fill = LIGHT_FILL
        ws6.cell(row=r, column=2).fill = LIGHT_FILL

apply_borders(ws6, 3, 3 + len(sources), 2)
ws6.column_dimensions["A"].width = 24
ws6.column_dimensions["B"].width = 90

# Methodology note
meth_row = 4 + len(sources) + 2
ws6.cell(row=meth_row, column=1, value="Methodology Note")
ws6.cell(row=meth_row, column=1).font = Font(bold=True, size=12, color="2F5496")
ws6.merge_cells(f"A{meth_row+1}:B{meth_row+5}")
meth_text = (
    "Aggregate time-series estimates are constructed by triangulating:\n"
    "1. OpenRouter traffic data showing programming's rise from 11% → 50%+ of total LLM tokens.\n"
    "2. Per-task token consumption medians from RelayPlane and iternal.ai production data.\n"
    "3. Enterprise adoption rates from Gartner and industry surveys.\n"
    "4. API pricing data tracked by TokenMix, AICostCheck, and provider documentation.\n"
    "5. Agentic coding task analysis from arxiv:2604.22750 (SWE-bench Verified benchmarks).\n\n"
    "Video editing figures represent LLM text tokens consumed for AI-assisted editing workflows "
    "(scene analysis, edit instructions, metadata generation), not raw video generation compute. "
    "Video generation token counts (e.g. 308K tokens for a 15s clip via Seedance 2.0) are included where applicable.\n\n"
    "All cost figures use March 2026 list pricing with no volume discounts unless otherwise noted."
)
ws6.cell(row=meth_row + 1, column=1, value=meth_text).alignment = Alignment(wrap_text=True, vertical="top")


# ═══════════════════════════════════════════════════════════════════════════════
# Save
# ═══════════════════════════════════════════════════════════════════════════════
output_path = "/workspace/token_consumption_by_vertical.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
