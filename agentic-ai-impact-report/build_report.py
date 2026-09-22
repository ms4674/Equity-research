"""
Builds the Agentic AI Market Impact Report workbook.

Covers agentic AI impact and 2025-2030 forecasts for five sectors:
broadband pricing, e-commerce traffic, delivery, ridesharing, and
online travel marketplaces.

Usage: python3 build_report.py
Output: Agentic_AI_Market_Impact_Report.xlsx (same directory)
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel, StrRef
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- styling ---

NAVY = "1F3864"
BLUE = "2E75B6"
LIGHT_BLUE = "DDEBF7"
GREY = "F2F2F2"
AMBER = "FFF2CC"

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=11, italic=True, color="595959")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=11)
BOLD_FONT = Font(name="Calibri", size=11, bold=True)

SECTION_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FILL = PatternFill("solid", fgColor=BLUE)
STRIPE_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
NOTE_FILL = PatternFill("solid", fgColor=AMBER)

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title_block(ws, row, title, subtitle=None):
    ws.cell(row=row, column=1, value=title).font = TITLE_FONT
    row += 1
    if subtitle:
        ws.cell(row=row, column=1, value=subtitle).font = SUBTITLE_FONT
        row += 1
    return row + 1


def section(ws, row, text, span):
    for c in range(1, span + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = SECTION_FILL
        cell.border = BOX
    ws.cell(row=row, column=1, value=text).font = SECTION_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    return row + 1


def table(ws, row, headers, rows, formats=None, stripe=True):
    """Write a styled table; formats maps column index (0-based) -> number format."""
    formats = formats or {}
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BOX
        cell.alignment = CENTER
    row += 1
    for i, data_row in enumerate(rows):
        for c, val in enumerate(data_row, start=1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.font = BODY_FONT
            cell.border = BOX
            cell.alignment = WRAP
            if stripe and i % 2 == 1:
                cell.fill = STRIPE_FILL
            if (c - 1) in formats and isinstance(val, (int, float)):
                cell.number_format = formats[c - 1]
        row += 1
    return row + 1


def note(ws, row, text, span):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="Calibri", size=10, italic=True)
    cell.fill = NOTE_FILL
    cell.alignment = WRAP
    ws.row_dimensions[row].height = 30
    return row + 2

PCT = "0.0%"
PCT1 = "0.00%"
USD_B = "$#,##0.0,, \"B\""
NUM = "#,##0.0"


# --------------------------------------------------------------- workbook ---

wb = Workbook()

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]

# Base-case agent-mediated share of sector activity (own scenario model,
# anchored to third-party 2026 data points and published 2030 ranges).
SHARE_BASE = {
    "E-commerce (share of US online retail sales via AI agents/platforms)":
        [0.004, 0.015, 0.035, 0.07, 0.11, 0.15],
    "Food & grocery delivery (share of orders initiated by AI agents)":
        [0.001, 0.01, 0.04, 0.08, 0.14, 0.20],
    "Ridesharing (share of bookings initiated by AI agents)":
        [0.0005, 0.005, 0.02, 0.05, 0.08, 0.12],
    "Online travel (share of online bookings mediated by AI agents)":
        [0.002, 0.01, 0.03, 0.06, 0.10, 0.15],
    "Broadband (AI inference share of consumer network traffic)":
        [0.02, 0.03, 0.05, 0.07, 0.09, 0.12],
}

# IHL Group agent-to-agent commerce forecast, global, $B.
IHL_A2A = {2024: 4, 2025: 29, 2026: 153, 2027: 574, 2028: 1567, 2029: 2796, 2030: 4000}

# ------------------------------------------------------------ 1. Overview ---

ws = wb.active
ws.title = "Overview"
ws.sheet_properties.tabColor = NAVY
set_widths(ws, [28, 60, 60])

r = title_block(
    ws, 1,
    "Agentic AI Market Impact Report: Five Consumer Marketplaces",
    "Impact assessment and 2025-2030 forecasts | Compiled September 2026 | All figures from public sources cited in 'Sources & Methodology'",
)

r = section(ws, r, "What this report covers", 3)
scope = (
    "This workbook assesses how agentic AI (autonomous AI agents that discover, compare, negotiate and "
    "transact on a user's behalf) is reshaping five consumer-facing markets: broadband pricing, e-commerce "
    "traffic, delivery, ridesharing, and online travel marketplaces. Each sector tab contains: (1) current "
    "impact evidence from published data, (2) sourced market statistics, and (3) a 2025-2030 forecast with "
    "bear / base / bull scenarios. The 'Forecast Summary' tab consolidates the cross-sector outlook."
)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
cell = ws.cell(row=r, column=1, value=scope)
cell.alignment = WRAP
ws.row_dimensions[r].height = 75
r += 2

r = section(ws, r, "Executive summary by sector", 3)
r = table(
    ws, r,
    ["Sector", "Current impact (2025-2026 evidence)", "2030 outlook (base case)"],
    [
        ["Broadband pricing",
         "AI agents generate up to 450% more network traffic per task than humans (Cisco). AI-inference flows "
         "are longer, more upstream-heavy and latency-sensitive than web/video traffic, pressuring the "
         "flat-rate speed-tier pricing model. Operators (Deutsche Telekom, Vodafone, Orange, Telefonica) are "
         "planning agent-consumable network APIs; IETF drafts propose on-demand, per-agent acceleration plans.",
         "Consumer traffic grows ~63% above the non-AI baseline by 2035 (6.6x vs ~4x). Pricing shifts from "
         "pure speed tiers toward usage-, event- and capability-based plans (guaranteed latency, uplink "
         "boosts, per-agent QoS). Base case: 8-12% of consumer broadband revenue from agentic/premium "
         "dynamic tiers by 2030, with modest ARPU uplift of $3-8/month on premium plans."],
        ["E-commerce traffic",
         "Fastest-moving sector. AI-referred traffic to US retail sites +138% YoY (May 2026) and +1,219% vs "
         "Oct 2024 (Adobe). AI-referred sessions convert 54-60% better, generate 53% more revenue per visit, "
         "and carry 14% higher AOV (Adobe/Shopify). Shopify saw ~13x YoY growth in AI-originated orders in "
         "Q1 2026. Still ~1.5% of US retail e-commerce sales ($20.6B, EMARKETER 2026).",
         "Agent-driven share of US e-commerce sales reaches ~15% by 2030 (Bain: 15-25%; Morgan Stanley: "
         "10-20%; J.P. Morgan: up to 25%). Global agentic commerce GMV $3-5T (McKinsey); agent-to-agent "
         "commerce ~$4T (IHL base). Traffic mix shifts from search/social referral to agent APIs; SEO "
         "budgets migrate to 'agent visibility' optimization."],
        ["Delivery",
         "First consumer category to go agentic at scale: Google Gemini's background-task agent launched "
         "with DoorDash, Uber Eats and Grubhub as debut partners (Mar 2026). Uber Eats AI cart assistant "
         "(Feb 2026) and Instacart's assistant (Nov 2024) already auto-build carts. On the supply side, AI "
         "routing/dispatch and autonomous last-mile (bots, drones, self-driving vans) cut cost per drop.",
         "~20% of food & grocery delivery orders initiated by AI agents by 2030 (recurring, low-risk "
         "purchases lead agentic adoption per J.P. Morgan). Autonomous last-mile delivery market grows from "
         "~$1.6-2.3B (2024) to ~$4.2-6.2B by 2030 (22-25% CAGR). Platforms with the best agent APIs and "
         "fulfillment density win order flow routed by third-party agents."],
        ["Ridesharing",
         "Agents now book rides: Gemini's multistep automation includes ride-hailing (Uber among first "
         "partners). Uber posted $193.4B gross bookings in 2025 with $85.4B paid to ~9.7M drivers - the cost "
         "pool AI-dispatched autonomous vehicles target. Uber has 20+ AV partnerships incl. Waymo; AI "
         "dispatch/pricing optimization already lifts utilization.",
         "~12% of rideshare bookings agent-initiated by 2030. AVs scale city-by-city; platforms that "
         "aggregate AV fleets keep the marketplace position but face take-rate pressure from vertically "
         "integrated AV operators (Waymo, Tesla). Driver-payout pool partially converts to fleet margin; "
         "pricing becomes more dynamic (agent-negotiated, subscription bundles)."],
        ["Online travel marketplaces",
         "Discovery is moving into ChatGPT/Gemini surfaces. Google's agentic hotel booking launched through "
         "(not around) incumbents: Booking.com, Expedia, Marriott, IHG, Choice, Wyndham are the founding "
         "partners. OTAs are building their own agents (Priceline Penny, Booking.com agent, Expedia's Layla "
         "acquisition). Bernstein calls agentic AI a 'paradigm shift' pressuring take rates and marketing "
         "efficiency at Booking ($21B rev) and Expedia ($13.7B rev).",
         "~15% of online travel bookings agent-mediated by 2030. Volume persists but economics compress: "
         "take rates on agent-mediated bookings settle at 8-12% vs 15-20% direct-OTA today; performance- "
         "marketing spend partially shifts to AI-platform placement fees. Bear case (Expedia): 30-40% of "
         "bookings bypass the OTA storefront, compressing EBITDA margins toward 8-14%."],
    ],
)
for rr in range(r - 6, r):
    ws.row_dimensions[rr].height = 118

r = section(ws, r, "Cross-cutting themes", 3)
r = table(
    ws, r,
    ["#", "Theme", "Implication"],
    [
        [1, "Traffic explodes, but changes shape",
         "Agentic workloads add volume (up to 450% more per task) and shift traffic upstream, longer-lived and "
         "latency-sensitive - favoring ISPs/platforms that can monetize quality, not just quantity."],
        [2, "Discovery decouples from transaction",
         "Agents compress search-compare-buy into one step. Whoever owns the agent surface (OpenAI, Google, "
         "Apple) taxes the funnel; merchants and marketplaces compete on machine-readability (APIs, feeds, UCP/AP2 protocols)."],
        [3, "Conversion quality leads volume",
         "AI-referred users convert 50-60% better today. Early channel economics are strongly positive, which "
         "pulls forward merchant investment even while volumes are <2% of total."],
        [4, "Recurring, low-risk categories adopt first",
         "Groceries, meal delivery, rides and repeat travel bookings are the beachhead (J.P. Morgan); "
         "high-consideration purchases stay human-in-the-loop longer."],
        [5, "Incumbents are being routed through, not around",
         "Google's agentic partners are DoorDash/Uber/Booking/Expedia - agents currently need incumbent "
         "inventory, payments and fulfillment. The risk is margin, not existence, this decade."],
    ],
)
for rr in range(r - 5, r):
    ws.row_dimensions[rr].height = 48
ws.freeze_panes = "A2"

# ---------------------------------------------------- 2. Forecast Summary ---

ws = wb.create_sheet("Forecast Summary")
ws.sheet_properties.tabColor = BLUE
set_widths(ws, [62, 12, 12, 12, 12, 12, 12, 14])

r = title_block(
    ws, 1,
    "Consolidated Forecast, 2025-2030",
    "Base-case scenario; anchored to published 2026 data points and 2030 analyst ranges (see Sources & Methodology)",
)

r = section(ws, r, "Agent-mediated share of sector activity (base case)", 8)
share_rows = [[k] + v + [None] for k, v in SHARE_BASE.items()]
share_hdr_row = r
r = table(
    ws, r,
    ["Metric"] + YEARS + ["2030 range (bear-bull)"],
    share_rows,
    formats={i: PCT for i in range(1, 7)},
)
ranges_2030 = ["8% - 25%", "10% - 30%", "6% - 20%", "8% - 25%", "9% - 15%"]
for i, txt in enumerate(ranges_2030):
    ws.cell(row=share_hdr_row + 1 + i, column=8, value=txt).alignment = CENTER
for rr in range(share_hdr_row + 1, share_hdr_row + 6):
    ws.row_dimensions[rr].height = 30

chart = LineChart()
chart.title = "Agent-mediated share of sector activity (base case)"
chart.y_axis.numFmt = "0%"
chart.y_axis.title = "Share"
chart.x_axis.title = "Year"
chart.height = 9
chart.width = 24
data = Reference(ws, min_col=2, max_col=7, min_row=share_hdr_row + 1, max_row=share_hdr_row + 5)
cats = Reference(ws, min_col=2, max_col=7, min_row=share_hdr_row)
chart.add_data(data, titles_from_data=False, from_rows=True)
for i, series in enumerate(chart.series):
    ref = f"'{ws.title}'!$A${share_hdr_row + 1 + i}"
    series.tx = SeriesLabel(strRef=StrRef(ref))
chart.set_categories(cats)
ws.add_chart(chart, f"A{r + 1}")
r += 20

r = section(ws, r, "Global agent-to-agent commerce GMV, $B (IHL Group)", 8)
a2a_hdr_row = r
r = table(
    ws, r,
    ["Scenario", 2024] + list(YEARS),
    [
        ["Base case ($B)"] + [IHL_A2A[y] for y in [2024] + YEARS],
        ["Bear 2030: $2,200B | Bull 2030: $4,400B", None, None, None, None, None, None, None],
    ],
    formats={i: "#,##0" for i in range(1, 8)},
)

bar = BarChart()
bar.title = "Agent-to-agent commerce GMV, global $B (IHL base case)"
bar.y_axis.title = "$B"
bar.height = 8
bar.width = 24
data = Reference(ws, min_col=2, max_col=8, min_row=a2a_hdr_row + 1, max_row=a2a_hdr_row + 1)
bar.add_data(data, titles_from_data=False, from_rows=True)
bar.series[0].tx = SeriesLabel(strRef=StrRef(f"'{ws.title}'!$A${a2a_hdr_row + 1}"))
bar.set_categories(Reference(ws, min_col=2, max_col=8, min_row=a2a_hdr_row))
ws.add_chart(bar, f"A{r + 1}")
r += 18

r = section(ws, r, "Published 2030 anchors used in this forecast", 8)
r = table(
    ws, r,
    ["Anchor", "Value", "Source", "", "", "", "", ""],
    [
        ["US e-commerce sales via AI agents, 2030", "15-25% ($300-500B)", "Bain", None, None, None, None, None],
        ["US e-commerce sales agent-driven, 2030", "10-20%", "Morgan Stanley", None, None, None, None, None],
        ["US online sales via agentic commerce, 2030", "up to 25%", "J.P. Morgan", None, None, None, None, None],
        ["Digital commerce transactions via AI platforms, 2030", "20%", "Gartner", None, None, None, None, None],
        ["Global agentic commerce GMV, 2030", "$3-5T", "McKinsey", None, None, None, None, None],
        ["B2B purchases intermediated by AI agents, 2028", "90% (>$15T)", "Gartner", None, None, None, None, None],
        ["AI inference share of total network traffic, 2035", "25%", "Cisco", None, None, None, None, None],
        ["Consumer network traffic multiple by 2035 (with AI)", "6.6x (vs ~4x without)", "Cisco", None, None, None, None, None],
        ["AI agents market size, 2030", "$52.6B (46.3% CAGR)", "MarketsandMarkets", None, None, None, None, None],
        ["Enterprise agentic AI market, 2030", "$24.5B (46.2% CAGR)", "Grand View Research", None, None, None, None, None],
    ],
)
ws.freeze_panes = "A2"

# -------------------------------------------------- 3. Broadband Pricing ----

ws = wb.create_sheet("Broadband Pricing")
ws.sheet_properties.tabColor = "70AD47"
set_widths(ws, [48, 16, 16, 16, 16, 16, 16])

r = title_block(
    ws, 1,
    "Broadband: Traffic Economics & Pricing Model Disruption",
    "How agentic AI changes ISP demand, cost and pricing structures",
)

r = section(ws, r, "Impact assessment", 7)
r = table(
    ws, r,
    ["Dimension", "Impact", "", "", "", "", ""],
    [
        ["Traffic volume",
         "AI agents generate up to 450% more traffic per task than human-paced use (Cisco). Consumer traffic grows "
         "6.6x by 2035 with agentic AI vs ~4x without - 63% incremental growth. Enterprise traffic: 9x with agentic "
         "AI vs 2.5x baseline.", None, None, None, None, None],
        ["Traffic shape",
         "Inference flows are longer-lived, token-streamed (median flow rate ~10x lower than web), upstream-heavy, "
         "and latency-sensitive. Networks optimized for bursty downstream video are mismatched; uplink capacity and "
         "tail-latency become the binding constraints.", None, None, None, None, None],
        ["Pricing model",
         "Flat monthly speed tiers underprice agentic households and can't monetize quality guarantees. Emerging "
         "models: usage-based agent traffic classes, on-demand acceleration (buy 2 hours of premium QoS for an "
         "agent), guaranteed-latency SLAs, per-agent authentication and billing via network APIs (GSMA Open "
         "Gateway; IETF agent-acceleration framework draft).", None, None, None, None, None],
        ["New customer type",
         "AI agents become direct consumers of network services - continuously purchasing capability (latency, "
         "uplink, priority) via APIs rather than humans buying plans annually. Operators gain a machine-to-machine "
         "revenue surface but must build agent identity, metering and billing rails.", None, None, None, None, None],
        ["Cost & operations",
         "Agentic AI also cuts ISP opex: proactive fault resolution, churn prediction, billing dispute automation. "
         "Data quality is the gating factor (up to ~40% of operator data has inconsistencies).", None, None, None, None, None],
    ],
)
for rr in range(r - 5, r):
    ws.row_dimensions[rr].height = 72

r = section(ws, r, "Forecast: consumer traffic index (2025 = 100)", 7)
tr_hdr = r
no_ai = [100]
with_ai = [100]
for _ in range(5):
    no_ai.append(no_ai[-1] * 1.15)     # ~4.05x by 2035 => ~15% CAGR
    with_ai.append(with_ai[-1] * 1.208)  # 6.6x by 2035 => ~20.8% CAGR
r = table(
    ws, r,
    ["Scenario"] + YEARS,
    [
        ["Without agentic AI (~15%/yr, path to ~4x by 2035)"] + [round(v) for v in no_ai],
        ["With agentic AI (~21%/yr, path to 6.6x by 2035 - Cisco)"] + [round(v) for v in with_ai],
        ["AI inference share of consumer traffic (base case)"] + SHARE_BASE[
            "Broadband (AI inference share of consumer network traffic)"],
    ],
    formats={i: "0" for i in range(1, 7)},
)
for c in range(2, 8):
    ws.cell(row=tr_hdr + 3, column=c).number_format = PCT

r = section(ws, r, "Forecast: pricing evolution (US consumer broadband, base case)", 7)
r = table(
    ws, r,
    ["Metric"] + YEARS,
    [
        ["Share of ISP consumer revenue from usage/event/capability-based plans",
         0.005, 0.01, 0.03, 0.05, 0.08, 0.10],
        ["Monthly ARPU premium on agentic/priority tiers ($, vs standard plan)",
         0, 1, 2, 4, 6, 8],
        ["Households on plans with agent-aware QoS features (share)",
         0.00, 0.01, 0.04, 0.09, 0.16, 0.25],
    ],
    formats={i: PCT for i in range(1, 7)},
)
for c in range(2, 8):
    ws.cell(row=r - 3, column=c).number_format = "$0"
ws.row_dimensions[r - 4].height = 30
ws.row_dimensions[r - 3].height = 30
ws.row_dimensions[r - 2].height = 30

r = note(ws, r,
         "Scenario logic: near-term (2025-2027) broadband pricing is largely unchanged - agentic traffic is small and "
         "absorbed by existing capacity. Inflection begins 2027-2029 as agent adoption compounds (Cisco models ~25% "
         "CAGR in AI inference traffic 2029-2032), pushing ISPs toward differentiated tiers. Bear case: regulation or "
         "slow agent adoption keeps flat-rate pricing dominant (premium revenue <5% by 2030). Bull case: agent-driven "
         "QoS purchasing via network APIs scales to ~15% of consumer revenue by 2030.", 7)
ws.freeze_panes = "A2"

# -------------------------------------------------- 4. E-commerce Traffic ---

ws = wb.create_sheet("E-commerce Traffic")
ws.sheet_properties.tabColor = "ED7D31"
set_widths(ws, [48, 16, 16, 16, 16, 16, 16])

r = title_block(
    ws, 1,
    "E-commerce: Traffic Mix Shift & Agentic Commerce",
    "AI-referred traffic is the fastest-growing, highest-converting channel in retail",
)

r = section(ws, r, "Current impact: published data points", 7)
r = table(
    ws, r,
    ["Metric", "Value", "Source / period", "", "", "", ""],
    [
        ["AI-referred traffic growth to US retail sites (YoY)", "+138% (May 2026); +62% (Jul 2026)", "Adobe Analytics", None, None, None, None],
        ["AI-referred traffic growth since Oct 2024", "+1,219% to +1,324% (>14x)", "Adobe Analytics", None, None, None, None],
        ["AI-referred session conversion vs non-AI traffic", "+54% (May 2026); +60% (Jul 2026)", "Adobe Analytics", None, None, None, None],
        ["Revenue per visit, AI-referred vs other", "+53%", "Adobe Analytics (Jul 2026)", None, None, None, None],
        ["Engagement: time on site / pages / bounce", "+59% time, +23% pages, -33% bounce", "Adobe Analytics", None, None, None, None],
        ["AI-referred sessions & orders growth (YoY, Q1 2026)", "Sessions >8x; orders ~13x", "Shopify", None, None, None, None],
        ["AI-referred conversion & AOV vs organic search", "Conversion +~50%; AOV +14%", "Shopify (Q1 2026)", None, None, None, None],
        ["US retail e-commerce sales via AI platforms, 2026", "$20.6B (~1.5% of total)", "EMARKETER (Dec 2025)", None, None, None, None],
        ["Cyber Week 2025 AI-influenced sales", "$67B of $336.6B (~20% of orders)", "Salesforce", None, None, None, None],
        ["AI-driven sessions share of total e-commerce traffic", "<0.2% (2025), fastest-growing channel", "commercetools synthesis", None, None, None, None],
    ],
)

r = section(ws, r, "Forecast: US e-commerce agentic penetration (2025-2030)", 7)
ec_hdr = r
r = table(
    ws, r,
    ["Metric"] + YEARS,
    [
        ["AI/agent-driven share of US e-commerce sales - BEAR", 0.002, 0.010, 0.02, 0.04, 0.06, 0.08],
        ["AI/agent-driven share of US e-commerce sales - BASE", 0.004, 0.015, 0.035, 0.07, 0.11, 0.15],
        ["AI/agent-driven share of US e-commerce sales - BULL", 0.006, 0.020, 0.06, 0.12, 0.18, 0.25],
        ["AI-referred share of US retail site traffic (base)", 0.002, 0.008, 0.02, 0.05, 0.09, 0.14],
        ["AI-referred conversion premium vs non-AI traffic (base)", 0.54, 0.60, 0.55, 0.45, 0.35, 0.25],
    ],
    formats={i: PCT for i in range(1, 7)},
)
for rr in range(ec_hdr + 1, ec_hdr + 6):
    ws.row_dimensions[rr].height = 30

r = section(ws, r, "Forecast: agentic commerce GMV anchors", 7)
r = table(
    ws, r,
    ["Metric", 2026, 2027, 2028, 2029, 2030, "Source"],
    [
        ["Global agent-to-agent commerce GMV ($B, base)", 153, 574, 1567, 2796, 4000, "IHL Group"],
        ["US agentic commerce GMV 2030 ($B, range)", None, None, None, None, "300-500", "Bain"],
        ["Global agentic commerce GMV 2030 ($T, range)", None, None, None, None, "3-5", "McKinsey"],
    ],
    formats={i: "#,##0" for i in range(1, 6)},
)

r = note(ws, r,
         "Scenario logic: base case interpolates from EMARKETER's 1.5% (2026) to ~15% by 2030 - the low end of Bain "
         "(15-25%) and midpoint of Morgan Stanley (10-20%). The conversion premium narrows over time as AI-referred "
         "traffic scales from early-adopter, high-intent users to the mainstream. Traffic-mix shift is the leading "
         "indicator: agents start >50% of sessions on product pages vs 20% for organic search (Shopify), collapsing "
         "the upper funnel. Bear case assumes trust/payment friction persists (autonomous-checkout trust is only "
         "~14% in the US today); bull case assumes agent checkout protocols (UCP, AP2) reach mainstream by 2028.", 7)
ws.freeze_panes = "A2"

# ------------------------------------------------------------ 5. Delivery ---

ws = wb.create_sheet("Delivery")
ws.sheet_properties.tabColor = "FFC000"
set_widths(ws, [48, 16, 16, 16, 16, 16, 16])

r = title_block(
    ws, 1,
    "Delivery: Agentic Ordering & Autonomous Last-Mile",
    "Food, grocery and parcel delivery - demand-side agents plus supply-side automation",
)

r = section(ws, r, "Current impact: published data points", 7)
r = table(
    ws, r,
    ["Metric / event", "Value / detail", "Source / period", "", "", "", ""],
    [
        ["Gemini agentic ordering beta (DoorDash, Uber, Grubhub)", "Multistep background ordering; first consumer agentic category", "Google / Retail TouchPoints (Mar 2026)", None, None, None, None],
        ["Uber Eats AI cart assistant (US launch)", "Auto-builds grocery carts from text, photos, handwritten lists", "Uber (Feb 2026)", None, None, None, None],
        ["Instacart AI cart assistant", "Launched Nov 2024 - grocery cart automation industry trend", "Instacart", None, None, None, None],
        ["Autonomous last-mile delivery market, 2024", "$1.6B (GVR) - $2.3B (GIA)", "Grand View Research / Global Industry Analysts", None, None, None, None],
        ["Autonomous last-mile delivery market, 2030", "$4.2B (MarketsandMarkets) - $6.2B (GIA)", "22.7% / 24.8% / 17.9% CAGR estimates", None, None, None, None],
        ["Agentic AI in logistics market, 2035", "$46.2B (17.5% CAGR)", "market synthesis via MarketsandMarkets insight", None, None, None, None],
    ],
)
for rr in range(r - 6, r):
    ws.row_dimensions[rr].height = 32

r = section(ws, r, "Forecast: agentic + autonomous delivery (2025-2030)", 7)
dl_hdr = r
r = table(
    ws, r,
    ["Metric"] + YEARS,
    [
        ["Share of food/grocery delivery orders initiated by AI agents - BEAR", 0.0005, 0.005, 0.02, 0.04, 0.07, 0.10],
        ["Share of food/grocery delivery orders initiated by AI agents - BASE", 0.001, 0.01, 0.04, 0.08, 0.14, 0.20],
        ["Share of food/grocery delivery orders initiated by AI agents - BULL", 0.002, 0.02, 0.06, 0.13, 0.21, 0.30],
        ["Autonomous last-mile delivery market, $B (blended base)", 2.0, 2.5, 3.1, 3.9, 4.8, 5.9],
        ["US deliveries fulfilled by autonomous vehicles/bots/drones (share, base)", 0.002, 0.005, 0.01, 0.02, 0.035, 0.05],
    ],
    formats={i: PCT for i in range(1, 7)},
)
for c in range(2, 8):
    ws.cell(row=dl_hdr + 4, column=c).number_format = NUM
for rr in range(dl_hdr + 1, dl_hdr + 6):
    ws.row_dimensions[rr].height = 30

r = note(ws, r,
         "Scenario logic: delivery is the beachhead for consumer agentic commerce - orders are recurring, low-risk "
         "and low-consideration (J.P. Morgan expects agentic volume to concentrate in exactly these categories), and "
         "the Gemini partnerships put agent ordering in front of mainstream Android users from 2026. Demand-side "
         "agent share therefore ramps fastest of the five sectors. Supply side is slower: autonomous last-mile "
         "market forecasts (blended from Grand View 24.8% CAGR and GIA 17.9% CAGR) imply low-single-digit share of "
         "actual deliveries by 2030, with campus/corridor deployments first. Margin impact: agent-routed orders "
         "commoditize the storefront, but platforms with dense fulfillment networks (DoorDash) capture agent order "
         "flow as infrastructure providers.", 7)
ws.freeze_panes = "A2"

# --------------------------------------------------------- 6. Ridesharing ---

ws = wb.create_sheet("Ridesharing")
ws.sheet_properties.tabColor = "7030A0"
set_widths(ws, [48, 16, 16, 16, 16, 16, 16])

r = title_block(
    ws, 1,
    "Ridesharing: Agent Booking & the Autonomous Vehicle Transition",
    "Near-term agent-initiated bookings; medium-term AV-driven cost restructuring",
)

r = section(ws, r, "Current impact: published data points", 7)
r = table(
    ws, r,
    ["Metric / event", "Value / detail", "Source / period", "", "", "", ""],
    [
        ["Gemini agentic ride booking (Uber among first partners)", "Book a ride via voice/single prompt; multistep background execution", "Google / Samsung (2026)", None, None, None, None],
        ["Uber gross bookings, 2025", "$193.4B (record)", "Uber FY2025", None, None, None, None],
        ["Uber revenue / adjusted profit, 2025", "$52B / $5.2B", "Uber FY2025", None, None, None, None],
        ["Driver earnings pool (the AV prize), 2025", "$85.4B paid to ~9.7M drivers", "Uber FY2025", None, None, None, None],
        ["Uber AV partnerships", "20+ self-driving firms incl. Waymo", "Uber / IndexBox", None, None, None, None],
        ["Monthly active platform users", ">200M", "Uber", None, None, None, None],
        ["Competitive threat", "Waymo commercial robotaxis live in multiple US cities; Tesla, WeRide, Pony.ai expanding", "PitchGrade analysis", None, None, None, None],
    ],
)
for rr in range(r - 7, r):
    ws.row_dimensions[rr].height = 30

r = section(ws, r, "Forecast: agentic + autonomous ridesharing (2025-2030)", 7)
rs_hdr = r
r = table(
    ws, r,
    ["Metric"] + YEARS,
    [
        ["Share of rideshare bookings initiated by AI agents - BEAR", 0.0002, 0.002, 0.01, 0.025, 0.04, 0.06],
        ["Share of rideshare bookings initiated by AI agents - BASE", 0.0005, 0.005, 0.02, 0.05, 0.08, 0.12],
        ["Share of rideshare bookings initiated by AI agents - BULL", 0.001, 0.01, 0.04, 0.08, 0.13, 0.20],
        ["AV share of US rideshare trips (base)", 0.005, 0.01, 0.025, 0.05, 0.08, 0.12],
        ["Platform take-rate direction on agent/AV bookings (vs ~28-30% today)",
         None, None, None, None, None, None],
    ],
    formats={i: PCT for i in range(1, 7)},
)
ws.cell(row=rs_hdr + 5, column=2,
        value="Take rates compress toward 15-25% on AV-partner trips as AV operators (Waymo et al.) capture fleet "
              "margin; agent platforms (Google/OpenAI) may add a 1-3% referral layer on top.")
ws.merge_cells(start_row=rs_hdr + 5, start_column=2, end_row=rs_hdr + 5, end_column=7)
ws.cell(row=rs_hdr + 5, column=2).alignment = WRAP
ws.row_dimensions[rs_hdr + 5].height = 44
for rr in range(rs_hdr + 1, rs_hdr + 5):
    ws.row_dimensions[rr].height = 30

r = note(ws, r,
         "Scenario logic: agent booking of rides lags food delivery (a ride request is time- and location-critical, "
         "so users confirm manually longer), reaching ~12% of bookings by 2030 in the base case. The larger economic "
         "story is autonomous vehicles: the $85.4B driver-payout pool is the margin prize. Base case has AVs at "
         "~12% of US trips by 2030, concentrated in ~20-30 metros. Risk to marketplaces: vertically integrated AV "
         "operators bypass Uber/Lyft; Uber's counter is aggregating 20+ AV fleets. Pricing gets more dynamic - agents "
         "comparison-shop across Uber/Lyft/Waymo in real time, compressing surge premiums on contested routes.", 7)
ws.freeze_panes = "A2"

# --------------------------------------------------------- 7. Online Travel -

ws = wb.create_sheet("Online Travel")
ws.sheet_properties.tabColor = "C00000"
set_widths(ws, [48, 16, 16, 16, 16, 16, 16])

r = title_block(
    ws, 1,
    "Online Travel Marketplaces: Agent-Mediated Booking & Take-Rate Pressure",
    "OTAs face margin compression, not extinction, as agents become the discovery layer",
)

r = section(ws, r, "Current impact: published data points", 7)
r = table(
    ws, r,
    ["Metric / event", "Value / detail", "Source / period", "", "", "", ""],
    [
        ["Google agentic hotel booking - founding partners", "Booking.com, Expedia, Marriott, IHG, Choice, Wyndham (routed through, not around, OTAs)", "Hospitality Net (2026)", None, None, None, None],
        ["Booking Holdings revenue (annual)", "~$21B", "Booking Holdings", None, None, None, None],
        ["Expedia Group revenue (annual)", "~$13.7B", "Expedia Group", None, None, None, None],
        ["Current OTA take rates", "10-20% commissions (hotels pay 15-25%)", "PitchGrade / industry", None, None, None, None],
        ["Bernstein assessment", "'Paradigm shift' - AI threatens take rates, marketing efficiency and the supply moat; Market-Perform on BKNG & EXPE, terminal-value risk key", "Bernstein (2026)", None, None, None, None],
        ["OTA AI counter-moves", "Priceline 'Penny', Booking.com booking agent, 'Lola' venture (Kayak founders); Expedia acquired AI-native Layla (2026)", "Hospitality Net / Business Times", None, None, None, None],
        ["Expedia bear case (analyst)", "30-40% of bookings bypass OTA; EBITDA margin compresses toward 8-14%", "PitchGrade (score 8/10 AI margin pressure)", None, None, None, None],
        ["Precedent: Google pressure on Booking", "Merchant-model shift cost >3pts of EBITDA margin vs 2015", "Bernstein", None, None, None, None],
    ],
)
for rr in range(r - 8, r):
    ws.row_dimensions[rr].height = 34

r = section(ws, r, "Forecast: agent-mediated travel booking (2025-2030)", 7)
tv_hdr = r
r = table(
    ws, r,
    ["Metric"] + YEARS,
    [
        ["Share of online travel bookings mediated by AI agents - BEAR", 0.001, 0.005, 0.015, 0.03, 0.055, 0.08],
        ["Share of online travel bookings mediated by AI agents - BASE", 0.002, 0.01, 0.03, 0.06, 0.10, 0.15],
        ["Share of online travel bookings mediated by AI agents - BULL", 0.004, 0.02, 0.05, 0.10, 0.17, 0.25],
        ["Effective take rate on agent-mediated bookings (base)", 0.15, 0.145, 0.135, 0.125, 0.115, 0.10],
        ["OTA performance-marketing spend shifting to AI-platform placement (base)", 0.00, 0.02, 0.05, 0.10, 0.16, 0.22],
    ],
    formats={i: PCT for i in range(1, 7)},
)
for rr in range(tv_hdr + 1, tv_hdr + 6):
    ws.row_dimensions[rr].height = 30

r = note(ws, r,
         "Scenario logic: travel is structurally harder for agents than retail (volatile inventory/pricing, complex "
         "multi-leg itineraries, payments and post-booking service), so agent-mediated share ramps slower than "
         "e-commerce - ~15% by 2030 base case. Near term, agents route through OTA inventory and APIs (Google's "
         "partner list proves it), so volume holds while economics erode: take rates on agent-mediated bookings "
         "drift from ~15% toward ~10% as AI platforms extract placement/referral fees and suppliers push direct "
         "connections. Bear case for OTAs is disintermediation via supplier-direct agent APIs (30-40% booking "
         "bypass); bull case for volume is agents expanding total bookings by lowering planning friction. Watch "
         "items: Google AI Overview click-through erosion; standardized travel-agent protocols; OTA in-house agent "
         "adoption (Penny, Layla).", 7)
ws.freeze_panes = "A2"

# ------------------------------------------------- 8. Sources & Methodology -

ws = wb.create_sheet("Sources & Methodology")
ws.sheet_properties.tabColor = "808080"
set_widths(ws, [34, 74, 52])

r = title_block(ws, 1, "Sources & Methodology", "All external figures trace to the sources below; forecasts are scenario models built on these anchors")

r = section(ws, r, "Methodology", 3)
method = (
    "1) Published data points (Adobe, Shopify, Cisco, EMARKETER, Salesforce, IHL, Bain, McKinsey, Morgan Stanley, "
    "J.P. Morgan, Gartner, Bernstein, company filings) are reported as-is with sources. "
    "2) Forecast rows labeled BEAR/BASE/BULL are this report's scenario interpolations: each base case is anchored "
    "to a published 2026 actual (e.g., EMARKETER's 1.5% AI-platform share of US e-commerce) and a published 2030 "
    "range (e.g., Bain 15-25%), with an S-curve ramp between them; bear/bull bracket the published analyst ranges. "
    "3) Sector adoption ordering follows J.P. Morgan's finding that recurring, low-risk categories (groceries, "
    "meals, rides) adopt agentic purchasing first, and high-consideration categories (travel) later. "
    "4) All dollar figures nominal USD. Compiled September 2026. These are planning scenarios, not investment advice."
)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
cell = ws.cell(row=r, column=1, value=method)
cell.alignment = WRAP
ws.row_dimensions[r].height = 120
r += 2

r = section(ws, r, "Sources", 3)
r = table(
    ws, r,
    ["Publisher / analyst", "What it provides", "URL"],
    [
        ["Cisco - AI Impact on Wide Area Networks (2026)", "Traffic multiples (6.6x consumer / 9x enterprise by 2035), 450% agent traffic per task, 25% inference share by 2035, flow-shape analysis", "https://www.cisco.com/c/dam/en/us/solutions/collateral/artificial-intelligence/mass-scale-infrastructure/ai-network-traffic-report.pdf"],
        ["Network World on Cisco study", "AI inference traffic ~25% CAGR 2029-2032; flow-rate comparisons", "https://www.networkworld.com/article/4175890/cisco-ai-traffic-is-radically-reshaping-wans.html"],
        ["Business Standard - AI economy & telecom networks", "Operator plans for agent-consumable network APIs (GSMA Open Gateway); uplink/latency requirements; FWA share", "https://www.business-standard.com/technology/artificial-intelligence/ai-agents-telecom-networks-uplink-traffic-5g-broadband-explained-126072400842_1.html"],
        ["IETF draft-li-opsawg-agent-accel-framework", "Framework for dynamic, on-demand broadband acceleration purchased per-agent (time/volume/event-based)", "https://datatracker.ietf.org/doc/html/draft-li-opsawg-agent-accel-framework-00"],
        ["Telecompetitor - Alcott interview", "ISP operational agentic AI use cases; data-quality constraint (~40% inconsistency)", "https://www.telecompetitor.com/scott-alcotts-agentic-ai-framework-for-broadband-providers-interview/"],
        ["Adobe Analytics via Digital Commerce 360 (Jun & Aug 2026)", "AI-referred retail traffic growth (+138% YoY May; +1,219% vs Oct 2024), conversion +54-60%, revenue/visit +53%, engagement metrics", "https://www.digitalcommerce360.com/2026/06/17/adobe-ai-referred-traffic-to-retail-sites-doubles-in-a-year/"],
        ["Shopify - AI search insights (Q1 2026)", "AI-referred sessions 8x / orders ~13x YoY; conversion +~50%; AOV +14%; >50% of AI sessions start on product pages", "https://www.shopify.com/enterprise/blog/ai-search-insights"],
        ["Axis Intelligence - Agentic commerce statistics 2026", "EMARKETER $20.6B / 1.5% US 2026; Salesforce Cyber Week $67B AI-influenced; Gartner 90% B2B by 2028 (>$15T); US autonomous-checkout trust 14%", "https://axis-intelligence.com/agentic-commerce-statistics/"],
        ["commercetools - Agentic commerce stats guide", "Bain $300-500B US by 2030 (15-25%); Gartner 20% of digital commerce by 2030; Morgan Stanley 10-20%; McKinsey $1T US / $3-5T global; J.P. Morgan up to 25%", "https://commercetools.com/blog/agentic-commerce-stats-enterprise-guide"],
        ["IHL Group - Agent-to-Agent Commerce Forecast 2030", "A2A commerce GMV: $4B (2024) to $4,000B (2030) base; bear $2.2T / bull $4.4T", "https://www.ihlservices.com/product/agent-to-agent-commerce-forecast/"],
        ["Stellagent / Retail TouchPoints (Mar 2026)", "Gemini agentic ordering beta with DoorDash, Uber, Grubhub; Uber Eats AI cart assistant; Instacart assistant", "https://stellagent.ai/insights/doordash-uber-gemini-agentic-ordering"],
        ["PYMNTS - Gemini agentic features", "Gemini multistep background tasks for food, grocery, rideshare; Pixel 10 / Galaxy S26 rollout", "https://www.pymnts.com/artificial-intelligence-2/2026/gemini-ai-will-soon-order-groceries-and-book-rides/"],
        ["MarketsandMarkets - Autonomous last-mile delivery", "$0.9B (2023) to $4.2B (2030), 22.7% CAGR; AI-impact phasing analysis", "https://www.marketsandmarkets.com/Market-Reports/autonomous-last-mile-delivery-market-41240862.html"],
        ["Grand View Research - Autonomous last-mile delivery", "$1.62B (2024) to $5.93B (2030), 24.8% CAGR", "https://www.grandviewresearch.com/industry-analysis/autonomous-last-mile-delivery-market"],
        ["Global Industry Analysts", "$2.3B (2024) to $6.2B (2030), 17.9% CAGR; ground vs aerial segment split", "https://www.marketresearch.com/Global-Industry-Analysts-v1039/Autonomous-Last-Mile-Delivery-38748689/"],
        ["IndexBox - Uber & DoorDash strategic analysis", "Uber FY2025: $193.4B gross bookings, $52B revenue, $5.2B adj. profit, $85.4B driver earnings, 9.7M drivers, 20+ AV deals", "https://www.indexbox.io/blog/uber-and-doordash-strategic-growth-analysis-autonomous-future/"],
        ["PitchGrade - Uber AI margin pressure", "AV threat assessment; Waymo/Tesla/WeRide/Pony.ai landscape; platform-durability question", "https://pitchgrade.com/research/uber-ai-margin-pressure"],
        ["Bernstein via Investing.com/Yahoo Finance", "OTA 'paradigm shift'; take-rate & marketing-efficiency pressure; Booking merchant-model margin precedent (-3pts vs 2015)", "https://finance.yahoo.com/news/booking-expedia-face-margin-pressure-072422511.html"],
        ["PitchGrade - Expedia & Booking AI margin pressure", "Expedia 8/10 AI-pressure score; bear case 30-40% booking bypass, EBITDA to 8-14%; Booking $21B revenue context", "https://pitchgrade.com/research/expedia-ai-margin-pressure"],
        ["Hospitality Net - 'Nobody gets bypassed'", "Google agentic hotel booking partners (Booking, Expedia, Marriott, IHG, Choice, Wyndham); OTA agent strategies (Penny, Lola)", "https://www.hospitalitynet.org/opinion/4132643/nobody-gets-bypassed-agentic-commerce-was-supposed-to-route-around-the-otas-googles-partner-list-routes-straight-through-them"],
        ["The Business Times - travel as AI's final frontier", "Travel-specific agent friction (real-time inventory/pricing); Expedia-Layla acquisition; incumbents' pacing advantage", "https://www.businesstimes.com.sg/startups-tech/technology/last-board-why-travel-may-be-ais-final-frontier"],
        ["MarketsandMarkets - AI agents market", "$7.84B (2025) to $52.62B (2030), 46.3% CAGR", "https://www.marketsandmarkets.com/Market-Reports/ai-agents-market-15761548.html"],
        ["Grand View Research - Enterprise agentic AI", "$3.67B (2025) to $24.5B (2030), 46.2% CAGR", "https://www.grandviewresearch.com/industry-analysis/enterprise-agentic-ai-market-report"],
    ],
)
for rr in range(r - 23, r):
    ws.row_dimensions[rr].height = 42
ws.freeze_panes = "A2"

# -------------------------------------------------------------------- save --

out = "Agentic_AI_Market_Impact_Report.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"Sheets: {wb.sheetnames}")
