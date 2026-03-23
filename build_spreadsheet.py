#!/usr/bin/env python3
"""
Build a comprehensive Excel spreadsheet aggregating hyperscale capex data
and the impact of a protracted Iran conflict on sentiment and operations.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

wb = openpyxl.Workbook()

# ── Shared style definitions ────────────────────────────────────────────────
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
SECTION_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SECTION_FONT = Font(name="Calibri", bold=True, size=11, color="1F4E79")
RISK_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RISK_AMBER = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RISK_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=12, color="1F4E79")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

USD_FMT = '#,##0'
PCT_FMT = '0.0%'


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def style_data_area(ws, start_row, end_row, max_col):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def auto_width(ws, max_col, max_width=45):
    for col in range(1, max_col + 1):
        best = 12
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=False):
            for cell in row:
                if cell.value:
                    best = max(best, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[get_column_letter(col)].width = best


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 1 – Hyperscaler CapEx Overview
# ═══════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Hyperscaler CapEx Overview"
ws1.sheet_properties.tabColor = "1F4E79"

ws1.merge_cells("A1:I1")
ws1["A1"] = "Hyperscaler Capital Expenditure Overview (2024-2026)"
ws1["A1"].font = TITLE_FONT
ws1["A1"].alignment = Alignment(horizontal="center")

ws1.merge_cells("A2:I2")
ws1["A2"] = "Source: Company earnings reports, analyst estimates | Prepared: March 2026"
ws1["A2"].font = Font(italic=True, size=9, color="666666")
ws1["A2"].alignment = Alignment(horizontal="center")

headers1 = [
    "Company", "2024 CapEx ($B)", "2025 CapEx ($B)",
    "2026 CapEx Guidance ($B)", "YoY Growth\n(2025→2026)",
    "Capital Intensity\n(% of Revenue)",
    "Primary Focus", "Key Constraint", "Middle East Exposure"
]
for c, h in enumerate(headers1, 1):
    ws1.cell(row=4, column=c, value=h)
style_header_row(ws1, 4, len(headers1))

data1 = [
    ["Amazon (AWS)", 83, 200, 200, "≈0% (fully guided)", "~50%",
     "AWS data centers; AI compute", "$244B backlog; supply-constrained",
     "UAE, Bahrain facilities struck by Iran (Mar 1 2026); 2 of 3 UAE AZs offline"],
    ["Microsoft (Azure)", 56, 96, "120+", "~25%", "~57%",
     "Azure cloud; AI infrastructure", "$80B unfulfilled orders (power shortage); GPUs idle",
     "$80B Saudi Arabia commitment; $15B UAE; 3 data centers in Saudi Eastern Province (Q4 2026)"],
    ["Alphabet (Google Cloud)", 52.5, 91.4, "175-185", "~100%", "~45%",
     "Google DeepMind; Cloud AI; TPUs", "Supply-constrained all 2026 per CEO",
     "$10B AI partnership with Saudi PIF (HUMAIN)"],
    ["Meta", 39, 72.2, "115-135", "~73%", "~48%",
     "AI training; Meta Compute division", "Cash-flow pressure; workforce cuts to offset",
     "Lower direct Middle East exposure; primarily US-focused DC buildout"],
    ["Oracle", 12, 25, "~50", "~100%", "~45%",
     "Cloud infrastructure; Stargate (UAE)", "$50B debt raised; $2.1B restructuring",
     "$14B UAE/Saudi commitment; Stargate campus (5GW, 10 sq mi) in UAE"],
]

for r, row_data in enumerate(data1, 5):
    for c, val in enumerate(row_data, 1):
        ws1.cell(row=r, column=c, value=val)

ws1.cell(row=10, column=1, value="TOTAL (Big Five)")
ws1.cell(row=10, column=1).font = Font(bold=True)
ws1.cell(row=10, column=2, value=242.5)
ws1.cell(row=10, column=3, value=484.6)
ws1.cell(row=10, column=4, value="660-690")
ws1.cell(row=10, column=5, value="~42%")
ws1.cell(row=10, column=6, value="45-57%")
ws1.cell(row=10, column=7, value="AI infrastructure (~75% of total)")
ws1.cell(row=10, column=8, value="Power supply; $108B debt raised in 2025; more debt than cash")
ws1.cell(row=10, column=9, value="Combined >$100B committed to Gulf region")
for c in range(1, len(headers1) + 1):
    ws1.cell(row=10, column=c).font = Font(bold=True)
    ws1.cell(row=10, column=c).fill = SECTION_FILL

style_data_area(ws1, 5, 10, len(headers1))
auto_width(ws1, len(headers1))

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 2 – Iran Conflict Impact Assessment
# ═══════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Iran Conflict Impact")
ws2.sheet_properties.tabColor = "C00000"

ws2.merge_cells("A1:G1")
ws2["A1"] = "Iran Conflict Impact on Hyperscale CapEx & Operations"
ws2["A1"].font = TITLE_FONT
ws2["A1"].alignment = Alignment(horizontal="center")

ws2.merge_cells("A2:G2")
ws2["A2"] = "Conflict began: US-Israel strikes on Iran (Operation Epic Fury) — Iran retaliated Mar 1, 2026"
ws2["A2"].font = Font(italic=True, size=9, color="666666")
ws2["A2"].alignment = Alignment(horizontal="center")

headers2 = [
    "Impact Category", "Risk Level", "Description",
    "Quantified Impact", "Affected Hyperscalers",
    "Mitigation / Response", "Protracted-Conflict Scenario (6-12 mo)"
]
for c, h in enumerate(headers2, 1):
    ws2.cell(row=4, column=c, value=h)
style_header_row(ws2, 4, len(headers2))

impacts = [
    # Direct infrastructure damage
    ["Direct Infrastructure\nDamage",
     "HIGH",
     "Iranian Shahed-136 drone strikes hit 3 AWS data centers in UAE & Bahrain on Mar 1 2026. "
     "2 of 3 UAE availability zones taken offline. Structural damage, fires, power disruption.",
     "$800M-$1B per facility; banks/payment processors disrupted",
     "Amazon (AWS) — directly hit;\nAll hyperscalers with Gulf presence at risk",
     "Hyperscalers can absorb single-facility losses financially; exploring missile defense; "
     "geographic diversification",
     "Repeated strikes could destroy $5-10B in infrastructure; "
     "region may become uninsurable for DC operations"],

    # Energy cost surge
    ["Energy Cost Surge",
     "HIGH",
     "Strait of Hormuz partially blocked; 10M bbl/day removed from global markets. "
     "Brent crude spiked from ~$70 to >$110/bbl (+57%). European gas prices up 40-50%. "
     "Energy = ~60% of DC operating costs.",
     "US electricity prices already up 6.9% in 2025; gas-fired DC power costs could double; "
     "operating margins pressured",
     "All hyperscalers — especially those reliant on gas-fired generation",
     "Some hedging in place; shift to renewables/nuclear long-term; "
     "short-term cost absorption hits margins",
     "Sustained $100+ oil could add $15-25B in annual energy costs across Big Five; "
     "operating margins compress 2-4pp; potential capex slowdown in H2 2026"],

    # Semiconductor supply chain
    ["Semiconductor &\nMaterial Supply\nDisruption",
     "HIGH",
     "Qatar produces 33% of global helium (critical for chip fab cooling/lithography, no substitute). "
     "Ras Laffan facility hit by drones. Taiwan holds only 11 days LNG reserves; "
     "37% of its grid power from Middle East energy. Bromine (97.5% from Dead Sea) also at risk.",
     "DDR5 memory & GPU prices rising; SK Hynix + Samsung lost >$200B market cap; "
     "25% of global helium supply at risk",
     "All hyperscalers (GPU/server procurement); TSMC/Samsung (chip fab)",
     "TSMC claims secured supply; US CHIPS Act diversifying fab capacity; "
     "but no near-term helium alternative",
     "Extended Hormuz closure could halt Taiwan chip production within 2-4 weeks; "
     "GPU delivery timelines extend 6-12 months; capex $ allocated but unspendable"],

    # Insurance & risk repricing
    ["Insurance &\nRisk Repricing",
     "MEDIUM-HIGH",
     "Data centers now classified as 'legitimate military targets.' "
     "Lloyd's warns major geopolitical conflict could cost global economy $14.5T over 5 years. "
     "War risk is 'harder to model than traditional perils.'",
     "Insurance premiums surging to record levels; shipping costs at records; "
     "150+ vessels delayed/rerouted",
     "All hyperscalers with Middle East infrastructure",
     "Some self-insurance; geographic diversification; "
     "facility hardening proposals",
     "Gulf DC facilities could become uninsurable; "
     "war-risk premiums could add 5-10% to regional operating costs"],

    # Geographic investment shift
    ["Geographic\nInvestment Shift",
     "MEDIUM",
     "Analysts expect 'shift in where next wave of capacity gets built.' "
     "Northern Europe, India, Southeast Asia gaining favor. "
     "Gulf had advantages: cheap energy, land, low latency to Asia.",
     ">$100B in committed Gulf investments at risk of delay/relocation",
     "All hyperscalers — especially Microsoft ($80B Saudi), "
     "Oracle ($14B UAE/Saudi), Google ($10B Saudi)",
     "Unlikely to abandon existing projects; future expansion shifted; "
     "accelerating US domestic buildout",
     "Gulf AI hub plans delayed 2-3 years; $30-50B in planned investment redirected; "
     "Northern Europe/India capacity brought forward"],

    # IT spending moderation
    ["Overall IT Spending\nModeration",
     "MEDIUM",
     "IDC forecasts global IT spending growth falls from 10% baseline to 9% "
     "in 3-month conflict scenario. Service providers expected to maintain AI investment "
     "despite pressures.",
     "~$50B reduction in global IT spending vs. baseline in downside case",
     "All hyperscalers (revenue growth); enterprise customers (budget pressure)",
     "AI demand remains structurally strong; hyperscalers unlikely to cut AI capex; "
     "may defer non-AI spending",
     "6-12 month conflict: IT spending growth could fall to 7-8%; "
     "non-AI cloud workloads see budget freezes; AI capex maintained but ROI scrutiny intensifies"],

    # Supply chain logistics
    ["Maritime & Logistics\nDisruption",
     "MEDIUM",
     "3,200 ships (4% global tonnage) idle in Persian Gulf; 500+ waiting outside. "
     "Rerouting around Africa adds ~19 days transit. Air cargo grounded from Middle East.",
     "$2-3B/week additional shipping costs globally; "
     "server/equipment delivery timelines extend",
     "All hyperscalers (equipment delivery); suppliers (component shipping)",
     "Diversifying supply routes; pre-positioning inventory; "
     "using air freight where possible",
     "Persistent rerouting adds $100-150B/year in global logistics costs; "
     "DC construction timelines extend 3-6 months"],

    # Debt & financial strain
    ["Financial Strain &\nDebt Burden",
     "MEDIUM",
     "Big Five collectively hold more debt than cash for first time. "
     "$108B debt issued in 2025. Capital intensity at 45-57% of revenue (vs. historical 40%). "
     "Nearly 100% of operating cash flow consumed by capex.",
     "Projected $1.5T total debt issuance needed over coming years; "
     "rising interest rates from inflation add $5-10B/year in debt service",
     "All hyperscalers — Meta and Oracle most leveraged relative to cash flow",
     "Meta: 20% workforce cuts + bonus/option reductions; "
     "Oracle: $2.1B restructuring; AI coding tools to reduce headcount",
     "Credit rating pressure if conflict persists; potential forced capex deferrals; "
     "some hyperscalers may raise equity or slow AI buildout pace"],
]

risk_fills = {"HIGH": RISK_RED, "MEDIUM-HIGH": RISK_AMBER, "MEDIUM": RISK_AMBER, "LOW": RISK_GREEN}

for r, row_data in enumerate(impacts, 5):
    for c, val in enumerate(row_data, 1):
        cell = ws2.cell(row=r, column=c, value=val)
    risk_cell = ws2.cell(row=r, column=2)
    risk_cell.alignment = CENTER
    risk_cell.font = Font(bold=True)
    risk_cell.fill = risk_fills.get(row_data[1], RISK_AMBER)

style_data_area(ws2, 5, 5 + len(impacts) - 1, len(headers2))
auto_width(ws2, len(headers2), max_width=50)
ws2.column_dimensions["C"].width = 55
ws2.column_dimensions["D"].width = 40
ws2.column_dimensions["F"].width = 40
ws2.column_dimensions["G"].width = 50

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 3 – Sentiment Tracker by Company
# ═══════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Sentiment by Company")
ws3.sheet_properties.tabColor = "2E75B6"

ws3.merge_cells("A1:H1")
ws3["A1"] = "Hyperscaler CapEx Sentiment Amid Protracted Iran Conflict"
ws3["A1"].font = TITLE_FONT
ws3["A1"].alignment = Alignment(horizontal="center")

headers3 = [
    "Company", "Pre-Conflict CapEx\nSentiment",
    "Current Sentiment\n(Mar 2026)",
    "CapEx Guidance\nChange",
    "Direct Conflict\nExposure",
    "Key Risk Factors",
    "Likely Response to\n6-12 Mo Conflict",
    "Analyst Consensus"
]
for c, h in enumerate(headers3, 1):
    ws3.cell(row=3, column=c, value=h)
style_header_row(ws3, 3, len(headers3))

sentiment_data = [
    ["Amazon (AWS)",
     "Very Bullish — record $200B commitment; $244B backlog",
     "Cautiously Bullish — maintaining guidance despite UAE facility strikes",
     "No change announced; $200B maintained",
     "HIGH — 3 facilities struck; 2 AZs offline in UAE",
     "Infrastructure vulnerability; Gulf expansion uncertainty; helium/energy supply",
     "Maintain domestic capex; defer new Gulf capacity; harden existing facilities; "
     "accelerate India/SE Asia buildout",
     "Overweight — AI demand structural; single-facility loss absorbable; "
     "geographic diversification underway"],

    ["Microsoft (Azure)",
     "Very Bullish — $120B+ guided; Azure demand exceeding supply",
     "Cautiously Bullish — $80B unfulfilled orders provide demand visibility",
     "No cut; power bottleneck more binding than conflict risk",
     "HIGH — $80B Saudi + $15B UAE commitments; 3 DCs planned in Saudi Eastern Province",
     "Power shortage already constraining growth; Gulf investments at physical risk; "
     "idle GPUs from power gaps",
     "Accelerate US/Europe power procurement; slow Saudi DC timeline; "
     "maintain AI capex priorities",
     "Overweight — demand far exceeds supply; conflict adds cost but doesn't reduce demand"],

    ["Alphabet (Google)",
     "Very Bullish — doubling capex to $175-185B; CEO says 'supply constrained all 2026'",
     "Bullish — less direct Gulf exposure than peers; demand intact",
     "No change; $175-185B guidance maintained",
     "MEDIUM — $10B Saudi AI partnership (HUMAIN); no facilities directly struck",
     "Helium supply for TPU/server procurement; energy cost inflation; "
     "Saudi project timeline risk",
     "Continue domestic buildout; potentially defer Saudi partnership milestones; "
     "TPU supply chain diversification",
     "Overweight — strongest balance sheet among peers; AI/Cloud growth >35%; "
     "least leveraged to Gulf risk"],

    ["Meta",
     "Bullish — $115-135B guided; Meta Compute division created; "
     "$600B US DC commitment by 2028",
     "Neutral-to-Cautious — workforce cuts signal margin pressure; "
     "limited Gulf exposure but energy costs rising",
     "No formal change but margin pressures mounting; $162-169B total expense guide",
     "LOW — primarily US-focused buildout; no known Gulf facilities struck",
     "Energy cost inflation; cash flow consumed by capex; Reality Labs losses ($6B/qtr); "
     "20% workforce reduction planned",
     "Continue AI capex but implement cost offsets (layoffs, bonus cuts); "
     "energy hedging; defer Reality Labs scaling",
     "Hold — AI capex justified but financial strain visible; "
     "workforce cuts help but don't resolve structural cash flow pressure"],

    ["Oracle",
     "Bullish — $50B capex guided; Stargate campus in UAE; "
     "OpenAI partnership",
     "Cautious — Stargate campus directly in conflict zone; "
     "$2.1B restructuring signals strain",
     "No formal change; $50B maintained but execution risk elevated",
     "VERY HIGH — Stargate UAE campus (5GW, 10 sq mi) in active conflict zone; "
     "$14B Gulf commitment",
     "Stargate construction timeline at risk; $50B debt load; "
     "most leveraged relative to revenue; restructuring accelerating",
     "Potentially defer UAE Stargate milestones; redirect some capacity to US/Japan; "
     "accelerate cost-cutting via AI coding tools",
     "Hold/Underweight — highest relative Gulf exposure; execution risk elevated; "
     "debt burden concerning"],
]

for r, row_data in enumerate(sentiment_data, 4):
    for c, val in enumerate(row_data, 1):
        ws3.cell(row=r, column=c, value=val)

style_data_area(ws3, 4, 4 + len(sentiment_data) - 1, len(headers3))
auto_width(ws3, len(headers3), max_width=50)
ws3.column_dimensions["B"].width = 40
ws3.column_dimensions["C"].width = 40
ws3.column_dimensions["F"].width = 45
ws3.column_dimensions["G"].width = 50
ws3.column_dimensions["H"].width = 45

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 4 – Scenario Analysis
# ═══════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Scenario Analysis")
ws4.sheet_properties.tabColor = "ED7D31"

ws4.merge_cells("A1:G1")
ws4["A1"] = "Protracted Iran Conflict — Scenario Analysis for Hyperscale CapEx"
ws4["A1"].font = TITLE_FONT
ws4["A1"].alignment = Alignment(horizontal="center")

headers4 = [
    "Scenario", "Duration", "Oil Price\n(Brent, $/bbl)",
    "Impact on Hyperscale\nCapEx ($B, vs baseline)",
    "IT Spending\nGrowth (Global)",
    "Semiconductor\nSupply Impact",
    "Key Assumptions & Implications"
]
for c, h in enumerate(headers4, 1):
    ws4.cell(row=3, column=c, value=h)
style_header_row(ws4, 3, len(headers4))

scenarios = [
    ["Baseline\n(No conflict escalation)",
     "Conflict resolved\n< 1 month",
     "$75-85",
     "$660-690B maintained\n(no material change)",
     "10%\n(baseline forecast)",
     "Minimal — temporary\nprice spikes only",
     "Quick ceasefire; Strait of Hormuz reopens; energy prices normalize; "
     "Gulf investments proceed on schedule; damaged AWS facilities rebuilt within 6 months"],

    ["Short Conflict",
     "1-3 months",
     "$90-110",
     "$630-660B\n(−$30-50B vs baseline;\nmostly Gulf deferrals)",
     "9%\n(IDC downside est.)",
     "Moderate — helium\nshortages begin;\nDDR5/GPU prices +10-15%",
     "Strait of Hormuz intermittently disrupted; Taiwan manages with reserves + rerouting; "
     "Gulf projects delayed 6-12 months; domestic capex maintained; "
     "energy cost increase of $5-10B across Big Five"],

    ["Protracted Conflict",
     "3-6 months",
     "$100-130",
     "$580-630B\n(−$60-100B vs baseline;\nGulf + energy-driven cuts)",
     "8%",
     "Severe — Taiwan LNG\nreserves depleted;\nchip production curtailed;\nGPU delivery +6mo",
     "Strait of Hormuz blockade sustained; Taiwan semiconductor output falls 15-25%; "
     "helium shortage forces fab slowdowns; Gulf projects suspended; "
     "hyperscalers redirect $30-50B to US/Europe/India; "
     "energy adds $15-25B annual cost; debt markets tighten"],

    ["Extended War",
     "6-12 months",
     "$120-150+",
     "$500-560B\n(−$100-150B vs baseline;\nbroad-based cuts)",
     "7%",
     "Critical — prolonged\nTaiwan energy crisis;\nglobal chip shortage;\nGPU prices +40-60%",
     "Full Hormuz blockade; global recession risk; Taiwan chip production falls 30-40%; "
     "GPU/server procurement severely constrained; $50-80B Gulf investments written down or suspended; "
     "credit ratings under review; some hyperscalers raise equity; "
     "AI capex maintained at reduced pace; non-AI IT spending frozen"],

    ["Worst Case\n(Regional war expansion)",
     "12+ months",
     "$150+\n(sustained)",
     "$400-500B\n(−$150-250B vs baseline;\nforced retrenchment)",
     "5-6%",
     "Catastrophic — global\nchip shortage rivaling\n2020-2022; fab closures",
     "Conflict expands to wider region; sustained energy crisis; global recession; "
     "sovereign debt crises in energy-importing nations; "
     "hyperscaler credit downgrades; forced capex prioritization — only highest-ROI AI projects proceed; "
     "Gulf AI hub vision abandoned for 5+ years"],
]

for r, row_data in enumerate(scenarios, 4):
    for c, val in enumerate(row_data, 1):
        cell = ws4.cell(row=r, column=c, value=val)

# Color-code scenarios
scenario_colors = [RISK_GREEN, RISK_AMBER, RISK_AMBER, RISK_RED, RISK_RED]
for r, color in enumerate(scenario_colors, 4):
    ws4.cell(row=r, column=1).fill = color
    ws4.cell(row=r, column=1).font = Font(bold=True)

style_data_area(ws4, 4, 4 + len(scenarios) - 1, len(headers4))
auto_width(ws4, len(headers4), max_width=50)
ws4.column_dimensions["G"].width = 60

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 5 – Energy & Supply Chain Data
# ═══════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Energy & Supply Chain")
ws5.sheet_properties.tabColor = "548235"

ws5.merge_cells("A1:F1")
ws5["A1"] = "Energy & Supply Chain Key Metrics — Iran Conflict Impact"
ws5["A1"].font = TITLE_FONT
ws5["A1"].alignment = Alignment(horizontal="center")

headers5 = [
    "Metric", "Pre-Conflict Value", "Current Value\n(Mar 2026)",
    "Change", "Relevance to Hyperscale CapEx", "Source / Notes"
]
for c, h in enumerate(headers5, 1):
    ws5.cell(row=3, column=c, value=h)
style_header_row(ws5, 3, len(headers5))

energy_data = [
    ["Brent Crude Oil ($/bbl)", "~$70", "$90-$110+", "+57% at peak",
     "Drives electricity costs (60% of DC opex); affects transportation/logistics",
     "Peak >$110; currently $90-103 range"],
    ["European Natural Gas Prices", "Baseline", "+40-50%", "+40-50%",
     "Gas-fired generation is primary DC power source in many regions",
     "Critical for European DC operations"],
    ["US Electricity Prices", "+6.9% YoY (2025)", "Further increases expected", "TBD",
     "Direct DC operating cost impact; US hyperscalers have some hedging",
     "EIA data; further increases as gas prices feed through"],
    ["Strait of Hormuz Traffic", "17-21M bbl/day; major LNG route",
     "Partially blocked/severely restricted", "~10M bbl/day removed",
     "Controls 20% global oil + significant LNG supply; affects Taiwan energy",
     "150+ vessels delayed; insurance premiums at records"],
    ["Global Helium Supply", "Qatar = 33% of global supply",
     "Ras Laffan facility hit; supply -33%", "-33%",
     "No substitute for helium in chip fab cooling/lithography; GPU/memory production at risk",
     "Qatar Energy facility struck by Iranian drones"],
    ["Taiwan LNG Reserves", "11 days", "Depleting", "Critical",
     "Taiwan produces majority of advanced semiconductors; 37% grid from ME energy",
     "vs. South Korea 52 days, Japan 21 days"],
    ["DDR5 Memory Prices", "Baseline", "Rising", "+10-15% (est.)",
     "Server memory costs affect DC buildout economics",
     "Supply disruption from helium/energy constraints"],
    ["GPU Prices", "Baseline", "Rising", "TBD (est. +10-20%)",
     "GPUs are ~60% of hyperscaler server capex; supply constrained pre-conflict",
     "NVIDIA H200/B200 delivery timelines extending"],
    ["Shipping Costs (Rerouting)", "Baseline", "Record highs",
     "+$2-3B/week globally",
     "Equipment delivery timelines extend; DC construction costs increase",
     "Rerouting around Africa adds ~19 days transit"],
    ["Ships Idle in Persian Gulf", "Normal flow", "3,200 ships (4% global tonnage)",
     "500+ waiting outside",
     "Supply chain congestion ripples globally",
     "Air cargo grounded from ME region"],
    ["SK Hynix + Samsung Market Cap", "Baseline", "Down >$200B combined", ">-$200B",
     "Reflects investor fear of sustained chip supply disruption",
     "CNBC, Mar 2026"],
    ["Data Center Insurance Premiums", "Standard commercial rates",
     "Surging to record levels", "Significant increase",
     "DCs now classified as 'legitimate military targets'; war risk hard to model",
     "Lloyd's; fintech.global"],
    ["Power Transformer Lead Times", "~100 weeks", "128 weeks", "+28%",
     "Critical bottleneck for new DC construction; already constraining Azure",
     "Pre-conflict constraint now worsening"],
]

for r, row_data in enumerate(energy_data, 4):
    for c, val in enumerate(row_data, 1):
        ws5.cell(row=r, column=c, value=val)

style_data_area(ws5, 4, 4 + len(energy_data) - 1, len(headers5))
auto_width(ws5, len(headers5), max_width=50)
ws5.column_dimensions["E"].width = 50

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 6 – Timeline of Key Events
# ═══════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Conflict Timeline")
ws6.sheet_properties.tabColor = "7030A0"

ws6.merge_cells("A1:E1")
ws6["A1"] = "Key Events Timeline — Iran Conflict & Hyperscale Impact"
ws6["A1"].font = TITLE_FONT
ws6["A1"].alignment = Alignment(horizontal="center")

headers6 = ["Date", "Event", "Impact on Hyperscale CapEx", "Oil Price Impact", "Source"]
for c, h in enumerate(headers6, 1):
    ws6.cell(row=3, column=c, value=h)
style_header_row(ws6, 3, len(headers6))

timeline = [
    ["Late Feb 2026", "US-Israel launch Operation Epic Fury strikes on Iran",
     "Initial market uncertainty; no immediate capex changes",
     "Brent jumps from ~$70 to $80+", "Reuters"],
    ["Mar 1, 2026", "Iran retaliates — Shahed-136 drones strike AWS data centers in UAE & Bahrain",
     "2 of 3 AWS UAE AZs offline; banks/payments disrupted; "
     "data centers classified as military targets",
     "Brent surges to $90-103", "Bisnow; Tom's Hardware; CNBC"],
    ["Mar 1, 2026", "Iranian drones strike Qatar Energy Ras Laffan facility",
     "33% of global helium supply threatened; chip fab at risk",
     "LNG prices spike", "abhs.in"],
    ["Early Mar 2026", "Strait of Hormuz partially blocked / severely restricted",
     "10M bbl/day removed; Taiwan LNG supply threatened; "
     "shipping rerouted around Africa (+19 days)",
     "Brent reaches $100-110+", "Reuters; AP News"],
    ["Mar 5, 2026", "South Korea warns Iran crisis could disrupt chipmaking materials",
     "Helium, bromine supply chains at risk; "
     "SK Hynix/Samsung lose >$200B combined market cap",
     "N/A", "Reuters"],
    ["Mar 6, 2026", "CNBC reports data centers have become military targets",
     "Insurance repricing begins; analyst questions about Gulf infrastructure safety",
     "N/A", "CNBC"],
    ["Mar 11, 2026", "CNBC analysis: Iran war could impact hyperscalers' Middle East AI plans",
     "Analysts suggest geographic shift to N. Europe, India, SE Asia; "
     "existing commitments likely maintained",
     "N/A", "CNBC"],
    ["Mid-Mar 2026", "Tom's Hardware: Strait of Hormuz blockade days away from crippling Taiwan semis",
     "Taiwan has 11 days LNG reserves; TSMC production at imminent risk; "
     "DDR5/GPU prices rising",
     "Brent $100-110 range", "Tom's Hardware; Taipei Times"],
    ["Mar 16, 2026", "Motley Fool: 'Can Hyperscalers Afford to Lose a Data Center to War?'",
     "Analysis: $800M-$1B per facility loss absorbable; $630B total capex provides buffer",
     "N/A", "The Motley Fool"],
    ["Mar 23, 2026", "Conflict enters 4th week — no ceasefire in sight",
     "Protracted conflict scenario increasingly likely; "
     "cumulative supply chain pressure building",
     "Brent $90-110 range", "Current"],
]

for r, row_data in enumerate(timeline, 4):
    for c, val in enumerate(row_data, 1):
        ws6.cell(row=r, column=c, value=val)

style_data_area(ws6, 4, 4 + len(timeline) - 1, len(headers6))
auto_width(ws6, len(headers6), max_width=55)
ws6.column_dimensions["B"].width = 50
ws6.column_dimensions["C"].width = 55

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 7 – Summary Dashboard
# ═══════════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Summary Dashboard")
ws7.sheet_properties.tabColor = "00B050"
# Move to front
wb.move_sheet(ws7, offset=-6)

ws7.merge_cells("A1:F1")
ws7["A1"] = "EXECUTIVE SUMMARY: Hyperscale CapEx Sentiment Under Protracted Iran Conflict"
ws7["A1"].font = Font(name="Calibri", bold=True, size=16, color="1F4E79")
ws7["A1"].alignment = Alignment(horizontal="center")

ws7.merge_cells("A2:F2")
ws7["A2"] = "Prepared: March 23, 2026 | Conflict Week 4 | Status: Protracted scenario increasingly likely"
ws7["A2"].font = Font(italic=True, size=10, color="666666")
ws7["A2"].alignment = Alignment(horizontal="center")

# Key metrics boxes
r = 4
metrics = [
    ("Aggregate 2026 CapEx (Baseline)", "$660-690B", "Big Five combined"),
    ("Aggregate 2026 CapEx (Protracted)", "$580-630B", "3-6 month conflict scenario"),
    ("Estimated CapEx at Risk", "$60-100B", "Gulf deferrals + energy cost drag"),
    ("Gulf Region Commitments", ">$100B", "Combined committed investments"),
    ("Oil Price (Current)", "$90-110/bbl", "+57% from pre-conflict"),
    ("Global IT Spending Growth", "8-9%", "vs. 10% baseline (IDC)"),
]

ws7.cell(row=r, column=1, value="KEY METRICS").font = SUBTITLE_FONT
r += 1
headers_m = ["Metric", "Value", "Context"]
for c, h in enumerate(headers_m, 1):
    ws7.cell(row=r, column=c, value=h)
style_header_row(ws7, r, 3)
r += 1
for metric in metrics:
    for c, val in enumerate(metric, 1):
        ws7.cell(row=r, column=c, value=val)
        ws7.cell(row=r, column=c).border = THIN_BORDER
    r += 1

r += 1
ws7.cell(row=r, column=1, value="OVERALL SENTIMENT ASSESSMENT").font = SUBTITLE_FONT
r += 1
headers_s = ["Company", "Sentiment", "Exposure", "Outlook"]
for c, h in enumerate(headers_s, 1):
    ws7.cell(row=r, column=c, value=h)
style_header_row(ws7, r, 4)
r += 1

sentiment_summary = [
    ["Amazon (AWS)", "Cautiously Bullish", "HIGH (facilities struck)", "Maintain guidance; diversify geography"],
    ["Microsoft (Azure)", "Cautiously Bullish", "HIGH ($95B Gulf commitments)", "Power bottleneck > conflict risk"],
    ["Alphabet (Google)", "Bullish", "MEDIUM ($10B Saudi)", "Strongest position; least Gulf exposure"],
    ["Meta", "Neutral-to-Cautious", "LOW (US-focused)", "Cost-cutting offsets energy pressure"],
    ["Oracle", "Cautious", "VERY HIGH (Stargate UAE)", "Highest relative risk; restructuring"],
]

for company_row in sentiment_summary:
    for c, val in enumerate(company_row, 1):
        cell = ws7.cell(row=r, column=c, value=val)
        cell.border = THIN_BORDER
    # Color-code exposure
    exp_cell = ws7.cell(row=r, column=3)
    if "VERY HIGH" in company_row[2]:
        exp_cell.fill = RISK_RED
    elif "HIGH" in company_row[2]:
        exp_cell.fill = RISK_RED
    elif "MEDIUM" in company_row[2]:
        exp_cell.fill = RISK_AMBER
    else:
        exp_cell.fill = RISK_GREEN
    r += 1

r += 1
ws7.cell(row=r, column=1, value="KEY FINDINGS").font = SUBTITLE_FONT
r += 1
findings = [
    "1. No hyperscaler has cut 2026 capex guidance as of March 23, 2026 — AI demand remains structurally strong.",
    "2. AWS facilities in UAE directly struck; first time data centers targeted as military infrastructure.",
    "3. Energy costs could add $15-25B annually to Big Five operating costs in a protracted scenario.",
    "4. Taiwan semiconductor production faces imminent risk — only 11 days of LNG reserves remain.",
    "5. 33% of global helium supply (essential for chip fab, no substitute) threatened by Qatar facility damage.",
    "6. Geographic investment shift away from Gulf toward Northern Europe, India, SE Asia is accelerating.",
    "7. Hyperscalers now hold more debt than cash collectively; rising energy costs compound financial strain.",
    "8. IDC projects IT spending growth falling from 10% to 8-9% under protracted conflict scenarios.",
    "9. Insurance industry repricing data center war risk; Gulf facilities may become uninsurable long-term.",
    "10. Net assessment: AI capex will be maintained but at reduced pace; non-AI spending bears the brunt of cuts.",
]
for finding in findings:
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws7.cell(row=r, column=1, value=finding).alignment = WRAP
    r += 1

auto_width(ws7, 6, max_width=50)
ws7.column_dimensions["A"].width = 35
ws7.column_dimensions["B"].width = 25
ws7.column_dimensions["C"].width = 30
ws7.column_dimensions["D"].width = 40

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 8 – CapEx Bar Chart
# ═══════════════════════════════════════════════════════════════════════════
ws_chart = wb.create_sheet("CapEx Chart Data")
ws_chart.sheet_properties.tabColor = "4472C4"

ws_chart["A1"] = "Company"
ws_chart["B1"] = "2024 CapEx ($B)"
ws_chart["C1"] = "2025 CapEx ($B)"
ws_chart["D1"] = "2026 Baseline ($B)"
ws_chart["E1"] = "2026 Protracted Conflict ($B)"
style_header_row(ws_chart, 1, 5)

chart_data = [
    ["Amazon", 83, 200, 200, 175],
    ["Microsoft", 56, 96, 120, 105],
    ["Alphabet", 52.5, 91.4, 180, 155],
    ["Meta", 39, 72.2, 125, 110],
    ["Oracle", 12, 25, 50, 40],
]
for r, row_data in enumerate(chart_data, 2):
    for c, val in enumerate(row_data, 1):
        ws_chart.cell(row=r, column=c, value=val)

chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.title = "Hyperscaler CapEx: Baseline vs. Protracted Iran Conflict ($B)"
chart.y_axis.title = "Capital Expenditure ($B)"
chart.x_axis.title = "Company"
chart.style = 10

data_ref = Reference(ws_chart, min_col=2, min_row=1, max_col=5, max_row=6)
cats = Reference(ws_chart, min_col=1, min_row=2, max_row=6)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.width = 22
chart.height = 14

ws_chart.add_chart(chart, "A9")

style_data_area(ws_chart, 2, 6, 5)
auto_width(ws_chart, 5)

# ═══════════════════════════════════════════════════════════════════════════
# Save
# ═══════════════════════════════════════════════════════════════════════════
OUTPUT = "/workspace/hyperscale_capex_iran_conflict_analysis.xlsx"
wb.save(OUTPUT)
print(f"Spreadsheet saved to {OUTPUT}")
