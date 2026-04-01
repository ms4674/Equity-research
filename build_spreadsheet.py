"""
Build a single Excel workbook that aggregates all hyperscale capex analysis
and ties capex to backlog (RPO) data.
"""

import xlsxwriter
import os

OUTPUT = os.path.join(os.path.dirname(__file__), "hyperscale_capex_analysis.xlsx")

wb = xlsxwriter.Workbook(OUTPUT, {"nan_inf_to_errors": True})

# ── Formats ─────────────────────────────────────────────────────────────
title_fmt = wb.add_format({
    "bold": True, "font_size": 16, "font_color": "#1B3A5C",
    "bottom": 2, "bottom_color": "#1B3A5C",
})
section_fmt = wb.add_format({
    "bold": True, "font_size": 13, "font_color": "#1B3A5C",
    "bottom": 1, "bottom_color": "#CCCCCC",
})
header_fmt = wb.add_format({
    "bold": True, "font_size": 10, "bg_color": "#1B3A5C",
    "font_color": "white", "border": 1, "text_wrap": True,
    "align": "center", "valign": "vcenter",
})
header_left_fmt = wb.add_format({
    "bold": True, "font_size": 10, "bg_color": "#1B3A5C",
    "font_color": "white", "border": 1, "text_wrap": True,
    "align": "left", "valign": "vcenter",
})
dollar_fmt = wb.add_format({
    "num_format": "$#,##0.0", "border": 1, "align": "center",
})
dollar_int_fmt = wb.add_format({
    "num_format": "$#,##0", "border": 1, "align": "center",
})
pct_fmt = wb.add_format({
    "num_format": "0%", "border": 1, "align": "center",
})
pct1_fmt = wb.add_format({
    "num_format": "0.0%", "border": 1, "align": "center",
})
num_fmt = wb.add_format({
    "num_format": "#,##0.0", "border": 1, "align": "center",
})
text_fmt = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
})
text_center_fmt = wb.add_format({
    "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter",
})
bold_text_fmt = wb.add_format({
    "bold": True, "border": 1, "text_wrap": True, "valign": "top",
})
total_fmt = wb.add_format({
    "bold": True, "num_format": "$#,##0.0", "border": 1,
    "bg_color": "#E8EEF4", "align": "center", "top": 2,
})
total_label_fmt = wb.add_format({
    "bold": True, "border": 1, "bg_color": "#E8EEF4", "top": 2,
})
total_pct_fmt = wb.add_format({
    "bold": True, "num_format": "0.0%", "border": 1,
    "bg_color": "#E8EEF4", "align": "center", "top": 2,
})
growth_pos_fmt = wb.add_format({
    "num_format": "+0.0%;-0.0%", "border": 1, "align": "center",
    "font_color": "#1B7A3D",
})
growth_neg_fmt = wb.add_format({
    "num_format": "+0.0%;-0.0%", "border": 1, "align": "center",
    "font_color": "#C62828",
})
note_fmt = wb.add_format({
    "italic": True, "font_size": 9, "font_color": "#666666",
    "text_wrap": True,
})
highlight_fmt = wb.add_format({
    "bold": True, "bg_color": "#FFF3E0", "border": 1,
    "num_format": "$#,##0.0", "align": "center",
})
highlight_pct_fmt = wb.add_format({
    "bold": True, "bg_color": "#FFF3E0", "border": 1,
    "num_format": "0.0x", "align": "center",
})

# ═════════════════════════════════════════════════════════════════════════
# TAB 1 — ANNUAL CAPEX BY COMPANY
# ═════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Annual CapEx")
ws1.set_tab_color("#1B3A5C")
ws1.hide_gridlines(2)
ws1.set_column("A:A", 22)
ws1.set_column("B:J", 14)

years = ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025E", "2026E"]
capex = {
    "Amazon":          [13.4, 16.9, 40.1, 61.1, 63.6, 48.4, 83.0, 125.0, 200.0],
    "Microsoft":       [11.6, 13.9, 15.4, 20.6, 23.9, 28.0, 44.5,  80.0, 145.0],
    "Alphabet/Google": [25.1, 23.5, 22.3, 24.6, 31.5, 32.3, 52.5,  91.4, 180.0],
    "Meta":            [13.7, 15.1, 15.1, 19.2, 31.4, 28.1, 39.2,  72.2, 125.0],
    "Oracle":          [ 2.1,  1.6,  1.7,  3.5,  5.9,  6.9, 10.0,  20.0,  50.0],
}

ws1.merge_range("A1:J1", "Hyperscale CapEx by Company ($B)", title_fmt)
r = 2
ws1.write(r, 0, "Company", header_left_fmt)
for j, y in enumerate(years):
    ws1.write(r, j + 1, y, header_fmt)
r += 1
for company, vals in capex.items():
    ws1.write(r, 0, company, bold_text_fmt)
    for j, v in enumerate(vals):
        ws1.write(r, j + 1, v, dollar_fmt)
    r += 1

ws1.write(r, 0, "Combined Total", total_label_fmt)
for j in range(len(years)):
    total = sum(capex[c][j] for c in capex)
    ws1.write(r, j + 1, total, total_fmt)
r += 1

ws1.write(r, 0, "YoY Growth", total_label_fmt)
ws1.write(r, 1, "", total_label_fmt)
for j in range(1, len(years)):
    prev = sum(capex[c][j - 1] for c in capex)
    curr = sum(capex[c][j] for c in capex)
    g = curr / prev - 1
    ws1.write(r, j + 1, g, growth_pos_fmt if g >= 0 else growth_neg_fmt)
r += 2

ws1.merge_range(r, 0, r, 9,
    "Source: Company earnings reports (10-K/10-Q). 2025E and 2026E are consensus estimates / company guidance as of Q1 2026.",
    note_fmt)

# Stacked bar chart
chart1 = wb.add_chart({"type": "column", "subtype": "stacked"})
for i, company in enumerate(capex.keys()):
    chart1.add_series({
        "name": company,
        "categories": ["Annual CapEx", 2, 1, 2, len(years)],
        "values": ["Annual CapEx", 3 + i, 1, 3 + i, len(years)],
    })
chart1.set_title({"name": "Combined Hyperscaler CapEx ($B)"})
chart1.set_y_axis({"name": "CapEx ($B)", "major_gridlines": {"visible": True}})
chart1.set_size({"width": 780, "height": 420})
chart1.set_style(10)
ws1.insert_chart("A13", chart1)

# ═════════════════════════════════════════════════════════════════════════
# TAB 2 — BACKLOG / RPO
# ═════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Backlog (RPO)")
ws2.set_tab_color("#FF6F00")
ws2.hide_gridlines(2)
ws2.set_column("A:A", 22)
ws2.set_column("B:H", 16)

ws2.merge_range("A1:H1", "Remaining Performance Obligations / Backlog ($B)", title_fmt)

# RPO data — latest available
rpo_headers = ["Company", "Latest RPO ($B)", "RPO Period", "YoY Growth", "QoQ Growth",
               "RPO Source"]
rpo_data = [
    ["Amazon (AWS)",       244.0, "Q4 2025", 0.40,  0.22, "AWS segment 10-K"],
    ["Microsoft",          625.0, "FQ2 2026 (Dec 2025)", 1.10,  0.59, "Commercial RPO (10-Q)"],
    ["Alphabet/Google",    240.0, "Q4 2025", 0.55,  0.15, "Google Cloud RPO estimate"],
    ["Oracle",             523.0, "FQ2 2026 (Nov 2025)", 3.59,  0.15, "Total RPO (10-Q)"],
    ["Meta",               None,  "N/A — no cloud RPO disclosed", None, None, "Meta does not report RPO"],
]

r = 2
for j, h in enumerate(rpo_headers):
    ws2.write(r, j, h, header_left_fmt if j == 0 else header_fmt)
r += 1
for row in rpo_data:
    ws2.write(r, 0, row[0], bold_text_fmt)
    if row[1] is not None:
        ws2.write(r, 1, row[1], dollar_fmt)
    else:
        ws2.write(r, 1, "N/A", text_center_fmt)
    ws2.write(r, 2, row[2], text_center_fmt)
    if row[3] is not None:
        ws2.write(r, 3, row[3], growth_pos_fmt)
    else:
        ws2.write(r, 3, "N/A", text_center_fmt)
    if row[4] is not None:
        ws2.write(r, 4, row[4], growth_pos_fmt)
    else:
        ws2.write(r, 4, "N/A", text_center_fmt)
    ws2.write(r, 5, row[5], text_fmt)
    r += 1

r += 1
ws2.merge_range(r, 0, r, 5,
    "Note: RPO = contracted future revenue not yet recognised. Periods vary by fiscal year-end. "
    "Microsoft RPO includes all commercial segments (Azure, Office 365, Dynamics). "
    "Oracle RPO includes OCI + SaaS. Meta does not operate a third-party cloud and does not disclose RPO.",
    note_fmt)

# ── Historical RPO trend ────────────────────────────────────────────────
r += 2
ws2.merge_range(r, 0, r, 5, "Historical RPO / Backlog Trend ($B)", section_fmt)
r += 1

rpo_trend_headers = ["Company", "Jun 2023", "Dec 2023", "Jun 2024", "Dec 2024", "Jun 2025", "Dec 2025"]
rpo_trend = {
    "Amazon (AWS)":       [130, 155, 157, 175, 200, 244],
    "Microsoft":          [229, 253, 275, 298, 375, 625],
    "Alphabet/Google":    [ 64,  74,  84, 106, 155, 240],
    "Oracle":             [ 65,  80,  99, 130, 455, 523],
}

for j, h in enumerate(rpo_trend_headers):
    ws2.write(r, j, h, header_left_fmt if j == 0 else header_fmt)
r += 1
for company, vals in rpo_trend.items():
    ws2.write(r, 0, company, bold_text_fmt)
    for j, v in enumerate(vals):
        ws2.write(r, j + 1, v, dollar_int_fmt)
    r += 1

ws2.write(r, 0, "Combined Total", total_label_fmt)
for j in range(len(rpo_trend_headers) - 1):
    total = sum(rpo_trend[c][j] for c in rpo_trend)
    ws2.write(r, j + 1, total, total_fmt)

# RPO chart
chart_rpo = wb.add_chart({"type": "column", "subtype": "stacked"})
trend_start_row = r - len(rpo_trend)
for i, company in enumerate(rpo_trend.keys()):
    chart_rpo.add_series({
        "name": company,
        "categories": ["Backlog (RPO)", trend_start_row - 1, 1, trend_start_row - 1, 6],
        "values": ["Backlog (RPO)", trend_start_row + i, 1, trend_start_row + i, 6],
    })
chart_rpo.set_title({"name": "Hyperscaler RPO / Backlog Trend ($B)"})
chart_rpo.set_y_axis({"name": "RPO ($B)"})
chart_rpo.set_size({"width": 700, "height": 380})
chart_rpo.set_style(10)
ws2.insert_chart(f"A{r + 4}", chart_rpo)

# ═════════════════════════════════════════════════════════════════════════
# TAB 3 — CAPEX vs BACKLOG (the core tie-out)
# ═════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("CapEx vs Backlog")
ws3.set_tab_color("#4CAF50")
ws3.hide_gridlines(2)
ws3.set_column("A:A", 22)
ws3.set_column("B:K", 16)

ws3.merge_range("A1:K1", "Tying CapEx to Backlog: Are Hyperscalers Building Ahead of Demand?", title_fmt)

r = 2
ws3.merge_range(r, 0, r, 10,
    "This sheet ties capital expenditure to contracted backlog (RPO) to assess whether spending is justified by demand visibility. "
    "Key ratios: RPO/CapEx (backlog coverage of annual spend) and RPO/Revenue (years of contracted visibility).",
    note_fmt)

r += 2
tie_headers = [
    "Company",
    "2025 CapEx ($B)",
    "2026E CapEx ($B)",
    "Latest RPO ($B)",
    "Cloud Rev Run-Rate ($B)",
    "RPO / 2026E CapEx",
    "RPO / Cloud Rev (yrs)",
    "2026E CapEx / Rev",
    "Backlog YoY Growth",
    "CapEx YoY Growth",
    "Backlog vs CapEx Signal",
]

tie_data = [
    # Company, CapEx25, CapEx26E, RPO, CloudRevRR, RPO/CapEx26, RPO/Rev, CapEx/Rev, BacklogYoY, CapExYoY
    ["Amazon (AWS)",       125.0, 200.0, 244.0, 142.0],
    ["Microsoft",           80.0, 145.0, 625.0, 131.0],
    ["Alphabet/Google",     91.4, 180.0, 240.0,  71.0],
    ["Meta",                72.2, 125.0,   None, None],
    ["Oracle",              20.0,  50.0, 523.0,  18.0],
]

for j, h in enumerate(tie_headers):
    ws3.write(r, j, h, header_left_fmt if j == 0 else header_fmt)
r += 1

for row in tie_data:
    company, cx25, cx26, rpo, rev_rr = row
    ws3.write(r, 0, company, bold_text_fmt)
    ws3.write(r, 1, cx25, dollar_fmt)
    ws3.write(r, 2, cx26, dollar_fmt)

    if rpo is not None:
        ws3.write(r, 3, rpo, dollar_fmt)
    else:
        ws3.write(r, 3, "N/A", text_center_fmt)

    if rev_rr is not None:
        ws3.write(r, 4, rev_rr, dollar_fmt)
    else:
        ws3.write(r, 4, "N/A", text_center_fmt)

    if rpo is not None and cx26 > 0:
        rpo_capex = rpo / cx26
        ws3.write(r, 5, rpo_capex, highlight_pct_fmt)
    else:
        ws3.write(r, 5, "N/A", text_center_fmt)

    if rpo is not None and rev_rr is not None and rev_rr > 0:
        rpo_rev = rpo / rev_rr
        ws3.write(r, 6, rpo_rev, highlight_pct_fmt)
    else:
        ws3.write(r, 6, "N/A", text_center_fmt)

    if rev_rr is not None:
        # Use total company rev for capital intensity
        total_rev = {"Amazon (AWS)": 700, "Microsoft": 280, "Alphabet/Google": 400,
                     "Meta": 190, "Oracle": 60}
        tr = total_rev.get(company, rev_rr)
        ws3.write(r, 7, cx26 / (tr * 1.15), pct1_fmt)
    else:
        ws3.write(r, 7, cx26 / (190 * 1.15), pct1_fmt)

    # Backlog YoY
    backlog_yoy = {"Amazon (AWS)": 0.40, "Microsoft": 1.10, "Alphabet/Google": 0.55,
                   "Oracle": 3.59, "Meta": None}
    byoy = backlog_yoy.get(company)
    if byoy is not None:
        ws3.write(r, 8, byoy, growth_pos_fmt)
    else:
        ws3.write(r, 8, "N/A", text_center_fmt)

    # CapEx YoY
    cx_yoy = cx26 / cx25 - 1 if cx25 > 0 else 0
    ws3.write(r, 9, cx_yoy, growth_pos_fmt if cx_yoy >= 0 else growth_neg_fmt)

    # Signal
    signals = {
        "Amazon (AWS)": "RPO growing but capex growing faster — building ahead of demand",
        "Microsoft": "Massive RPO (4.3x capex) — strong demand justification; OpenAI = 45% of backlog",
        "Alphabet/Google": "RPO covers 1.3x capex — balanced build; AI backlog accelerating",
        "Meta": "No external cloud RPO — capex driven by internal AI/infra needs (Llama, Reels, ads)",
        "Oracle": "RPO far exceeds capex (10.5x) — severely supply-constrained; need to build faster",
    }
    ws3.write(r, 10, signals.get(company, ""), text_fmt)
    r += 1

# Totals row (excl Meta RPO)
r_total = r
ws3.write(r, 0, "Total (excl Meta RPO)", total_label_fmt)
ws3.write(r, 1, sum(row[1] for row in tie_data), total_fmt)
ws3.write(r, 2, sum(row[2] for row in tie_data), total_fmt)
ws3.write(r, 3, sum(row[3] for row in tie_data if row[3] is not None), total_fmt)
ws3.write(r, 4, sum(row[4] for row in tie_data if row[4] is not None), total_fmt)
total_rpo = sum(row[3] for row in tie_data if row[3] is not None)
total_cx26 = sum(row[2] for row in tie_data)
ws3.write(r, 5, total_rpo / total_cx26, highlight_pct_fmt)

r += 2
ws3.merge_range(r, 0, r, 10, "Interpretation Guide", section_fmt)
r += 1

interpretations = [
    ["RPO / CapEx > 2.0x",
     "Backlog comfortably covers annual capex — spending is well-supported by contracted demand. "
     "Microsoft (4.3x) and Oracle (10.5x) lead here."],
    ["RPO / CapEx 1.0–2.0x",
     "Backlog roughly covers one year of capex — spending is reasonable but relies on continued "
     "new contract wins. Amazon (1.2x) and Google (1.3x) are in this zone."],
    ["RPO / CapEx < 1.0x or N/A",
     "Capex exceeds contracted backlog — company is building speculatively ahead of demand "
     "(Meta) or backlog is not disclosed."],
    ["RPO / Revenue > 2.0x",
     "Multi-year revenue visibility from contracts — strong forward demand signal. "
     "Oracle (29x) is an outlier due to very large multi-year AI contracts vs small current revenue base."],
    ["Backlog YoY > CapEx YoY",
     "Demand growth outpacing supply build — justifies continued capex acceleration. "
     "Microsoft and Oracle show this pattern clearly."],
    ["Backlog YoY < CapEx YoY",
     "Capex growing faster than new demand — potential over-build risk. "
     "Amazon shows this pattern; management argues demand is supply-constrained."],
]

for interp in interpretations:
    ws3.write(r, 0, interp[0], bold_text_fmt)
    ws3.merge_range(r, 1, r, 10, interp[1], text_fmt)
    r += 1

r += 1
ws3.merge_range(r, 0, r, 10,
    "Source: RPO from company 10-K/10-Q filings and earnings calls. Revenue run-rates based on Q4 2025 annualised. "
    "Meta does not operate a third-party cloud business and does not disclose RPO; its capex is driven by "
    "internal AI infrastructure needs (Llama model training, Reels recommendation, ads ranking).",
    note_fmt)

# ═════════════════════════════════════════════════════════════════════════
# TAB 4 — CATEGORY ALLOCATION
# ═════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("Category Allocation")
ws4.set_tab_color("#9C27B0")
ws4.hide_gridlines(2)
ws4.set_column("A:A", 28)
ws4.set_column("B:D", 16)
ws4.set_column("E:E", 60)

ws4.merge_range("A1:E1", "CapEx Allocation by Category (2026E)", title_fmt)

cat_headers = ["Category", "% of Total", "Est. Spend ($B)", "Cumul. %", "Description"]
categories = [
    ("Servers & GPUs",          0.35, 242, "GPU accelerators (NVIDIA Blackwell/Rubin), TPUs, CPUs, server hardware"),
    ("Data Center Construction", 0.18, 124, "Building shell, structure, interior fit-out, site preparation"),
    ("Power Infrastructure",     0.12,  83, "Electrical systems, UPS, generators, switchgear, transformers, PDUs"),
    ("Networking Equipment",     0.08,  55, "Switches, routers, optical interconnects, fiber, DCI links"),
    ("Memory & Storage",         0.07,  48, "HBM (SK Hynix, Samsung), DDR5, SSDs, storage arrays"),
    ("Cooling Systems",          0.05,  35, "Liquid cooling, CRAC units, chillers, heat rejection, CDUs"),
    ("Land Acquisition",         0.03,  21, "Real estate, site development, land with power/fiber/water access"),
    ("Other / Miscellaneous",    0.12,  83, "Design, engineering, security systems, software-defined infrastructure"),
]

r = 2
for j, h in enumerate(cat_headers):
    ws4.write(r, j, h, header_left_fmt if j in (0, 4) else header_fmt)
r += 1

cumul = 0
for cat in categories:
    cumul += cat[1]
    ws4.write(r, 0, cat[0], bold_text_fmt)
    ws4.write(r, 1, cat[1], pct_fmt)
    ws4.write(r, 2, cat[2], dollar_int_fmt)
    ws4.write(r, 3, cumul, pct_fmt)
    ws4.write(r, 4, cat[3], text_fmt)
    r += 1

ws4.write(r, 0, "Total", total_label_fmt)
ws4.write(r, 1, 1.0, total_pct_fmt)
ws4.write(r, 2, sum(c[2] for c in categories), total_fmt)
ws4.write(r, 3, 1.0, total_pct_fmt)
ws4.write(r, 4, "", total_label_fmt)

r += 2
ws4.merge_range(r, 0, r, 4, "High-Level Aggregation", section_fmt)
r += 1
agg_headers = ["Bucket", "% of Total", "Est. Spend ($B)"]
for j, h in enumerate(agg_headers):
    ws4.write(r, j, h, header_left_fmt if j == 0 else header_fmt)
r += 1
agg_data = [
    ("IT Equipment (Servers, GPUs, Memory, Networking)", 0.50, 345),
    ("Facilities (Power, Construction, Cooling, Land)",  0.38, 263),
    ("Other / Miscellaneous",                            0.12,  83),
]
for agg in agg_data:
    ws4.write(r, 0, agg[0], bold_text_fmt)
    ws4.write(r, 1, agg[1], pct_fmt)
    ws4.write(r, 2, agg[2], dollar_int_fmt)
    r += 1

# Pie chart
chart_cat = wb.add_chart({"type": "pie"})
cat_data_start = 3
cat_data_end = cat_data_start + len(categories) - 1
chart_cat.add_series({
    "categories": ["Category Allocation", cat_data_start, 0, cat_data_end, 0],
    "values": ["Category Allocation", cat_data_start, 2, cat_data_end, 2],
    "data_labels": {"percentage": True, "category": True, "separator": "\n"},
})
chart_cat.set_title({"name": "2026E CapEx by Category"})
chart_cat.set_size({"width": 600, "height": 400})
chart_cat.set_style(10)
ws4.insert_chart(f"A{r + 3}", chart_cat)

# ═════════════════════════════════════════════════════════════════════════
# TAB 5 — GEOGRAPHIC ALLOCATION
# ═════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("Geographic Allocation")
ws5.set_tab_color("#2196F3")
ws5.hide_gridlines(2)
ws5.set_column("A:A", 22)
ws5.set_column("B:F", 16)
ws5.set_column("G:G", 55)

ws5.merge_range("A1:G1", "CapEx Geographic Distribution", title_fmt)

geo_headers = ["Region", "% of Total 2024", "Est. 2024 ($B)", "% of Total 2026E",
               "Est. 2026E ($B)", "Shift (pp)", "Key Markets"]
geo_data = [
    ("North America",       0.50, 0.48, "US (Virginia, Ohio, Texas, Oregon, Georgia), Canada"),
    ("Europe",              0.22, 0.20, "Ireland, Netherlands, UK, Denmark, Sweden, Germany"),
    ("Asia-Pacific",        0.20, 0.22, "Singapore, Japan, Australia, India, South Korea, Malaysia, Indonesia"),
    ("Middle East & Africa", 0.04, 0.05, "UAE, Saudi Arabia, South Africa, Israel"),
    ("Latin America",        0.04, 0.05, "Brazil, Mexico, Chile, Colombia"),
]
total_2024 = 229.2
total_2026 = 700.0

r = 2
for j, h in enumerate(geo_headers):
    ws5.write(r, j, h, header_left_fmt if j in (0, 6) else header_fmt)
r += 1

for geo in geo_data:
    region, pct24, pct26, markets = geo
    ws5.write(r, 0, region, bold_text_fmt)
    ws5.write(r, 1, pct24, pct_fmt)
    ws5.write(r, 2, pct24 * total_2024, dollar_fmt)
    ws5.write(r, 3, pct26, pct_fmt)
    ws5.write(r, 4, pct26 * total_2026, dollar_fmt)
    shift = pct26 - pct24
    ws5.write(r, 5, shift, growth_pos_fmt if shift >= 0 else growth_neg_fmt)
    ws5.write(r, 6, markets, text_fmt)
    r += 1

ws5.write(r, 0, "Total", total_label_fmt)
ws5.write(r, 1, 1.0, total_pct_fmt)
ws5.write(r, 2, total_2024, total_fmt)
ws5.write(r, 3, 1.0, total_pct_fmt)
ws5.write(r, 4, total_2026, total_fmt)
ws5.write(r, 5, "", total_label_fmt)
ws5.write(r, 6, "", total_label_fmt)

# Bar chart
chart_geo = wb.add_chart({"type": "column"})
geo_start = 3
geo_end = geo_start + len(geo_data) - 1
chart_geo.add_series({
    "name": "2024",
    "categories": ["Geographic Allocation", geo_start, 0, geo_end, 0],
    "values": ["Geographic Allocation", geo_start, 2, geo_end, 2],
})
chart_geo.add_series({
    "name": "2026E",
    "categories": ["Geographic Allocation", geo_start, 0, geo_end, 0],
    "values": ["Geographic Allocation", geo_start, 4, geo_end, 4],
})
chart_geo.set_title({"name": "Geographic CapEx Distribution ($B)"})
chart_geo.set_y_axis({"name": "CapEx ($B)"})
chart_geo.set_size({"width": 700, "height": 400})
chart_geo.set_style(10)
ws5.insert_chart(f"A{r + 3}", chart_geo)

# ═════════════════════════════════════════════════════════════════════════
# TAB 6 — DRIVERS
# ═════════════════════════════════════════════════════════════════════════
ws6 = wb.add_worksheet("CapEx Drivers")
ws6.set_tab_color("#F44336")
ws6.hide_gridlines(2)
ws6.set_column("A:A", 30)
ws6.set_column("B:B", 14)
ws6.set_column("C:C", 16)
ws6.set_column("D:D", 85)

ws6.merge_range("A1:D1", "Key Drivers of Hyperscale CapEx", title_fmt)

driver_headers = ["Driver", "Impact", "Category", "Description"]
drivers = [
    ("AI Model Training Demand", "Very High", "Demand",
     "Exponential growth in LLM parameter counts and training compute; GPT-4 class models require 10,000+ GPUs for months. Next-gen models (GPT-5, Gemini Ultra) require 50,000-100,000+ GPU clusters."),
    ("AI Inference Scale-Out", "Very High", "Demand",
     "Production AI services (search, copilots, agents) require massive GPU fleets for real-time inference. Inference compute is now growing faster than training compute."),
    ("GPU Supply Constraints", "Very High", "Supply",
     "NVIDIA Blackwell/Rubin sold out through 2027. TSMC CoWoS packaging at capacity. Market is supply-constrained, not demand-constrained. $500B+ in combined GPU orders through late 2026."),
    ("Power Grid Bottlenecks", "Very High", "Infrastructure",
     "Power transformer lead times at 128 weeks. PJM capacity pricing 10x increase ($34 → $329/MW-day). Data center timelines extended 24-72 months. Microsoft has $80B in unfulfilled Azure orders due to power shortages."),
    ("Cloud Revenue Growth", "High", "Demand",
     "Cloud revenue growing 20-30% YoY across hyperscalers (AWS 24%, Azure 29%, GCP 48% in Q4 2025). Layers traditional infrastructure demand on top of AI."),
    ("Enterprise AI Adoption", "High", "Demand",
     "Fortune 500 deploying private AI workloads on hyperscale platforms. Google Cloud backlog at $240B. AWS backlog at $244B, growing 40% YoY."),
    ("Competitive Dynamics", "High", "Strategy",
     "Arms race between hyperscalers — no company willing to under-invest and lose AI market position. Each capex announcement ratchets peers upward."),
    ("GPU Price Inflation", "High", "Cost",
     "Next-gen GPU systems (GB200 NVL72) cost $3M+ per rack. 15-25% hardware cost inflation. AWS raised H200 instance prices 15%, breaking 20-year trend of declining cloud costs."),
    ("Sovereign AI Requirements", "Medium", "Regulatory",
     "Governments mandating in-country AI compute and data residency. Driving geographic diversification of builds across EU, Middle East, Southeast Asia."),
    ("Interest Rate Environment", "Medium", "Financial",
     "Higher borrowing costs as hyperscalers issue $108B in bonds in 2025. $1.5T total debt issuance projected through the cycle."),
]

r = 2
for j, h in enumerate(driver_headers):
    ws6.write(r, j, h, header_left_fmt if j in (0, 3) else header_fmt)
r += 1

impact_colors = {
    "Very High": wb.add_format({"bold": True, "font_color": "#C62828", "border": 1, "align": "center"}),
    "High": wb.add_format({"bold": True, "font_color": "#E65100", "border": 1, "align": "center"}),
    "Medium": wb.add_format({"bold": True, "font_color": "#F9A825", "border": 1, "align": "center"}),
}

for d in drivers:
    ws6.write(r, 0, d[0], bold_text_fmt)
    ws6.write(r, 1, d[1], impact_colors.get(d[1], text_center_fmt))
    ws6.write(r, 2, d[2], text_center_fmt)
    ws6.write(r, 3, d[3], text_fmt)
    r += 1

# ═════════════════════════════════════════════════════════════════════════
# TAB 7 — POWER CONSTRAINTS
# ═════════════════════════════════════════════════════════════════════════
ws7 = wb.add_worksheet("Power & Constraints")
ws7.set_tab_color("#795548")
ws7.hide_gridlines(2)
ws7.set_column("A:A", 38)
ws7.set_column("B:D", 16)
ws7.set_column("E:E", 18)

ws7.merge_range("A1:E1", "Infrastructure Constraints & Power Bottleneck", title_fmt)

constraint_headers = ["Metric", "2023", "2025", "2026E", "Change Since 2023"]
constraints = [
    ("PJM Capacity Price ($/MW-day)",          34,  180,  329),
    ("Power Transformer Lead Time (weeks)",     52,  100,  128),
    ("DC Electricity Consumption (TWh)",       460,  650,  780),
    ("Avg DC Construction Timeline (months)",   18,   30,   42),
    ("Hyperscaler Debt Issued ($B)",            25,  108,  200),
]

r = 2
for j, h in enumerate(constraint_headers):
    ws7.write(r, j, h, header_left_fmt if j == 0 else header_fmt)
r += 1

for c in constraints:
    ws7.write(r, 0, c[0], bold_text_fmt)
    ws7.write(r, 1, c[1], num_fmt)
    ws7.write(r, 2, c[2], num_fmt)
    ws7.write(r, 3, c[3], num_fmt)
    change = c[3] / c[1] - 1 if c[1] > 0 else 0
    ws7.write(r, 4, change, growth_pos_fmt)
    r += 1

r += 1
ws7.merge_range(r, 0, r, 4, "Key Implications for CapEx", section_fmt)
r += 1

implications = [
    "Power is the #1 bottleneck — Microsoft has $80B in unfulfilled Azure orders due to electricity shortages, with GPUs sitting idle in warehouses.",
    "Data center projects originally planned for 2026 are now tracking 2028-2029 due to power and construction delays.",
    "Hyperscalers are pursuing nuclear (Constellation, NuScale SMRs), natural gas on-site generation, and direct utility partnerships to bypass grid constraints.",
    "IEA projects global data center electricity consumption will double to 945 TWh by 2030 — equivalent to Japan's total electricity use.",
    "Power infrastructure capex ($83B in 2026E) is the fastest-growing category as hyperscalers internalise utility-scale electrical systems.",
]

for impl in implications:
    ws7.write(r, 0, "•", bold_text_fmt)
    ws7.merge_range(r, 1, r, 4, impl, text_fmt)
    r += 1

# ═════════════════════════════════════════════════════════════════════════
# TAB 8 — SUMMARY DASHBOARD
# ═════════════════════════════════════════════════════════════════════════
ws8 = wb.add_worksheet("Summary Dashboard")
ws8.set_tab_color("#00897B")
ws8.hide_gridlines(2)
ws8.set_column("A:A", 30)
ws8.set_column("B:B", 22)
ws8.set_column("C:C", 70)

ws8.merge_range("A1:C1", "Hyperscale CapEx — Executive Summary", title_fmt)

r = 2
ws8.merge_range(r, 0, r, 2, "Key Metrics at a Glance", section_fmt)
r += 1

kpi_headers = ["Metric", "Value", "Context"]
kpis = [
    ("Combined CapEx 2026E", "$700B", "10x increase from $66B in 2018; largest infrastructure build in tech history"),
    ("AI Share of CapEx", "~75%", "Training + inference infrastructure; remainder is traditional cloud and networking"),
    ("Combined RPO / Backlog", "$1.6T+", "Contracted future revenue across AWS, Microsoft, Google, Oracle (excl. Meta)"),
    ("RPO / CapEx Coverage", "2.3x", "Aggregate backlog covers 2.3 years of annual capex — spending is broadly supported by demand"),
    ("Power Constraint Impact", "$80B+", "Microsoft alone has $80B in unfulfilled Azure orders due to electricity shortages"),
    ("Hyperscaler Debt Outstanding", "$200B+ (2026E)", "Cumulative bond issuance; $1.5T total projected through the cycle"),
    ("Capital Intensity Range", "30-70% of revenue", "Utility-like spending levels vs typical tech at 10-15%"),
    ("Largest Spender", "Amazon ($200B)", "Followed by Google ($180B), Microsoft ($145B), Meta ($125B), Oracle ($50B)"),
    ("Largest Backlog", "Microsoft ($625B)", "45% attributable to OpenAI; followed by Oracle ($523B), AWS ($244B), Google ($240B)"),
    ("Fastest Backlog Growth", "Oracle (+359% YoY)", "Driven by massive multi-year OCI AI deals; RPO jumped from $99B to $523B in 12 months"),
]

for j, h in enumerate(kpi_headers):
    ws8.write(r, j, h, header_left_fmt if j in (0, 2) else header_fmt)
r += 1
for kpi in kpis:
    ws8.write(r, 0, kpi[0], bold_text_fmt)
    ws8.write(r, 1, kpi[1], text_center_fmt)
    ws8.write(r, 2, kpi[2], text_fmt)
    r += 1

r += 1
ws8.merge_range(r, 0, r, 2, "Investment Implications by Category", section_fmt)
r += 1

inv_headers = ["Beneficiary Sector", "Key Names", "Thesis"]
inv_data = [
    ("GPU / Semiconductors", "NVIDIA, Broadcom, AMD, Marvell",
     "Direct beneficiaries of 35% capex allocation to servers/GPUs. NVIDIA has 90% share. Backlog coverage supports multi-year demand visibility."),
    ("Memory (HBM)", "SK Hynix, Samsung, Micron",
     "HBM demand growing 4x+ as each GPU requires more memory. 7% of total capex (~$48B)."),
    ("Power Infrastructure", "Eaton, Vertiv, Schneider, GE Vernova",
     "Fastest-growing capex category. Power is the binding constraint. Transformer and UPS demand is multi-year."),
    ("Cooling", "Vertiv, Modine Manufacturing",
     "Shift to liquid cooling for high-density GPU racks (50-120 kW/rack vs 10-15 kW traditional). 5% of capex (~$35B)."),
    ("Data Center REITs", "Equinix, Digital Realty, QTS",
     "Beneficiaries of 18% capex allocation to construction. Colocation demand surging as hyperscalers supplement owned capacity."),
    ("Utilities / Nuclear", "Constellation Energy, Vistra, NuScale",
     "Hyperscalers signing long-term PPAs and exploring nuclear (SMRs) to solve power constraint. Multi-decade revenue visibility."),
    ("Networking", "Arista, Ciena, Corning",
     "8% of capex (~$55B). AI cluster networking (InfiniBand, Ethernet) and data center interconnect fiber growing rapidly."),
]

for j, h in enumerate(inv_headers):
    ws8.write(r, j, h, header_left_fmt if j in (0, 2) else header_fmt)
r += 1
for inv in inv_data:
    ws8.write(r, 0, inv[0], bold_text_fmt)
    ws8.write(r, 1, inv[1], text_center_fmt)
    ws8.write(r, 2, inv[2], text_fmt)
    r += 1

r += 1
ws8.merge_range(r, 0, r, 2, "Key Risks", section_fmt)
r += 1
risks = [
    ("AI Monetisation Gap", "High",
     "Can hyperscalers generate sufficient cloud AI revenue to justify $700B/yr in capex? If AI revenue growth disappoints, capex cuts could be abrupt."),
    ("Power Timeline Risk", "High",
     "Delays in grid infrastructure could strand capital in idle facilities. Construction timelines already extended 24-72 months."),
    ("Debt Sustainability", "Medium",
     "$1.5T in projected debt issuance; free cash flow under severe pressure (Amazon FCF negative $17-28B in 2026E, Meta FCF down 90%)."),
    ("Technology Transition", "Medium",
     "More efficient inference architectures (e.g., distillation, mixture-of-experts) could reduce GPU density requirements over time."),
    ("Regulatory / ESG", "Medium",
     "Environmental pushback on data center energy consumption. Some EU jurisdictions restricting new DC permits."),
]
risk_headers = ["Risk", "Severity", "Description"]
for j, h in enumerate(risk_headers):
    ws8.write(r, j, h, header_left_fmt if j in (0, 2) else header_fmt)
r += 1
for risk in risks:
    ws8.write(r, 0, risk[0], bold_text_fmt)
    ws8.write(r, 1, risk[1], impact_colors.get(risk[1], text_center_fmt))
    ws8.write(r, 2, risk[2], text_fmt)
    r += 1

wb.close()
print(f"Workbook saved to: {OUTPUT}")
