#!/usr/bin/env python3
"""
Broadcom-Google TPU vs NVIDIA vs AMD — AI Accelerator Market Share in Training & Inference

Generates an Excel workbook with:
  1. Overall AI accelerator market share (2023-2028E)
  2. Training-specific market share breakdown
  3. Inference-specific market share breakdown
  4. Revenue projections by vendor (Training vs Inference)
  5. ASIC vs GPU growth rate comparison
  6. Chip-by-chip 2026 spec comparison
  7. TCO comparison for inference workloads
  8. Broadcom custom ASIC revenue breakdown
  9. Charts: market share trajectories, training vs inference splits, CAGR comparison

Sources cited in the "Sources" sheet.
"""

import math
from dataclasses import dataclass, field
from typing import List, Dict, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import (
    BarChart, LineChart, PieChart, Reference, Series, BarChart3D,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Color palette
# ---------------------------------------------------------------------------
NVIDIA_GREEN = "76B900"
AMD_RED = "ED1C24"
GOOGLE_BLUE = "4285F4"
BROADCOM_RED = "CC0000"
ASIC_PURPLE = "7B2D8E"
OTHER_GRAY = "888888"

HEADER_FILL = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1B2A4A")
SUBTITLE_FONT = Font(name="Calibri", size=11, bold=True, color="1B2A4A")
DATA_FONT = Font(name="Calibri", size=10)
SOURCE_FONT = Font(name="Calibri", size=9, italic=True, color="666666")
THIN_BORDER = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)
PCT_FMT = '0.0%'
USD_FMT = '$#,##0.0'
USD_B_FMT = '$#,##0.0"B"'
NUM_FMT = '#,##0'


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_data_area(ws, start_row, end_row, max_col, fmt=None):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if fmt and c > 1:
                cell.number_format = fmt


def add_title(ws, row, title, merge_end=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = TITLE_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")


def add_subtitle(ws, row, text, merge_end=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBTITLE_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")


def add_source_note(ws, row, text, merge_end=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SOURCE_FONT


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

YEARS = [2023, 2024, 2025, 2026, 2027, 2028]

# Overall AI Accelerator Market Size ($B) — Total addressable market
TOTAL_MARKET_SIZE = [55, 160, 250, 350, 450, 604]

# --- OVERALL MARKET SHARE (% of total AI accelerator revenue) ---
OVERALL_SHARE = {
    "NVIDIA":                  [0.92, 0.90, 0.86, 0.80, 0.72, 0.65],
    "AMD":                     [0.03, 0.05, 0.07, 0.10, 0.12, 0.14],
    "Google/Broadcom TPU":     [0.02, 0.03, 0.04, 0.06, 0.09, 0.12],
    "Other Custom ASIC":       [0.02, 0.01, 0.02, 0.03, 0.05, 0.06],
    "Other (Intel, etc.)":     [0.01, 0.01, 0.01, 0.01, 0.02, 0.03],
}

# --- TRAINING MARKET SHARE (% of training compute spend) ---
TRAINING_SHARE = {
    "NVIDIA":                  [0.95, 0.93, 0.91, 0.88, 0.85, 0.82],
    "AMD":                     [0.02, 0.04, 0.05, 0.07, 0.08, 0.10],
    "Google/Broadcom TPU":     [0.02, 0.02, 0.02, 0.03, 0.04, 0.04],
    "Other Custom ASIC":       [0.005, 0.005, 0.01, 0.01, 0.02, 0.025],
    "Other (Intel, etc.)":     [0.005, 0.005, 0.01, 0.01, 0.01, 0.015],
}

# Training as % of total AI compute
TRAINING_PCT_OF_TOTAL = [0.67, 0.55, 0.42, 0.33, 0.28, 0.25]

# --- INFERENCE MARKET SHARE (% of inference compute spend) ---
INFERENCE_SHARE = {
    "NVIDIA":                  [0.88, 0.85, 0.78, 0.68, 0.50, 0.30],
    "AMD":                     [0.04, 0.06, 0.08, 0.12, 0.15, 0.18],
    "Google/Broadcom TPU":     [0.04, 0.05, 0.08, 0.12, 0.18, 0.25],
    "Other Custom ASIC":       [0.03, 0.03, 0.04, 0.06, 0.13, 0.22],
    "Other (Intel, etc.)":     [0.01, 0.01, 0.02, 0.02, 0.04, 0.05],
}

# Inference as % of total AI compute
INFERENCE_PCT_OF_TOTAL = [0.33, 0.45, 0.58, 0.67, 0.72, 0.75]

# --- REVENUE PROJECTIONS ($B) by vendor ---
def compute_revenue(share_dict, pct_of_total):
    rev = {}
    for vendor, shares in share_dict.items():
        rev[vendor] = []
        for i, yr in enumerate(YEARS):
            segment_size = TOTAL_MARKET_SIZE[i] * pct_of_total[i]
            rev[vendor].append(round(shares[i] * segment_size, 1))
    return rev


# --- ASIC vs GPU CAGR ---
CAGR_DATA = [
    ("Custom ASICs (Hyperscaler)", "44.6%", "$18B", "$165B", "Inference-optimized"),
    ("General-purpose GPUs", "16.1%", "$130B", "$290B", "Training + flexible inference"),
    ("AMD Accelerators", "~18%", "$12B", "$55B", "Cost-sensitive training/cloud"),
    ("Total AI Accelerator Market", "~16%", "$160B", "$604B", "All AI compute"),
]

# --- 2026 CHIP SPEC COMPARISON ---
CHIP_SPECS = [
    # (Name, Vendor, Process, FP8 PFLOPS, FP4 PFLOPS, Memory GB, Mem BW TB/s, TDP W, Interconnect, Primary Use)
    ("Vera Rubin GPU", "NVIDIA", "3nm", "~25", "50", 288, ">12", "~1000", "NVLink 6 (3.6 TB/s)", "Training + Inference"),
    ("B200 Blackwell", "NVIDIA", "4nm", "4.5", "9", 192, "8", "1000", "NVLink 5 (1.8 TB/s)", "Training + Inference"),
    ("TPU v7 Ironwood", "Google/Broadcom", "3nm", "4.6", "N/D", 192, ">7.2", "~500", "Optical Mesh (4.8 Tbps)", "Inference-optimized"),
    ("TPU v6 Trillium", "Google/Broadcom", "5nm", "~0.9", "N/D", 32, "~1.6", "~300", "ICI (proprietary)", "Training + Inference"),
    ("MI455X", "AMD", "2nm", "~10", "~20", 432, "19.6", "~700", "Infinity Fabric", "Training + Inference"),
    ("MI350X", "AMD", "3nm", "~5", "~10", 288, "8", "~600", "Infinity Fabric", "Training + Inference"),
    ("Maia 200", "Microsoft/Broadcom", "3nm", "~5", ">10", 216, "~8", "750", "Azure Custom Fabric", "Inference (GPT-optimized)"),
    ("Trainium 3", "AWS", "3nm", "2.52", "N/D", 144, "~5", "~600", "EFA (3.2 Tbps)", "Training + Inference"),
]

# --- TCO COMPARISON: 1,000-chip inference cluster, 3-year ---
TCO_DATA = [
    # (Cost Factor, NVIDIA H100/B200, Google TPU v6/v7, AMD MI350/MI400, Winner)
    ("Hardware (CapEx)", "$100M", "$52M", "$75M", "TPU (-48%)"),
    ("Electricity (3yr)", "$47M", "$16M", "$32M", "TPU (-66%)"),
    ("Cooling infrastructure", "$12M", "$4M", "$8M", "TPU (-67%)"),
    ("Software licenses", "$0 (CUDA free)", "$0 (JAX free)", "$0 (ROCm free)", "Tie"),
    ("Support & maintenance", "$8M", "$3M", "$6M", "TPU (-63%)"),
    ("Network infrastructure", "$6M", "$2M", "$5M", "TPU (-67%)"),
    ("Real estate (rack space)", "$4M", "$1.5M", "$3M", "TPU (-63%)"),
    ("TOTAL 3-YEAR TCO", "$177M", "$78.5M", "$129M", "TPU (-56%)"),
]

# --- BROADCOM AI REVENUE (ACTUAL REPORTED + ESTIMATES) ---
# Broadcom fiscal years end in ~October/November; approx align to calendar year
BROADCOM_YEARS = ["FY2023", "FY2024", "FY2025", "FY2026E", "FY2027E"]
BROADCOM_DATA = {
    "Google TPU (est.)":             [3.8,  6.0, 10.0, 14.0, 21.0],
    "Other XPU Customers (est.)":    [0.2,  1.3,  3.5, 11.0, 30.0],
    "AI Networking":                 [1.0,  5.0,  6.5, 13.0, 35.0],
    "Total AI Revenue":              [5.0, 12.3, 20.0, 38.0, 86.0],
}

# --- BROADCOM QUARTERLY ACTUALS + RUN RATE DERIVATION ---
# Used to derive current TPU revenue from reported AI revenue guidance
BROADCOM_QUARTERLY = {
    "headers": ["Quarter", "AI Revenue ($B)", "YoY Growth", "Networking Mix",
                "XPU Revenue (est.)", "Google TPU (est.)", "TPU Run Rate (ann.)"],
    "rows": [
        ("Q1 FY2025", 4.1, "77%",  "~30%", "~$2.9B", "~$2.1B", "~$8.4B"),
        ("Q2 FY2025", 4.4, "44%",  "~30%", "~$3.1B", "~$2.2B", "~$8.8B"),
        ("Q3 FY2025", 5.2, "63%",  "~30%", "~$3.6B", "~$2.6B", "~$10.4B"),
        ("Q4 FY2025", 6.5, "74%",  "~33%", "~$4.4B", "~$3.0B", "~$12.0B"),
        ("Q1 FY2026 (actual)", 8.4, "106%", "~33%", "~$5.6B", "~$3.4B", "~$13.6B"),
        ("Q2 FY2026 (guide)", 10.7, "140%", "~40%", "~$6.4B", "~$3.5B", "~$14.0B"),
    ],
    "notes": [
        "AI Revenue = Broadcom's reported 'AI semiconductor revenue' per earnings calls",
        "Networking Mix = % of AI revenue from AI networking (Tomahawk, DSP, etc.) per Hock Tan commentary",
        "XPU Revenue = AI Revenue × (1 - Networking Mix); includes all custom accelerators",
        "Google TPU est. = Analyst consensus ~50% of XPU in FY2025, declining to ~40-45% as Meta/Anthropic/OpenAI ramp",
        "TPU Run Rate = Google TPU quarterly × 4; represents annualized Google TPU revenue from Broadcom",
        "Broadcom has 6 XPU customers: Google, Meta, Anthropic, OpenAI + 2 undisclosed (as of Q1 FY2026)",
        "Hock Tan (Q1 FY2026 call): 'Line of sight to AI revenue from chips in excess of $100B in 2027'",
        "$100B target is for chips only (XPUs, switches, DSPs) — confirmed by Blayne Curtis/Jefferies Q&A",
    ],
}

# --- REAL-WORLD MIGRATION CASE STUDIES ---
CASE_STUDIES = [
    ("Midjourney", "NVIDIA → TPU v6e", "65% cost reduction", "$2.1M → $700K/mo", "6 weeks, 3 engineers"),
    ("Anthropic", "Multi-provider → TPU v7", "1M+ TPU chips by 2027", "Tens of $B commitment", "Largest TPU deal ever"),
    ("Meta", "NVIDIA GPU → TPU talks", "Multibillion-dollar TPU talks", "600K+ chip infra by 2026", "Advanced negotiations"),
    ("Character.AI", "NVIDIA → TPU", "3.8× cost improvement", "Full inference stack", "8 weeks, 2 engineers"),
    ("Perplexity AI", "Mixed → TPU v5e/v6", "Entire inference on TPU", "Default TPU stack", "4 weeks, 2 engineers"),
    ("Stability AI", "NVIDIA → TPU v6", "40% of image gen to TPU", "Q3 2025 migration", "Ongoing"),
    ("OpenAI", "NVIDIA → Broadcom ASIC", "10GW capacity by 2029", "~$10B investment", "Custom chip in design"),
]

# --- SOURCES ---
SOURCES = [
    ("Broadcom Q1 FY2026 Earnings", "AI semi revenue $8.4B (+106% YoY); total revenue $19.3B (+29%); 6 XPU customers", "March 4, 2026"),
    ("Broadcom Q1 FY2026 Guidance", "Q2 AI semi revenue $10.7B (+140% YoY); total revenue $22B; networking 40% of AI rev", "March 4, 2026"),
    ("Broadcom CEO Hock Tan", "'Line of sight to AI revenue from chips >$100B in 2027'; chips = XPUs + switches + DSPs", "Q1 FY2026 Call"),
    ("Broadcom Q4 FY2025 Earnings", "FY2025 total AI revenue ~$20B (+65% YoY); AI semi revenue $6.5B in Q4 (+74%)", "December 2025"),
    ("Reuters", "Broadcom sees >$100B AI chip sales by 2027; Google TPU orders $21B in two quarters", "March 2026"),
    ("JPMorgan (Harlan Sur)", "Expects Broadcom AI revenue >$9B/quarter; projects >$65B total AI rev FY2026", "March 2026"),
    ("Bloomberg Intelligence", "AI Accelerator Market $604B by 2033; ASIC 44.6% CAGR vs GPU 16.1% CAGR", "January 2026"),
    ("New Street Research", "Inference Compute Share Analysis 2024-2028; NVIDIA inference share 90%→20-30% by 2028", "December 2025"),
    ("SemiAnalysis", "NVIDIA Market Share and Competitive Landscape; 90%+ current market share", "Q4 2025"),
    ("Morgan Stanley", "AI Compute Economics: Training vs Inference; TPU v6 to generate >$150B lifetime rev", "November 2025"),
    ("Goldman Sachs", "Hyperscaler AI Capex $660-690B in 2026; TPU 35% inference share by Q4 2026", "February 2026"),
    ("Google Cloud", "TPU v7 Ironwood specs: 4,614 TFLOPS, 192GB HBM3e, 9,216-chip pods", "April 2025"),
    ("NVIDIA GTC 2026", "Vera Rubin: 336B transistors, 50 PFLOPS FP4, 288GB HBM4, NVLink 6", "March 2026"),
    ("The Information", "Midjourney 65% cost reduction with TPU migration; Meta TPU talks", "September 2025"),
    ("Anthropic Blog", "1M+ Google TPU v7 chips for Claude inference; tens of $B commitment", "2025"),
    ("Zylos Research", "AI Chip Hardware Acceleration Trends 2026; ASIC share 15%→40% in inference", "February 2026"),
    ("Introl Blog", "Custom Silicon Inflection 2026; hyperscaler ASIC roadmap comparison", "February 2026"),
    ("AMD Newsroom", "MI400 series (2nm CDNA 5); MI355X 30% faster inference than B200", "CES 2026"),
    ("Silicon Analysts", "NVIDIA GPU Market Share 2024-2026: 87% peak, declining trajectory", "2026"),
    ("Epoch AI", "Inference projected at 75-80% of AI compute by 2030", "2025"),
]


# ---------------------------------------------------------------------------
# Workbook builder
# ---------------------------------------------------------------------------
def build_workbook():
    wb = Workbook()

    # ===== Sheet 1: Overall Market Share =====
    ws = wb.active
    ws.title = "Overall Market Share"
    ws.sheet_properties.tabColor = "1B2A4A"

    add_title(ws, 1, "AI Accelerator Market Share by Vendor (2023-2028E)")
    add_subtitle(ws, 2, "Broadcom-Google TPU vs NVIDIA vs AMD — Overall Market")

    row = 4
    headers = ["Vendor"] + [str(y) for y in YEARS]
    for c, h in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, len(headers))

    row = 5
    for vendor, shares in OVERALL_SHARE.items():
        ws.cell(row=row, column=1, value=vendor)
        for c, s in enumerate(shares, 2):
            ws.cell(row=row, column=c, value=s)
        row += 1

    ws.cell(row=row, column=1, value="Total Market Size ($B)")
    ws.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=True)
    for c, val in enumerate(TOTAL_MARKET_SIZE, 2):
        cell = ws.cell(row=row, column=c, value=val)
        cell.number_format = USD_FMT
        cell.font = Font(name="Calibri", size=10, bold=True)
    row += 1

    style_data_area(ws, 5, row - 2, len(headers), PCT_FMT)
    for r in range(5, row - 1):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")

    add_source_note(ws, row + 1, "Sources: Bloomberg Intelligence, SemiAnalysis, New Street Research, Silicon Analysts (2025-2026)")

    # Line chart for overall market share
    chart = LineChart()
    chart.title = "AI Accelerator Market Share Trajectory (2023-2028E)"
    chart.x_axis.title = "Year"
    chart.y_axis.title = "Market Share (%)"
    chart.y_axis.numFmt = '0%'
    chart.style = 10
    chart.width = 22
    chart.height = 14

    cats = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=4)
    for i, vendor in enumerate(OVERALL_SHARE.keys()):
        vals = Reference(ws, min_col=2, max_col=len(YEARS) + 1, min_row=5 + i)
        s = Series(vals, title=vendor)
        chart.append(s)
    chart.set_categories(cats)

    colors = [NVIDIA_GREEN, AMD_RED, GOOGLE_BLUE, ASIC_PURPLE, OTHER_GRAY]
    for i, s in enumerate(chart.series):
        s.graphicalProperties.line.width = 28000
        if i < len(colors):
            s.graphicalProperties.line.solidFill = colors[i]

    ws.add_chart(chart, "A" + str(row + 3))

    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 20

    # ===== Sheet 2: Training Market Share =====
    ws2 = wb.create_sheet("Training Share")
    ws2.sheet_properties.tabColor = NVIDIA_GREEN

    add_title(ws2, 1, "AI Training Market Share by Vendor (2023-2028E)")
    add_subtitle(ws2, 2, "NVIDIA dominates training — CUDA ecosystem + NVLink maintain strong moat")

    row = 4
    headers = ["Vendor"] + [str(y) for y in YEARS]
    for c, h in enumerate(headers, 1):
        ws2.cell(row=row, column=c, value=h)
    style_header_row(ws2, row, len(headers))

    row = 5
    for vendor, shares in TRAINING_SHARE.items():
        ws2.cell(row=row, column=1, value=vendor)
        for c, s in enumerate(shares, 2):
            ws2.cell(row=row, column=c, value=s)
        row += 1

    ws2.cell(row=row, column=1, value="Training % of Total AI Compute")
    ws2.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=True)
    for c, val in enumerate(TRAINING_PCT_OF_TOTAL, 2):
        ws2.cell(row=row, column=c, value=val)
    row += 1

    training_rev = compute_revenue(TRAINING_SHARE, TRAINING_PCT_OF_TOTAL)
    row += 1
    add_subtitle(ws2, row, "Training Revenue by Vendor ($B)")
    row += 1
    for c, h in enumerate(headers, 1):
        ws2.cell(row=row, column=c, value=h)
    style_header_row(ws2, row, len(headers))
    rev_start = row + 1
    row += 1
    for vendor, revs in training_rev.items():
        ws2.cell(row=row, column=1, value=vendor)
        for c, r in enumerate(revs, 2):
            ws2.cell(row=row, column=c, value=r)
        row += 1

    style_data_area(ws2, 5, 5 + len(TRAINING_SHARE) - 1, len(headers), PCT_FMT)
    ws2.cell(row=5 + len(TRAINING_SHARE), column=1).alignment = Alignment(horizontal="left")
    for r in range(5, 5 + len(TRAINING_SHARE)):
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left")
    style_data_area(ws2, rev_start, rev_start + len(training_rev) - 1, len(headers), USD_FMT)
    for r in range(rev_start, rev_start + len(training_rev)):
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left")

    add_source_note(ws2, row + 1, "Sources: Bloomberg Intelligence, SemiAnalysis, Morgan Stanley, NVIDIA GTC 2026")

    chart2 = BarChart()
    chart2.type = "col"
    chart2.grouping = "stacked"
    chart2.title = "Training Revenue by Vendor ($B, 2023-2028E)"
    chart2.x_axis.title = "Year"
    chart2.y_axis.title = "Revenue ($B)"
    chart2.style = 10
    chart2.width = 22
    chart2.height = 14

    cats2 = Reference(ws2, min_col=2, max_col=len(YEARS) + 1, min_row=rev_start - 1)
    for i, vendor in enumerate(training_rev.keys()):
        vals = Reference(ws2, min_col=2, max_col=len(YEARS) + 1, min_row=rev_start + i)
        s = Series(vals, title=vendor)
        chart2.append(s)
    chart2.set_categories(cats2)

    for i, s in enumerate(chart2.series):
        if i < len(colors):
            s.graphicalProperties.solidFill = colors[i]

    ws2.add_chart(chart2, "A" + str(row + 3))

    for c in range(1, len(headers) + 1):
        ws2.column_dimensions[get_column_letter(c)].width = 22

    # ===== Sheet 3: Inference Market Share =====
    ws3 = wb.create_sheet("Inference Share")
    ws3.sheet_properties.tabColor = GOOGLE_BLUE

    add_title(ws3, 1, "AI Inference Market Share by Vendor (2023-2028E)")
    add_subtitle(ws3, 2, "NVIDIA's moat eroding — Google/Broadcom TPU + Custom ASICs gaining rapidly")

    row = 4
    for c, h in enumerate(headers, 1):
        ws3.cell(row=row, column=c, value=h)
    style_header_row(ws3, row, len(headers))

    row = 5
    for vendor, shares in INFERENCE_SHARE.items():
        ws3.cell(row=row, column=1, value=vendor)
        for c, s in enumerate(shares, 2):
            ws3.cell(row=row, column=c, value=s)
        row += 1

    ws3.cell(row=row, column=1, value="Inference % of Total AI Compute")
    ws3.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=True)
    for c, val in enumerate(INFERENCE_PCT_OF_TOTAL, 2):
        ws3.cell(row=row, column=c, value=val)
    row += 1

    inference_rev = compute_revenue(INFERENCE_SHARE, INFERENCE_PCT_OF_TOTAL)
    row += 1
    add_subtitle(ws3, row, "Inference Revenue by Vendor ($B)")
    row += 1
    for c, h in enumerate(headers, 1):
        ws3.cell(row=row, column=c, value=h)
    style_header_row(ws3, row, len(headers))
    rev_start3 = row + 1
    row += 1
    for vendor, revs in inference_rev.items():
        ws3.cell(row=row, column=1, value=vendor)
        for c, r in enumerate(revs, 2):
            ws3.cell(row=row, column=c, value=r)
        row += 1

    style_data_area(ws3, 5, 5 + len(INFERENCE_SHARE) - 1, len(headers), PCT_FMT)
    for r in range(5, 5 + len(INFERENCE_SHARE)):
        ws3.cell(row=r, column=1).alignment = Alignment(horizontal="left")
    style_data_area(ws3, rev_start3, rev_start3 + len(inference_rev) - 1, len(headers), USD_FMT)
    for r in range(rev_start3, rev_start3 + len(inference_rev)):
        ws3.cell(row=r, column=1).alignment = Alignment(horizontal="left")

    add_source_note(ws3, row + 1, "Sources: New Street Research, Goldman Sachs, Zylos Research, Epoch AI (2025-2026)")

    chart3 = BarChart()
    chart3.type = "col"
    chart3.grouping = "stacked"
    chart3.title = "Inference Revenue by Vendor ($B, 2023-2028E)"
    chart3.x_axis.title = "Year"
    chart3.y_axis.title = "Revenue ($B)"
    chart3.style = 10
    chart3.width = 22
    chart3.height = 14

    cats3 = Reference(ws3, min_col=2, max_col=len(YEARS) + 1, min_row=rev_start3 - 1)
    for i, vendor in enumerate(inference_rev.keys()):
        vals = Reference(ws3, min_col=2, max_col=len(YEARS) + 1, min_row=rev_start3 + i)
        s = Series(vals, title=vendor)
        chart3.append(s)
    chart3.set_categories(cats3)

    for i, s in enumerate(chart3.series):
        if i < len(colors):
            s.graphicalProperties.solidFill = colors[i]

    ws3.add_chart(chart3, "A" + str(row + 3))

    # Side-by-side inference share line chart
    chart3b = LineChart()
    chart3b.title = "Inference Market Share: NVIDIA Decline vs TPU/ASIC Rise"
    chart3b.x_axis.title = "Year"
    chart3b.y_axis.title = "Inference Market Share (%)"
    chart3b.y_axis.numFmt = '0%'
    chart3b.style = 10
    chart3b.width = 22
    chart3b.height = 14

    for i, vendor in enumerate(INFERENCE_SHARE.keys()):
        vals = Reference(ws3, min_col=2, max_col=len(YEARS) + 1, min_row=5 + i)
        s = Series(vals, title=vendor)
        chart3b.append(s)
    chart3b.set_categories(cats3)

    for i, s in enumerate(chart3b.series):
        s.graphicalProperties.line.width = 28000
        if i < len(colors):
            s.graphicalProperties.line.solidFill = colors[i]

    ws3.add_chart(chart3b, "A" + str(row + 20))

    for c in range(1, len(headers) + 1):
        ws3.column_dimensions[get_column_letter(c)].width = 22

    # ===== Sheet 4: Training vs Inference Split =====
    ws4 = wb.create_sheet("Training vs Inference")
    ws4.sheet_properties.tabColor = "FF9900"

    add_title(ws4, 1, "Training vs Inference: The Great Compute Shift")
    add_subtitle(ws4, 2, "Inference now dominates — 67% of all AI compute in 2026, projected 75% by 2028")

    row = 4
    split_headers = ["Metric"] + [str(y) for y in YEARS]
    for c, h in enumerate(split_headers, 1):
        ws4.cell(row=row, column=c, value=h)
    style_header_row(ws4, row, len(split_headers))

    rows_data = [
        ("Training % of Total Compute", TRAINING_PCT_OF_TOTAL),
        ("Inference % of Total Compute", INFERENCE_PCT_OF_TOTAL),
        ("Total Market Size ($B)", TOTAL_MARKET_SIZE),
    ]

    training_market = [TOTAL_MARKET_SIZE[i] * TRAINING_PCT_OF_TOTAL[i] for i in range(len(YEARS))]
    inference_market = [TOTAL_MARKET_SIZE[i] * INFERENCE_PCT_OF_TOTAL[i] for i in range(len(YEARS))]
    rows_data.append(("Training Market Size ($B)", training_market))
    rows_data.append(("Inference Market Size ($B)", inference_market))

    row = 5
    for label, data in rows_data:
        ws4.cell(row=row, column=1, value=label)
        ws4.cell(row=row, column=1).alignment = Alignment(horizontal="left")
        for c, val in enumerate(data, 2):
            cell = ws4.cell(row=row, column=c, value=val)
            if "%" in label:
                cell.number_format = PCT_FMT
            else:
                cell.number_format = USD_FMT
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
        row += 1

    add_source_note(ws4, row + 1, "Sources: Morgan Stanley, Epoch AI, New Street Research (2025)")

    # Stacked bar chart: training vs inference market size
    chart4 = BarChart()
    chart4.type = "col"
    chart4.grouping = "stacked"
    chart4.title = "AI Compute Market: Training vs Inference ($B)"
    chart4.x_axis.title = "Year"
    chart4.y_axis.title = "Market Size ($B)"
    chart4.style = 10
    chart4.width = 22
    chart4.height = 14

    cats4 = Reference(ws4, min_col=2, max_col=len(YEARS) + 1, min_row=4)
    training_ref = Reference(ws4, min_col=2, max_col=len(YEARS) + 1, min_row=8)
    inference_ref = Reference(ws4, min_col=2, max_col=len(YEARS) + 1, min_row=9)

    s1 = Series(training_ref, title="Training")
    s1.graphicalProperties.solidFill = NVIDIA_GREEN
    s2 = Series(inference_ref, title="Inference")
    s2.graphicalProperties.solidFill = GOOGLE_BLUE

    chart4.append(s1)
    chart4.append(s2)
    chart4.set_categories(cats4)

    ws4.add_chart(chart4, "A" + str(row + 3))

    for c in range(1, len(split_headers) + 1):
        ws4.column_dimensions[get_column_letter(c)].width = 28

    # ===== Sheet 5: ASIC vs GPU Growth =====
    ws5 = wb.create_sheet("ASIC vs GPU Growth")
    ws5.sheet_properties.tabColor = ASIC_PURPLE

    add_title(ws5, 1, "Custom ASIC vs GPU Growth Rate Comparison (through 2033)")
    add_subtitle(ws5, 2, "ASICs growing at 44.6% CAGR vs GPUs at 16.1% — driven by inference economics")

    row = 4
    cagr_headers = ["Market Segment", "CAGR", "2024 Revenue", "2033 Projected", "Primary Use Case"]
    for c, h in enumerate(cagr_headers, 1):
        ws5.cell(row=row, column=c, value=h)
    style_header_row(ws5, row, len(cagr_headers))

    row = 5
    for data_row in CAGR_DATA:
        for c, val in enumerate(data_row, 1):
            cell = ws5.cell(row=row, column=c, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left")
        row += 1

    add_source_note(ws5, row + 1, "Source: Bloomberg Intelligence, AI Accelerator Market Forecast (January 2026)")

    # Key takeaways
    row += 3
    add_subtitle(ws5, row, "Key Drivers of ASIC Outperformance")
    row += 1
    drivers = [
        "1. Inference workloads now 67% of compute (2026) — ASICs purpose-built for predictable, cost-sensitive inference",
        "2. Hyperscaler vertical integration — Google, Microsoft, Amazon, Meta each investing $60-80B in AI capex (2026)",
        "3. Cost advantage: Midjourney cut inference costs 65% migrating from NVIDIA GPUs to Google TPUs",
        "4. Power efficiency: TPU v7 ~500W vs NVIDIA Vera Rubin ~1000W — 2× better at data center scale",
        "5. TSMC 3nm capacity running at 100% utilization — demand 3× supply; favors committed hyperscaler orders",
    ]
    for d in drivers:
        ws5.cell(row=row, column=1, value=d)
        ws5.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws5.cell(row=row, column=1).font = DATA_FONT
        row += 1

    for c in range(1, len(cagr_headers) + 1):
        ws5.column_dimensions[get_column_letter(c)].width = 30

    # ===== Sheet 6: 2026 Chip Specs =====
    ws6 = wb.create_sheet("2026 Chip Specs")
    ws6.sheet_properties.tabColor = "FF6600"

    add_title(ws6, 1, "2026-Class AI Accelerator Spec Comparison", merge_end=10)
    add_subtitle(ws6, 2, "Comprehensive chip-by-chip comparison: NVIDIA, Google/Broadcom, AMD, Microsoft, AWS", merge_end=10)

    row = 4
    spec_headers = ["Chip", "Vendor", "Process", "FP8 (PFLOPS)", "FP4 (PFLOPS)",
                     "Memory (GB)", "Mem BW (TB/s)", "TDP (W)", "Interconnect", "Primary Workload"]
    for c, h in enumerate(spec_headers, 1):
        ws6.cell(row=row, column=c, value=h)
    style_header_row(ws6, row, len(spec_headers))

    row = 5
    vendor_fills = {
        "NVIDIA": PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"),
        "Google/Broadcom": PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid"),
        "AMD": PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid"),
        "Microsoft/Broadcom": PatternFill(start_color="F3E5F5", end_color="F3E5F5", fill_type="solid"),
        "AWS": PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid"),
    }

    for spec_row in CHIP_SPECS:
        vendor = spec_row[1]
        fill = vendor_fills.get(vendor, None)
        for c, val in enumerate(spec_row, 1):
            cell = ws6.cell(row=row, column=c, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c > 2 else "left", wrap_text=True)
            if fill:
                cell.fill = fill
        row += 1

    add_source_note(ws6, row + 1, "Sources: NVIDIA GTC 2026, Google Cloud Next 2025, AMD CES 2026, AWS re:Invent 2025, Microsoft Azure Blog (2026)", merge_end=10)

    for c in range(1, len(spec_headers) + 1):
        ws6.column_dimensions[get_column_letter(c)].width = 20

    # ===== Sheet 7: Inference TCO =====
    ws7 = wb.create_sheet("Inference TCO")
    ws7.sheet_properties.tabColor = "009688"

    add_title(ws7, 1, "3-Year TCO Comparison: 1,000-Chip Inference Cluster")
    add_subtitle(ws7, 2, "NVIDIA H100/B200 vs Google TPU v6/v7 vs AMD MI350/MI400 — 24/7 inference at 80% utilization")

    row = 4
    tco_headers = ["Cost Factor", "NVIDIA H100/B200", "Google TPU v6/v7", "AMD MI350/MI400", "Winner"]
    for c, h in enumerate(tco_headers, 1):
        ws7.cell(row=row, column=c, value=h)
    style_header_row(ws7, row, len(tco_headers))

    row = 5
    for tco_row in TCO_DATA:
        is_total = "TOTAL" in tco_row[0]
        for c, val in enumerate(tco_row, 1):
            cell = ws7.cell(row=row, column=c, value=val)
            cell.font = Font(name="Calibri", size=10, bold=is_total)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left")
            if is_total:
                cell.fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
        row += 1

    add_source_note(ws7, row + 1, "Sources: Google Cloud TCO calculators, NVIDIA DGX pricing, Uptime Institute datacenter energy audits")

    row += 3
    add_subtitle(ws7, row, "Key TCO Takeaways")
    row += 1
    takeaways = [
        "Google TPU delivers 56% lower 3-year TCO vs NVIDIA for dedicated inference workloads",
        "Power consumption is the largest differentiator: TPU v7 ~500W vs NVIDIA B200 1000W",
        "AMD positioned between NVIDIA and TPU — 27% lower TCO than NVIDIA, 65% higher than TPU",
        "Software migration cost typically $80K-$200K; payback period 18-48 days at $105K/mo savings",
        "At Meta scale (600K+ chips), the TCO delta represents ~$59B over hardware lifecycle",
    ]
    for t in takeaways:
        ws7.cell(row=row, column=1, value=t)
        ws7.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws7.cell(row=row, column=1).font = DATA_FONT
        row += 1

    for c in range(1, len(tco_headers) + 1):
        ws7.column_dimensions[get_column_letter(c)].width = 28

    # ===== Sheet 8: Broadcom Revenue =====
    ws8 = wb.create_sheet("Broadcom ASIC Revenue")
    ws8.sheet_properties.tabColor = BROADCOM_RED

    add_title(ws8, 1, "Broadcom AI Revenue & Google TPU Run-Rate Derivation", merge_end=8)
    add_subtitle(ws8, 2, "Based on Broadcom reported AI semiconductor revenue, earnings call commentary, and analyst estimates", merge_end=8)

    # Part 1: Annual revenue breakdown
    row = 4
    add_subtitle(ws8, row, "Annual AI Revenue Breakdown ($B)", merge_end=8)
    row += 1
    bc_headers = ["Revenue Segment"] + BROADCOM_YEARS
    for c, h in enumerate(bc_headers, 1):
        ws8.cell(row=row, column=c, value=h)
    style_header_row(ws8, row, len(bc_headers))

    row += 1
    for segment, vals in BROADCOM_DATA.items():
        is_total = "Total" in segment
        ws8.cell(row=row, column=1, value=segment)
        for c, val in enumerate(vals, 2):
            cell = ws8.cell(row=row, column=c, value=val)
            cell.number_format = USD_FMT
            cell.font = Font(name="Calibri", size=10, bold=is_total)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            if is_total:
                cell.fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
        ws8.cell(row=row, column=1).alignment = Alignment(horizontal="left")
        ws8.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=is_total)
        ws8.cell(row=row, column=1).border = THIN_BORDER
        row += 1

    add_source_note(ws8, row, "FY2023-FY2025 = Broadcom reported actuals; FY2026E-FY2027E = guidance + analyst est.", merge_end=8)
    row += 2

    # Part 2: Quarterly run-rate derivation
    add_subtitle(ws8, row, "Quarterly Run-Rate Derivation: How to Estimate Current TPU Revenue", merge_end=8)
    row += 1
    q_headers = BROADCOM_QUARTERLY["headers"]
    for c, h in enumerate(q_headers, 1):
        ws8.cell(row=row, column=c, value=h)
    style_header_row(ws8, row, len(q_headers))

    q_data_start = row + 1
    row += 1
    highlight_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    for q_row in BROADCOM_QUARTERLY["rows"]:
        is_latest = "actual" in q_row[0] or "guide" in q_row[0]
        for c, val in enumerate(q_row, 1):
            cell = ws8.cell(row=row, column=c, value=val)
            cell.font = Font(name="Calibri", size=10, bold=is_latest)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left")
            if is_latest:
                cell.fill = highlight_fill
        row += 1

    row += 1
    add_subtitle(ws8, row, "Derivation Notes", merge_end=8)
    row += 1
    for note in BROADCOM_QUARTERLY["notes"]:
        ws8.cell(row=row, column=1, value=note)
        ws8.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(q_headers))
        ws8.cell(row=row, column=1).font = Font(name="Calibri", size=9, italic=True)
        row += 1

    row += 1
    add_subtitle(ws8, row, "Key Conclusion", merge_end=8)
    row += 1
    conclusions = [
        "Current Google TPU revenue through Broadcom: ~$3.4B/quarter = ~$13.6B annualized run rate (Q1 FY2026)",
        "Based on: $8.4B AI revenue × (1 - 33% networking) × ~60% Google share of XPU = ~$3.4B",
        "This is Broadcom's revenue for designing/manufacturing TPUs — not the total value of TPU compute deployed",
        "Broadcom total AI run rate: $8.4B/qtr = $33.6B annualized; guided to $10.7B/qtr = $42.8B ann. by Q2",
        "FY2025 full-year AI revenue was $20B (reported); FY2026E consensus ~$38-46B; FY2027 mgmt target >$100B",
    ]
    for conc in conclusions:
        ws8.cell(row=row, column=1, value=conc)
        ws8.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(q_headers))
        ws8.cell(row=row, column=1).font = DATA_FONT
        row += 1

    add_source_note(ws8, row + 1, "Sources: Broadcom Q1 FY2026 earnings (March 4, 2026), Q4 FY2025 earnings, Reuters, JPMorgan", merge_end=8)
    row += 2

    chart8 = BarChart()
    chart8.type = "col"
    chart8.grouping = "stacked"
    chart8.title = "Broadcom AI Revenue by Segment ($B)"
    chart8.x_axis.title = "Fiscal Year"
    chart8.y_axis.title = "Revenue ($B)"
    chart8.style = 10
    chart8.width = 20
    chart8.height = 13

    cats8 = Reference(ws8, min_col=2, max_col=len(BROADCOM_YEARS) + 1, min_row=5)
    bc_colors = [GOOGLE_BLUE, ASIC_PURPLE, "FF9900"]
    for i, (segment, _) in enumerate(list(BROADCOM_DATA.items())[:3]):
        vals = Reference(ws8, min_col=2, max_col=len(BROADCOM_YEARS) + 1, min_row=6 + i)
        s = Series(vals, title=segment)
        if i < len(bc_colors):
            s.graphicalProperties.solidFill = bc_colors[i]
        chart8.append(s)
    chart8.set_categories(cats8)
    ws8.add_chart(chart8, "A" + str(row + 1))

    for c in range(1, max(len(bc_headers), len(q_headers)) + 1):
        ws8.column_dimensions[get_column_letter(c)].width = 22

    # ===== Sheet 9: Migration Case Studies =====
    ws9 = wb.create_sheet("Migration Case Studies")
    ws9.sheet_properties.tabColor = "00BCD4"

    add_title(ws9, 1, "Real-World Migration Case Studies: NVIDIA GPU → TPU/ASIC", merge_end=5)
    add_subtitle(ws9, 2, "Major AI companies voting with their wallets — inference economics drive migration", merge_end=5)

    row = 4
    cs_headers = ["Company", "Migration Path", "Cost Impact", "Scale", "Timeline"]
    for c, h in enumerate(cs_headers, 1):
        ws9.cell(row=row, column=c, value=h)
    style_header_row(ws9, row, len(cs_headers))

    row = 5
    for cs_row in CASE_STUDIES:
        for c, val in enumerate(cs_row, 1):
            cell = ws9.cell(row=row, column=c, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left", wrap_text=True)
        row += 1

    add_source_note(ws9, row + 1, "Sources: The Information, Anthropic Blog, Reuters, company disclosures (2025-2026)", merge_end=5)

    for c in range(1, len(cs_headers) + 1):
        ws9.column_dimensions[get_column_letter(c)].width = 28

    # ===== Sheet 10: Sources =====
    ws10 = wb.create_sheet("Sources")
    ws10.sheet_properties.tabColor = "607D8B"

    add_title(ws10, 1, "Data Sources and References")

    row = 3
    src_headers = ["Source", "Data Point / Coverage", "Date"]
    for c, h in enumerate(src_headers, 1):
        ws10.cell(row=row, column=c, value=h)
    style_header_row(ws10, row, len(src_headers))

    row = 4
    for src in SOURCES:
        for c, val in enumerate(src, 1):
            cell = ws10.cell(row=row, column=c, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left", wrap_text=True)
        row += 1

    for c in range(1, len(src_headers) + 1):
        ws10.column_dimensions[get_column_letter(c)].width = 45 if c == 2 else 30

    return wb


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("Building Broadcom-Google TPU vs NVIDIA vs AMD market share workbook...")
    wb = build_workbook()
    output_path = "tpu_market_share_comparison.xlsx"
    wb.save(output_path)
    print(f"Saved to {output_path}")
    print("\nSheets created:")
    for ws in wb.worksheets:
        print(f"  - {ws.title}")


if __name__ == "__main__":
    main()
