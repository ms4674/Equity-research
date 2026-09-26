#!/usr/bin/env python3
"""Build the aggregated token-pricing spreadsheet from data/token_pricing_events.csv.

Outputs:
  - token_pricing_frontier_labs.xlsx  (multi-sheet workbook with charts)
  - data/flagship_timeseries_monthly.csv  (wide-format monthly flagship matrix)
"""

import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
EVENTS_CSV = ROOT / "data" / "token_pricing_events.csv"
MONTHLY_CSV = ROOT / "data" / "flagship_timeseries_monthly.csv"
XLSX_OUT = ROOT / "token_pricing_frontier_labs.xlsx"

AS_OF = date(2026, 9, 26)
LABS = ["OpenAI", "Anthropic", "Google", "xAI", "Meta", "DeepSeek", "Mistral", "Alibaba"]

# Each lab's frontier flagship price in force over time: (effective date, model, input, output).
# A new tuple is added whenever the flagship changes model or the incumbent's price moves.
FLAGSHIP_SERIES = {
    "OpenAI": [
        (date(2020, 6, 11), "GPT-3 (davinci)", 60.00, 60.00),
        (date(2022, 9, 1), "GPT-3 (davinci, cut)", 20.00, 20.00),
        (date(2023, 3, 14), "GPT-4", 30.00, 60.00),
        (date(2023, 11, 6), "GPT-4 Turbo", 10.00, 30.00),
        (date(2024, 5, 13), "GPT-4o", 5.00, 15.00),
        (date(2024, 8, 6), "GPT-4o (Aug 2024 cut)", 2.50, 10.00),
        (date(2025, 4, 14), "GPT-4.1", 2.00, 8.00),
        (date(2025, 8, 7), "GPT-5", 1.25, 10.00),
        (date(2025, 11, 13), "GPT-5.1", 1.25, 10.00),
        (date(2025, 12, 11), "GPT-5.2", 1.75, 14.00),
        (date(2026, 2, 5), "GPT-5.3", 1.75, 14.00),
        (date(2026, 4, 23), "GPT-5.5", 5.00, 30.00),
        (date(2026, 7, 9), "GPT-5.6 Sol", 5.00, 30.00),
        (date(2026, 8, 21), "GPT-5.6 Sol (cut)", 4.00, 20.00),
        (date(2026, 9, 3), "GPT-6 Astra", 10.00, 50.00),
    ],
    "Anthropic": [
        (date(2023, 3, 14), "Claude 1", 11.02, 32.68),
        (date(2023, 7, 11), "Claude 2", 11.02, 32.68),
        (date(2023, 11, 21), "Claude 2.1", 8.00, 24.00),
        (date(2024, 3, 4), "Claude 3 Opus", 15.00, 75.00),
        (date(2025, 5, 22), "Claude Opus 4", 15.00, 75.00),
        (date(2025, 8, 5), "Claude Opus 4.1", 15.00, 75.00),
        (date(2025, 11, 24), "Claude Opus 4.5", 5.00, 25.00),
        (date(2026, 2, 5), "Claude Opus 4.6", 5.00, 25.00),
        (date(2026, 4, 16), "Claude Opus 4.7", 5.00, 25.00),
        (date(2026, 5, 28), "Claude Opus 4.8", 5.00, 25.00),
        (date(2026, 6, 9), "Claude Fable 5", 10.00, 50.00),
        (date(2026, 9, 1), "Claude Fable 5.1", 10.00, 50.00),
    ],
    "Google": [
        (date(2024, 2, 15), "Gemini 1.0 Pro", 0.50, 1.50),
        (date(2024, 5, 14), "Gemini 1.5 Pro", 3.50, 10.50),
        (date(2024, 10, 1), "Gemini 1.5 Pro (cut)", 1.25, 5.00),
        (date(2025, 3, 25), "Gemini 2.5 Pro", 1.25, 10.00),
        (date(2025, 11, 18), "Gemini 3 Pro", 2.00, 12.00),
        (date(2026, 5, 19), "Gemini 3.1 Pro", 2.00, 12.00),
        (date(2026, 8, 13), "Gemini 3.7 Flash", 0.75, 3.75),
        (date(2026, 9, 2), "Gemini 3.8 Flash", 0.75, 3.75),
    ],
    "xAI": [
        (date(2024, 12, 12), "Grok 2", 2.00, 10.00),
        (date(2025, 2, 17), "Grok 3", 3.00, 15.00),
        (date(2025, 7, 9), "Grok 4", 3.00, 15.00),
        (date(2026, 4, 30), "Grok 4.3", 1.25, 2.50),
        (date(2026, 7, 8), "Grok 4.5", 2.00, 6.00),
        (date(2026, 8, 12), "Grok 4.6", 2.00, 6.00),
        (date(2026, 9, 21), "Grok 4.7", 2.00, 6.00),
    ],
    "Meta": [
        (date(2023, 10, 1), "Llama 2 70B (Together)", 0.90, 0.90),
        (date(2024, 4, 18), "Llama 3 70B (Together)", 0.90, 0.90),
        (date(2024, 7, 23), "Llama 3.1 405B (Together)", 3.50, 3.50),
        (date(2025, 4, 5), "Llama 4 Maverick (Together)", 0.27, 0.85),
        (date(2026, 7, 9), "Muse Spark 1.1", 1.25, 4.25),
        (date(2026, 8, 5), "Muse Spark 1.2", 1.25, 4.25),
        (date(2026, 9, 2), "Muse Spark 1.3", 1.25, 4.25),
    ],
    "DeepSeek": [
        (date(2024, 5, 6), "DeepSeek-V2", 0.14, 0.28),
        (date(2024, 12, 26), "DeepSeek-V3 (promo)", 0.14, 0.28),
        (date(2025, 2, 8), "DeepSeek-V3", 0.27, 1.10),
        (date(2025, 8, 21), "DeepSeek-V3.1", 0.56, 1.68),
        (date(2025, 9, 29), "DeepSeek-V3.2-Exp", 0.28, 0.42),
        (date(2026, 4, 24), "DeepSeek-V4-Pro", 1.74, 3.48),
        (date(2026, 6, 1), "DeepSeek-V4-Pro (cut)", 0.43, 0.87),
        (date(2026, 8, 16), "DeepSeek-V4-Pro (raise, off-peak)", 0.66, 1.98),
        (date(2026, 9, 10), "DeepSeek-V4.1-Flash (off-peak)", 0.15, 0.60),
    ],
    "Mistral": [
        (date(2023, 12, 11), "Mistral Medium (original)", 2.50, 7.50),
        (date(2024, 2, 26), "Mistral Large", 8.00, 24.00),
        (date(2024, 7, 24), "Mistral Large 2", 3.00, 9.00),
        (date(2025, 12, 2), "Mistral Large 3", 2.00, 6.00),
        (date(2026, 4, 28), "Mistral Medium 3.5", 1.50, 7.50),
    ],
    "Alibaba": [
        (date(2025, 1, 25), "Qwen Max (intl)", 1.60, 6.40),
        (date(2026, 5, 15), "Qwen3.7 Max", 2.50, 7.50),
        (date(2026, 8, 3), "Qwen3.8 Max", 2.00, 6.00),
    ],
}

# Cheapest generally-available tier per lab as of AS_OF (from provider pricing pages).
CHEAPEST_TIER = {
    "OpenAI": ("GPT-5 nano", 0.05, 0.40),
    "Anthropic": ("Claude Haiku 4.5", 1.00, 5.00),
    "Google": ("Gemini 2.5 Flash-Lite", 0.10, 0.40),
    "xAI": ("Grok Composer 2.5", 0.50, 2.50),
    "Meta": ("Llama 4 Scout (Together)", 0.18, 0.59),
    "DeepSeek": ("DeepSeek-V4.1-Flash (off-peak)", 0.15, 0.60),
    "Mistral": ("Mistral Small 4", 0.15, 0.60),
    "Alibaba": ("Qwen3.8-Flash", 0.15, 0.47),
}

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def month_ends(start: date, end: date):
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
        last = date(ny, nm, 1).toordinal() - 1
        yield date.fromordinal(min(last, end.toordinal()))
        y, m = ny, nm


def flagship_at(lab: str, day: date):
    """Return the (model, input, output) in force for a lab on the given day, or None."""
    current = None
    for eff, model, inp, out in FLAGSHIP_SERIES[lab]:
        if eff <= day:
            current = (model, inp, out)
        else:
            break
    return current


def style_header(ws, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER


def build_readme(wb: Workbook):
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 118
    lines = [
        ("Token pricing for frontier AI labs — global time series", TITLE_FONT),
        ("", None),
        (f"Compiled {AS_OF.isoformat()}. All prices are USD per 1 million tokens, public list price on the provider's own API.", None),
        ("", None),
        ("Sheets", Font(bold=True, size=12)),
        ("  1. Price Events — every published price point (launches, cuts, raises) in long format: one row per event.", None),
        ("  2. Flagship Input $-per-1M — monthly matrix (Jan 2023 to Sep 2026) of each lab's frontier flagship input price, with chart.", None),
        ("  3. Flagship Output $-per-1M — same matrix for output prices, with chart.", None),
        ("  4. Lab Summary — current flagship and cheapest tier per lab as of the compile date.", None),
        ("", None),
        ("Labs covered", Font(bold=True, size=12)),
        ("  US: OpenAI, Anthropic, Google (DeepMind), xAI, Meta.  Europe: Mistral (France).  China: DeepSeek, Alibaba (Qwen).", None),
        ("", None),
        ("Methodology and caveats", Font(bold=True, size=12)),
        ("  - Prices are synchronous-API list rates. Batch tiers (typically 50% off), negotiated/enterprise rates, and free tiers are out of scope.", None),
        ("  - 'Flagship' means the model the lab itself positions as its frontier offering at that date. Where a flagship has tiered pricing", None),
        ("    (e.g. long-context surcharges), the base/short-context tier is recorded.", None),
        ("  - Meta had no first-party API before the Meta Model API (July 2026); earlier Llama rows use Together AI's hosted rate as a reference proxy.", None),
        ("  - Google shifted its flagship designation from the Pro line to the Flash line in mid-2026; the flagship series follows that shift at Gemini 3.7 Flash (Aug 2026).", None),
        ("  - DeepSeek bills peak/off-peak since Aug 2026; off-peak (133 of 168 hours per week) is recorded.", None),
        ("  - Qwen prices are Alibaba's International (Singapore) USD list, not the Chinese-mainland table.", None),
        ("  - Cached-input rates are prompt-cache READ rates where the provider publishes one.", None),
        ("  - Claude 1/2 pricing was originally quoted per character; the standard token-equivalent conversion ($11.02/$32.68) is used.", None),
        ("  - Tokenizers differ across (and within) providers, so $1/M tokens does not buy identical text volume across labs.", None),
        ("", None),
        ("Sources", Font(bold=True, size=12)),
        ("  Provider announcements and pricing pages (openai.com/api/pricing, platform.claude.com pricing docs, ai.google.dev, docs.x.ai,", None),
        ("    api-docs.deepseek.com, docs.mistral.ai, Alibaba Model Studio, Meta Model API), Google Developers Blog posts,", None),
        ("    cross-checked against independent trackers: mungomash.com/ai/pricing-history, tokenomy.ai/price-timeline,", None),
        ("    kickllm.com and zerlo.net pricing histories, and Wayback Machine snapshots cited therein.", None),
        ("", None),
        ("  Raw data lives in data/token_pricing_events.csv in this repository; this workbook is generated by scripts/build_spreadsheet.py.", None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=text)
        if font:
            cell.font = font


def build_events(wb: Workbook, rows):
    ws = wb.create_sheet("Price Events")
    headers = ["Date", "Lab", "HQ", "Model", "Tier", "Event",
               "Input $/1M", "Output $/1M", "Cached input $/1M", "Notes", "Source"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for r in rows:
        ws.append([
            r["date"], r["lab"], r["lab_hq"], r["model"], r["tier"], r["event"],
            float(r["input_usd_per_1m"]) if r["input_usd_per_1m"] else None,
            float(r["output_usd_per_1m"]) if r["output_usd_per_1m"] else None,
            float(r["cached_input_usd_per_1m"]) if r["cached_input_usd_per_1m"] else None,
            r["notes"], r["source"],
        ])
    widths = [11, 10, 8, 34, 12, 24, 11, 11, 13, 68, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = BORDER
            if cell.column in (7, 8, 9) and cell.value is not None:
                cell.number_format = "$#,##0.00###"
            if cell.column in (10, 11):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def build_flagship_sheet(wb: Workbook, title: str, price_index: int, months):
    """price_index: 1 for input price, 2 for output price."""
    ws = wb.create_sheet(title)
    headers = ["Month"] + LABS
    ws.append(headers)
    style_header(ws, 1, len(headers))
    matrix_rows = []
    for m in months:
        label = f"{m.year}-{m.month:02d}"
        row = [label]
        for lab in LABS:
            snap = flagship_at(lab, m)
            row.append(round(snap[price_index], 4) if snap else None)
        ws.append(row)
        matrix_rows.append(row)
    ws.column_dimensions["A"].width = 10
    for i in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = BORDER
            if cell.column > 1 and cell.value is not None:
                cell.number_format = "$#,##0.00"
    ws.freeze_panes = "B2"

    chart = LineChart()
    chart.title = f"Frontier flagship {'input' if price_index == 1 else 'output'} price, USD per 1M tokens (log scale)"
    chart.style = 12
    chart.height = 12
    chart.width = 30
    chart.y_axis.title = "USD per 1M tokens"
    chart.y_axis.scaling.logBase = 10
    chart.x_axis.title = "Month"
    data = Reference(ws, min_col=2, max_col=1 + len(LABS), min_row=1, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"{get_column_letter(len(headers) + 2)}2")
    return matrix_rows


def build_summary(wb: Workbook):
    ws = wb.create_sheet("Lab Summary")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value=f"Frontier labs, state of pricing as of {AS_OF.isoformat()}").font = TITLE_FONT
    headers = ["Lab", "HQ", "Current flagship", "Flagship input $/1M", "Flagship output $/1M",
               "Cheapest GA tier", "Cheapest input $/1M", "Cheapest output $/1M"]
    hq = {"OpenAI": "US", "Anthropic": "US", "Google": "US", "xAI": "US",
          "Meta": "US", "DeepSeek": "China", "Mistral": "Europe", "Alibaba": "China"}
    ws.append([])
    ws.append(headers)
    style_header(ws, 3, len(headers))
    for lab in LABS:
        model, inp, out = flagship_at(lab, AS_OF)
        cmodel, cin, cout = CHEAPEST_TIER[lab]
        ws.append([lab, hq[lab], model, inp, out, cmodel, cin, cout])
    widths = [11, 9, 32, 17, 18, 32, 17, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for cell in row:
            cell.border = BORDER
            if cell.column in (4, 5, 7, 8):
                cell.number_format = "$#,##0.00"


def main():
    with EVENTS_CSV.open() as f:
        rows = sorted(csv.DictReader(f), key=lambda r: (r["date"], r["lab"], r["model"]))

    wb = Workbook()
    build_readme(wb)
    build_events(wb, rows)
    months = list(month_ends(date(2023, 1, 1), AS_OF))
    input_matrix = build_flagship_sheet(wb, "Flagship Input $-per-1M", 1, months)
    output_matrix = build_flagship_sheet(wb, "Flagship Output $-per-1M", 2, months)
    build_summary(wb)
    wb.save(XLSX_OUT)

    with MONTHLY_CSV.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month"] + [f"{lab}_input" for lab in LABS] + [f"{lab}_output" for lab in LABS])
        for in_row, out_row in zip(input_matrix, output_matrix):
            w.writerow(in_row + out_row[1:])

    print(f"Wrote {XLSX_OUT.relative_to(ROOT)} ({len(rows)} price events, {len(months)} monthly samples)")
    print(f"Wrote {MONTHLY_CSV.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
