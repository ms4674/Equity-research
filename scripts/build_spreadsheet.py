#!/usr/bin/env python3
"""Build an XLSX spreadsheet comparing open-weights vs closed LLM token pricing,
including weekly token volume per model.

Inputs (snapshots in data/raw/):
  - openrouter_rankings_week_*.json : per-model weekly token totals
      (https://openrouter.ai/api/frontend/v1/rankings/models?period=week)
  - openrouter_models_*.json        : model catalog with pricing + hugging_face_id
      (https://openrouter.ai/api/v1/models)

Outputs:
  - data/llm_pricing_vs_token_volume.xlsx
  - data/csv/model_comparison.csv
  - data/csv/open_vs_closed_summary.csv
  - data/csv/by_developer.csv
"""

import csv
import glob
import json
import os
import re
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
OUT_XLSX = os.path.join(ROOT, "data", "llm_pricing_vs_token_volume.xlsx")
OUT_CSV_DIR = os.path.join(ROOT, "data", "csv")

WEEK_ENDING = "2026-07-16"

# Models whose openness cannot be inferred from the catalog signals below.
# True = open-weights, False = closed. Keyed by catalog id (without :variant).
MANUAL_OVERRIDES = {
    # Description states "open-weight" but no hugging_face_id in catalog yet.
    "moonshotai/kimi-k3": True,
    # MiniMax M1 was released with open weights (Apache 2.0) in 2025.
    "minimax/minimax-m1": True,
}

OPEN_KEYWORDS = re.compile(
    r"open[- ]?(weight|source)|MIT licen[cs]e|Apache[- ]2", re.IGNORECASE
)


def load_raw():
    rankings_path = sorted(glob.glob(os.path.join(RAW, "openrouter_rankings_week_*.json")))[-1]
    models_path = sorted(glob.glob(os.path.join(RAW, "openrouter_models_*.json")))[-1]
    with open(rankings_path) as f:
        rankings = json.load(f)["data"]
    with open(models_path) as f:
        models = json.load(f)["data"]
    return rankings, models


def index_catalog(models):
    """Map canonical_slug -> catalog entry, preferring the standard (non-free) variant."""
    by_slug = {}
    for entry in models:
        slug = entry["canonical_slug"]
        is_variant = ":" in entry["id"]
        if slug not in by_slug or (":" in by_slug[slug]["id"] and not is_variant):
            by_slug[slug] = entry
    return by_slug


def classify_open(entry):
    base_id = entry["id"].split(":")[0]
    if base_id in MANUAL_OVERRIDES:
        return MANUAL_OVERRIDES[base_id]
    if entry.get("hugging_face_id"):
        return True
    if OPEN_KEYWORDS.search(entry.get("description") or ""):
        return True
    return False


def build_rows(rankings, catalog):
    agg = defaultdict(lambda: {"prompt": 0, "completion": 0, "free": 0, "requests": 0})
    for row in rankings:
        a = agg[row["model_permaslug"]]
        tokens = row["total_prompt_tokens"] + row["total_completion_tokens"]
        a["prompt"] += row["total_prompt_tokens"]
        a["completion"] += row["total_completion_tokens"]
        a["requests"] += row["count"]
        if row["variant"] == "free":
            a["free"] += tokens

    rows, excluded_tokens = [], 0
    for slug, a in agg.items():
        total = a["prompt"] + a["completion"]
        entry = catalog.get(slug)
        if entry is None:
            # Embedding / image / audio models not in the text-model catalog.
            excluded_tokens += total
            continue
        price_in = float(entry["pricing"].get("prompt", 0) or 0) * 1e6
        price_out = float(entry["pricing"].get("completion", 0) or 0) * 1e6
        blended = (3 * price_in + price_out) / 4
        paid_tokens = total - a["free"]
        paid_share = paid_tokens / total if total else 0
        # Spend estimate uses list prices on paid-variant traffic only, with the
        # paid prompt/completion split assumed equal to the overall split.
        est_spend = (a["prompt"] * paid_share * price_in + a["completion"] * paid_share * price_out) / 1e6
        rows.append(
            {
                "model": entry["name"],
                "developer": entry["id"].split("/")[0],
                "slug": slug,
                "open": classify_open(entry),
                "price_in": price_in,
                "price_out": price_out,
                "blended": blended,
                "context": entry.get("context_length") or 0,
                "tokens_total": total,
                "tokens_prompt": a["prompt"],
                "tokens_completion": a["completion"],
                "free_share": a["free"] / total if total else 0,
                "requests": a["requests"],
                "est_spend": est_spend,
            }
        )
    rows.sort(key=lambda r: -r["tokens_total"])
    return rows, excluded_tokens


DEVELOPER_NAMES = {
    "anthropic": "Anthropic", "openai": "OpenAI", "google": "Google",
    "x-ai": "xAI", "deepseek": "DeepSeek", "qwen": "Alibaba (Qwen)",
    "z-ai": "Z.ai (Zhipu)", "moonshotai": "Moonshot AI", "minimax": "MiniMax",
    "xiaomi": "Xiaomi", "tencent": "Tencent", "nvidia": "NVIDIA",
    "meta-llama": "Meta", "mistralai": "Mistral AI", "stepfun": "StepFun",
    "poolside": "Poolside", "cohere": "Cohere", "amazon": "Amazon",
    "microsoft": "Microsoft", "perplexity": "Perplexity", "bytedance": "ByteDance",
}


def summarize(rows):
    groups = {True: [], False: []}
    for r in rows:
        groups[r["open"]].append(r)
    summary = []
    grand_total = sum(r["tokens_total"] for r in rows)
    for is_open, label in ((True, "Open-weights"), (False, "Closed / proprietary")):
        g = groups[is_open]
        tokens = sum(r["tokens_total"] for r in g)
        spend = sum(r["est_spend"] for r in g)
        paid_tokens = sum(r["tokens_total"] * (1 - r["free_share"]) for r in g)
        w_blended = (
            sum(r["blended"] * r["tokens_total"] for r in g) / tokens if tokens else 0
        )
        w_in = sum(r["price_in"] * r["tokens_total"] for r in g) / tokens if tokens else 0
        w_out = sum(r["price_out"] * r["tokens_total"] for r in g) / tokens if tokens else 0
        prices = sorted(r["blended"] for r in g if r["blended"] > 0)
        median = prices[len(prices) // 2] if prices else 0
        summary.append(
            {
                "class": label,
                "models": len(g),
                "tokens": tokens,
                "share": tokens / grand_total if grand_total else 0,
                "w_in": w_in,
                "w_out": w_out,
                "w_blended": w_blended,
                "median_blended": median,
                "est_spend": spend,
                "eff_price": spend / (paid_tokens / 1e6) if paid_tokens else 0,
            }
        )
    return summary, grand_total


def by_developer(rows):
    devs = defaultdict(lambda: {"tokens": 0, "models": 0, "spend": 0, "open": 0, "closed": 0})
    for r in rows:
        d = devs[r["developer"]]
        d["tokens"] += r["tokens_total"]
        d["models"] += 1
        d["spend"] += r["est_spend"]
        d["open" if r["open"] else "closed"] += 1
    out = []
    for dev, d in sorted(devs.items(), key=lambda kv: -kv[1]["tokens"]):
        if d["open"] and d["closed"]:
            mix = "Mixed"
        elif d["open"]:
            mix = "Open-weights"
        else:
            mix = "Closed"
        out.append(
            {
                "developer": DEVELOPER_NAMES.get(dev, dev.title()),
                "slug": dev,
                "class": mix,
                "models": d["models"],
                "tokens": d["tokens"],
                "spend": d["spend"],
            }
        )
    return out


# ---------------------------------------------------------------- XLSX styling

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
OPEN_FILL = PatternFill("solid", fgColor="E2EFDA")
CLOSED_FILL = PatternFill("solid", fgColor="FCE4EC")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_workbook(rows, summary, devs, grand_total, excluded_tokens):
    wb = Workbook()

    # ---- Sheet 1: Summary (open vs closed)
    ws = wb.active
    ws.title = "Open vs Closed Summary"
    ws["A1"] = "Open-weights vs closed LLMs — token pricing and volume"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Weekly usage across OpenRouter, week ending {WEEK_ENDING} (trailing 7 days)"
    ws["A3"] = "Source: OpenRouter (openrouter.ai/rankings), as of 2026-07-16. Prices are OpenRouter list prices per 1M tokens."
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="595959")

    headers = [
        "Class", "# models", "Weekly tokens (B)", "Volume share",
        "Vol-weighted input $/1M", "Vol-weighted output $/1M", "Vol-weighted blended $/1M*",
        "Median blended $/1M (paid models)", "Est. weekly spend ($M)**", "Effective $/1M (paid traffic)",
    ]
    hr = 5
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    for i, s in enumerate(summary):
        r = hr + 1 + i
        fill = OPEN_FILL if s["class"].startswith("Open") else CLOSED_FILL
        vals = [
            s["class"], s["models"], s["tokens"] / 1e9, s["share"],
            s["w_in"], s["w_out"], s["w_blended"], s["median_blended"],
            s["est_spend"] / 1e6, s["eff_price"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = fill
            cell.border = BORDER
        ws.cell(row=r, column=3).number_format = "#,##0.0"
        ws.cell(row=r, column=4).number_format = "0.0%"
        for c in (5, 6, 7, 8, 10):
            ws.cell(row=r, column=c).number_format = "$#,##0.000"
        ws.cell(row=r, column=9).number_format = "$#,##0.0"
    trow = hr + 3
    ws.cell(row=trow, column=1, value="Total (matched text LLMs)").font = Font(bold=True)
    ws.cell(row=trow, column=2, value=sum(s["models"] for s in summary)).font = Font(bold=True)
    ws.cell(row=trow, column=3, value=grand_total / 1e9).font = Font(bold=True)
    ws.cell(row=trow, column=3).number_format = "#,##0.0"
    ws.cell(row=trow, column=4, value=1.0).number_format = "0.0%"
    ws.cell(row=trow, column=9, value=sum(s["est_spend"] for s in summary) / 1e6).number_format = "$#,##0.0"

    notes = [
        "* Blended price = (3 × input + 1 × output) / 4 per 1M tokens; volume-weighted rows weight each model's list price by its total weekly tokens (including free-tier traffic).",
        "** Estimated spend applies list prices to paid-variant traffic only (free-tier tokens excluded); actual spend is lower due to prompt-caching and batch discounts.",
        f"Excluded from all sheets: {excluded_tokens / 1e9:.0f}B tokens ({excluded_tokens / (excluded_tokens + grand_total):.1%} of raw total) from embedding, image, and audio models that are not text LLMs.",
        "Open-weights = model weights are publicly downloadable (Hugging Face listing or explicit open-weight/OSS license). Closed = API-only proprietary models.",
        "Token counts use each upstream provider's own tokenizer, so cross-provider token comparisons are approximate.",
    ]
    for i, n in enumerate(notes):
        cell = ws.cell(row=trow + 2 + i, column=1, value=n)
        cell.font = Font(size=9, color="595959")
    autosize(ws, [24, 10, 16, 12, 18, 18, 18, 22, 16, 18])

    pie = PieChart()
    pie.title = "Weekly token volume share"
    data = Reference(ws, min_col=3, min_row=hr, max_row=hr + 2)
    cats = Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + 2)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)
    pie.height, pie.width = 8, 12
    ws.add_chart(pie, f"A{trow + 9}")

    # ---- Sheet 2: Model comparison
    ws2 = wb.create_sheet("Model Comparison")
    headers2 = [
        "Rank", "Model", "Developer", "Class", "Input $/1M", "Output $/1M",
        "Blended $/1M (3:1)", "Context (tokens)", "Weekly tokens (B)", "Volume share",
        "Prompt tokens (B)", "Completion tokens (B)", "Free-tier share",
        "Weekly requests (M)", "Est. weekly spend ($M)", "OpenRouter slug",
    ]
    for c, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header(ws2, 1, len(headers2))
    ws2.freeze_panes = "A2"
    for i, r in enumerate(rows, start=2):
        fill = OPEN_FILL if r["open"] else CLOSED_FILL
        vals = [
            i - 1, r["model"], DEVELOPER_NAMES.get(r["developer"], r["developer"].title()),
            "Open-weights" if r["open"] else "Closed",
            r["price_in"], r["price_out"], r["blended"], r["context"],
            r["tokens_total"] / 1e9, r["tokens_total"] / grand_total,
            r["tokens_prompt"] / 1e9, r["tokens_completion"] / 1e9,
            r["free_share"], r["requests"] / 1e6, r["est_spend"] / 1e6, r["slug"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.border = BORDER
            if c == 4:
                cell.fill = fill
        for c in (5, 6, 7):
            ws2.cell(row=i, column=c).number_format = "$#,##0.000"
        ws2.cell(row=i, column=8).number_format = "#,##0"
        for c in (9, 11, 12):
            ws2.cell(row=i, column=c).number_format = "#,##0.00"
        ws2.cell(row=i, column=10).number_format = "0.00%"
        ws2.cell(row=i, column=13).number_format = "0.0%"
        ws2.cell(row=i, column=14).number_format = "#,##0.0"
        ws2.cell(row=i, column=15).number_format = "$#,##0.00"
    ws2.auto_filter.ref = f"A1:P{len(rows) + 1}"
    autosize(ws2, [6, 34, 16, 13, 11, 11, 13, 13, 14, 11, 13, 15, 11, 13, 14, 40])

    # ---- Sheet 3: Top 20 chart
    ws3 = wb.create_sheet("Top 20 Volume Chart")
    ws3.cell(row=1, column=1, value="Model")
    ws3.cell(row=1, column=2, value="Weekly tokens (B)")
    ws3.cell(row=1, column=3, value="Class")
    style_header(ws3, 1, 3)
    top = rows[:20]
    for i, r in enumerate(top, start=2):
        ws3.cell(row=i, column=1, value=r["model"])
        ws3.cell(row=i, column=2, value=r["tokens_total"] / 1e9).number_format = "#,##0"
        cell = ws3.cell(row=i, column=3, value="Open-weights" if r["open"] else "Closed")
        cell.fill = OPEN_FILL if r["open"] else CLOSED_FILL
    autosize(ws3, [34, 18, 14])
    chart = BarChart()
    chart.type = "bar"
    chart.title = f"Top 20 LLMs by weekly token volume (week ending {WEEK_ENDING})"
    chart.y_axis.title = "Tokens (billions)"
    data = Reference(ws3, min_col=2, min_row=1, max_row=len(top) + 1)
    cats = Reference(ws3, min_col=1, min_row=2, max_row=len(top) + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 14, 24
    chart.legend = None
    ws3.add_chart(chart, "E2")

    # ---- Sheet 4: By developer
    ws4 = wb.create_sheet("By Developer")
    headers4 = ["Developer", "Class", "# models", "Weekly tokens (B)", "Volume share", "Est. weekly spend ($M)"]
    for c, h in enumerate(headers4, 1):
        ws4.cell(row=1, column=c, value=h)
    style_header(ws4, 1, len(headers4))
    ws4.freeze_panes = "A2"
    for i, d in enumerate(devs, start=2):
        fill = (
            OPEN_FILL if d["class"] == "Open-weights"
            else CLOSED_FILL if d["class"] == "Closed"
            else PatternFill("solid", fgColor="FFF2CC")
        )
        vals = [d["developer"], d["class"], d["models"], d["tokens"] / 1e9,
                d["tokens"] / grand_total, d["spend"] / 1e6]
        for c, v in enumerate(vals, 1):
            cell = ws4.cell(row=i, column=c, value=v)
            cell.border = BORDER
            if c == 2:
                cell.fill = fill
        ws4.cell(row=i, column=4).number_format = "#,##0.0"
        ws4.cell(row=i, column=5).number_format = "0.00%"
        ws4.cell(row=i, column=6).number_format = "$#,##0.00"
    ws4.auto_filter.ref = f"A1:F{len(devs) + 1}"
    autosize(ws4, [22, 14, 10, 16, 12, 18])

    wb.save(OUT_XLSX)


def write_csvs(rows, summary, devs, grand_total):
    os.makedirs(OUT_CSV_DIR, exist_ok=True)
    with open(os.path.join(OUT_CSV_DIR, "model_comparison.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "rank", "model", "developer", "class", "input_usd_per_1m", "output_usd_per_1m",
            "blended_usd_per_1m_3to1", "context_tokens", "weekly_tokens", "volume_share",
            "prompt_tokens", "completion_tokens", "free_tier_share", "weekly_requests",
            "est_weekly_spend_usd", "openrouter_slug",
        ])
        for i, r in enumerate(rows, 1):
            w.writerow([
                i, r["model"], DEVELOPER_NAMES.get(r["developer"], r["developer"].title()),
                "open-weights" if r["open"] else "closed",
                round(r["price_in"], 4), round(r["price_out"], 4), round(r["blended"], 4),
                r["context"], r["tokens_total"], round(r["tokens_total"] / grand_total, 6),
                r["tokens_prompt"], r["tokens_completion"], round(r["free_share"], 4),
                r["requests"], round(r["est_spend"], 2), r["slug"],
            ])
    with open(os.path.join(OUT_CSV_DIR, "open_vs_closed_summary.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "class", "models", "weekly_tokens", "volume_share",
            "vol_weighted_input_usd_per_1m", "vol_weighted_output_usd_per_1m",
            "vol_weighted_blended_usd_per_1m", "median_blended_usd_per_1m_paid",
            "est_weekly_spend_usd", "effective_usd_per_1m_paid",
        ])
        for s in summary:
            w.writerow([
                s["class"], s["models"], s["tokens"], round(s["share"], 6),
                round(s["w_in"], 4), round(s["w_out"], 4), round(s["w_blended"], 4),
                round(s["median_blended"], 4), round(s["est_spend"], 2), round(s["eff_price"], 4),
            ])
    with open(os.path.join(OUT_CSV_DIR, "by_developer.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["developer", "class", "models", "weekly_tokens", "volume_share", "est_weekly_spend_usd"])
        for d in devs:
            w.writerow([
                d["developer"], d["class"], d["models"], d["tokens"],
                round(d["tokens"] / grand_total, 6), round(d["spend"], 2),
            ])


def main():
    rankings, models = load_raw()
    catalog = index_catalog(models)
    rows, excluded = build_rows(rankings, catalog)
    summary, grand_total = summarize(rows)
    devs = by_developer(rows)
    write_workbook(rows, summary, devs, grand_total, excluded)
    write_csvs(rows, summary, devs, grand_total)

    print(f"Matched text LLMs: {len(rows)}  |  weekly tokens: {grand_total / 1e12:.2f}T")
    print(f"Excluded (embeddings/media): {excluded / 1e9:.0f}B tokens")
    for s in summary:
        print(
            f"  {s['class']:<22} models={s['models']:>3}  tokens={s['tokens'] / 1e12:6.2f}T "
            f"({s['share']:5.1%})  w-blended=${s['w_blended']:.3f}/1M  "
            f"median=${s['median_blended']:.3f}/1M  est-spend=${s['est_spend'] / 1e6:.1f}M"
        )
    print(f"Wrote {OUT_XLSX}")


if __name__ == "__main__":
    main()
