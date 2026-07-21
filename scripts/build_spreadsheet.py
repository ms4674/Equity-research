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
  - data/csv/params_benchmarks.csv
  - data/csv/tool_use_benchmarks.csv
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

# ------------------------------------------------------------------ Tool use
# Hand-curated tool-use / agentic benchmark comparison (researched 2026-07-21).
# Primary source for K3 / Fable 5 / Sol / Opus 4.8 / GLM-5.2 rows: Moonshot's
# Kimi K3 launch evaluation table (kimi.com/blog/kimi-k3) - the only published
# single-methodology table covering all of them (MCP Atlas: 500-task public
# subset, 100-turn limit, Gemini 3.1 Pro judge; AutomationBench: 600-task
# public subset; Toolathlon-Verified; BrowseComp with 300K context compaction;
# GDPval-AA Elo cited from artificialanalysis.ai). DeepSeek / MiniMax / Kimi
# K2.6 / Nemotron rows come from their own vendor reports and AA/BenchLM
# mirrors, so cross-row comparisons outside the Moonshot block are looser.
TOOL_USE = [
    {
        "slug": "moonshotai/kimi-k3-20260715",
        "model": "Kimi K3", "developer": "Moonshot AI", "open": True,
        "mcp_atlas": 84.2, "toolathlon": 73.2, "automation": 30.8,
        "browsecomp": 91.2, "gdpval_elo": 1668, "tau2": None,
        "notes": "Moonshot launch table (reasoning max, KimiCode/Claude Code "
                 "harnesses). #1 on AutomationBench-AA (53%) per AA; also leads "
                 "SpreadsheetBench 2 (34.8) and DeepSearchQA (95.0 F1). "
                 "BrowseComp hits 90.4 even without context compaction.",
    },
    {
        "slug": "anthropic/claude-5-fable-20260609",
        "model": "Claude Fable 5", "developer": "Anthropic", "open": False,
        "mcp_atlas": 84.7, "toolathlon": 77.9, "automation": 29.1,
        "browsecomp": 88.0, "gdpval_elo": 1760, "tau2": None,
        "notes": "Best-in-table on Toolathlon-Verified, GDPval-AA, Job Bench "
                 "(57.4), OfficeQA Pro (69.9), APEX-Agents (43.3). OpenAI's own "
                 "table also shows Fable leading Toolathlon (61.7 vs Sol 58.0).",
    },
    {
        "slug": "openai/gpt-5.6-sol-20260709",
        "model": "GPT-5.6 Sol", "developer": "OpenAI", "open": False,
        "mcp_atlas": 83.6, "toolathlon": 74.9, "automation": 29.7,
        "browsecomp": 90.4, "gdpval_elo": 1748, "tau2": 85.1,
        "notes": "Moonshot table (Codex harness). OpenAI reports SOTA computer "
                 "use (OSWorld 2.0 62.6%) and Agents' Last Exam 52.7% (+12.2 "
                 "over Fable); BrowseComp 92.2 in Sol Ultra 4-agent mode. "
                 "Tau2 from AA leaderboard.",
    },
    {
        "slug": "anthropic/claude-4.8-opus-20260528",
        "model": "Claude Opus 4.8", "developer": "Anthropic", "open": False,
        "mcp_atlas": 83.6, "toolathlon": 76.2, "automation": 27.2,
        "browsecomp": 84.3, "gdpval_elo": 1600, "tau2": None,
        "notes": "Closed reference point below the Fable/Sol tier; Moonshot "
                 "launch table.",
    },
    {
        "slug": "z-ai/glm-5.2-20260616",
        "model": "GLM-5.2", "developer": "Z.ai (Zhipu)", "open": True,
        "mcp_atlas": 82.6, "toolathlon": 59.9, "automation": 12.9,
        "browsecomp": None, "gdpval_elo": 1514, "tau2": None,
        "notes": "Moonshot launch table. Z.ai's own blog reports MCP Atlas 77.0 "
                 "and HLE-with-tools 54.7; built for 'thousands of tool calls' "
                 "long-horizon sessions. Only row with parallel tool-call API "
                 "support besides Kimi K2.6.",
    },
    {
        "slug": "deepseek/deepseek-v4-pro-20260423",
        "model": "DeepSeek V4 Pro", "developer": "DeepSeek", "open": True,
        "mcp_atlas": 73.6, "toolathlon": 51.8, "automation": None,
        "browsecomp": 83.4, "gdpval_elo": 1554, "tau2": None,
        "notes": "DeepSeek V4 technical report, Think-Max mode (Toolathlon "
                 "standard variant, not Verified).",
    },
    {
        "slug": "deepseek/deepseek-v4-flash-20260423",
        "model": "DeepSeek V4 Flash", "developer": "DeepSeek", "open": True,
        "mcp_atlas": 69.0, "toolathlon": 47.8, "automation": None,
        "browsecomp": 73.2, "gdpval_elo": 1395, "tau2": None,
        "notes": "DeepSeek V4 technical report, Think-Max mode. Highest "
                 "real-world tool-call volume of any paid model on OpenRouter.",
    },
    {
        "slug": "minimax/minimax-m3-20260531",
        "model": "MiniMax M3", "developer": "MiniMax", "open": True,
        "mcp_atlas": 74.2, "toolathlon": None, "automation": None,
        "browsecomp": 83.5, "gdpval_elo": 1395, "tau2": 88.9,
        "notes": "Vendor report / AA-BenchLM mirrors. Native computer use and "
                 "multimodal input; reported 24-hour autonomous run with "
                 "~2,000 tool calls.",
    },
    {
        "slug": "moonshotai/kimi-k2.6-20260420",
        "model": "Kimi K2.6", "developer": "Moonshot AI", "open": True,
        "mcp_atlas": 69.4, "toolathlon": 50.0, "automation": None,
        "browsecomp": 83.2, "gdpval_elo": 1190, "tau2": 95.9,
        "notes": "Moonshot K2.7 launch comparison / BenchLM. Best tau2-bench in "
                 "this set; native agent swarms (300 sub-agents, 4,000 steps); "
                 "HLE-with-tools 54.0 led all frontier models at release.",
    },
    {
        "slug": "nvidia/nemotron-3-ultra-550b-a55b-20260604",
        "model": "NVIDIA Nemotron 3 Ultra", "developer": "NVIDIA", "open": True,
        "mcp_atlas": None, "toolathlon": None, "automation": None,
        "browsecomp": 44.4, "gdpval_elo": None, "tau2": 70.9,
        "notes": "NVIDIA model card (BF16). Tau2 column shows TauBench V3 "
                 "average (airline/retail/telecom/banking), a different version "
                 "from the AA tau2 numbers above. No MCP Atlas or Toolathlon "
                 "published.",
    },
]

# OpenRouter weekly tool calls per model (rankings 'Tool Calls' chart), last
# completed ISO week. Snapshot: data/raw/openrouter_tool_calls_week_2026-07-20.json
TOOL_CALLS_WEEK = "2026-07-13"

# ------------------------------------------------------- Params & Benchmarks
# Hand-curated deep-dive for selected model families (researched 2026-07-20).
# Parameter counts are official unless marked "undisclosed". Benchmark scores
# are vendor-reported unless the note says otherwise; the AA Intelligence Index
# and cost-per-task columns come from the artificialanalysis.ai leaderboard
# snapshot of 2026-07-20 (highest-effort configuration of each model).
# None = not published / not applicable.
PARAMS_BENCHMARKS = [
    {
        "catalog_id": "nvidia/nemotron-3-ultra-550b-a55b",
        "model": "NVIDIA Nemotron 3 Ultra",
        "developer": "NVIDIA",
        "open": True,
        "released": "2026-06-04",
        "total_b": 550, "active_b": 55,
        "arch": "Hybrid Mamba-Transformer MoE (90% sparsity)",
        "aa_index": 38, "aa_cost_task": 0.24,
        "gpqa": 87.0, "swe_verified": 71.9, "swe_pro": None,
        "terminal_bench": 56.4, "tb_ver": "2.1", "hle": 26.7,
        "notes": "Largest US open-weights model; NVIDIA model card (BF16). "
                 "AA scored it 47.7-48.2 on Index v4.0 at release.",
    },
    {
        "catalog_id": "nvidia/nemotron-3-super-120b-a12b",
        "model": "NVIDIA Nemotron 3 Super",
        "developer": "NVIDIA",
        "open": True,
        "released": "2026-03-10",
        "total_b": 120, "active_b": 12,
        "arch": "Hybrid Mamba-Transformer LatentMoE, NVFP4 pretraining",
        "aa_index": 25, "aa_cost_task": 0.21,
        "gpqa": 79.2, "swe_verified": 60.5, "swe_pro": None,
        "terminal_bench": 31.0, "tb_ver": "2.0", "hle": 18.3,
        "notes": "SWE-bench Verified via OpenHands harness (NVIDIA model card).",
    },
    {
        "catalog_id": "anthropic/claude-fable-5",
        "model": "Claude Fable 5",
        "developer": "Anthropic",
        "open": False,
        "released": "2026-06-09",
        "total_b": None, "active_b": None,
        "arch": "Undisclosed (Mythos-class; third-party estimates ~3T)",
        "aa_index": 60, "aa_cost_task": 2.75,
        "gpqa": 92.6, "swe_verified": 95.0, "swe_pro": 80.3,
        "terminal_bench": 88.0, "tb_ver": "2.1", "hle": 53.3,
        "notes": "Anthropic launch materials / system card; OpenAI's comparison "
                 "table reports Terminal-Bench 83.1 and SWE-bench Pro 80.0. "
                 "AA config: Adaptive Reasoning, Max Effort, Opus 4.8 fallback.",
    },
    {
        "catalog_id": "openai/gpt-5.6-sol",
        "model": "GPT-5.6 Sol",
        "developer": "OpenAI",
        "open": False,
        "released": "2026-07-09",
        "total_b": None, "active_b": None,
        "arch": "Undisclosed (third-party estimates ~3T)",
        "aa_index": 59, "aa_cost_task": 1.04,
        "gpqa": 94.6, "swe_verified": None, "swe_pro": 64.6,
        "terminal_bench": 88.8, "tb_ver": "2.1", "hle": 47.2,
        "notes": "OpenAI launch table (Codex harness); no SWE-bench Verified "
                 "published. AA config: max reasoning effort. HLE is AA-HLE.",
    },
    {
        "catalog_id": "moonshotai/kimi-k3",
        "model": "Kimi K3",
        "developer": "Moonshot AI",
        "open": True,
        "released": "2026-07-16",
        "total_b": 2800, "active_b": 50,
        "arch": "Stable LatentMoE, 16/896 experts active; KDA attention "
                "(active params ~50B, third-party estimate)",
        "aa_index": 57, "aa_cost_task": 0.95,
        "gpqa": 93.5, "swe_verified": None, "swe_pro": None,
        "terminal_bench": 88.3, "tb_ver": "2.1", "hle": 43.5,
        "notes": "First open 3T-class model; weights promised by 2026-07-27. "
                 "Moonshot tech blog (reasoning_effort max, KimiCode harness); "
                 "no SWE-bench results published.",
    },
    {
        "catalog_id": "z-ai/glm-5.2",
        "model": "GLM-5.2",
        "developer": "Z.ai (Zhipu)",
        "open": True,
        "released": "2026-06-13",
        "total_b": 753, "active_b": 40,
        "arch": "MoE, MIT license; trained on Huawei Ascend 910B",
        "aa_index": 51, "aa_cost_task": 0.47,
        "gpqa": 91.2, "swe_verified": None, "swe_pro": 62.1,
        "terminal_bench": 81.0, "tb_ver": "2.1", "hle": 40.5,
        "notes": "Z.ai release blog (Terminus-2 harness); strongest open-weights "
                 "model on the AA Index. No SWE-bench Verified published.",
    },
    {
        "catalog_id": "qwen/qwen3.7-max",
        "model": "Qwen3.7 Max",
        "developer": "Alibaba (Qwen)",
        "open": False,
        "released": "2026-05-20",
        "total_b": None, "active_b": None,
        "arch": "Undisclosed (API-only; Qwen's top tier is closed since late 2025)",
        "aa_index": 46, "aa_cost_task": 1.03,
        "gpqa": 92.4, "swe_verified": 80.4, "swe_pro": 60.6,
        "terminal_bench": 69.7, "tb_ver": "2.0", "hle": 41.4,
        "notes": "Alibaba Qwen3.7 release blog (xhigh reasoning).",
    },
    {
        "catalog_id": "deepseek/deepseek-v4-pro",
        "model": "DeepSeek V4 Pro",
        "developer": "DeepSeek",
        "open": True,
        "released": "2026-04-24",
        "total_b": 1600, "active_b": 49,
        "arch": "MoE with hybrid CSA+HCA attention, MIT license",
        "aa_index": 44, "aa_cost_task": 0.04,
        "gpqa": 90.1, "swe_verified": 80.6, "swe_pro": 55.4,
        "terminal_bench": 67.9, "tb_ver": "2.0", "hle": 37.7,
        "notes": "DeepSeek V4 technical report, Think-Max mode.",
    },
    {
        "catalog_id": "deepseek/deepseek-v4-flash",
        "model": "DeepSeek V4 Flash",
        "developer": "DeepSeek",
        "open": True,
        "released": "2026-04-24",
        "total_b": 284, "active_b": 13,
        "arch": "MoE with hybrid CSA+HCA attention, MIT license",
        "aa_index": 40, "aa_cost_task": 0.02,
        "gpqa": 88.1, "swe_verified": 79.0, "swe_pro": 52.6,
        "terminal_bench": 56.9, "tb_ver": "2.0", "hle": 34.8,
        "notes": "DeepSeek V4 technical report, Think-Max mode.",
    },
]


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


def enrich_tool_use(catalog, rows):
    """Join the tool-use table with catalog API flags, weekly volume, and
    real-world weekly tool-call counts from the OpenRouter snapshot."""
    tc_path = sorted(glob.glob(os.path.join(RAW, "openrouter_tool_calls_week_*.json")))[-1]
    with open(tc_path) as f:
        weeks = json.load(f)["data"]
    week = next(w for w in weeks if w["x"] == TOOL_CALLS_WEEK)
    calls = defaultdict(int)
    for key, n in week["ys"].items():
        calls[key.split(":")[0]] += n
    total_calls = sum(week["ys"].values())

    by_slug = {r["slug"]: r for r in rows}
    out = []
    for tu in TOOL_USE:
        tu = dict(tu)
        entry = catalog.get(tu["slug"])
        sp = (entry or {}).get("supported_parameters") or []
        tu["parallel_tools"] = "parallel_tool_calls" in sp
        row = by_slug.get(tu["slug"])
        tu["tokens_total"] = row["tokens_total"] if row else None
        tu["weekly_tool_calls"] = calls.get(tu["slug"]) or None
        out.append(tu)
    top_callers = sorted(week["ys"].items(),
                         key=lambda kv: (kv[0] == "Others", -kv[1]))
    return out, top_callers, total_calls


def write_tool_use_sheet(wb, tu_rows, top_callers, total_calls):
    ws = wb.create_sheet("Tool Use")
    ws["A1"] = "Tool use: Kimi K3 vs closed frontier (Fable 5, GPT-5.6 Sol) and open-weights peers"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Agentic tool-calling benchmarks (researched 2026-07-21) plus real-world "
                "tool-call traffic on OpenRouter")
    ws["A3"] = ("Rows 1-5 (K3, Fable 5, Sol, Opus 4.8, GLM-5.2) share one methodology: Moonshot's Kimi K3 launch "
                "table. Remaining rows are vendor-reported under their own harnesses, so compare them loosely.")
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="595959")

    headers = [
        "Model", "Developer", "Class",
        "MCP Atlas (%)", "Toolathlon-Verified (%)", "AutomationBench (%)",
        "BrowseComp (%)", "GDPval-AA (Elo)", "tau2-bench (%)",
        "Parallel tool-call API", "OpenRouter tool calls, wk of "
        + TOOL_CALLS_WEEK + " (M)", "Weekly tokens (B)", "Notes / sources",
    ]
    hr = 5
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"A{hr + 1}"

    for i, tu in enumerate(tu_rows):
        r = hr + 1 + i
        vals = [
            tu["model"], tu["developer"],
            "Open-weights" if tu["open"] else "Closed",
            tu["mcp_atlas"], tu["toolathlon"], tu["automation"],
            tu["browsecomp"], tu["gdpval_elo"], tu["tau2"],
            "Yes" if tu["parallel_tools"] else "No",
            tu["weekly_tool_calls"] / 1e6 if tu["weekly_tool_calls"] else None,
            tu["tokens_total"] / 1e9 if tu["tokens_total"] else None,
            tu["notes"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v if v is not None else "—")
            cell.border = BORDER
            if c == 3:
                cell.fill = OPEN_FILL if tu["open"] else CLOSED_FILL
            if c == 13:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for c in (4, 5, 6, 7, 9, 11):
            if isinstance(ws.cell(row=r, column=c).value, (int, float)):
                ws.cell(row=r, column=c).number_format = "0.0"
        if isinstance(ws.cell(row=r, column=8).value, (int, float)):
            ws.cell(row=r, column=8).number_format = "#,##0"
        if isinstance(ws.cell(row=r, column=12).value, (int, float)):
            ws.cell(row=r, column=12).number_format = "#,##0.0"
    autosize(ws, [24, 15, 13, 11, 15, 14, 12, 12, 11, 13, 18, 13, 70])

    # Real-world tool-call leaderboard (OpenRouter 'Tool Calls' chart)
    lb_start = hr + len(tu_rows) + 2
    ws.cell(row=lb_start, column=1,
            value=f"Real-world tool-call traffic on OpenRouter — completed week of {TOOL_CALLS_WEEK} "
                  f"(total {total_calls / 1e6:.0f}M tool calls)").font = Font(bold=True, size=12)
    for c, h in enumerate(["Model (permaslug)", "Tool calls (M)", "Share"], 1):
        ws.cell(row=lb_start + 1, column=c, value=h)
    style_header(ws, lb_start + 1, 3)
    for i, (slug, n) in enumerate(top_callers):
        r = lb_start + 2 + i
        ws.cell(row=r, column=1, value=slug).border = BORDER
        ws.cell(row=r, column=2, value=n / 1e6).number_format = "#,##0.0"
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=3, value=n / total_calls).number_format = "0.0%"
        ws.cell(row=r, column=3).border = BORDER

    nrow = lb_start + 2 + len(top_callers) + 1
    notes = [
        "Benchmarks: MCP Atlas = multi-step tool use over Model Context Protocol servers (500-task public subset, "
        "100-turn limit, Gemini 3.1 Pro judge). Toolathlon-Verified = multi-tool office/personal workflows; DeepSeek "
        "rows use the standard Toolathlon variant. AutomationBench = Zapier-style SaaS automation (600-task subset). "
        "BrowseComp = agentic web research. GDPval-AA = Artificial Analysis's Elo for economically valuable "
        "professional tasks. tau2-bench = conversational tool use; the Nemotron figure is TauBench V3 (different version).",
        "Kimi K3's launch table is vendor-published: K3 runs at max reasoning effort on Moonshot's preferred harness "
        "per benchmark, while competitor numbers mix harnesses and third-party citations. OpenAI's own table shows the "
        "same ordering on Toolathlon (Fable 61.7 > Opus 59.9 > Sol 58.0 on the standard variant).",
        "Takeaway: on tool use the three frontier models are within ~1 point on MCP Atlas (Fable 84.7, K3 84.2, Sol 83.6). "
        "Fable 5 leads orchestrated professional work (Toolathlon, GDPval-AA, OfficeQA); Sol leads computer use and "
        "long-horizon agent runs (OSWorld 62.6, Agents' Last Exam 52.7); K3 leads web research (BrowseComp 91.2) and "
        "Zapier-style automation, at roughly half Fable's cost per task ($0.94 vs $2.75 per AA).",
        "Real-world traffic tells a different story: OpenRouter's tool-call chart is dominated by cheap open-weights "
        "models (Hy3, MiMo-V2.5, DeepSeek V4 Flash, GLM-5.2, MiniMax M3). Fable 5, Sol, and K3 are outside the top 9 "
        "tool-callers - premium closed models are used more via first-party APIs and coding products than through "
        "OpenRouter tool-calling.",
        "Weekly tokens = OpenRouter total volume, week ending " + WEEK_ENDING + ". Kimi K3 launched 2026-07-16, so its "
        "volume covers <1 day. Source: OpenRouter (openrouter.ai/rankings), as of 2026-07-21.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")

    chart = BarChart()
    chart.type = "col"
    chart.title = "Tool-use benchmarks: MCP Atlas vs Toolathlon-Verified"
    data = Reference(ws, min_col=4, max_col=5, min_row=hr, max_row=hr + len(tu_rows))
    cats = Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + len(tu_rows))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 10, 24
    ws.add_chart(chart, f"E{lb_start + 1}")


def enrich_params_benchmarks(models, rows):
    """Join the hand-curated table with catalog pricing and weekly token volume."""
    id_to_canon = {e["id"]: e["canonical_slug"] for e in models if ":" not in e["id"]}
    by_slug = {r["slug"]: r for r in rows}
    out = []
    for pb in PARAMS_BENCHMARKS:
        pb = dict(pb)
        row = by_slug.get(id_to_canon.get(pb["catalog_id"], ""))
        pb["price_in"] = row["price_in"] if row else None
        pb["price_out"] = row["price_out"] if row else None
        pb["tokens_total"] = row["tokens_total"] if row else None
        out.append(pb)
    return out


def write_params_benchmarks_sheet(wb, pb_rows):
    ws = wb.create_sheet("Params & Benchmarks")
    ws["A1"] = "Parameter and benchmark deep-dive: selected model families"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("NVIDIA Nemotron, Claude Fable, OpenAI GPT-5.6 Sol, Kimi K3, Z.ai GLM, "
                "Qwen, DeepSeek - researched 2026-07-20")
    ws["A3"] = ("Benchmark scores are vendor-reported (highest reasoning effort) unless noted; harnesses differ "
                "across labs, so treat cross-model comparisons as approximate. AA = Artificial Analysis "
                "Intelligence Index v4.1 leaderboard snapshot, 2026-07-20.")
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="595959")

    headers = [
        "Model", "Developer", "Class", "Released",
        "Total params (B)", "Active params (B)", "Architecture / license",
        "AA Intelligence Index", "AA cost per task ($)",
        "GPQA Diamond (%)", "SWE-bench Verified (%)", "SWE-bench Pro (%)",
        "Terminal-Bench (%)", "TB ver.", "HLE (%)",
        "Input $/1M", "Output $/1M", "Weekly tokens (B)", "Notes / sources",
    ]
    hr = 5
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"A{hr + 1}"

    for i, pb in enumerate(pb_rows):
        r = hr + 1 + i
        vals = [
            pb["model"], pb["developer"],
            "Open-weights" if pb["open"] else "Closed",
            pb["released"],
            pb["total_b"] if pb["total_b"] is not None else "undisclosed",
            pb["active_b"] if pb["active_b"] is not None else "undisclosed",
            pb["arch"],
            pb["aa_index"], pb["aa_cost_task"],
            pb["gpqa"], pb["swe_verified"], pb["swe_pro"],
            pb["terminal_bench"], pb["tb_ver"], pb["hle"],
            pb["price_in"], pb["price_out"],
            pb["tokens_total"] / 1e9 if pb["tokens_total"] else None,
            pb["notes"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v if v is not None else "—")
            cell.border = BORDER
            if c == 3:
                cell.fill = OPEN_FILL if pb["open"] else CLOSED_FILL
            if c == 19:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for c in (5, 6):
            if isinstance(ws.cell(row=r, column=c).value, (int, float)):
                ws.cell(row=r, column=c).number_format = "#,##0"
        ws.cell(row=r, column=9).number_format = "$0.00"
        for c in (10, 11, 12, 13, 15):
            if isinstance(ws.cell(row=r, column=c).value, (int, float)):
                ws.cell(row=r, column=c).number_format = "0.0"
        for c in (16, 17):
            if isinstance(ws.cell(row=r, column=c).value, (int, float)):
                ws.cell(row=r, column=c).number_format = "$#,##0.000"
        if isinstance(ws.cell(row=r, column=18).value, (int, float)):
            ws.cell(row=r, column=18).number_format = "#,##0.0"
    autosize(ws, [24, 15, 13, 11, 13, 13, 40, 12, 11, 12, 13, 12, 13, 7, 9, 10, 10, 13, 60])

    nrow = hr + len(pb_rows) + 2
    notes = [
        "Total/active parameters: official vendor disclosures for open-weights models. Anthropic, OpenAI, and Alibaba "
        "do not disclose parameter counts for Fable 5, GPT-5.6 Sol, or Qwen3.7 Max; ~3T figures are third-party estimates only.",
        "Kimi K3's ~50B active is a community estimate from the disclosed 16-of-896 expert activation; Moonshot has not "
        "published an official active-parameter figure. Weights promised by 2026-07-27 (API-only as of 2026-07-20).",
        "Terminal-Bench versions differ (2.0 vs 2.1) and each lab uses its own agent harness (KimiCode, Claude Code, Codex, "
        "Terminus-2), so scores are not strictly comparable across rows.",
        "SWE-bench Pro validity is contested: OpenAI's 2026-07-08 audit estimates ~30% of tasks are flawed.",
        "AA cost per task = Artificial Analysis's measured cost to run its Intelligence Index suite, divided per task; "
        "reflects token efficiency as well as list price.",
        "HLE = Humanity's Last Exam, no tools (AA-HLE for Fable 5 / Sol). Weekly tokens = OpenRouter volume, week ending "
        + WEEK_ENDING + " (understates first-party API usage of closed models).",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")

    chart = BarChart()
    chart.type = "col"
    chart.title = "GPQA Diamond vs AA Intelligence Index"
    data_start = hr + 1
    data_end = hr + len(pb_rows)
    data = Reference(ws, min_col=8, max_col=8, min_row=hr, max_row=data_end)
    data2 = Reference(ws, min_col=10, max_col=10, min_row=hr, max_row=data_end)
    cats = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    chart.add_data(data, titles_from_data=True)
    chart.add_data(data2, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 10, 24
    ws.add_chart(chart, f"A{nrow + len(notes) + 2}")


def write_workbook(rows, summary, devs, grand_total, excluded_tokens, pb_rows,
                   tu_rows, top_callers, total_calls):
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

    # ---- Sheet 5: Params & benchmarks deep-dive
    write_params_benchmarks_sheet(wb, pb_rows)

    # ---- Sheet 6: Tool use comparison
    write_tool_use_sheet(wb, tu_rows, top_callers, total_calls)

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


def write_params_benchmarks_csv(pb_rows):
    with open(os.path.join(OUT_CSV_DIR, "params_benchmarks.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "model", "developer", "class", "released",
            "total_params_b", "active_params_b", "architecture",
            "aa_intelligence_index", "aa_cost_per_task_usd",
            "gpqa_diamond_pct", "swe_bench_verified_pct", "swe_bench_pro_pct",
            "terminal_bench_pct", "terminal_bench_version", "hle_pct",
            "input_usd_per_1m", "output_usd_per_1m", "weekly_tokens", "notes",
        ])
        for pb in pb_rows:
            w.writerow([
                pb["model"], pb["developer"],
                "open-weights" if pb["open"] else "closed",
                pb["released"], pb["total_b"], pb["active_b"], pb["arch"],
                pb["aa_index"], pb["aa_cost_task"],
                pb["gpqa"], pb["swe_verified"], pb["swe_pro"],
                pb["terminal_bench"], pb["tb_ver"], pb["hle"],
                pb["price_in"], pb["price_out"], pb["tokens_total"], pb["notes"],
            ])


def write_tool_use_csv(tu_rows):
    with open(os.path.join(OUT_CSV_DIR, "tool_use_benchmarks.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "model", "developer", "class",
            "mcp_atlas_pct", "toolathlon_verified_pct", "automationbench_pct",
            "browsecomp_pct", "gdpval_aa_elo", "tau2_bench_pct",
            "parallel_tool_call_api", "openrouter_tool_calls_week_" + TOOL_CALLS_WEEK.replace("-", ""),
            "weekly_tokens", "notes",
        ])
        for tu in tu_rows:
            w.writerow([
                tu["model"], tu["developer"],
                "open-weights" if tu["open"] else "closed",
                tu["mcp_atlas"], tu["toolathlon"], tu["automation"],
                tu["browsecomp"], tu["gdpval_elo"], tu["tau2"],
                tu["parallel_tools"], tu["weekly_tool_calls"],
                tu["tokens_total"], tu["notes"],
            ])


def main():
    rankings, models = load_raw()
    catalog = index_catalog(models)
    rows, excluded = build_rows(rankings, catalog)
    summary, grand_total = summarize(rows)
    devs = by_developer(rows)
    pb_rows = enrich_params_benchmarks(models, rows)
    tu_rows, top_callers, total_calls = enrich_tool_use(catalog, rows)
    write_workbook(rows, summary, devs, grand_total, excluded, pb_rows,
                   tu_rows, top_callers, total_calls)
    write_csvs(rows, summary, devs, grand_total)
    write_params_benchmarks_csv(pb_rows)
    write_tool_use_csv(tu_rows)

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
