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
  - data/csv/cloud_provider_split.csv
  - data/csv/harness_rl_environments.csv
  - data/csv/quantization_kld.csv
  - data/csv/training_data.csv
  - data/csv/model_harnesses_tools.csv
  - data/llm_data_tables.xlsx (all CSV tables compiled into one data-only workbook)
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

# ------------------------------------------------- Per-model harnesses & tools
# Hand-curated harness/tooling columns per model (researched 2026-07-29);
# API feature flags and modalities are joined live from the OpenRouter catalog
# snapshot. catalog_id keys into the models catalog.
MODEL_HARNESS_TOOLS = [
    {
        "catalog_id": "moonshotai/kimi-k3",
        "harness": "Kimi Code CLI; Kimi.com / Kimi Work apps; K3 Swarm Max "
                   "orchestrator (up to 300 parallel sub-agents)",
        "third_party": "Claude Code, Cursor, OpenClaw via Anthropic/OpenAI-"
                       "compatible API - but requires preserved thinking history "
                       "(verified harnesses only; no mid-session model switches)",
        "tools": "Web search/fetch, terminal, code execution, vision-in-the-loop "
                 "(live screenshots), video editing, Widgets + Dashboard in Kimi "
                 "Work; MCP",
        "notes": "Weights live on HF since 2026-07-27 (96 shards, ~99K downloads "
                 "in 2 days); day-0 vLLM support with Moonshot-contributed KDA "
                 "prefix caching.",
    },
    {
        "catalog_id": "moonshotai/kimi-k2.6",
        "harness": "Kimi CLI; native Agent Swarm (300 sub-agents, 4,000 steps)",
        "third_party": "Claude Code, OpenClaw, Cline (Anthropic/OpenAI-compatible)",
        "tools": "Web search, terminal, code execution, swarm orchestration; MCP; "
                 "one of only two models here exposing parallel tool calls on "
                 "OpenRouter",
        "notes": "Apache-2.0-style open weights available now; HLE-with-tools "
                 "54.0 led all frontier models at release.",
    },
    {
        "catalog_id": "anthropic/claude-fable-5",
        "harness": "Claude Code (CLI + SDK) - the industry reference harness; "
                   "Claude Cowork for knowledge work",
        "third_party": "Anthropic API only; broad ecosystem targets it (Cursor, "
                       "OpenClaw, Cline all speak Anthropic protocol)",
        "tools": "Memory tool, code execution, computer use, web search, "
                 "programmatic tool calling, context editing / tool-result "
                 "clearing, compaction, task budgets (beta); MCP. Adaptive "
                 "thinking always on (effort parameter); no assistant prefill",
        "notes": "Safety classifiers can reroute cyber/bio/distillation requests "
                 "to Opus 4.8 (fallback affects <5% of sessions per Anthropic).",
    },
    {
        "catalog_id": "anthropic/claude-sonnet-5",
        "harness": "Claude Code; Claude.ai apps",
        "third_party": "Anthropic API ecosystem (same as Fable 5)",
        "tools": "Same Claude tool suite: memory tool, code execution, computer "
                 "use, web search, compaction; MCP",
        "notes": "Mid-tier workhorse at intro pricing ($2/$10 through Aug 2026).",
    },
    {
        "catalog_id": "openai/gpt-5.6-sol",
        "harness": "Codex suite: CLI, Cloud, VS Code extension, App Server "
                   "(JSON-RPC); ChatGPT agent surfaces",
        "third_party": "Responses API; community proxies run it inside Claude "
                       "Code (some report it codes better there - model/harness "
                       "entanglement debate)",
        "tools": "Web search, code interpreter, file search, computer use "
                 "(OSWorld 2.0 SOTA 62.6%), subagent orchestration (Sol as "
                 "orchestrator over Terra/Luna; config bug openai/codex#31814), "
                 "automatic compaction; MCP",
        "notes": "Models explicitly tuned for the Codex harness; ~25h "
                 "uninterrupted runs reported.",
    },
    {
        "catalog_id": "openai/gpt-5.5",
        "harness": "Codex suite (previous flagship)",
        "third_party": "Responses API ecosystem",
        "tools": "Web search, code interpreter, file search, computer use; MCP",
        "notes": "Superseded by GPT-5.6 family for agentic work.",
    },
    {
        "catalog_id": "z-ai/glm-5.2",
        "harness": "None of its own - by design. GLM Coding Plan ($3-80/mo) "
                   "plugs into third-party harnesses",
        "third_party": "Claude Code (Anthropic-native API), OpenClaw, Cline, "
                       "Cursor; drop-in config documented by Z.ai",
        "tools": "Function calling with parallel tool calls; dynamic working-"
                 "memory management for thousands-of-tool-call sessions; MCP",
        "notes": "Strongest open-weights MCP Atlas score (82.6 on Moonshot's "
                 "table; 77.0 per Z.ai's own blog).",
    },
    {
        "catalog_id": "deepseek/deepseek-v4-pro",
        "harness": "None - dual-mode OpenAI + Anthropic-compatible API",
        "third_party": "Codex-style and Claude Code-style harnesses both work; "
                       "popular in Kilo Code, OpenClaw, aggregators",
        "tools": "Function calling via new XML-based DSML schema (replaces "
                 "JSON); three reasoning modes (non-think/high/max); MCP",
        "notes": "Agentic tool-use data injected in mid-training, not just "
                 "post-training.",
    },
    {
        "catalog_id": "deepseek/deepseek-v4-flash",
        "harness": "None - same dual-mode API as V4 Pro",
        "third_party": "Same as V4 Pro; the default budget model in many "
                       "third-party coding agents",
        "tools": "Same DSML function calling and reasoning modes; MCP",
        "notes": "Highest real-world tool-call volume of any paid model on "
                 "OpenRouter (~70M calls/week).",
    },
    {
        "catalog_id": "minimax/minimax-m3",
        "harness": "MiniMax Agent product; no dedicated CLI",
        "third_party": "OpenAI/Anthropic-compatible API",
        "tools": "Native computer use (only open model here with it), video "
                 "input in tool workflows, function calling; MCP; 24h autonomous "
                 "runs with ~2,000 tool calls reported",
        "notes": "Multimodal agent focus (text+image+video input).",
    },
    {
        "catalog_id": "nvidia/nemotron-3-ultra-550b-a55b",
        "harness": "NVIDIA API / NIM; NeMo Gym ships out-of-the-box harnesses "
                   "(incl. Claude Code and Hermes) for eval/training",
        "third_party": "OpenAI-compatible endpoints; runs in standard open "
                       "harnesses (OpenHands, mini-SWE-agent)",
        "tools": "Function calling, structured outputs; tuned for TauBench-"
                 "style conversational tool use (V3 avg 70.9); MCP via harness",
        "notes": "Optimized for high-throughput agentic serving (300+ tok/s on "
                 "Blackwell).",
    },
    {
        "catalog_id": "qwen/qwen3.7-max",
        "harness": "Qwen Chat; DashScope / Model Studio API",
        "third_party": "Native OpenAI + Anthropic API compatibility (works in "
                       "Claude Code / Codex-style harnesses)",
        "tools": "Function calling, extended thinking with preserve_thinking "
                 "across turns, prompt caching; MCP (leads MCP-Mark per "
                 "Alibaba); 35h autonomous kernel-optimization run (1,158 tool "
                 "calls) reported",
        "notes": "Agent-first positioning; API-only.",
    },
    {
        "catalog_id": "xiaomi/mimo-v2.5",
        "harness": "Xiaomi AI Studio; MiMo API (Token Plans)",
        "third_party": "OpenAI-compatible; SGLang/vLLM ship a dedicated 'mimo' "
                       "tool-call parser for self-hosting",
        "tools": "Function calling; omnimodal input (text+image+audio+video) "
                 "usable inside tool loops; MCP",
        "notes": "Highest-volume open model on OpenRouter after Hy3; ~110M "
                 "tool calls/week.",
    },
    {
        "catalog_id": "tencent/hy3",
        "harness": "Tencent Yuanbao apps; Tencent Cloud API",
        "third_party": "OpenAI-compatible; hosted by GMICloud, Novita, "
                       "DeepInfra, SiliconFlow etc.",
        "tools": "Function calling, structured outputs; MCP via harness",
        "notes": "Free promo tier drove ~140M tool calls/week at peak; 295B/21B "
                 "MoE (192 experts, top-8).",
    },
    {
        "catalog_id": "google/gemini-3-flash-preview",
        "harness": "Gemini CLI; AI Studio; Vertex AI Agent Builder",
        "third_party": "Vertex/AI Studio APIs; Gemini CLI is open-source",
        "tools": "Google Search grounding, code execution, URL context, "
                 "function calling; MCP",
        "notes": "Google's high-volume agentic workhorse tier.",
    },
    {
        "catalog_id": "google/gemini-3.1-pro-preview",
        "harness": "Gemini CLI; AI Studio; Vertex AI",
        "third_party": "Same Google API surfaces",
        "tools": "Google Search grounding, code execution, computer use "
                 "(preview), function calling; MCP",
        "notes": "Tops hard-reasoning benchmarks (GPQA 94.3) but trails on "
                 "agentic harness benchmarks (Terminal-Bench 74.0).",
    },
    {
        "catalog_id": "x-ai/grok-4.5",
        "harness": "Grok Build (xAI's coding-agent harness); Grok apps / X "
                   "integration",
        "third_party": "OpenAI-compatible xAI API",
        "tools": "Real-time X/web search, code execution, function calling; MCP",
        "notes": "AA benchmarks it as the Grok 4.5 + Grok Build pair (ties Sol "
                 "on SWE-Atlas-QnA).",
    },
    {
        "catalog_id": "meta-llama/llama-4-maverick",
        "harness": "None (no first-party agent product); Meta AI consumer apps "
                   "use tuned variants",
        "third_party": "Served by 20+ clouds; used in open harnesses "
                       "(OpenHands, custom agents)",
        "tools": "Function calling and structured outputs only - no reasoning-"
                 "effort controls exposed (pre-reasoning-era design)",
        "notes": "2025-generation model; still significant volume as a cheap "
                 "workhorse.",
    },
    {
        "catalog_id": "openai/gpt-oss-120b",
        "harness": "None first-party; reference Harmony chat format",
        "third_party": "Runs in most open harnesses; served by 20 providers",
        "tools": "Function calling, structured outputs, reasoning effort "
                 "controls; browser/python tool support in the Harmony format",
        "notes": "OpenAI's open-weights line (Apache 2.0); no first-party "
                 "agent product.",
    },
    {
        "catalog_id": "stepfun/step-3.7-flash",
        "harness": "StepFun API / apps",
        "third_party": "OpenAI-compatible",
        "tools": "Function calling, structured outputs, reasoning; video input; "
                 "MCP via harness",
        "notes": "Budget agentic tier; 894B tokens/week on OpenRouter.",
    },
]


def enrich_harness_tools(models, rows):
    byid = {x["id"]: x for x in models if ":" not in x["id"]}
    canon_rows = {r["slug"]: r for r in rows}
    out = []
    for m in MODEL_HARNESS_TOOLS:
        m = dict(m)
        e = byid.get(m["catalog_id"])
        if not e:
            continue
        sp = set(e.get("supported_parameters") or [])
        m["model"] = e["name"]
        m["developer"] = DEVELOPER_NAMES.get(e["id"].split("/")[0], e["id"].split("/")[0].title())
        row = canon_rows.get(e["canonical_slug"])
        m["open"] = row["open"] if row else bool(e.get("hugging_face_id"))
        m["modalities"] = "+".join(sorted(e["architecture"]["input_modalities"]))
        m["fc"] = "tools" in sp
        m["parallel"] = "parallel_tool_calls" in sp
        m["structured"] = "structured_outputs" in sp or "response_format" in sp
        m["reasoning"] = "reasoning" in sp or "reasoning_effort" in sp
        m["context"] = e.get("context_length") or 0
        m["tokens"] = row["tokens_total"] if row else None
        out.append(m)
    return out


def write_harness_tools_sheet(wb, ht_rows):
    ws = wb.create_sheet("Model Harnesses & Tools")
    ws["A1"] = "Harnesses and tools used by each model"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Per-model view: the harness each model ships with / runs in, and the tools it supports. API feature "
                "flags and input modalities from the OpenRouter catalog snapshot (2026-07-16); harness and tool-suite "
                "columns hand-curated from lab docs, researched 2026-07-29.")
    ws["A2"].font = Font(italic=True, color="595959")

    headers = [
        "Model", "Developer", "Class", "First-party harness / agent surface",
        "Third-party harness compatibility", "Built-in / first-party tools",
        "Input modalities", "Function calling", "Parallel tool calls",
        "Structured outputs", "Reasoning controls", "Context (tokens)",
        "Weekly tokens (B)", "Notes",
    ]
    hr = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"B{hr + 1}"
    for i, m in enumerate(ht_rows):
        r = hr + 1 + i
        vals = [
            m["model"], m["developer"],
            "Open-weights" if m["open"] else "Closed",
            m["harness"], m["third_party"], m["tools"], m["modalities"],
            "Yes" if m["fc"] else "No",
            "Yes" if m["parallel"] else "No",
            "Yes" if m["structured"] else "No",
            "Yes" if m["reasoning"] else "No",
            m["context"],
            m["tokens"] / 1e9 if m["tokens"] else None,
            m["notes"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v if v is not None else "—")
            cell.border = BORDER
            if c in (4, 5, 6, 14):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 3:
                cell.fill = OPEN_FILL if m["open"] else CLOSED_FILL
            if c in (9,) and v == "Yes":
                cell.fill = OPEN_FILL
        ws.cell(row=r, column=12).number_format = "#,##0"
        if isinstance(ws.cell(row=r, column=13).value, (int, float)):
            ws.cell(row=r, column=13).number_format = "#,##0.0"
        ws.row_dimensions[r].height = 92
    ws.auto_filter.ref = f"A{hr}:N{hr + len(ht_rows)}"
    autosize(ws, [26, 14, 12, 34, 32, 44, 18, 10, 10, 10, 10, 12, 12, 40])

    nrow = hr + len(ht_rows) + 2
    notes = [
        "Function calling / parallel tool calls / structured outputs / reasoning flags reflect what each model "
        "exposes through the OpenRouter API (catalog snapshot " + WEEK_ENDING + "); first-party APIs may expose more "
        "(e.g. Anthropic's memory tool and computer use are Claude API features, not OpenRouter parameters).",
        "Only GLM-5.2 and Kimi K2.6 expose parallel tool calls via OpenRouter; Llama 4 is the only row without "
        "reasoning-effort controls (pre-reasoning-era design).",
        "Pattern: closed labs bundle model + harness + tool suite as one product (Claude Code, Codex, Gemini CLI, "
        "Grok Build); Chinese open-weights labs except Moonshot ship API compatibility instead of harnesses and let "
        "Claude Code / OpenClaw / Cursor be the surface; Moonshot is the only open-weights lab shipping a full "
        "first-party harness + swarm orchestrator (Kimi Code, K3 Swarm Max).",
        "MCP (Model Context Protocol) is supported across effectively all rows, natively or via the harness layer - "
        "it has become the cross-vendor tool-integration standard.",
        "Kimi K3 weights shipped on Hugging Face 2026-07-27 as promised (96 safetensors shards); self-hosted K3 can "
        "now run in any compatible harness, subject to the preserved-thinking requirement.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")


def write_harness_tools_csv(ht_rows):
    with open(os.path.join(OUT_CSV_DIR, "model_harnesses_tools.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "model", "developer", "class", "first_party_harness",
            "third_party_harness_compatibility", "builtin_tools", "input_modalities",
            "function_calling", "parallel_tool_calls", "structured_outputs",
            "reasoning_controls", "context_tokens", "weekly_tokens", "notes",
        ])
        for m in ht_rows:
            w.writerow([
                m["model"], m["developer"],
                "open-weights" if m["open"] else "closed",
                m["harness"], m["third_party"], m["tools"], m["modalities"],
                m["fc"], m["parallel"], m["structured"], m["reasoning"],
                m["context"], m["tokens"], m["notes"],
            ])


# ---------------------------------------------------------- Training dataset sizes
# Pretraining corpus sizes as disclosed by each lab (researched 2026-07-24).
# tokens_t = trillions of pretraining tokens, None = undisclosed.
TRAINING_DATA = [
    {
        "model": "Kimi K3", "developer": "Moonshot AI", "open": True,
        "params_b": 2800, "tokens_t": None,
        "status": "Undisclosed (tech report due with weights, 2026-07-27)",
        "notes": "Pipeline disclosed without quantities: web+code+vision "
                 "pretraining, long-context mid-training toward 1M, SFT, RLVR. "
                 "Predecessor Kimi K2 (1T params) disclosed 15.5T tokens, so "
                 "K3's corpus is presumably well beyond that; treat any "
                 "specific figure circulating now as an estimate.",
        "sources": "kimi.com/blog/kimi-k3; Kimi K2 paper (arXiv:2507.20534)",
    },
    {
        "model": "Kimi K2 / K2.6", "developer": "Moonshot AI", "open": True,
        "params_b": 1000, "tokens_t": 15.5,
        "status": "Disclosed for K2 base (K2.6/K2.7 increments not broken out)",
        "notes": "15.5T curated tokens across web text, code, mathematics, "
                 "knowledge; MuonClip optimizer, 4,096-token pretraining window.",
        "sources": "Kimi K2 paper (arXiv:2507.20534)",
    },
    {
        "model": "Claude Fable 5", "developer": "Anthropic", "open": False,
        "params_b": None, "tokens_t": None,
        "status": "Undisclosed",
        "notes": "Anthropic discloses neither corpus size nor parameter count "
                 "for any frontier model.",
        "sources": "Fable 5 system card (no training-data quantities)",
    },
    {
        "model": "GPT-5.6 Sol", "developer": "OpenAI", "open": False,
        "params_b": None, "tokens_t": None,
        "status": "Undisclosed",
        "notes": "OpenAI stopped disclosing training-data scale after GPT-3 "
                 "(300B tokens, 2020). gpt-oss open-weight models also ship "
                 "without a precise token count ('trillions of tokens').",
        "sources": "GPT-5.6 launch materials (no training-data quantities)",
    },
    {
        "model": "Gemini 3.x", "developer": "Google", "open": False,
        "params_b": None, "tokens_t": None,
        "status": "Undisclosed",
        "notes": "No corpus quantities disclosed for the Gemini series.",
        "sources": "Gemini model cards",
    },
    {
        "model": "GLM-5 / 5.1 / 5.2", "developer": "Z.ai (Zhipu)", "open": True,
        "params_b": 753, "tokens_t": 27,
        "status": "Disclosed (GLM-5 base, which 5.1/5.2 build on)",
        "notes": "27T-token corpus prioritizing code and reasoning early; "
                 "distinct mid-training phase extends context 4K->200K with "
                 "long-context agentic data.",
        "sources": "GLM-5 paper (arXiv:2602.15763)",
    },
    {
        "model": "DeepSeek V4 Pro", "developer": "DeepSeek", "open": True,
        "params_b": 1600, "tokens_t": 32,
        "status": "Disclosed ('more than 32T tokens', shared corpus with Flash)",
        "notes": "Agentic tool-use data injected already in mid-training; "
                 "FP8-mixed pretraining with MXFP4 on expert weights.",
        "sources": "DeepSeek-V4 report (arXiv:2606.19348)",
    },
    {
        "model": "DeepSeek V4 Flash", "developer": "DeepSeek", "open": True,
        "params_b": 284, "tokens_t": 32,
        "status": "Disclosed ('more than 32T tokens', shared corpus with Pro)",
        "notes": "Same >32T corpus as V4 Pro; the small variant is heavily "
                 "over-trained relative to its size (~113 tokens per param).",
        "sources": "DeepSeek-V4 report (arXiv:2606.19348)",
    },
    {
        "model": "Qwen3.7 Max", "developer": "Alibaba (Qwen)", "open": False,
        "params_b": None, "tokens_t": None,
        "status": "Undisclosed (closed tier)",
        "notes": "Alibaba disclosed 36T tokens for the open Qwen3 generation "
                 "(2025) but publishes no quantities for the closed Max tier.",
        "sources": "Qwen3 technical report (2025); Qwen3.7 launch blog",
    },
    {
        "model": "Nemotron 3 (Nano/Super/Ultra)", "developer": "NVIDIA", "open": True,
        "params_b": None, "tokens_t": 25,
        "status": "Disclosed AND largely released - the only lab that ships "
                  "much of the corpus itself",
        "notes": "25T tokens in two phases (23.5T diverse + 1.5T high-quality) "
                 "+ 121B long-context phase; 16 data categories. ~8-10T tokens "
                 "(~40-50% of the blend) released as open datasets "
                 "(Nemotron-CC-v2.x, CC-Math, Pretraining-Code, Specialized). "
                 "Documented for Nano/Super; Ultra follows the same data program.",
        "sources": "NVIDIA Nemotron 3 pretraining docs + Nano tech report; "
                   "HF Nemotron pretraining dataset collections",
    },
    {
        "model": "MiMo-V2.5", "developer": "Xiaomi", "open": True,
        "params_b": 310, "tokens_t": 48,
        "status": "Disclosed - largest disclosed corpus of any model here",
        "notes": "~48T tokens, FP8 mixed precision, five-stage pipeline (text "
                 "pretraining, projector warmup, multimodal pretraining, "
                 "SFT/agentic post-training with 32K->256K->1M context "
                 "extension, RL+MOPD). ~155 tokens per param - the most "
                 "over-trained large model on this list.",
        "sources": "XiaomiMiMo/MiMo-V2.5 HF card; mimo.xiaomi.com",
    },
    {
        "model": "Hy3", "developer": "Tencent", "open": True,
        "params_b": 295, "tokens_t": None,
        "status": "Undisclosed",
        "notes": "295B/21B MoE (192 experts, top-8 routing); Tencent's model "
                 "card discusses post-training scale-up but no corpus count.",
        "sources": "Tencent Hy3 model card; SiliconFlow explainer",
    },
    {
        "model": "MiniMax M3", "developer": "MiniMax", "open": True,
        "params_b": None, "tokens_t": None,
        "status": "Undisclosed",
        "notes": "No corpus size published for M3; total params also not "
                 "officially stated (third-party estimates ~200-400B).",
        "sources": "M3 launch materials",
    },
    {
        "model": "Llama 4 Scout", "developer": "Meta", "open": True,
        "params_b": 109, "tokens_t": 40,
        "status": "Disclosed (2025)",
        "notes": "~40T multimodal tokens - highest tokens-per-param ratio here "
                 "(~367x); 10M-token context via iRoPE.",
        "sources": "Meta Llama 4 launch blog (2025)",
    },
    {
        "model": "Llama 4 Maverick", "developer": "Meta", "open": True,
        "params_b": 400, "tokens_t": 22,
        "status": "Disclosed (2025)",
        "notes": "~22T multimodal tokens.",
        "sources": "Meta Llama 4 launch blog (2025)",
    },
]


def write_training_data_sheet(wb):
    ws = wb.create_sheet("Training Data")
    ws["A1"] = "Pretraining dataset sizes: what each lab disclosed"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Vendor-disclosed pretraining corpus sizes (trillions of tokens) for the models featured in this "
                "workbook; researched 2026-07-24. Closed labs disclose nothing, so their rows document that "
                "asymmetry rather than a number.")
    ws["A2"].font = Font(italic=True, color="595959")

    headers = ["Model", "Developer", "Class", "Total params (B)",
               "Pretraining tokens (T)", "Tokens per total param",
               "Disclosure status", "Data composition / pipeline notes", "Sources"]
    hr = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"A{hr + 1}"
    for i, m in enumerate(TRAINING_DATA):
        r = hr + 1 + i
        ratio = (m["tokens_t"] * 1000 / m["params_b"]
                 if m["tokens_t"] and m["params_b"] else None)
        vals = [m["model"], m["developer"],
                "Open-weights" if m["open"] else "Closed",
                m["params_b"] if m["params_b"] is not None else "undisclosed",
                m["tokens_t"] if m["tokens_t"] is not None else "undisclosed",
                ratio if ratio is not None else "—",
                m["status"], m["notes"], m["sources"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 3:
                cell.fill = OPEN_FILL if m["open"] else CLOSED_FILL
        if isinstance(ws.cell(row=r, column=4).value, (int, float)):
            ws.cell(row=r, column=4).number_format = "#,##0"
        if isinstance(ws.cell(row=r, column=5).value, (int, float)):
            ws.cell(row=r, column=5).number_format = "0.0"
        if isinstance(ws.cell(row=r, column=6).value, (int, float)):
            ws.cell(row=r, column=6).number_format = "#,##0x"
        ws.row_dimensions[r].height = 68
    autosize(ws, [22, 14, 12, 13, 14, 13, 34, 52, 30])

    nrow = hr + len(TRAINING_DATA) + 2
    notes = [
        "Disclosure asymmetry mirrors the rest of this workbook: every disclosed corpus size belongs to an "
        "open-weights model; Anthropic, OpenAI, Google, and Alibaba's closed tier publish nothing. NVIDIA goes "
        "furthest, releasing ~8-10T tokens of the actual corpus as open datasets.",
        "Tokens-per-parameter is a rough over-training indicator (Chinchilla-optimal is ~20x): DeepSeek V4 Pro sits "
        "near 20x while small-model releases (MiMo-V2.5 ~155x, Llama 4 Scout ~367x, Nemotron 3 Super ~208x) are "
        "heavily over-trained to maximize quality per active parameter at inference time.",
        "Counts are vendor-reported and not independently verifiable; token counts also are not comparable "
        "apples-to-apples across labs (different tokenizers, dedup policies, and multimodal counting).",
        "Kimi K3's corpus size should be published with its technical report alongside the weights (due 2026-07-27); "
        "this row can then be updated from 'undisclosed' to a disclosed figure.",
        "Pretraining corpus size is distinct from post-training/RL data, which is covered in the Harnesses & RL "
        "Envs tab.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")

    chart = BarChart()
    chart.type = "col"
    chart.title = "Disclosed pretraining corpus size (T tokens)"
    disclosed = [(m["model"], m["tokens_t"]) for m in TRAINING_DATA if m["tokens_t"]]
    start = nrow + len(notes) + 2
    ws.cell(row=start, column=1, value="Model")
    ws.cell(row=start, column=2, value="Pretraining tokens (T)")
    for i, (name, t) in enumerate(disclosed):
        ws.cell(row=start + 1 + i, column=1, value=name)
        ws.cell(row=start + 1 + i, column=2, value=t)
    data = Reference(ws, min_col=2, min_row=start, max_row=start + len(disclosed))
    cats = Reference(ws, min_col=1, min_row=start + 1, max_row=start + len(disclosed))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 9, 20
    chart.legend = None
    ws.add_chart(chart, f"D{start}")


def write_training_data_csv():
    with open(os.path.join(OUT_CSV_DIR, "training_data.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["model", "developer", "class", "total_params_b",
                    "pretraining_tokens_t", "tokens_per_total_param",
                    "disclosure_status", "notes", "sources"])
        for m in TRAINING_DATA:
            ratio = (round(m["tokens_t"] * 1000 / m["params_b"], 1)
                     if m["tokens_t"] and m["params_b"] else None)
            w.writerow([m["model"], m["developer"],
                        "open-weights" if m["open"] else "closed",
                        m["params_b"], m["tokens_t"], ratio,
                        m["status"], m["notes"], m["sources"]])


def write_csv_compilation_workbook():
    """Compile every CSV in data/csv/ into one data-only Excel workbook."""
    out_path = os.path.join(ROOT, "data", "llm_data_tables.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    for path in sorted(glob.glob(os.path.join(OUT_CSV_DIR, "*.csv"))):
        name = os.path.splitext(os.path.basename(path))[0][:31]
        ws = wb.create_sheet(name)
        with open(path, newline="") as f:
            for ri, row in enumerate(csv.reader(f), start=1):
                for ci, v in enumerate(row, start=1):
                    if v == "":
                        val = None
                    else:
                        try:
                            val = int(v)
                        except ValueError:
                            try:
                                val = float(v)
                            except ValueError:
                                val = v
                    ws.cell(row=ri, column=ci, value=val)
        style_header(ws, 1, ws.max_column)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
    wb.save(out_path)
    return out_path


# ------------------------------------------------------- Quantization fidelity (KLD)
# KL divergence measures how far a compressed/quantized variant's output
# distribution is from ITS OWN full-precision reference - it is not defined
# across different models (different tokenizers/vocabularies). This table
# therefore compares each model's quantization-fidelity story, which is the
# only meaningful "KLD comparison" across models. Researched 2026-07-22.
QUANT_KLD = [
    {
        "model": "Kimi K3", "developer": "Moonshot AI", "open": True,
        "precision": "Native MXFP4 weights + MXFP8 activations; QAT from the SFT "
                     "stage (the 4-bit checkpoint IS the model - no higher-"
                     "precision reference exists)",
        "measurable": "Not yet - weights due 2026-07-27; nothing on HF as of "
                      "2026-07-22",
        "kld_native": "~0 by construction (QAT native format)",
        "kld_4bit": "n/a - the native format is already 4-bit",
        "kld_2bit": "Pending community quants (Unsloth gates sub-4-bit releases "
                    "on measured KLD)",
        "notes": "Expected to mirror DeepSeek V4 Flash: native MXFP4 experts "
                 "repack bit-for-bit into GGUF (KLD ~0, 100% top-token "
                 "agreement). ~1.4TB at MXFP4 vs ~5.6TB BF16-equivalent.",
        "sources": "kimi.com/blog/kimi-k3; HF community analysis",
    },
    {
        "model": "Kimi K2.6 / K2.7", "developer": "Moonshot AI", "open": True,
        "precision": "Native INT4 QAT for MoE weights, BF16 elsewhere (1T params)",
        "measurable": "Yes (weights public)",
        "kld_native": "UD-Q8_K_XL 'truly lossless' (KLD ~0; PPL 1.8419)",
        "kld_4bit": "UD-Q4_K_XL near-lossless: PPL 1.8420 vs 1.8419 lossless "
                    "(584GB vs 595GB)",
        "kld_2bit": "Dynamic 2-bit: PPL 2.4131 (340GB, -43% size); 1/3-bit "
                    "uploads gated on KLD scores",
        "notes": "Because MoE weights are already INT4-native, Q8 repack is "
                 "bit-exact; only the BF16 non-expert tensors can degrade in "
                 "smaller quants.",
        "sources": "Unsloth Kimi K2.6 guide; unsloth/Kimi-K2.6-GGUF",
    },
    {
        "model": "DeepSeek V4 Flash", "developer": "DeepSeek", "open": True,
        "precision": "Native MXFP4 routed experts (96% of params) + FP8/BF16 "
                     "rest; QAT (284B/13B)",
        "measurable": "Yes - the best-measured KLD ladder of any frontier model "
                      "(see ladder table below)",
        "kld_native": "UD-Q8_K_XL: KLD ~0, 100% top-token agreement, all 1,328 "
                      "tensors bit-identical to official weights",
        "kld_4bit": "UD-Q4_K_XL mean KLD 0.0102 (96.3% top-token); bartowski "
                    "MXFP4 0.0105; third-party imatrix Q4 0.0290-0.0291 (93.9%)",
        "kld_2bit": "0.36-0.42 mean KLD, only ~78% top-token agreement "
                    "(87-98GB) - the quality cliff",
        "notes": "Canonical example of QAT killing the quantization tax at "
                 "native precision; KLD only reappears in sub-4-bit community "
                 "compressions.",
        "sources": "Unsloth DeepSeek-V4 guide (wikitext-2, ctx 512, 4x B200)",
    },
    {
        "model": "DeepSeek V4 Pro", "developer": "DeepSeek", "open": True,
        "precision": "Same family: native MXFP4 experts + FP8 mixed (1.6T/49B)",
        "measurable": "Yes (weights public)",
        "kld_native": "Native repack lossless by same mechanism as Flash",
        "kld_4bit": "No published ladder (community measured Flash first)",
        "kld_2bit": "No published ladder",
        "notes": "Expect Flash-like profile; 1.6T scale makes full KLD runs "
                 "expensive, so trackers prioritized Flash.",
        "sources": "DeepSeek V4 HF card; Unsloth docs",
    },
    {
        "model": "GLM-5.2", "developer": "Z.ai (Zhipu)", "open": True,
        "precision": "BF16/FP8 release (753B/40B); no QAT claimed",
        "measurable": "Yes (weights public, MIT)",
        "kld_native": "No first-party KLD published",
        "kld_4bit": "Community GGUF quants follow the standard Unsloth Dynamic "
                    "pattern; no headline KLD table published",
        "kld_2bit": "Not published",
        "notes": "Without QAT, 4-bit quants carry the normal ~0.004-0.012 KLD "
                 "format tax (see format table below) rather than ~0.",
        "sources": "z.ai/blog/glm-5.2; community GGUF repos",
    },
    {
        "model": "MiniMax M3", "developer": "MiniMax", "open": True,
        "precision": "Open weights; MSA sparse-attention architecture",
        "measurable": "Yes (weights public)",
        "kld_native": "No KLD published",
        "kld_4bit": "Not published",
        "kld_2bit": "Not published",
        "notes": "No first-party or major community KLD measurements found as "
                 "of 2026-07-22.",
        "sources": "M3 launch materials",
    },
    {
        "model": "NVIDIA Nemotron 3 Ultra", "developer": "NVIDIA", "open": True,
        "precision": "BF16 reference + NVFP4 post-training-quantized release "
                     "(550B/55B)",
        "measurable": "Yes, but NVIDIA reports benchmark deltas instead of KLD",
        "kld_native": "BF16 = reference (KLD 0 by definition)",
        "kld_4bit": "NVFP4 quality cost measured on benchmarks: AA Intelligence "
                    "Index 48.2 -> 47.7; GPQA 87.0 -> 87.9; SWE-bench Verified "
                    "71.9 -> 69.7; Terminal-Bench 56.4 -> 53.9",
        "kld_2bit": "Not published",
        "notes": "The PTQ (not QAT) counter-example: a real, measurable quality "
                 "delta between precisions - exactly what QAT-native releases "
                 "avoid.",
        "sources": "NVIDIA model card (BF16 vs NVFP4 columns); AA launch article",
    },
    {
        "model": "Claude Fable 5", "developer": "Anthropic", "open": False,
        "precision": "Undisclosed (serving precision unknown)",
        "measurable": "No - no weight access; API exposes no logprobs; no "
                      "reference distribution exists for outsiders",
        "kld_native": "Unmeasurable",
        "kld_4bit": "Unmeasurable",
        "kld_2bit": "Unmeasurable",
        "notes": "Any internal quantization/distillation between training and "
                 "serving is invisible; users cannot audit serving fidelity.",
        "sources": "Anthropic API docs (no logprobs surface)",
    },
    {
        "model": "GPT-5.6 Sol", "developer": "OpenAI", "open": False,
        "precision": "Undisclosed",
        "measurable": "No - no weights; API returns top-k logprobs only, and "
                      "there is no full-precision reference to diverge from",
        "kld_native": "Unmeasurable",
        "kld_4bit": "Unmeasurable",
        "kld_2bit": "Unmeasurable",
        "notes": "Top-k logprobs allow drift monitoring over time, but not KLD "
                 "against a reference checkpoint.",
        "sources": "OpenAI API docs",
    },
    {
        "model": "Qwen3.7 Max", "developer": "Alibaba (Qwen)", "open": False,
        "precision": "Undisclosed (API-only tier)",
        "measurable": "No - closed weights since late 2025 for the Max tier",
        "kld_native": "Unmeasurable",
        "kld_4bit": "Unmeasurable",
        "kld_2bit": "Unmeasurable",
        "notes": "Alibaba's open Qwen3.x models are measurable (community GGUF "
                 "KLD tables exist, e.g. Qwen3.6-27B used as the format "
                 "reference below); the Max tier is not.",
        "sources": "Qwen release notes",
    },
]

# Measured KLD ladder for DeepSeek V4 Flash (closest public proxy for what
# K3's post-release ladder will look like). Source: Unsloth, wikitext-2,
# ctx 512, vs official weights.
KLD_LADDER = [
    ("Official checkpoint (reference)", 156.4, 4.5319, 0.0, 0.0, 100.0),
    ("Unsloth UD-Q8_K_XL", 161.9, 4.5319, 0.0, 0.0, 100.0),
    ("Unsloth UD-Q4_K_XL", 155.1, 4.5335, 0.0102, 3.40, 96.28),
    ("bartowski MXFP4", 156.0, 4.5351, 0.0105, 3.42, 96.18),
    ("antirez Q4KExperts-F16 (imatrix)", 164.6, 4.5743, 0.0291, 5.87, 93.95),
    ("antirez mixed L37-42-Q4K (imatrix)", 97.6, 5.8169, 0.3605, 21.15, 79.74),
    ("antirez IQ2XXS (imatrix)", 86.7, 6.0808, 0.4079, 22.23, 78.15),
]

# 4-bit-class format comparison: mean KLD vs FP16 reference measured on
# Qwen3.6-27B (SpecPicks 2026 format review). MXFP4 is K3's native format.
KLD_FORMATS = [
    (4.0, 0.0123, 0.0078, 0.0061, 0.0042),
    (4.5, 0.0094, 0.0058, 0.0047, 0.0031),
    (6.0, 0.0028, 0.0019, 0.0015, 0.0017),
    (8.0, 0.0009, 0.0008, 0.0007, 0.0006),
]


def write_quant_kld_sheet(wb):
    ws = wb.create_sheet("Quantization & KLD")
    ws["A1"] = "KLD (quantization fidelity): Kimi K3 vs open-weights peers and closed frontier models"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("KL divergence measures a quantized variant against ITS OWN full-precision reference; it is not "
                "defined across different models (different tokenizers/vocabularies). The comparable question is "
                "each model's quantization-fidelity profile. Researched 2026-07-22.")
    ws["A2"].font = Font(italic=True, color="595959")

    headers = [
        "Model", "Developer", "Class", "Release precision / QAT",
        "KLD measurable by outsiders?", "KLD at native / lossless tier",
        "KLD at 4-bit tier", "KLD at ~2-bit tier", "Notes", "Sources",
    ]
    hr = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"A{hr + 1}"
    for i, m in enumerate(QUANT_KLD):
        r = hr + 1 + i
        vals = [m["model"], m["developer"],
                "Open-weights" if m["open"] else "Closed",
                m["precision"], m["measurable"], m["kld_native"],
                m["kld_4bit"], m["kld_2bit"], m["notes"], m["sources"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 3:
                cell.fill = OPEN_FILL if m["open"] else CLOSED_FILL
        ws.row_dimensions[r].height = 95
    autosize(ws, [19, 14, 12, 34, 26, 30, 34, 26, 36, 26])

    # Measured ladder (DeepSeek V4 Flash)
    lr = hr + len(QUANT_KLD) + 2
    ws.cell(row=lr, column=1, value="Measured KLD ladder - DeepSeek V4 Flash (closest public proxy for K3's "
                                    "future ladder; Unsloth, wikitext-2)").font = Font(bold=True, size=12)
    ladder_headers = ["Quant", "Size (GB)", "Perplexity", "Mean KLD", "RMS delta-p (%)", "Same top token (%)"]
    for c, h in enumerate(ladder_headers, 1):
        ws.cell(row=lr + 1, column=c, value=h)
    style_header(ws, lr + 1, len(ladder_headers))
    for i, (name, gb, ppl, kld, dp, tt) in enumerate(KLD_LADDER):
        r = lr + 2 + i
        for c, v in enumerate([name, gb, ppl, kld, dp, tt], 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
        ws.cell(row=r, column=3).number_format = "0.0000"
        ws.cell(row=r, column=4).number_format = "0.0000"
        ws.cell(row=r, column=5).number_format = "0.00"
        ws.cell(row=r, column=6).number_format = "0.00"

    # Format-level KLD reference
    fr = lr + 2 + len(KLD_LADDER) + 1
    ws.cell(row=fr, column=1, value="4-bit-class format KLD vs FP16 (measured on Qwen3.6-27B; lower = closer to "
                                    "full precision; MXFP4 is K3's native format)").font = Font(bold=True, size=12)
    fmt_headers = ["Bits per weight", "Q4_K_M (classic GGUF)", "oQ ('optimal Q')", "MXFP4/MXFP6", "UD-MLX"]
    for c, h in enumerate(fmt_headers, 1):
        ws.cell(row=fr + 1, column=c, value=h)
    style_header(ws, fr + 1, len(fmt_headers))
    for i, row_ in enumerate(KLD_FORMATS):
        r = fr + 2 + i
        for c, v in enumerate(row_, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c > 1:
                cell.number_format = "0.0000"

    nrow = fr + 2 + len(KLD_FORMATS) + 1
    notes = [
        "Why there is no 'K3 KLD vs Fable 5 KLD' number: KL divergence compares two probability distributions over "
        "the same vocabulary. Different models have different tokenizers and no shared reference, so cross-model KLD "
        "is undefined. The comparable dimension is how faithfully each model's served/compressed form reproduces its "
        "own reference - and whether outsiders can measure that at all.",
        "The open-weights frontier has engineered the quantization tax away: K3, K2.x, and DeepSeek V4 are "
        "quantization-aware-trained so the low-precision checkpoint IS the model (KLD ~0 at native precision). KLD "
        "only becomes a live issue for sub-4-bit community compressions, where ~2-bit costs 0.36+ KLD and ~22% "
        "top-token disagreement (the quality cliff).",
        "Closed models are unauditable on this axis: Anthropic exposes no logprobs, OpenAI only top-k, and neither "
        "releases a reference checkpoint - so serving fidelity is a trust relationship, not a measurement. With open "
        "models you can verify token-level fidelity yourself (Unsloth verified DeepSeek V4 Flash bit-identical).",
        "Interpretation caveat (July 2026 community finding): KLD has a 'silent zone' near baseline where further "
        "KLD reductions do not translate into quality gains; Unsloth and the 'Accuracy is Not All You Need' paper "
        "recommend pairing KLD with flip-sensitive benchmarks (e.g. Aider) rather than reading tiny KLD deltas.",
        "Kimi K3 status: weights promised by 2026-07-27 (not on Hugging Face as of 2026-07-22). Expect an "
        "Unsloth-style measured ladder within days of release; until then the K3 row is by-construction reasoning, "
        "not measurement.",
        "Nemotron 3 Ultra is the post-training-quantization counter-example: NVIDIA ships BF16 + NVFP4 and reports "
        "the quality delta on benchmarks (AA Index 48.2 vs 47.7) rather than KLD.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")


def write_quant_kld_csv():
    with open(os.path.join(OUT_CSV_DIR, "quantization_kld.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "model", "developer", "class", "release_precision_qat", "kld_measurable",
            "kld_native_lossless_tier", "kld_4bit_tier", "kld_2bit_tier", "notes", "sources",
        ])
        for m in QUANT_KLD:
            w.writerow([
                m["model"], m["developer"],
                "open-weights" if m["open"] else "closed",
                m["precision"], m["measurable"], m["kld_native"],
                m["kld_4bit"], m["kld_2bit"], m["notes"], m["sources"],
            ])


# ----------------------------------------------------- Harness & RL environments
# Hand-curated comparison of agent harnesses and RL training environments
# (researched 2026-07-22). Sources per row; closed labs disclose harness
# architecture but not RL environments, so those cells describe what is
# publicly known or stated by the lab.
HARNESS_RL = [
    {
        "model": "Kimi K3", "developer": "Moonshot AI", "open": True,
        "harness": "Kimi Code CLI (KimiCode harness); Kimi.com / Kimi Work apps; "
                   "K3 Swarm Max orchestrator for up to 300 parallel sub-agents",
        "compat": "OpenAI- and Anthropic-compatible APIs; MCP tools. Hard "
                  "requirement: preserved thinking history - harnesses must pass "
                  "back all prior reasoning or quality destabilizes; Moonshot "
                  "recommends verified harnesses (Kimi Code) and warns against "
                  "mid-session model switches",
        "rl_stack": "Parallel Agent RL (PARL): trainable orchestrator with frozen "
                    "sub-agents for credit assignment (Agent Swarm lineage from "
                    "K2.5); per-head Muon optimizer; quantile-balanced MoE "
                    "routing; trained in preserved-thinking mode; QAT from SFT "
                    "stage (MXFP4/MXFP8)",
        "rl_envs": "Not released. Disclosed examples: 24h GPU kernel-optimization "
                   "sandboxes (H200 + non-NVIDIA GPGPU), 48h chip-design run on "
                   "open-source EDA tools, internal knowledge-work environments "
                   "derived from real user-agent workflows. Technical report due "
                   "with weights (by 2026-07-27)",
        "openness": "Weights open (due 2026-07-27); RL framework and environments "
                    "not released; contributed KDA prefix-caching to vLLM",
        "eval_harness": "KimiCode (own evals); mini-SWE-agent on DeepSWE "
                        "leaderboard; Claude Code for SWE Marathon/OfficeQA",
        "sources": "kimi.com/blog/kimi-k3; Kimi API docs",
    },
    {
        "model": "Claude Fable 5", "developer": "Anthropic", "open": False,
        "harness": "Claude Code (CLI + SDK) - the reference harness most labs "
                   "benchmark against; Claude Cowork for knowledge work",
        "compat": "Anthropic API only (adaptive thinking always on; effort "
                  "parameter replaces thinking budgets). Harness-level features "
                  "co-designed with the model: memory tool, context editing / "
                  "tool-result clearing, compaction, task budgets, code execution, "
                  "programmatic tool calling; no assistant prefill",
        "rl_stack": "Undisclosed. Known: RLHF + RL on long-horizon agentic tasks; "
                    "adaptive reasoning with effort levels (low->max); Opus 4.8 "
                    "fallback via safety classifiers (Fable) on cyber/bio/"
                    "distillation requests",
        "rl_envs": "Undisclosed. System card documents evaluation (not training) "
                   "environments across SWE, terminal, multi-agent, computer-use "
                   "(OSWorld), finance/legal/healthcare agents. Anthropic "
                   "publishes no RL training data or environments",
        "openness": "Fully closed (weights, data, RL stack); harness (Claude "
                    "Code) is a product with open plugin/MCP ecosystem",
        "eval_harness": "Claude Code / Terminus-2 (AA); vendor system card",
        "sources": "Anthropic Fable 5 launch + system card; platform.claude.com docs",
    },
    {
        "model": "GPT-5.6 Sol", "developer": "OpenAI", "open": False,
        "harness": "Codex suite: Codex CLI (agent loop), Codex Cloud, VS Code "
                   "extension, App Server (JSON-RPC protocol exposing the "
                   "harness to products)",
        "compat": "Responses API drives the agent loop (structured context: repo "
                  "metadata, file tree, diffs, command outputs; automatic "
                  "compaction; sandboxed execution). Sol defaults to subagent-"
                  "orchestrator mode; known open issue: subagent config fields "
                  "hidden, blocking cheap Terra/Luna subagent routing "
                  "(openai/codex#31814)",
        "rl_stack": "Undisclosed. Known: RL optimized for multi-step execution "
                    "(plan->implement->validate->repair) and long uninterrupted "
                    "runs (~25h reported); reasoning modes incl. 'pro'; models "
                    "explicitly tuned for the Codex harness",
        "rl_envs": "Undisclosed. Public signals: Agents' Last Exam-style "
                   "long-horizon professional workflows (55 fields), OSWorld "
                   "computer use, terminal tasks. No RL environments or data "
                   "released",
        "openness": "Fully closed for GPT-5.x; separately ships open-weight "
                    "gpt-oss models (Apache 2.0, no RL stack)",
        "eval_harness": "Codex (own evals and AA Coding Agent Index)",
        "sources": "OpenAI 'Unrolling the Codex agent loop', 'Unlocking the "
                   "Codex harness', GPT-5.6 launch post",
    },
    {
        "model": "GLM-5 / GLM-5.2", "developer": "Z.ai (Zhipu)", "open": True,
        "harness": "No own CLI - ships Anthropic-native API designed to drop into "
                   "Claude Code, OpenClaw, Cline etc. (GLM Coding Plan $3-80/mo); "
                   "positions third-party harnesses as the product surface",
        "compat": "Anthropic-native + OpenAI-compatible; parallel tool-call API "
                  "support (rare among peers); 1M context with dynamic working-"
                  "memory management for thousands-of-tool-call sessions",
        "rl_stack": "slime (open-source, Megatron+SGLang) - the disclosed RL "
                    "framework behind GLM-4.5->5.2. Sequential pipeline: Reasoning "
                    "RL -> Agentic RL -> General RL with on-policy cross-stage "
                    "distillation; fully async decoupled rollout/training; TITO "
                    "gateway (token-in-token-out); double-sided importance "
                    "sampling; hybrid rewards (rule-based + ORM + GRM)",
        "rl_envs": "Most detailed public disclosure of any lab: 10K+ verifiable "
                   "SWE environments built with RepoLaunch from real issue-PR "
                   "pairs across 9 languages; synthesized terminal environments "
                   "in Harbor format (Dockerized, >90% build accuracy); slide-"
                   "generation environment with executable HTML rendering "
                   "verification; multi-hop search tasks. Environments themselves "
                   "not released, but pipeline is documented",
        "openness": "Weights open (MIT); RL framework open (slime on GitHub); "
                    "environment-building pipelines documented in GLM-5 paper; "
                    "environment data not released",
        "eval_harness": "Claude Code / Terminus-2",
        "sources": "GLM-5 paper (arXiv:2602.15763); THUDM/slime; z.ai/blog/glm-5.2",
    },
    {
        "model": "DeepSeek V4 Pro / Flash", "developer": "DeepSeek", "open": True,
        "harness": "No own harness - dual-mode OpenAI + Anthropic-compatible API "
                   "(drops into Codex- or Claude Code-style harnesses); new "
                   "XML-based DSML tool-call schema replacing JSON",
        "compat": "OpenAI + Anthropic API compatibility; three reasoning modes "
                  "(non-think / high / max); 1M context",
        "rl_stack": "Two-stage post-training that replaces mixed RL: (1) domain "
                    "specialists (math, code, agent, instruction following) each "
                    "trained with SFT + GRPO where the actor natively functions "
                    "as its own generative reward model (no separate RM); (2) "
                    "10+ specialist teachers merged into one student via "
                    "multi-teacher on-policy distillation (weighted reverse-KL)",
        "rl_envs": "Sandbox infrastructure for agentic RL documented in the V4 "
                   "report (Sec 5.2.5) incl. million-token-context RL rollouts; "
                   "agentic tool-use data injected already in mid-training. "
                   "Environments and RL data not released",
        "openness": "Weights open (MIT) incl. base checkpoints; detailed "
                    "technical report; RL sandbox infra and data not released",
        "eval_harness": "Own harness + mini-SWE-agent (DeepSWE); Terminus (TB 2.0)",
        "sources": "DeepSeek-V4 report (arXiv:2606.19348)",
    },
    {
        "model": "Nemotron 3 (Nano/Super/Ultra)", "developer": "NVIDIA", "open": True,
        "harness": "NeMo Gym ships out-of-the-box harnesses (incl. Claude Code "
                   "and Hermes) plus bring-your-own-agent interfaces; served via "
                   "NVIDIA API / NIM",
        "compat": "OpenAI-compatible endpoints; models tuned for high-throughput "
                  "agentic serving on Blackwell (NVFP4)",
        "rl_stack": "Fully open: NeMo RL + NeMo Gym (RLVR - RL from verifiable "
                    "rewards). Super trained in 6 disclosed RL stages (RLVR 1-3, "
                    "SWE 1-2, RLHF) with pass-rate-ordered curriculum; Ultra adds "
                    "multi-teacher on-policy distillation (MOPD). Also usable "
                    "with VeRL/Unsloth",
        "rl_envs": "The only lab that RELEASES its RL environments and data: "
                   "NeMo Gym environment hub (70+ environments - coding, math, "
                   "tool use, workplace-assistant agents, SWE-Gym, DAPO-Math, "
                   "Skywork-OR1...) plus published per-stage RL training blends "
                   "for Nano/Super/Ultra on Hugging Face with mixing ratios and "
                   "recipes",
        "openness": "Maximal: open weights + open RL environments + open RL "
                    "data blends + training recipes + tech reports (NVIDIA Open "
                    "Model License)",
        "eval_harness": "NeMo Evaluator SDK / NeMo Skills; OpenHands for SWE-bench",
        "sources": "NVIDIA-NeMo/Gym; HF nvidia/Nemotron-RL-*-Training-Blends",
    },
    {
        "model": "MiniMax M3", "developer": "MiniMax", "open": True,
        "harness": "MiniMax Agent product; native computer use (only open model "
                   "here with it); no dedicated CLI - OpenAI/Anthropic-compatible "
                   "API for third-party harnesses",
        "compat": "OpenAI + Anthropic-compatible; text+image+video input; 1M "
                  "context (MSA sparse attention)",
        "rl_stack": "CISPO lineage (clips importance-sampling weights instead of "
                    "token updates - M1 paper); large-scale RL on sandbox-based "
                    "real-world SWE environments; M1 full RL run cost ~$535K on "
                    "512 H800s in 3 weeks. M3-specific recipe not yet published",
        "rl_envs": "Sandbox-based real-world software engineering environments "
                   "(disclosed in M1 paper); 24-hour autonomous runs with ~2,000 "
                   "tool calls reported for M3. Environments not released",
        "openness": "Weights open; CISPO algorithm published; environments and "
                    "RL data not released",
        "eval_harness": "Own agent product; Terminus-2 (AA)",
        "sources": "MiniMax-M1 paper (arXiv:2506.13585); M3 launch materials",
    },
    {
        "model": "Qwen3.7 Max", "developer": "Alibaba (Qwen)", "open": False,
        "harness": "Qwen Chat + DashScope API; agent-first positioning with "
                   "native OpenAI- and Anthropic-API compatibility (works in "
                   "Claude Code / Codex-style harnesses); retains reasoning "
                   "across turns",
        "compat": "OpenAI + Anthropic-compatible; extended-thinking mode "
                  "(enable_thinking / preserve_thinking); 1M context",
        "rl_stack": "Undisclosed for the closed Max tier. Public signals: "
                    "QwenWorldBench - internal benchmark using LLMs as world "
                    "models to simulate agentic environments across 7 domains "
                    "(Terminal, SWE, MCP, Search, OS, Android, Web) - implies "
                    "world-model-driven agent training; 35h autonomous "
                    "kernel-optimization sandbox (Docker + H100, CUTLASS/CUDA "
                    "docs only) with GPT-5.4 as anti-reward-hacking judge",
        "rl_envs": "Not released for Max. Open Qwen3.x models ship weights "
                   "(Apache 2.0) but not RL environments",
        "openness": "Max tier fully closed since late 2025; smaller Qwen models "
                    "open-weight",
        "eval_harness": "Claude Code harness for agentic evals (per Qwen3.7 blog)",
        "sources": "Qwen3.7 blog (alibabacloud.com); deeplearning.ai The Batch",
    },
]


def write_harness_rl_sheet(wb):
    ws = wb.create_sheet("Harnesses & RL Envs")
    ws["A1"] = "Agent harnesses and RL training environments: Kimi K3 vs open-weights peers and frontier labs"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("What each lab ships as its agent harness, what is known about its RL post-training stack and "
                "environments, and how open that stack is (researched 2026-07-22)")
    ws["A2"].font = Font(italic=True, color="595959")

    headers = [
        "Model / family", "Developer", "Class", "First-party agent harness",
        "API / harness compatibility", "RL algorithm & training stack",
        "RL environments (disclosed)", "Openness of post-training stack",
        "Eval harness used", "Sources",
    ]
    hr = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(headers))
    ws.freeze_panes = f"A{hr + 1}"

    for i, m in enumerate(HARNESS_RL):
        r = hr + 1 + i
        vals = [
            m["model"], m["developer"],
            "Open-weights" if m["open"] else "Closed",
            m["harness"], m["compat"], m["rl_stack"], m["rl_envs"],
            m["openness"], m["eval_harness"], m["sources"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 3:
                cell.fill = OPEN_FILL if m["open"] else CLOSED_FILL
        ws.row_dimensions[r].height = 150
    autosize(ws, [17, 13, 12, 38, 38, 44, 44, 30, 24, 26])

    # Shared third-party infrastructure block
    lb = hr + len(HARNESS_RL) + 2
    ws.cell(row=lb, column=1, value="Shared harness & RL-environment infrastructure referenced above").font = Font(bold=True, size=12)
    infra = [
        ("Claude Code", "Anthropic's CLI harness - de-facto industry reference; GLM, Kimi, and Qwen all "
                        "advertise compatibility with it and use it in their own evals"),
        ("Codex CLI / App Server", "OpenAI's harness; agent loop on the Responses API; also the eval harness for GPT models"),
        ("Kimi Code", "Moonshot's CLI; only harness with verified preserved-thinking support for K3"),
        ("Terminus-2", "Terminal-Bench's neutral reference harness, used by Artificial Analysis for cross-model evals"),
        ("mini-SWE-agent / OpenHands", "Open community harnesses used by DeepSWE and SWE-bench leaderboards"),
        ("Harbor", "Dockerized task format for terminal/agent environments; used by Z.ai's synthesized "
                   "environments and PostTrain Bench"),
        ("RepoLaunch", "Pipeline that turns real GitHub issue-PR pairs into executable, verifiable SWE "
                       "environments; basis of Z.ai's 10K+ RL environments"),
        ("NeMo Gym", "NVIDIA's open RL-environment hub (70+ environments, RLVR datasets, harness "
                     "integrations incl. Claude Code and Hermes) - the largest open RL environment release"),
        ("slime", "Z.ai/THUDM's open RL post-training framework (Megatron + SGLang); battle-tested on "
                  "GLM-4.5 through GLM-5.2; also supports Qwen, DeepSeek, Llama, Kimi K2"),
    ]
    for c, h in enumerate(["Component", "Role"], 1):
        ws.cell(row=lb + 1, column=c, value=h)
    style_header(ws, lb + 1, 2)
    for i, (name, desc) in enumerate(infra):
        r = lb + 2 + i
        ws.cell(row=r, column=1, value=name).border = BORDER
        cell = ws.cell(row=r, column=2, value=desc)
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 28

    nrow = lb + 2 + len(infra) + 1
    notes = [
        "Openness spectrum (least to most): OpenAI / Anthropic / Qwen Max disclose harness architecture but no RL "
        "environments or data -> Moonshot / DeepSeek / MiniMax release weights + papers but keep environments "
        "private -> Z.ai releases weights + the RL framework (slime) and documents its environment pipelines -> "
        "NVIDIA releases the entire stack (weights, environments, RL data blends, recipes).",
        "A structural trend across labs: models and harnesses are co-trained and entangled. K3 requires preserved "
        "thinking history (unstable in non-verified harnesses), GPT-5.6 Sol is tuned for Codex and Fable 5 for "
        "Claude Code, and AA's Coding Agent Index now benchmarks model+harness pairs rather than bare models.",
        "RL environment scale is becoming a disclosed competitive metric: Z.ai reports 10K+ verifiable SWE "
        "environments; NVIDIA's NeMo Gym ships 70+ public environments; DeepSeek and Moonshot describe sandbox "
        "infrastructure without counts; closed labs disclose nothing.",
        "Closed-lab rows describe publicly known information only; actual RL stacks are proprietary and may differ.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")


def write_harness_rl_csv():
    with open(os.path.join(OUT_CSV_DIR, "harness_rl_environments.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "model", "developer", "class", "first_party_harness", "api_compatibility",
            "rl_algorithm_stack", "rl_environments_disclosed", "post_training_openness",
            "eval_harness", "sources",
        ])
        for m in HARNESS_RL:
            w.writerow([
                m["model"], m["developer"],
                "open-weights" if m["open"] else "closed",
                m["harness"], m["compat"], m["rl_stack"], m["rl_envs"],
                m["openness"], m["eval_harness"], m["sources"],
            ])


# ------------------------------------------------------------- Cloud providers
# Category map for every serving provider seen in the endpoint-stats snapshot.
# "Hyperscaler" = general-purpose public cloud; "Model lab" = the model
# developer's own first-party API; "AI inference cloud" = independent
# GPU/inference-specialist provider.
PROVIDER_CATEGORY = {
    "Amazon Bedrock": "Hyperscaler (US)", "Azure": "Hyperscaler (US)",
    "Google": "Hyperscaler (US)", "Google AI Studio": "Hyperscaler (US)",
    "Cloudflare": "Hyperscaler (US)", "DigitalOcean": "Hyperscaler (US)",
    "Alibaba": "Hyperscaler (China)", "Baidu": "Hyperscaler (China)",
    "Tencent": "Hyperscaler (China)", "StreamLake": "Hyperscaler (China)",
    "OpenAI": "Model lab", "Anthropic": "Model lab", "DeepSeek": "Model lab",
    "Moonshot AI": "Model lab", "Minimax": "Model lab", "Mistral": "Model lab",
    "xAI": "Model lab", "Z.AI": "Model lab", "Xiaomi": "Model lab",
    "StepFun": "Model lab", "Poolside": "Model lab", "Nvidia": "Model lab",
    "Cohere": "Model lab", "Morph": "Model lab",
}
DEFAULT_PROVIDER_CATEGORY = "AI inference cloud"


def build_cloud_split(rows):
    """Estimate weekly tokens per serving provider by distributing each
    model-variant's weekly volume across its endpoints proportionally to the
    endpoints' live 30-minute request counts."""
    stats_path = sorted(glob.glob(os.path.join(RAW, "openrouter_endpoint_stats_*.json")))[-1]
    prov_path = sorted(glob.glob(os.path.join(RAW, "openrouter_providers_*.json")))[-1]
    with open(stats_path) as f:
        snapshot = json.load(f)
    with open(prov_path) as f:
        providers_dir = json.load(f)
    hq = {p["displayName"]: p.get("headquarters") or "" for p in providers_dir}
    hq.update({p["name"]: p.get("headquarters") or "" for p in providers_dir})

    open_by_slug = {r["slug"]: r["open"] for r in rows}
    prov = defaultdict(lambda: {"open": 0.0, "closed": 0.0, "models": set()})
    covered = unattributed = 0
    for m in snapshot["data"]:
        tokens = m["weekly_tokens"]
        is_open = open_by_slug.get(m["model_permaslug"])
        if is_open is None:
            unattributed += tokens
            continue
        total_req = sum(ep["request_count_30m"] for ep in m["endpoints"])
        if total_req == 0:
            unattributed += tokens
            continue
        covered += tokens
        cls = "open" if is_open else "closed"
        for ep in m["endpoints"]:
            share = ep["request_count_30m"] / total_req
            p = prov[ep["provider_name"]]
            p[cls] += tokens * share
            if share > 0:
                p["models"].add(m["model_permaslug"])

    out = []
    for name, p in sorted(prov.items(), key=lambda kv: -(kv[1]["open"] + kv[1]["closed"])):
        out.append(
            {
                "provider": name,
                "category": PROVIDER_CATEGORY.get(name, DEFAULT_PROVIDER_CATEGORY),
                "hq": hq.get(name, ""),
                "open": p["open"],
                "closed": p["closed"],
                "total": p["open"] + p["closed"],
                "models": len(p["models"]),
            }
        )
    cats = defaultdict(lambda: {"open": 0.0, "closed": 0.0, "providers": 0})
    for r in out:
        c = cats[r["category"]]
        c["open"] += r["open"]
        c["closed"] += r["closed"]
        c["providers"] += 1
    meta = {
        "fetched_at": snapshot["fetched_at"],
        "covered": covered,
        "unattributed": unattributed,
        "n_pairs": len(snapshot["data"]),
    }
    return out, dict(cats), meta


def write_cloud_sheet(wb, cloud_rows, cloud_cats, meta, grand_total):
    ws = wb.create_sheet("By Cloud Provider")
    ws["A1"] = "Where LLM tokens are served: open vs closed volume by cloud / inference provider"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"OpenRouter-routed traffic. Weekly volume (week ending {WEEK_ENDING}) distributed across each "
                f"model's serving endpoints by live request share (endpoint snapshot {meta['fetched_at']}).")
    ws["A3"] = (f"Covers the top {meta['n_pairs']} model-variant pairs = {meta['covered'] / 1e12:.1f}T of "
                f"{grand_total / 1e12:.1f}T weekly tokens ({meta['covered'] / grand_total:.0%}); "
                "long-tail models are not attributed.")
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="595959")

    # --- Section 1: by provider category
    hr = 5
    cat_headers = ["Provider category", "# providers", "Open-weights tokens (B)",
                   "Closed tokens (B)", "Total (B)", "Share of covered",
                   "Open share within category"]
    for c, h in enumerate(cat_headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, len(cat_headers))
    order = ["Model lab", "AI inference cloud", "Hyperscaler (US)", "Hyperscaler (China)"]
    covered = meta["covered"]
    for i, cat in enumerate(order):
        d = cloud_cats.get(cat, {"open": 0, "closed": 0, "providers": 0})
        r = hr + 1 + i
        tot = d["open"] + d["closed"]
        vals = [cat, d["providers"], d["open"] / 1e9, d["closed"] / 1e9, tot / 1e9,
                tot / covered if covered else 0, d["open"] / tot if tot else 0]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
        for c in (3, 4, 5):
            ws.cell(row=r, column=c).number_format = "#,##0.0"
        for c in (6, 7):
            ws.cell(row=r, column=c).number_format = "0.0%"

    # --- Section 2: per provider
    hr2 = hr + len(order) + 2
    ws.cell(row=hr2 - 1, column=1, value="Per provider (estimated weekly tokens)").font = Font(bold=True, size=12)
    headers = ["Provider", "Category", "HQ", "Open-weights tokens (B)", "Closed tokens (B)",
               "Total (B)", "Share of covered", "Open share", "# models served"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr2, column=c, value=h)
    style_header(ws, hr2, len(headers))
    for i, r_ in enumerate(cloud_rows):
        r = hr2 + 1 + i
        tot = r_["total"]
        vals = [r_["provider"], r_["category"], r_["hq"],
                r_["open"] / 1e9, r_["closed"] / 1e9, tot / 1e9,
                tot / covered if covered else 0,
                r_["open"] / tot if tot else 0, r_["models"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c == 2:
                cell.fill = (
                    PatternFill("solid", fgColor="DDEBF7") if v.startswith("Hyperscaler")
                    else PatternFill("solid", fgColor="FFF2CC") if v == "Model lab"
                    else PatternFill("solid", fgColor="EDEDED")
                )
        for c in (4, 5, 6):
            ws.cell(row=r, column=c).number_format = "#,##0.0"
        for c in (7, 8):
            ws.cell(row=r, column=c).number_format = "0.0%"
    ws.auto_filter.ref = f"A{hr2}:I{hr2 + len(cloud_rows)}"
    ws.freeze_panes = f"A{hr2 + 1}"
    autosize(ws, [20, 19, 6, 20, 16, 12, 14, 11, 13])

    nrow = hr2 + len(cloud_rows) + 2
    notes = [
        "Methodology: OpenRouter does not publish per-provider token volume. Each model-variant's weekly tokens are "
        "split across its serving endpoints in proportion to each endpoint's live 30-minute request count "
        "(openrouter.ai frontend endpoint stats). Requests are a proxy for tokens - providers with atypical request "
        "sizes will be over/under-stated - and the 30-minute window is extrapolated to the week, so treat rows as "
        "order-of-magnitude estimates.",
        "Categories: 'Model lab' = first-party API of the model's developer (OpenAI, Anthropic, DeepSeek, Moonshot...). "
        "'AI inference cloud' = independent GPU/inference specialists (DeepInfra, Fireworks, Together, Baseten, Novita, "
        "Groq...). 'Hyperscaler' = general-purpose public clouds (AWS Bedrock, Azure, Google Vertex/AI Studio, "
        "Cloudflare, DigitalOcean; Alibaba, Baidu, Tencent, StreamLake/Kuaishou).",
        "Closed models can only be served by their lab or licensed hyperscalers, so their provider mix is structurally "
        "narrow; open-weights models are served competitively by many independent clouds.",
        "Google is counted as a hyperscaler (Vertex AI / AI Studio are cloud services) even though it is also the lab "
        "behind the Gemini models it serves; the same applies to Alibaba/Tencent/Baidu serving their own models.",
        "Scope caveat: this covers only traffic routed through OpenRouter. Direct first-party API usage (e.g. most "
        "OpenAI/Anthropic enterprise traffic) and direct cloud contracts (Bedrock/Vertex/Azure enterprise deals) are "
        "not visible here, so hyperscaler and lab shares of the GLOBAL market are far larger than these rows suggest.",
        "Free-tier variant endpoints sometimes expose no live stats (e.g. Tencent Hy3 free promo); those tokens are "
        "distributed using the model's standard-variant endpoint mix as a proxy.",
        "Source: OpenRouter (openrouter.ai/rankings), as of " + meta["fetched_at"] + ".",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=nrow + i, column=1, value=n).font = Font(size=9, color="595959")

    pie = PieChart()
    pie.title = "Covered tokens by provider category"
    data = Reference(ws, min_col=5, min_row=hr, max_row=hr + len(order))
    cats_ref = Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + len(order))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.height, pie.width = 8, 12
    ws.add_chart(pie, f"K{hr}")


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
                   tu_rows, top_callers, total_calls, cloud_rows, cloud_cats, cloud_meta,
                   ht_rows):
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

    # ---- Sheet 7: Cloud provider split
    write_cloud_sheet(wb, cloud_rows, cloud_cats, cloud_meta, grand_total)

    # ---- Sheet 8: Harnesses & RL environments
    write_harness_rl_sheet(wb)

    # ---- Sheet 9: Quantization & KLD
    write_quant_kld_sheet(wb)

    # ---- Sheet 10: Training dataset sizes
    write_training_data_sheet(wb)

    # ---- Sheet 11: Per-model harnesses & tools
    write_harness_tools_sheet(wb, ht_rows)

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


def write_cloud_csv(cloud_rows, meta):
    with open(os.path.join(OUT_CSV_DIR, "cloud_provider_split.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "provider", "category", "hq",
            "est_weekly_tokens_open", "est_weekly_tokens_closed",
            "est_weekly_tokens_total", "share_of_covered", "open_share", "models_served",
        ])
        for r in cloud_rows:
            w.writerow([
                r["provider"], r["category"], r["hq"],
                round(r["open"]), round(r["closed"]), round(r["total"]),
                round(r["total"] / meta["covered"], 6) if meta["covered"] else 0,
                round(r["open"] / r["total"], 4) if r["total"] else 0,
                r["models"],
            ])


def main():
    rankings, models = load_raw()
    catalog = index_catalog(models)
    rows, excluded = build_rows(rankings, catalog)
    summary, grand_total = summarize(rows)
    devs = by_developer(rows)
    pb_rows = enrich_params_benchmarks(models, rows)
    tu_rows, top_callers, total_calls = enrich_tool_use(catalog, rows)
    cloud_rows, cloud_cats, cloud_meta = build_cloud_split(rows)
    ht_rows = enrich_harness_tools(models, rows)
    write_workbook(rows, summary, devs, grand_total, excluded, pb_rows,
                   tu_rows, top_callers, total_calls, cloud_rows, cloud_cats, cloud_meta,
                   ht_rows)
    write_csvs(rows, summary, devs, grand_total)
    write_params_benchmarks_csv(pb_rows)
    write_tool_use_csv(tu_rows)
    write_cloud_csv(cloud_rows, cloud_meta)
    write_harness_rl_csv()
    write_quant_kld_csv()
    write_training_data_csv()
    write_harness_tools_csv(ht_rows)
    tables_path = write_csv_compilation_workbook()
    print(f"Wrote {tables_path}")

    print(f"\nCloud split (covered {cloud_meta['covered'] / 1e12:.1f}T of {grand_total / 1e12:.1f}T):")
    for cat, d in sorted(cloud_cats.items(), key=lambda kv: -(kv[1]['open'] + kv[1]['closed'])):
        tot = d['open'] + d['closed']
        print(f"  {cat:22s} {tot / 1e12:6.2f}T ({tot / cloud_meta['covered']:5.1%})  "
              f"open={d['open'] / 1e12:5.2f}T closed={d['closed'] / 1e12:5.2f}T  providers={d['providers']}")

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
