# Open-weights vs closed LLMs: token pricing and volume

Aggregates per-model API token pricing and real-world token volume for open-weights
and closed (proprietary, API-only) LLMs.

**Main deliverable:** [`llm_pricing_vs_token_volume.xlsx`](llm_pricing_vs_token_volume.xlsx)

## Data source

All data comes from OpenRouter, the largest multi-provider LLM gateway and the only
first-party public source of per-model token volume:

- **Token volume:** `https://openrouter.ai/api/frontend/v1/rankings/models?period=week`
  — per-model weekly token totals (prompt + completion), trailing 7 days ending
  **2026-07-16**. Matches the public rankings at openrouter.ai/rankings.
- **Pricing and metadata:** `https://openrouter.ai/api/v1/models` — list price per
  token (input/output), context length, and `hugging_face_id`.

Raw snapshots are stored in [`raw/`](raw/). Citation required by OpenRouter's data
terms: *Source: OpenRouter (openrouter.ai/rankings), as of 2026-07-16.*

## Headline numbers (week ending 2026-07-16)

| Class | Models | Weekly tokens | Volume share | Vol-weighted blended price* | Est. weekly spend |
|---|---|---|---|---|---|
| Open-weights | 143 | 44.2T | 74.9% | $0.50 / 1M | ~$9.7M |
| Closed / proprietary | 161 | 14.8T | 25.1% | $5.99 / 1M | ~$47.0M |

Open-weights models carry ~3x the token volume of closed models on OpenRouter, at
roughly 1/12th the volume-weighted price — while closed models still capture ~83%
of estimated dollar spend.

\* Blended = (3 × input + 1 × output) / 4, weighted by each model's weekly tokens.

## Workbook contents

| Sheet | Contents |
|---|---|
| Open vs Closed Summary | Aggregate token volume, volume-weighted and median prices, estimated spend per class, volume-share pie chart |
| Model Comparison | All 304 matched text LLMs: input/output/blended $ per 1M tokens, context window, weekly token volume (prompt/completion split), free-tier share, requests, estimated spend, class color-coding; filterable |
| Top 20 Volume Chart | Bar chart of the 20 highest-volume models |
| By Developer | Volume, model count, and estimated spend rolled up per developer (Anthropic, DeepSeek, Google, OpenAI, ...) |
| Params & Benchmarks | Deep-dive on 9 selected models (NVIDIA Nemotron 3 Ultra/Super, Claude Fable 5, GPT-5.6 Sol, Kimi K3, GLM-5.2, Qwen3.7 Max, DeepSeek V4 Pro/Flash): total/active parameters, architecture, AA Intelligence Index and cost per task, GPQA Diamond, SWE-bench Verified/Pro, Terminal-Bench, HLE, pricing, and weekly volume |
| Tool Use | Tool-calling comparison of Kimi K3 vs closed frontier (Claude Fable 5, GPT-5.6 Sol, Opus 4.8) and open-weights peers (GLM-5.2, DeepSeek V4 Pro/Flash, MiniMax M3, Kimi K2.6, Nemotron 3 Ultra): MCP Atlas, Toolathlon-Verified, AutomationBench, BrowseComp, GDPval-AA Elo, tau2-bench, parallel tool-call API support, plus OpenRouter's real-world weekly tool-call leaderboard |

CSV versions of every table are in [`csv/`](csv/).

## Methodology

- **Open vs closed classification:** a model is open-weights if its catalog entry has
  a `hugging_face_id` (weights publicly downloadable) or its description explicitly
  states an open-weight / MIT / Apache-2.0 release. Two manual overrides
  (Kimi K3, MiniMax M1) are documented in `scripts/build_spreadsheet.py`.
- **Prices** are OpenRouter list prices per 1M tokens for the standard (paid) variant.
  Provider-direct prices can differ slightly; caching and batch discounts are ignored.
- **Estimated spend** applies list prices to paid-variant traffic only (free-tier
  tokens are excluded), assuming the paid traffic has the same prompt/completion
  split as total traffic. Treat it as an upper-bound order-of-magnitude estimate.
- **Exclusions:** 418B tokens (~0.7% of raw total) from embedding, image, TTS, and
  transcription models that are not text LLMs.
- **Caveats:** OpenRouter routes a large share of open-model and Chinese-lab traffic
  but only a fraction of first-party OpenAI/Anthropic/Google API traffic, so closed
  models' global volume share is understated relative to what these numbers show.
  Token counts use each upstream provider's own tokenizer, so cross-provider token
  comparisons are approximate.
- **Tool Use sheet:** the K3 / Fable 5 / Sol / Opus 4.8 / GLM-5.2 rows come from
  Moonshot's Kimi K3 launch evaluation table (kimi.com/blog/kimi-k3), the only
  published single-methodology table covering all five; remaining rows are
  vendor-reported under their own harnesses. Real-world tool-call counts come from
  OpenRouter's public "Tool Calls" chart (snapshot in `raw/`), which only lists the
  top 9 models per week plus an aggregated "Others" row.
- **Params & Benchmarks sheet:** parameter counts are official vendor disclosures for
  open-weights models; Anthropic, OpenAI, and Alibaba do not disclose parameter counts
  (~3T figures for Fable 5 / GPT-5.6 Sol are third-party estimates, and Kimi K3's ~50B
  active is a community estimate from the disclosed 16-of-896 expert activation).
  Benchmark scores are vendor-reported at highest reasoning effort, researched
  2026-07-20; labs use different agent harnesses and Terminal-Bench versions, so
  cross-model comparisons are approximate. Sources are cited per row in the sheet.

## Reproducing

```bash
pip install openpyxl
python3 scripts/build_spreadsheet.py
```

The script reads the newest snapshots in `data/raw/` and regenerates the XLSX and CSVs.
To refresh the data, re-download the two endpoints above into `data/raw/` with the
current date in the filename.
