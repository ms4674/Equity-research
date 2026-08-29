# Blended Token Pricing — Frontier & Open-Weight Models

A global time-series dataset of LLM API token prices (2020–2026) with blended
pricing at the standard 3:1 input:output weighting, aggregated into an Excel
workbook.

## Deliverable

**[`Blended_Token_Pricing_Frontier_and_Open_Weight_Models.xlsx`](Blended_Token_Pricing_Frontier_and_Open_Weight_Models.xlsx)**

| Sheet | Contents |
| --- | --- |
| README | Definitions, methodology, caveats, headline findings |
| Price History | Long-format time series: 157 price events (launches, cuts, increases) across 142 models |
| Current Snapshot | Latest list price per model as of 2026-08-29, ranked by blended price |
| Quarterly Series | Carry-forward blended price of 13 vendor lineages, 2023 Q1 – 2026 Q3 |
| Charts | Log-scale blended price trajectories: frontier flagships, open-weight flagships, budget tier |
| Sources | Pricing pages, announcements, and trackers used |

## Definitions

- **Blended price** = (3 × input + 1 × output) / 4 per 1M tokens — the 3:1
  usage weighting used by Artificial Analysis, ModelPriceWatch, and AIMultiple.
- **Frontier (proprietary)**: closed API models (OpenAI, Anthropic, Google,
  xAI, Meta Muse Spark, Amazon, Cohere, Mistral commercial, Alibaba Max line).
- **Open-weight**: downloadable-weights models (Llama, DeepSeek, Qwen, Kimi,
  GLM, MiniMax, Mistral open line, gpt-oss), priced at the first-party API
  rate where one exists, otherwise a reference hosted rate.

## Files

- `data/token_price_history.py` — source-of-truth dataset (one record per price event)
- `data/token_price_history.csv` — flat CSV export with computed blended prices
- `scripts/build_workbook.py` — regenerates the workbook and CSV

## Regenerate

```bash
pip install openpyxl
python3 scripts/build_workbook.py
```
