# Equity-research

## GLM 5.2 vs Leading Frontier Models — Monthly Token Consumption

Aggregated, sourced view of monthly token consumption for **Z.ai's GLM 5.2** versus
the leading frontier models, built from OpenRouter's public usage rankings and model
pages (as of **30 Jun 2026**). The goal is an analyst-ready spreadsheet that puts
GLM 5.2's adoption in context and translates token volumes into implied inference
spend.

### Deliverables

| File | Description |
| --- | --- |
| `output/GLM-5.2_vs_Frontier_Token_Consumption.xlsx` | The workbook (4 sheets, live formulas). |
| `data/openrouter_token_consumption_june2026.csv` | Raw, version-controlled source data. |
| `scripts/build_workbook.py` | Regenerates the workbook from the CSV. |

The workbook contains:

- **Assumptions** — toggle cells (weeks/month, input vs. output token mix) that drive every formula.
- **Monthly_Consumption** — the main comparison table (token volume, growth, pricing, implied spend).
- **GLM5.2_RunRate** — GLM 5.2's weekly trajectory and monthly run-rate.
- **Sources** — full source list, methodology and caveats.

### Headline read

- The OpenRouter trailing-month top 10 is dominated by Chinese open-weight models
  (DeepSeek V4 Flash ~10.9T and Tencent Hy3 Preview ~10.7T lead), with Anthropic's
  Claude Opus 4.7 (~7.48T) and Sonnet 4.6 (~7.45T) the highest closed models.
- **GLM 5.2 launched mid-June** (13 Jun Coding Plan, 16 Jun MIT weights + API), so it
  is absent from the *trailing-month* top 10. On a **weekly run-rate** basis it is
  already **#7** at **2.11T tokens (week ending 28 Jun, +66% WoW)** — which annualizes
  to a **~9.2T monthly-equivalent** run-rate, placing it among the top 3–5 models.
- GLM 5.2's realized blended pricing (~$0.45 / $3.31 per Mtok) sits well below
  closed frontier coding models, supporting the "open, Opus-style coding" thesis.

### Methodology (summary)

- `Monthly run-rate (T) = latest weekly tokens × weeks-per-month (4.345)`.
- `Blended $/Mtok = input_share × input$ + output_share × output$` (default 50/50, editable).
- `Implied monthly spend ($M) = tokens (trillions) × blended $/Mtok` (since 1T tokens × $1/Mtok = $1M).
- Prompt-cache **realized** pricing is used where OpenRouter publishes it; otherwise provider **list** pricing.
- OpenRouter measures API/developer usage routed through that platform only — it excludes
  first-party traffic (ChatGPT, Gemini, Claude.ai), so absolute totals understate ecosystem usage.
  Implied spend estimates inference dollars on OpenRouter, **not** lab revenue.

See the **Sources** sheet for per-figure citations and caveats.

### Regenerate

```bash
pip install -r requirements.txt
python scripts/build_workbook.py
```
