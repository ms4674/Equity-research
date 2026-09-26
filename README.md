# Equity-research

## Token pricing for frontier AI labs — global time series

A time series of public API token pricing (USD per 1M tokens) for the major frontier AI labs
worldwide, from GPT-3's 2020 launch through September 2026.

**Labs covered:** OpenAI, Anthropic, Google, xAI, Meta (US) · Mistral (Europe) · DeepSeek, Alibaba/Qwen (China)

### Files

| File | Contents |
| --- | --- |
| `token_pricing_frontier_labs.xlsx` | Aggregated spreadsheet: Read Me, Price Events (137 events, long format), monthly Flagship Input/Output matrices with log-scale charts, Lab Summary |
| `data/token_pricing_events.csv` | Source of truth: one row per published price event (launch, cut, raise) with input/output/cached rates, notes, and source |
| `data/flagship_timeseries_monthly.csv` | Generated wide-format matrix: each lab's frontier-flagship input & output price, sampled monthly Jan 2023 – Sep 2026 |
| `scripts/build_spreadsheet.py` | Regenerates the XLSX and monthly CSV from the events CSV (`pip install openpyxl`, then `python3 scripts/build_spreadsheet.py`) |

### Methodology (short version)

- Prices are public synchronous-API list rates on the provider's own platform; batch tiers,
  enterprise discounts, and free tiers are out of scope.
- "Flagship" = the model each lab positions as its frontier offering at that date; base /
  short-context tier recorded where pricing is tiered.
- Meta rows before its first-party Meta Model API (July 2026) use Together AI hosted rates as a
  reference proxy for open-weights Llama models.
- DeepSeek's post-Aug-2026 peak/off-peak split is recorded at the off-peak leg; Qwen prices use
  Alibaba's International (Singapore) USD list.
- Compiled 2026-09-26 from provider announcements and pricing pages, cross-checked against
  independent price trackers and archived pricing snapshots. Full caveats are on the workbook's
  Read Me sheet.
