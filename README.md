# Equity-research

## Agentic AI Time Series Databook

`data/Agentic_AI_Time_Series.xlsx` contains time-series data on agentic AI:

- **Token Consumption** — monthly platform token volumes (Google, Microsoft Foundry, OpenRouter), Jan 2024 – Dec 2026, plus an estimated agentic share and agentic token consumption.
- **Agents – Number & Type** — annual deployed-agent counts by type (in-application, custom-configured, standalone, bespoke), 2025–2030, calibrated to IDC's Agent Economics model.
- **Token Loads** — tokens-per-task benchmarks by workload archetype, agentic-vs-chat multipliers, and derived annual agentic token loads.
- **Actions per Agent** — turns / model calls per task by archetype and derived annual daily-action volumes (IDC anchor: 217B daily actions by 2029).
- **Sources & Notes** — every disclosed anchor with source and URL, plus methodology caveats.

Disclosed figures are labeled `Disclosed`; interpolated/extrapolated values are labeled `Modeled`.

Regenerate the workbook with:

```bash
pip install openpyxl
python3 scripts/build_agentic_ai_time_series.py
```
