# Equity Research — Hyperscale CapEx Analysis

Analysis of capital expenditure patterns across the five largest hyperscale cloud/AI infrastructure companies (Amazon, Microsoft, Alphabet/Google, Meta, Oracle), covering drivers, category allocation, and geographic distribution.

## Contents

| File | Description |
|------|-------------|
| `hyperscale_capex_analysis.ipynb` | Main analysis notebook with 7 visualisations and commentary |
| `data/hyperscale_capex_annual.csv` | Annual capex by company (2018–2026E) |
| `data/capex_category_allocation.csv` | Breakdown by category (servers, power, construction, land, etc.) |
| `data/capex_geographic_allocation.csv` | Regional allocation (North America, Europe, APAC, MEA, LatAm) |
| `data/capex_drivers.csv` | Key drivers ranked by impact level |

## Key Findings

- Combined hyperscaler capex is projected to reach **~$700B in 2026**, up from $66B in 2018 (10x in 8 years)
- **AI infrastructure** accounts for ~75% of total capex
- **Servers & GPUs** are the largest single category (~35%), followed by **facilities** (power + construction + cooling + land, ~38%)
- **North America** captures ~48% of capex, with **APAC gaining share** (22%) driven by emerging markets
- **Power availability** — not GPU supply — is the binding constraint on deployment speed
- Capital intensity has reached **30–70% of revenue**, resembling utility companies rather than traditional tech

## Setup

```bash
pip install pandas matplotlib seaborn jupyter numpy
jupyter notebook hyperscale_capex_analysis.ipynb
```
