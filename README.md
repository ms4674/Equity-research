# Equity Research — Hyperscale CapEx Analysis

Analysis of capital expenditure patterns across the five largest hyperscale cloud/AI infrastructure companies (Amazon, Microsoft, Alphabet/Google, Meta, Oracle), covering drivers, category allocation, geographic distribution, and the tie-out between capex and contracted backlog (RPO).

## Contents

| File | Description |
|------|-------------|
| `hyperscale_capex_analysis.xlsx` | **Single aggregated workbook** (8 tabs — see below) |
| `hyperscale_capex_analysis.ipynb` | Interactive notebook with 8 visualisations and commentary |
| `build_spreadsheet.py` | Script to regenerate the Excel workbook |
| `data/hyperscale_capex_annual.csv` | Annual capex by company (2018–2026E) |
| `data/capex_category_allocation.csv` | Breakdown by category (servers, power, construction, land, etc.) |
| `data/capex_geographic_allocation.csv` | Regional allocation (North America, Europe, APAC, MEA, LatAm) |
| `data/capex_drivers.csv` | Key drivers ranked by impact level |

### Spreadsheet Tabs

| Tab | Content |
|-----|---------|
| Annual CapEx | Company-level capex 2018–2026E with stacked bar chart |
| Backlog (RPO) | RPO by company, historical trend, and stacked bar chart |
| CapEx vs Backlog | Core tie-out: RPO/CapEx coverage ratios, backlog vs capex growth signals, interpretation guide |
| Category Allocation | Servers/GPUs, power, construction, land, cooling, networking breakdown with pie chart |
| Geographic Allocation | Regional split (NA, Europe, APAC, MEA, LatAm) for 2024 vs 2026E with bar chart |
| CapEx Drivers | 10 key drivers ranked by impact level with detailed descriptions |
| Power & Constraints | Infrastructure bottleneck metrics (PJM pricing, transformer lead times, DC construction) |
| Summary Dashboard | Executive KPIs, investment implications by sector, and key risks |

## Key Findings

- Combined hyperscaler capex is projected to reach **~$700B in 2026**, up from $66B in 2018 (10x in 8 years)
- **AI infrastructure** accounts for ~75% of total capex
- **Servers & GPUs** are the largest single category (~35%), followed by **facilities** (power + construction + cooling + land, ~38%)
- **North America** captures ~48% of capex, with **APAC gaining share** (22%) driven by emerging markets
- **Power availability** — not GPU supply — is the binding constraint on deployment speed
- Capital intensity has reached **30–70% of revenue**, resembling utility companies rather than traditional tech
- **Combined backlog (RPO) of ~$1.6T** provides 2.3x coverage of annual capex — spending is broadly demand-justified
- **Microsoft** (4.3x) and **Oracle** (10.5x) have the strongest backlog coverage; **Amazon** (1.2x) is building more speculatively

## Setup

```bash
pip install pandas matplotlib seaborn jupyter numpy openpyxl xlsxwriter
jupyter notebook hyperscale_capex_analysis.ipynb
python3 build_spreadsheet.py  # regenerate the Excel workbook
```
