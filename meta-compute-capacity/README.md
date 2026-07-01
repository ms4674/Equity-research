# Meta — Total Compute Capacity: Owned vs. Leased

Aggregation of Meta Platforms' data-center (compute) capacity, split between
company-**owned** facilities and **leased / colocation** capacity.

## Deliverable
- `Meta_Compute_Capacity_Owned_vs_Leased.xlsx` — the spreadsheet (7 tabs).
- `build_workbook.py` — reproducible generator (`python3 build_workbook.py`, needs `openpyxl`).

## Why power / electricity is the capacity metric
Data-center capacity is measured in electric power (MW / GW): power is the binding
constraint on how much compute (servers / GPUs) a site can host. Meta does not
publish a single "owned vs leased MW" number, but its **Environmental Data Index**
reports actual electricity consumption (MWh/yr) for **every owned online site
individually** and a single line for **"Leased data center facilities."** That is
the cleanest apples-to-apples owned-vs-leased split of realized compute and
anchors the model.

## Headline (FY2024, by data-center electricity consumed)
| Category | MWh/yr | Share |
|---|---:|---:|
| Owned online data centers | 14,890,259 | ~82% |
| Leased / colocation facilities | 3,069,504 | ~17% |
| Other DC-related (warehouse/network/small colo) | 102,016 | ~1% |
| **Total data-center electricity** | **18,061,781** | 100% |

Leased share of data-center power rose from ~11% (2020) to ~17% (2024) as Meta
pre-leased heavily for AI. Forward AI build-out (targets, not installed): trackers
count **~15.8 GW** across ~20 sites (~50% operational); flagships **Prometheus**
(Ohio, ~1 GW) and **Hyperion** (Louisiana, 2 GW scaling to 5 GW; financed via a
~$27–30B Blue Owl JV with Meta ~20% equity, leased back to Meta).

## Tabs
1. **Summary - Owned vs Leased** — authoritative energy-based split + trend + charts.
2. **Owned - Site Energy** — per-site MWh, 2020–2024.
3. **Owned - Campus Detail** — buildings, area, investment, nameplate-MW estimates.
4. **Leased & Colocation** — leased energy trend + known markets + relationships.
5. **AI Buildout (GW)** — forward capacity, owned vs JV/leased.
6. **Sources & Methodology** — all sources with URLs and caveats.

## Caveats
- Electricity consumption ≠ nameplate MW, but is the best public realized-capacity proxy (Meta-reported, third-party assured).
- Nameplate MW in Campus Detail are third-party estimates (usdatamap), incomplete, and predate the AI expansion.
- AI GW figures are announcements/targets with multi-year phasing, not installed capacity.

See the workbook's **Sources & Methodology** tab for full source list.
