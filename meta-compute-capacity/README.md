# Meta — Total Compute Capacity: Owned vs. Colocation vs. Neocloud

Aggregation of Meta Platforms' data-center (compute) capacity across the three
ways Meta sources compute:
- **Owned** — Meta builds and owns the data center and the compute.
- **Leased colocation** — Meta leases space & power (powered shells / wholesale
  colo) and installs its **own** servers/GPUs (measured in MW / MWh; appears in
  Meta's leased-facility electricity).
- **Neocloud / cloud** — Meta rents GPU **compute-as-a-service**; a third party
  owns the GPUs and the building (measured in $ multi-year commitments; opex, not
  in Meta's owned-electricity reporting).

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

## Headline #1 — physical footprint (FY2024, by data-center electricity consumed)
| Category | MWh/yr | Share |
|---|---:|---:|
| Owned online data centers | 14,890,259 | ~82% |
| Leased colocation facilities (Meta owns the compute) | 3,069,504 | ~17% |
| Other DC-related (warehouse/network/small colo) | 102,016 | ~1% |
| **Total data-center electricity** | **18,061,781** | 100% |

Leased colocation share of data-center power rose from ~11% (2020) to ~17% (2024).

## Headline #2 — neocloud / cloud commitments ($, separate from the split above)
| Counterparty | Value | Term | Avg $B/yr (straight-line) |
|---|---:|---|---:|
| CoreWeave — initial | $14.2B | Sep 2025 – Dec 2031 | ~$2.3B |
| CoreWeave — expansion | $21B | 2027 – 2032 | ~$3.5B |
| Nebius — Order #1 | $2.9B | 2026 – 2030 (5-yr) | ~$0.6B |
| Nebius — Vera Rubin dedicated | $12B | 2027 – 2031 (5-yr) | ~$2.4B |
| Nebius — unsold-capacity backstop | up to $15B | ~2027 – 2031 | up to ~$3.0B (contingent) |
| Google Cloud | $10B+ | 6 years | ~$1.7B |
| AWS, Microsoft Azure | undisclosed | ongoing | n/d |

**~$60B firm disclosed** (up to ~$75B if the Nebius $15B backstop triggers).
Neocloud is compute rented as a service — a **different unit** ($) than the
electricity split, so the two are not additive.

### How much per year on CoreWeave + Nebius?
Straight-line (contract total ÷ term), the combined **CoreWeave + Nebius** run-rate
ramps to roughly **$8–9B/yr firm in 2028–2031** (up to ~$12B/yr if the Nebius
backstop triggers). Actual cash spend is back-loaded — small at each contract's
start and rising as GPU capacity is deployed. See the per-year matrix on the
**Neocloud & Cloud** tab.

Forward AI build-out (targets, not installed): ~**15.8 GW** across ~20 sites (~50%
operational); flagships **Prometheus** (Ohio, ~1 GW) and **Hyperion** (Louisiana,
2 GW → 5 GW; ~$27–30B Blue Owl JV, Meta ~20% equity, leased back to Meta).

## Tabs
1. **Summary** — owned vs colocation vs neocloud (energy split + neocloud $ + charts).
2. **Owned - Site Energy** — per-site MWh, 2020–2024.
3. **Owned - Campus Detail** — buildings, area, investment, nameplate-MW estimates.
4. **Colocation (space & power)** — leased-facility energy trend + markets + relationships.
5. **Neocloud & Cloud** — GPU-as-a-service commitments (CoreWeave, Nebius, Google, AWS, Azure).
6. **Utilization & Rental Rates** — colocation ($/kW/mo by tier & metro, vacancy/occupancy) and neocloud ($/GPU-hr by SKU/provider, utilization economics).
7. **AI Buildout (GW)** — forward capacity, owned vs JV/leased.
8. **Sources & Methodology** — all sources with URLs and caveats.

## Caveats
- Electricity consumption ≠ nameplate MW, but is the best public realized-capacity proxy (Meta-reported, third-party assured).
- Colocation (MWh) and neocloud ($) are different units and not additive.
- Neocloud $ are multi-year contract values, not annual spend or installed capacity.
- Rental rates & utilization (tab 6) are market benchmarks (CBRE/JLL/CREFC/neocloud disclosures), **not** Meta's confidential contract terms.
- Nameplate MW in Campus Detail are third-party estimates (usdatamap), incomplete, and predate the AI expansion.
- AI GW figures are announcements/targets with multi-year phasing, not installed capacity.

See the workbook's **Sources & Methodology** tab for full source list.
