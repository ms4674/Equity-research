# Datacenter Infrastructure Model

A comprehensive Excel-based datacenter model covering hyperscalers, neocloud vendors, and colocation/REIT players with historical data (2018-2024) and forecasts through 2030.

## Files

- **`Datacenter_Infrastructure_Model.xlsx`** — The main Excel workbook (12 sheets, 43 charts)
- **`build_datacenter_model.py`** — Python script to regenerate the model

## Model Contents

### Sheets

| Sheet | Description |
|-------|-------------|
| **Dashboard** | Executive summary with key metrics, 3-segment splits, supply chain TAM, and 4 overview charts |
| **Capex Spend** | Capital expenditure ($B) by company across all 3 segments |
| **Revenue** | Annual revenue ($B) by company |
| **Servers** | Server counts (thousands) by company |
| **Datacenters** | Datacenter/facility counts by company |
| **GW Capacity** | IT load capacity in gigawatts |
| **Power & Generation** | Power contracted, utilization rates, and renewable energy mix |
| **Revenue per MW** | Revenue/MW and Capex/MW efficiency analysis |
| **Power Supply & Ramp** | Power demand vs supply, generation additions by source (gas, solar, wind, nuclear, battery, SMR), grid constraints, interconnection queues, build timelines, and demand drivers |
| **DC Bill of Materials** | Per-MW cost breakdown across 26 line items (land, building, power, cooling, networking, IT, security, services), 100MW reference facility costs, vendor mapping, and cost trend analysis |
| **Capex to Revenue** | How DC capex translates to supply chain revenue: TAM by segment ($B), capex flow allocation, YoY growth, revenue multipliers, and key vendor beneficiaries |
| **Assumptions & Sources** | Data sources, methodology, definitions, power/BOM/supply chain assumptions, and caveats |

### Companies Covered

**Hyperscalers (6):** Amazon (AWS), Microsoft (Azure), Google (GCP), Meta, Oracle Cloud, Apple

**Neocloud Vendors (6):** CoreWeave, Lambda, Crusoe Energy, Voltage Park, Together AI, Applied Digital

**Colocation / DC REITs (6):** Equinix, Digital Realty, CyrusOne, QTS Realty, Vantage Data Centers, Switch

### Key Outputs

- Capex spend trends and forecasts across all 3 segments
- Server and datacenter buildout trajectories
- GW-scale power capacity and utilization
- Renewable energy adoption curves
- Revenue per MW efficiency metrics ($M/MW)
- Capex per MW investment intensity
- Power supply vs demand gap analysis with generation source breakdown
- DC build timeline (land-to-live) tracking
- Power demand drivers (AI training, inference, cloud IaaS, edge, crypto/HPC)
- Bill of materials cost per MW with vendor mapping
- Supply chain TAM by segment (Servers/GPUs, Networking, Power, Cooling, Construction, Real Estate, Software, Fiber)
- Capex-to-revenue flow showing how each $1 of DC capex is allocated
- Key vendor beneficiaries per supply chain segment

## Regenerating the Model

```bash
pip install openpyxl
python build_datacenter_model.py
```

## Data Notes

- Historical data (2018-2024) sourced from public SEC filings, earnings calls, sustainability reports, REIT supplements, and industry analyst estimates
- Forecasts (2025E-2030E) represent base-case estimates
- Forecast columns are highlighted in yellow in the spreadsheet
- CyrusOne (KKR/GIP) and QTS (Blackstone) are now private; post-acquisition data are estimates
- All figures in USD
