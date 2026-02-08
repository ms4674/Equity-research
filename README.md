# Datacenter Infrastructure Model

A comprehensive Excel-based datacenter model covering hyperscalers and neocloud vendors with historical data (2018-2024) and forecasts through 2030.

## Files

- **`Datacenter_Infrastructure_Model.xlsx`** — The main Excel workbook (9 sheets, 24 charts)
- **`build_datacenter_model.py`** — Python script to regenerate the model

## Model Contents

### Sheets

| Sheet | Description |
|-------|-------------|
| **Dashboard** | Executive summary with key metrics, segment splits, and 4 overview charts |
| **Capex Spend** | Capital expenditure ($B) by company with line and stacked bar charts |
| **Revenue** | Annual revenue ($B) by company |
| **Servers** | Server counts (thousands) by company |
| **Datacenters** | Datacenter/facility counts by company |
| **GW Capacity** | IT load capacity in gigawatts |
| **Power & Generation** | Power contracted, utilization rates, and renewable energy mix |
| **Revenue per MW** | Revenue/MW and Capex/MW efficiency analysis |
| **Assumptions & Sources** | Data sources, methodology, definitions, and caveats |

### Companies Covered

**Hyperscalers:** Amazon (AWS), Microsoft (Azure), Google (GCP), Meta, Oracle Cloud, Apple

**Neocloud Vendors:** CoreWeave, Lambda, Crusoe Energy, Voltage Park, Together AI, Applied Digital

### Key Outputs

- Capex spend trends and forecasts
- Server and datacenter buildout trajectories
- GW-scale power capacity and utilization
- Renewable energy adoption curves
- Revenue per MW efficiency metrics ($M/MW)
- Capex per MW investment intensity
- Hyperscaler vs Neocloud segment comparisons

## Regenerating the Model

```bash
pip install openpyxl
python build_datacenter_model.py
```

## Data Notes

- Historical data (2018-2024) sourced from public SEC filings, earnings calls, sustainability reports, and industry analyst estimates
- Forecasts (2025E-2030E) represent base-case estimates
- Forecast columns are highlighted in yellow in the spreadsheet
- All figures in USD
