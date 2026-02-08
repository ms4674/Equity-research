# Snap Inc. (SNAP) – Three-Statement Financial Model

## Overview

A comprehensive three-statement Excel financial model for **Snap Inc. (SNAP US Equity)** with historical data and forward projections.

## Files

| File | Description |
|------|-------------|
| `Snap_Inc_Three_Statement_Model.xlsx` | Complete Excel workbook (9 tabs) |
| `snap_financial_model.py` | Python script that generates the model |

## Model Structure (9 Tabs)

### 1. Revenue Build
- DAU (Daily Active Users) growth assumptions with historical trend
- ARPU (Average Revenue Per User) by year
- Revenue = DAU × ARPU decomposition
- Geographic segmentation: North America, Europe, Rest of World
- Implied ad metrics: impressions/DAU/day, eCPM

### 2. Income Statement
- Full GAAP P&L: FY2019–FY2024 (historical) + FY2025E–FY2028E (forecast)
- Gross margin by year with cost-of-revenue detail
- Operating expenses (R&D, S&M, G&A) as % of revenue
- Stock-based compensation tracked separately
- EBITDA and EBITDA margin
- EPS (diluted)

### 3. Balance Sheet
- Full balance sheet: assets, liabilities, stockholders' equity
- Balance-check row confirms A = L + E
- Working capital items driven by DSO/DPO ratios

### 4. Cash Flow & FCF Bridge
- Full cash flow statement (CFO / CFI / CFF)
- **GAAP Operating Loss → FCF Bridge**: reconciles operating income through D&A, SBC, working capital, interest, taxes, and CapEx to arrive at free cash flow
- FCF margin tracked over time

### 5. Quarterly Detail
- 16 quarters: Q1-2023 through Q4-2026E
- Revenue, COGS, gross profit, OpEx breakdown, EBITDA, net income
- Seasonality-adjusted forecasts (Q4 heaviest)

### 6. Share Count & Dilution
- Annual diluted shares: FY2019–FY2028E
- **8-quarter forward dilution schedule** (Q1-25E through Q4-26E)
- RSU vesting, new grants, buybacks, option exercises
- Unvested RSU pool and vested unexercised options

### 7. Sensitivity Analysis
- **Table 1**: DAU Growth × ARPU Growth → FY2026E Revenue
- **Table 2**: Revenue Growth × EBITDA Margin → FY2026E EBITDA
- **Table 3**: WACC × Terminal Growth → Implied Share Price (DCF)
- Base cases highlighted in green

### 8. Valuation Cross-Checks
- **DCF**: Full discounted cash flow with terminal value
- **Cohort NPV**: User-level LTV/CAC analysis with 8-year cohort cash flows
- **Adoption S-Curve**: DAU penetration vs. TAM with phase identification
- **Unit Economics → EV**: EV/DAU, EV/Revenue, FCF yield benchmarked vs. peers
- Valuation summary table with upside/downside vs. current price

### 9. Assumptions & Drivers
- All key model inputs in one place (yellow-highlighted = editable)
- DAU growth, ARPU growth, margin assumptions, CapEx, tax rate, dilution
- Quarterly seasonality weights

## Key Design Principles

- **Yellow cells** = adjustable assumptions/inputs
- **Blue italic** = forecast values
- **Historical** = FY2019–FY2024 (6 years)
- **Forecast** = FY2025E–FY2028E (4 years annual), 8 quarters forward
- All amounts in $M unless otherwise noted

## How to Regenerate

```bash
pip install openpyxl
python3 snap_financial_model.py
```
