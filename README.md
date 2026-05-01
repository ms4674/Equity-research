# Equity Research

## Storage Market Supply-Demand-Capacity Analysis

This repository contains time-series data on supply, demand, and capacity for the four major storage/memory markets:

### Spreadsheet: `storage_supply_demand_capacity.xlsx`

Contains 5 tabs:

| Tab | Contents |
|-----|----------|
| **DRAM** | Annual market data (revenue, bit supply/demand growth, wafer starts, CapEx), quarterly revenue, vendor market share (Q4 2025) |
| **NAND** | Annual market data, capacity structural shift (Korea/USA bloc wafer starts), demand breakdown by application |
| **SSD** | Annual shipments/revenue/capacity, quarterly data, enterprise SSD forecast (McKinsey), NVIDIA Vera Rubin demand impact |
| **HDD** | Annual data, quarterly exabyte shipments by vendor (Seagate/WDC/Toshiba), supply-demand dynamics, vendor technology comparison |
| **Summary** | Cross-market revenue comparison, supply-demand balance status (2020-2026E), key investment themes |

### Key Findings (2025-2026)

- All four storage markets are in supply-constrained conditions simultaneously
- AI infrastructure (HBM, enterprise SSD, nearline HDD) is the dominant demand driver
- New DRAM/NAND fab capacity not expected until 2027-2028
- DRAM+NAND combined revenue exceeded $200B for the first time in 2025
- HDD exabytes growing >20% annually despite declining units
- NVIDIA Vera Rubin could add ~9.3% incremental NAND demand in 2027

### Data Sources

TrendForce, IDC, Trendfocus, Omdia, Isaiah Research, Atlas Peak Research, McKinsey, Citigroup Securities, Statista, DRAMeXchange, MemoryMarket, StorageNewsletter, company earnings/filings.

### Generator Script

`generate_spreadsheet.py` - Python script using openpyxl to regenerate the spreadsheet with formatting.
