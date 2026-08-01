# Equity-research

## 1999 vs 2008 bubble comparison

`Tech_Bubble_Comparison_1999_vs_2008.xlsx` compares the dot-com bubble
(peak March 2000) and the credit bubble / Global Financial Crisis
(bust 2008-09) across three lenses:

- **Capex** — semiconductor industry capital spending (Gartner/Dataquest),
  fab equipment billings (SEMI WWSEMS), US telecom carrier capex (FRBSF)
  and US business fixed investment (BEA/FRED).
- **Semiconductor demand** — worldwide semiconductor sales 1995-2011
  (WSTS/SIA), DRAM industry revenue, PC and mobile-phone unit shipments.
- **Component pricing** — DRAM/NAND price timelines for both busts
  (128Mb SDRAM in 2000-01; DDR2 1Gb and NAND in 2007-09), sourced from
  Dataquest and DRAMeXchange/TrendForce.

The workbook contains seven sheets (Overview, Bubble Comparison,
Semiconductor Demand, Capex, Component Pricing, Market Context, and
Sources & Notes with full citations and caveats). Growth rates and
peak-to-trough declines are live Excel formulas.

To regenerate the workbook:

```bash
pip install openpyxl
python3 build_bubble_comparison_workbook.py
```
