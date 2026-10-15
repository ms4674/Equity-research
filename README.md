# Equity-research

## AI model labs: gross vs. net ARR

`ai-model-arr/AI_Model_ARR_Gross_vs_Net.xlsx` compares annualized revenue run rates for frontier
(closed-weight) and open-weight model labs on both a gross basis (end-customer billings including
cloud-marketplace / partner channel sales) and a net basis (revenue retained after partner revenue share).

Sheets: `Summary`, `Frontier`, `Open_Weight`, `Assumptions`, `Anthropic_Sensitivity`, `Methodology`, `Sources`.
Blue cells are editable inputs; yellow cells are estimates. Regenerate with:

```bash
pip install openpyxl
cd ai-model-arr && python build_arr_workbook.py
```
