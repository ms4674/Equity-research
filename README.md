# Equity-research

## Models

### Anthropic – three-statement model (`models/Anthropic_3_Statement_Model.xlsx`)

Integrated income statement, balance sheet and cash flow for Anthropic, PBC (FY2023A–FY2030E),
built off the Form S-1 (filed 26 Oct 2026) and subsequent investor disclosures. The core analysis is
revenue versus the compute bill, split between **inference** (customer-serving, free-tier, internal)
and **training & research** compute, with a take-or-pay floor from the contracted compute minimums.

- Scenario switch on `Assumptions!C4` (1 = Base, 2 = Bull, 3 = Bear). Blue cells are inputs; every
  forecast cell is a live formula.
- Period layout (identical on every sheet): FY2023A | FY2024A | FY2025A | H1'26A | H2'26E | FY2026E | FY2027E–FY2030E.
- Sheets: Cover, Summary (charts + valuation snapshot), Assumptions, Revenue Build, Compute Build,
  Opex Build, Income Statement, Balance Sheet, Cash Flow, Compute Commitments, Historical Data, Checks.

Regenerate the workbook with `python3 models/build_anthropic_model.py` (requires `openpyxl`).
