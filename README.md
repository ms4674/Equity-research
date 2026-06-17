# Equity-research

## AI Infrastructure Financing — Hyperscalers & Neoclouds (2025–2030)

`AI_Infrastructure_Financing_2025-2030.xlsx` is an Excel model of the **equity and
debt financing** raised by the major **hyperscalers** and **neocloud / AI-native
infrastructure** companies to build out AI infrastructure, cumulative **2025–2030**.
Total debt is broken down by instrument type (loans, alternate private credit,
insurance-linked lending, bonds, securitization/ABS, convertibles, and
vendor/equipment-lease financing), with an off-balance-sheet SPV/JV memo line.

### Workbook contents
| Sheet | Description |
|-------|-------------|
| Cover & Methodology | Scope, definitions, debt-bucket taxonomy, caveats |
| Summary (by Company) | Equity / internal funding vs. total debt, % debt-funded, chart |
| Debt by Type | Main table — total debt split into mutually-exclusive instrument buckets |
| Capex Projections | Annual AI-infrastructure capex 2025–2030 by company |
| Key Deals & Sources | Notable disclosed transactions with citations |

### Companies covered
- **Hyperscalers:** Microsoft, Amazon (AWS), Alphabet (Google), Meta, Oracle (OCI)
- **Neoclouds / AI-native:** CoreWeave, SpaceX/xAI (Colossus), Nebius, IREN, Crusoe, Lambda,
  Fluidstack, Nscale, and an "Other" aggregate.
  - SpaceX acquired xAI (effective 2 Feb 2026), so Colossus I & II are now SpaceX-owned; the two
    are shown as one consolidated **SpaceX / xAI** line to avoid double-counting the same clusters.

### Headline figures (research estimates, US$bn)
- Total financing ≈ **$5,374bn** (≈ Goldman Sachs' ~$5.3tn 2025–2030 AI + data-center capex)
- Equity / internal funding ≈ **$3,531bn**; total debt ≈ **$1,843bn** (≈ 34% debt-funded)
- Debt mix: corporate bonds $803bn · private credit $356bn · bank/syndicated loans $249bn ·
  insurance-linked lending $174bn · vendor/lease $142bn · securitization/ABS $95bn ·
  convertibles $24bn

> **Note:** Figures are research estimates that blend disclosed transactions (2023–early 2026)
> with forward projections. They are illustrative, not audited, and not investment advice.
> See the Cover & Methodology and Key Deals & Sources sheets for definitions and citations.

### Rebuild
```bash
pip install -r requirements.txt
python scripts/build_financing_model.py
```
