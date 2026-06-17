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

### Companies / segments covered
- **Hyperscalers:** Microsoft, Amazon (AWS), Alphabet (Google), Meta, Oracle (OCI)
- **Neoclouds / AI-native:** CoreWeave, SpaceX/xAI (Colossus), Nebius, IREN, Crusoe, Lambda,
  Fluidstack, Nscale, and an "Other" aggregate.
  - SpaceX acquired xAI (effective 2 Feb 2026), so Colossus I & II are now SpaceX-owned; the two
    are shown as one consolidated **SpaceX / xAI** line to avoid double-counting the same clusters.
- **Colocation / data-center REITs:** Equinix, Digital Realty, QTS (Blackstone), Vantage, Aligned
  (MGX/BlackRock), CyrusOne (KKR/GIP), Switch (DigitalBridge), Stack, and an "Other" aggregate.
- **Sovereign clouds:** HUMAIN (Saudi PIF), MGX/Mubadala, Stargate UAE/G42/Khazna, Qatar (QIA),
  EU/France (Mistral), India (IndiaAI Mission), and an "Other sovereign AI" aggregate.

### Headline figures (research estimates, US$bn)
- Total (gross) financing ≈ **$6,084bn**; equity/internal ≈ **$3,870bn**, total debt ≈ **$2,214bn**
- By segment (equity + debt): Hyperscalers **$4,708** · Neoclouds **$666** · Colocation **$460** ·
  Sovereign **$250**
- Debt mix: corporate bonds $863bn · private credit $426bn · bank/syndicated loans $332bn ·
  insurance-linked lending $204bn · securitization/ABS $203bn · vendor/lease $160bn · convertibles $26bn

> **Cross-segment overlap:** Hyperscalers lease much colocation capacity and sovereign-wealth funds
> also invest in neoclouds/labs & colo (e.g. MGX/BlackRock bought Aligned; PIF put ~$3bn into xAI),
> so the all-segment grand total is a **gross** figure across distinct balance sheets, not strictly
> additive. Hyperscaler + Neocloud (~$5.4tn) ties to Goldman's ~$5.3tn big-tech capex; adding
> third-party colo + sovereign brings the gross total toward broader ~$6–7tn AI-infrastructure estimates.

> **Note:** Figures are research estimates that blend disclosed transactions (2023–early 2026)
> with forward projections. They are illustrative, not audited, and not investment advice.
> See the Cover & Methodology and Key Deals & Sources sheets for definitions and citations.

### Rebuild
```bash
pip install -r requirements.txt
python scripts/build_financing_model.py
```
