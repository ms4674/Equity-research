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
| Summary (by Company) | Funding = operating cash flow (self-funded) + external equity + total debt; % debt-funded, chart |
| Debt by Type | Main table — total debt split into mutually-exclusive instrument buckets |
| Cash-Flow Bridge | Operating cash flow → free cash flow → external-funding gap (separates OCF & FCF from equity) |
| Capex Projections | Annual AI-infrastructure capex 2025–2030 — every company listed separately, grouped by segment with subtotals |
| Deals & Sources | Unified transaction log — (1) private-credit / SPV / 144A deals (itemized + total), (2) context/frameworks, (3) other key deals (bonds, equity, securitization, sovereign, leases), plus sources |

Funding sources are split three ways: **operating cash flow (self-funded)**, **external equity**
(IPOs, strategic stakes, ATM, sovereign equity, mandatory convertible preferred), and **total debt**.
The Cash-Flow Bridge sheet shows operating cash flow (CFO), the portion paid out as
dividends/buybacks, the self-funded capex, and **free cash flow = CFO − AI capex**.

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
- Total (gross) financing ≈ **$6,084bn**: operating cash flow (self-funded) ≈ **$3,372bn** +
  external equity ≈ **$498bn** + total debt ≈ **$2,214bn**
- Cumulative operating cash flow (CFO) ≈ **$5,069bn**; aggregate free cash flow (CFO − capex) ≈ **−$1,015bn**
  (the structural gap that drives external equity + debt issuance)
- By segment (total financing): Hyperscalers **$4,708** · Neoclouds **$666** · Colocation **$460** ·
  Sovereign **$250**
- Debt mix: corporate bonds $863bn · private credit $426bn · bank/syndicated loans $332bn ·
  insurance-linked lending $204bn · securitization/ABS $203bn · vendor/lease $160bn · convertibles $26bn
- Deals & Sources tab logs **17 itemized private-credit transactions ≈ $204bn** (Meta-Hyperion, Anthropic TPU SPV,
  xAI Colossus 2, Oracle SPVs, CoreWeave facilities, etc.) — consistent with reported >$200bn of
  outstanding AI private credit — plus market frameworks/projections (KKR-ECP $50bn, Morgan Stanley ~$800bn)
  and other key deals (corporate bonds, equity raises, securitization/CMBS, sovereign equity, finance leases)

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
