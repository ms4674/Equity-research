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
| Deals & Sources | Debt transaction log grouped by instrument — corporate bonds · bank & syndicated loans · alternate private credit/SPV/144A · insurance-linked lending · securitization/ABS/CMBS & secured data-center bonds (incl. full Seaport Global desk-note universe) · convertibles (+ vendor/equity/context memos + sources). Debt amount column is numeric only |

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
- Deals & Sources tab is a debt transaction log grouped by instrument (numeric-only amounts): itemized debt
  transactions total **≈ $506bn** — corporate bonds ~$242bn (incl. SpaceX's $25bn inaugural bond + CoreWeave
  HY notes), alternate private credit/SPV ~$186bn, securitization/secured data-center bonds ~$55bn,
  bank & syndicated loans ~$18bn, convertibles ~$6bn — with separate memo sections for **insurance-linked
  lending** (e.g. Anthropic A2 notes $24bn to Athene/insurers), vendor/lease & backstops (Microsoft's $196.6bn
  finance leases), equity/sovereign commitments (Microsoft's BlackRock/GIP/MGX AIP up to $100bn), and context
- Securitization section now includes the full Seaport Global desk-note universe of **secured data-center
  project bonds** by tenant-credit tier — e.g. APLD/Applied Digital, Core Scientific, Cipher, Terawulf (WULF/Flash),
  Tract (SV/PR RNO), Meridian, Hut 8 (Beacon Point/River Bend), QTS Fayetteville, Edged, SE Cosmos, Elk Grove

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
