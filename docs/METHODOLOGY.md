# Methodology & Assumptions

This document explains the structure, data points, and assumptions behind the
SpaceX × Cursor pro-forma three-statement model (`build_model.py` →
`SpaceX_Cursor_Pro_Forma_Model.xlsx`). All figures are in **US$ millions** unless
stated otherwise. Inputs are highlighted in **blue** on the `Assumptions` tab; every
other projected cell is a formula.

---

## 1. The transaction

| Item | Value used | Basis |
| --- | --- | --- |
| Announcement | 16-Jun-2026 | Public reporting of the binding agreement |
| Structure | 100% stock | All-stock consideration |
| Equity value | **$60,000M** | Reported ~$60bn deal value |
| Expected close | Q3-2026 | Per the announcement |
| Modelling convention | Effective **1-Jan-2026** | Full-year pro-forma comparability |

Because the deal is **all-stock**, no cash leaves the combined entity to fund the
purchase. The transaction is captured entirely in the **opening pro-forma balance
sheet** (the "PF Open 1/1/26" column of the Balance Sheet tab): SpaceX issues ~$60bn
of new Class A stock, recognises Cursor's identifiable assets and liabilities at fair
value, books identifiable intangibles, a deferred tax liability and goodwill, and
brings Cursor's cash onto the combined balance sheet. There are therefore **no
acquisition-related cash flows** in the Cash Flow Statement.

### Why this scenario is plausible

* SpaceX completed its IPO in June 2026 (the model includes ~$50bn of primary
  proceeds), giving it a highly valued Class A currency to fund stock deals.
* SpaceX absorbed **xAI** in February 2026; acquiring Cursor — among the
  fastest-growing AI developer-tools companies — extends that AI strategy and gives a
  natural **cost synergy**: in-sourcing Cursor's model-inference spend onto
  xAI/Starlink compute.

---

## 2. Standalone forecasts

### 2.1 SpaceX (consolidated: Starlink, launch, Starshield, xAI)

Anchored to reported FY2025 figures:

* **FY2025 revenue ≈ $18,674M** (up ~33–43% YoY), with **Starlink ≈ $11.4bn** the
  primary engine, plus launch services (~$4.2bn), Starshield (~$1.8bn) and xAI/other.
* FY2025 was **GAAP loss-making** (operating loss ≈ $(2.6)bn; net loss ≈ $(4.9)bn),
  driven by Starship R&D (~$3bn) and the absorption of xAI losses.

Drivers (Assumptions tab):

| Driver | 2025A | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue growth | — | 40% | 35% | 30% | 25% | 22% |
| COGS % (ex-D&A) | 55.0% | 53.5% | 52.0% | 50.5% | 49.0% | 48.0% |
| R&D % | 22.0% | 19.0% | 17.0% | 15.0% | 13.5% | 12.5% |
| SG&A % | 14.7% | 13.5% | 12.5% | 11.5% | 10.5% | 10.0% |

Cost ratios are expressed **excluding depreciation & amortization**; D&A is added
back as a separate line so the model can show a clean EBITDA-to-EBIT bridge. The
2025 ratios are calibrated to reproduce roughly the reported operating loss.

### 2.2 Cursor (Anysphere)

The fastest application-layer SaaS ramp on record: ARR of ~$100M (Jan-2025),
~$500M (Jun-2025), ~$1bn (Nov-2025) and >$2bn annualised (early 2026); the $29.3bn
Series D priced in November 2025. Full-year **recognised** revenue lags exit ARR, so:

| Driver | 2025A | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue ($M) | 450 | 3,000 | 6,500 | 10,000 | 13,500 | 17,000 |
| COGS % (inference/cloud) | 55% | 50% | 46% | 43% | 40% | 38% |
| R&D % | 40% | 32% | 28% | 25% | 23% | 22% |
| Sales & marketing % | 35% | 30% | 26% | 22% | 20% | 18% |
| G&A % | 15% | 13% | 11% | 10% | 9% | 8% |

Cursor's high COGS reflects model-inference / cloud costs; margins improve as it
scales and (per the synergy assumption) in-sources inference onto SpaceX/xAI compute.

---

## 3. Purchase price allocation (`Deal & PPA` tab)

| Component | $M |
| --- | ---: |
| Equity purchase price (consideration) | 60,000 |
| Net tangible assets acquired (cash, AR, PP&E less liabilities) | 2,900 |
| Identifiable intangibles | 15,000 |
| — Developed technology (IDE + agent models), 8-yr life | 8,000 |
| — Customer relationships (enterprise contracts), 10-yr life | 5,000 |
| — Trade name (Cursor), 15-yr life | 2,000 |
| Deferred tax liability on intangibles (21%) | (3,150) |
| **FV of identifiable net assets** | **14,750** |
| **Goodwill** (= 60,000 − 14,750) | **45,250** |

* New SpaceX shares issued = $60,000M ÷ ~$450 share price ≈ **133.3M shares**.
* Annual intangible amortization ≈ **$1,633M** (8,000/8 + 5,000/10 + 2,000/15).
* As the intangibles amortize, the deferred tax liability unwinds, producing a
  **deferred tax benefit** of ≈ 21% × amortization each year.

---

## 4. Pro-forma adjustments

| Adjustment | 2026E | 2027E | 2028E | 2029E | 2030E | Treatment |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| Revenue synergies ($M) | 50 | 200 | 350 | 450 | 550 | Added to revenue |
| Cost synergies ($M) | 100 | 350 | 550 | 700 | 850 | Reduce COGS (in-source inference) |
| Integration / one-time costs ($M) | 500 | 0 | 0 | 0 | 0 | Operating expense |
| Intangible amortization ($M) | 1,633 | 1,633 | 1,633 | 1,633 | 1,633 | Below EBITDA |

---

## 5. Schedules & balance-sheet mechanics

* **PP&E:** `ending = beginning + capex − depreciation`. Capex is a % of combined
  revenue (26% → 16.5%); depreciation is 9% of beginning net PP&E.
* **Acquired intangibles:** amortize on a straight-line basis; deferred tax liability
  unwinds in proportion.
* **Debt & spectrum obligation:** rolled forward with net-draw and repayment
  assumptions; interest expense is charged at 6.5% on beginning interest-bearing
  balances (debt + spectrum obligation).
* **Tax / NOL:** statutory rate 21%. A large opening net operating loss carryforward
  (~$26bn, reflecting SpaceX's accumulated deficit plus Cursor) shelters early
  profits, subject to the 80%-of-taxable-income usage limitation. Intangible
  amortization is treated as non-deductible (it created the deferred tax liability).
* **Working capital:** recurring operating items (receivables, inventory, prepaids,
  payables, accrued liabilities, current deferred revenue, customer deposits) scale
  with revenue at ratios calibrated to the opening balance sheet. Large **structural**
  balances (long-term deferred revenue & deposits, other long-term items) are held
  flat so they do not generate artificial cash swings.
* **Financing:** the model includes ~$50bn of SpaceX primary equity issuance (the
  June-2026 IPO) in 2026; stock-based compensation (≈ 4–6% of revenue) is a non-cash
  add-back that also builds paid-in capital.

### Opening balance sheet (FY2025A, SpaceX standalone)

Constructed to be consistent with reported anchors (**total assets ≈ $92.1bn**,
**total equity ≈ $2.6bn** — a thin equity base reflecting heavy deposit/deferred-
revenue and debt funding). Individual line items within those totals are modelled
estimates.

---

## 6. Integrity & conventions

The `Checks` tab confirms, for every projected year:

1. **Balance sheet balances** — total assets = total liabilities + equity (≈ 0).
2. **Cash flow ties out** — ending cash on the Cash Flow Statement equals balance-
   sheet cash (≈ 0).
3. **PPA bridge balances** — the opening pro-forma column balances (≈ 0).

By construction the Cash Flow Statement is derived from the change in every non-cash
balance-sheet line, so the balance sheet balances automatically. Interest is computed
on beginning balances to avoid circular references.

---

## 7. Sources / data points

Public reporting and estimates used to anchor the model (all approximate):

* SpaceX FY2025 revenue ≈ $18.7bn; Starlink ≈ $11.4bn; GAAP net loss ≈ $4.9bn;
  total assets ≈ $92.1bn; total equity ≈ $2.6bn (S-1 / secondary coverage, 2026).
* SpaceX June-2026 IPO targeting a ~$1.5–1.75tn valuation, raising on the order of
  ~$50bn of primary capital; xAI absorbed February 2026.
* Cursor / Anysphere: $29.3bn Series D (Nov-2025); ARR trajectory $100M → $500M →
  $1bn → >$2bn (Jan-2025 through early-2026); ~$60bn all-stock acquisition by SpaceX
  announced 16-Jun-2026.

These are used only to set reasonable starting points; the projections, margins,
synergies, share count and opening balance-sheet composition are **illustrative
assumptions**, not company-disclosed figures.
