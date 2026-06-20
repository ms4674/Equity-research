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

## 2. Segment forecasts

The combined entity is modelled as **three operating segments** on the `Segments`
tab. Each is driven by a revenue base, a growth path, and an EBITDA margin.

### 2.1 Space — launch services (Falcon/Starship) + Starshield

* FY2025 base **$6,000M** (≈ launch ~$4.2bn + Starshield ~$1.8bn).
* Heavy Starship R&D keeps early margins negative; reusability and cadence drive
  margin expansion over time.

| Driver | 2025A | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue growth | — | 28% | 30% | 28% | 24% | 20% |
| EBITDA margin | (8)% | 0% | 8% | 14% | 19% | 24% |

### 2.2 Starlink — satellite connectivity

* FY2025 base **$11,400M** — already the group's profit engine (reported segment
  operating margin ~39%; estimated EBITDA margin ~50%+).
* Subscriber growth + Direct-to-Cell drive continued scale economics.

| Driver | 2025A | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue growth | — | 42% | 36% | 30% | 25% | 21% |
| EBITDA margin | 45% | 48% | 51% | 53% | 55% | 57% |

### 2.3 xAI-Cursor — AI segment (xAI/Grok + acquired Cursor)

* FY2025 base **$1,750M** (xAI/other ~$1,300M + Cursor ~$450M recognised revenue).
* Reflects the fastest application-layer SaaS ramp on record (Cursor ARR $100M →
  >$2bn between Jan-2025 and early-2026) plus xAI/Grok. Margins are deeply negative
  early (investment + integration) and improve as inference is in-sourced.
* **All deal effects are carried in this segment:** revenue synergies are added to
  segment revenue; cost synergies and one-time integration costs adjust segment
  EBITDA; the acquired-intangible amortization is charged here in the EPS bridge.

| Driver | 2025A | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Organic revenue growth | — | 85% | 70% | 50% | 38% | 30% |
| EBITDA margin (pre-synergies) | (60)% | (25)% | (2)% | 8% | 16% | 22% |

EBITDA margins are expressed **excluding depreciation & amortization**; D&A is added
back separately so the model shows a clean EBITDA-to-EBIT bridge.

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

## 5. Segment EPS contribution

The `Segments` tab (and the bottom of the `Income Statement`) bridge each segment
from EBITDA to a net-income and **EPS contribution**, using a transparent allocation
that **sums exactly to consolidated net income and diluted EPS**:

```
Segment EBITDA
  − Depreciation        (each segment's own depreciation from the PP&E schedule)
  − Intangible amort.   (100% to xAI-Cursor — it is the Cursor deal)
  = Segment EBIT
  + Net interest        (allocated pro-rata to segment revenue share)
  = Segment pre-tax
  − Income tax          (allocated pro-rata to segment revenue share)
  = Segment net income
  ÷ Diluted shares      = Segment EPS contribution
```

Because segment depreciation sums to total depreciation, amortization sits in a
single segment, and net interest and tax are allocated by revenue share (which sums
to 100%), the segment
net-income and EPS contributions **add up to the consolidated totals by construction**
(verified on the `Checks` tab). The resulting picture: **Starlink** is the dominant
EPS driver, **Space** turns accretive as Starship matures, and **xAI-Cursor** is the
largest near-term drag (integration costs + intangible amortization) before scaling
into positive contribution.

## 6. Schedules & balance-sheet mechanics

* **PP&E — forecast by segment (through 2030):** the `Schedules` tab rolls each
  segment forward as `ending = beginning + capex − depreciation`, with totals that
  tie to the balance sheet and feed consolidated depreciation:
  * **Capex** is a percentage of each segment's own revenue, reflecting differing
    capital intensity — Space 28% → 18%, Starlink 36% → 22% (satellite
    constellation), xAI-Cursor 25% → 18% (AI compute, then moderating).
  * **Depreciation** is a percentage of each segment's beginning net PP&E,
    reflecting differing asset lives — Space 9%, Starlink 12%, xAI-Cursor 16%.
  * **Opening (FY2025) net PP&E** is split Space $14.0bn / Starlink $29.0bn /
    xAI-Cursor $3.3bn (Starlink is the balancing item so segment openings always
    sum to total opening PP&E of $46.3bn).
  * Each segment's depreciation flows into its EPS contribution; total capex and
    total depreciation flow to the cash flow statement and income statement.

  Illustrative output (US$M): total net PP&E grows from $46.3bn (FY2025) to ~$75.6bn
  (2030); total capex rises from ~$6.2bn to ~$15.7bn; total depreciation from
  ~$5.3bn to ~$7.9bn.
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

## 7. Integrity & conventions

The `Checks` tab confirms, for every projected year:

1. **Balance sheet balances** — total assets = total liabilities + equity (≈ 0).
2. **Cash flow ties out** — ending cash on the Cash Flow Statement equals balance-
   sheet cash (≈ 0).
3. **PPA bridge balances** — the opening pro-forma column balances (≈ 0).
4. **Segment net income** sums to consolidated net income (≈ 0).
5. **Segment EPS contributions** sum to diluted EPS (≈ 0).

By construction the Cash Flow Statement is derived from the change in every non-cash
balance-sheet line, so the balance sheet balances automatically. Interest is computed
on beginning balances to avoid circular references.

The **Summary** tab is a live, formula-linked output dashboard that consolidates the
key results (revenue/EBITDA/EPS by segment, the segment PP&E forecast, balance-sheet
and cash-flow highlights, deal terms, and the integrity checks). After `build_model.py`
writes the live formulas, `recalc.py` opens the workbook in LibreOffice with
"recalculate on load" forced on and re-saves it, so **every tab carries cached
computed values** (visible in any viewer) while the formulas stay live.

---

## 8. Sources / data points

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
