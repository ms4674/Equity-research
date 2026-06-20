# SpaceX × Cursor — Pro-Forma Three-Statement Model

An integrated, fully formula-linked **three-statement financial model** (Income
Statement, Balance Sheet, Cash Flow Statement) for **SpaceX following its
all-stock acquisition of Cursor (Anysphere, Inc.)**.

> **Deal context.** On **16-Jun-2026** SpaceX announced a binding agreement to
> acquire Cursor in an **all-stock transaction valued at ~US$60bn**, with close
> expected in Q3-2026. The deal followed SpaceX's June-2026 IPO and its February-2026
> absorption of xAI, and is aimed at strengthening SpaceX's AI / developer-tools
> position. Figures and assumptions are illustrative — see the disclaimer below.

## What you get

Running the builder produces **`SpaceX_Cursor_Pro_Forma_Model.xlsx`**, a workbook
with the following tabs:

| Tab | Contents |
| --- | --- |
| **Cover** | Transaction summary, contents, disclaimer |
| **Summary** | **Output dashboard** consolidating all key results (revenue/EBITDA/EPS by segment, PP&E, balance-sheet & cash-flow highlights, deal terms, integrity checks) — all live-linked |
| **Assumptions** | Every input driver (blue cells are editable) |
| **Segments** | **Space**, **Starlink** and **xAI-Cursor**: revenue, EBITDA, and a full per-segment bridge to net income and **EPS contribution** |
| **Deal & PPA** | Consideration, purchase price allocation, goodwill, intangibles |
| **Schedules** | **Segment PP&E forecast** (capex & depreciation by segment through 2030), intangible amortization & deferred tax, debt & spectrum, tax / NOL |
| **Income Statement** | Segment-driven P&L (revenue → net income, EPS) with **EPS contribution by segment** |
| **Balance Sheet** | Purchase-accounting bridge + projected balance sheet |
| **Cash Flow** | Integrated cash flow statement |
| **Checks** | Balance, cash-flow, and segment NI / EPS tie-out integrity checks |

### Segments

The combined entity is modelled as three reporting segments:

| Segment | Contents |
| --- | --- |
| **Space** | Launch services (Falcon / Starship) + Starshield (government/defense) |
| **Starlink** | Satellite connectivity |
| **xAI-Cursor** | AI segment — xAI/Grok **+ the acquired Cursor** business; carries all deal synergies, integration costs and acquired-intangible amortization |

The model is **live**: every projected cell is an Excel formula that references the
`Assumptions` tab, so changing any blue input flows through all three statements and
they stay in balance.

## How the three statements link

```
Income Statement ──► Net income ──────────────► Retained earnings (Balance Sheet)
        │                                              ▲
        ├─ D&A, amortization, SBC, deferred tax ──► Cash Flow (add-backs)
        │                                              │
Balance Sheet ◄── ending cash ◄────────────── Cash Flow (Δ working capital,
        │                                              capex, financing)
        └─ debt / PP&E / intangibles / DTL  ◄──  Supporting Schedules
```

* Interest is charged on **beginning-of-period** balances, so the model has **no
  circular references**.
* The acquisition is **all-stock and therefore non-cash**: it is reflected in the
  opening pro-forma balances (goodwill, intangibles, deferred tax liability, $60bn of
  newly issued equity, and Cursor's acquired cash), not as a cash outflow.

## Headline output (US$ millions, illustrative)

| | 2025PF | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue — Space | 6,000 | 7,680 | 9,984 | 12,780 | 15,847 | 19,016 |
| Revenue — Starlink | 11,400 | 16,188 | 22,016 | 28,620 | 35,775 | 43,288 |
| Revenue — xAI-Cursor | 1,750 | 3,288 | 5,704 | 8,606 | 11,843 | 15,361 |
| **Total revenue** | **19,150** | **27,156** | **37,703** | **50,006** | **63,465** | **77,665** |
| Total EBITDA | 3,600 | 6,548 | 12,263 | 18,196 | 25,282 | 33,468 |
| Total D&A | 5,268 | 6,901 | 7,273 | 7,858 | 8,621 | 9,486 |
| EBIT | (1,668) | (353) | 4,989 | 10,338 | 16,661 | 23,981 |
| Net income | (3,242) | (1,464) | 5,546 | 10,878 | 16,908 | 20,361 |
| **Diluted EPS ($)** | **(1.52)** | **(0.69)** | **2.60** | **5.10** | **7.93** | **9.54** |

### EPS contribution by segment ($ / diluted share)

| Segment | 2025PF | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Space | (1.05) | (0.74) | (0.18) | 0.24 | 0.71 | 0.93 |
| Starlink | 0.33 | 1.70 | 3.67 | 5.35 | 7.18 | 8.28 |
| xAI-Cursor | (0.81) | (1.65) | (0.89) | (0.49) | 0.04 | 0.34 |
| **Total** | **(1.52)** | **(0.69)** | **2.60** | **5.10** | **7.93** | **9.54** |

### PP&E forecast by segment (US$M, net)

| | 2025PF | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Space — capex / dep'n | 1,680 / 1,260 | 1,997 / 1,260 | 2,396 / 1,326 | 2,811 / 1,423 | 3,169 / 1,548 | 3,423 / 1,694 |
| Starlink — capex / dep'n | 4,104 / 3,480 | 5,504 / 3,480 | 6,605 / 3,723 | 7,728 / 4,069 | 8,586 / 4,508 | 9,523 / 4,997 |
| xAI-Cursor — capex / dep'n | 438 / 528 | 921 / 528 | 1,483 / 591 | 1,979 / 734 | 2,369 / 933 | 2,765 / 1,163 |
| **Total capex** | **6,222** | **8,421** | **10,484** | **12,518** | **14,124** | **15,711** |
| **Total depreciation** | **5,268** | **5,268** | **5,640** | **6,225** | **6,988** | **7,853** |
| **Total ending net PP&E** | **46,300** | **49,453** | **54,297** | **60,591** | **67,726** | **75,584** |

Starlink is the EPS engine; Space turns accretive around 2028 as Starship matures;
xAI-Cursor is the largest drag through the integration years (integration costs +
intangible amortization) before turning positive as the AI segment scales and
synergies build. Segment contributions **sum exactly to total diluted EPS** (see the
`Checks` tab). Goodwill recognised on the deal: **~$45.3bn**; identifiable
intangibles: **$15.0bn**.

## Running it

```bash
pip install -r requirements.txt
python build_model.py          # writes the workbook (live formulas)
python recalc.py               # recalculates with LibreOffice so every cell carries a cached value
python verify_model.py         # recomputes and prints integrity checks (optional)
```

* `build_model.py` writes the workbook with live formulas (openpyxl does not cache
  computed results).
* `recalc.py` opens the workbook in LibreOffice with "recalculate on load" forced on
  and re-saves it, so **every tab shows its computed values in any viewer** while the
  formulas remain live. (Requires LibreOffice; the committed workbook is already
  recalculated.)
* `verify_model.py` recomputes the workbook independently with the `formulas` engine
  and confirms the balance sheet balances, the cash flow ties to balance-sheet cash,
  and the segment net-income / EPS contributions sum to the consolidated totals.

## Documentation

See [`docs/METHODOLOGY.md`](docs/METHODOLOGY.md) for the full set of assumptions,
the purchase price allocation, the data points behind the standalone forecasts, and
the modelling conventions.

## Disclaimer

This is an **illustrative analytical model** built from publicly reported estimates
and reasonable assumptions about a privately negotiated transaction. It is **not
investment advice**, is **not verified by SpaceX or Cursor/Anysphere**, and should not
be relied upon for any investment or other decision. Many inputs (margins, share
count, opening balance sheet, synergies) are modelled estimates rather than disclosed
figures.
