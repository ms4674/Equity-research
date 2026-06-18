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
| **Assumptions** | Every input driver (blue cells are editable) |
| **IS SpaceX** | SpaceX standalone operating build (revenue → EBITDA) |
| **IS Cursor** | Cursor standalone operating build (revenue → EBITDA) |
| **Deal & PPA** | Consideration, purchase price allocation, goodwill, intangibles |
| **Schedules** | PP&E roll, intangible amortization & deferred tax, debt & spectrum, tax / NOL |
| **Income Statement** | Pro-forma combined P&L (revenue → net income, EPS) |
| **Balance Sheet** | Purchase-accounting bridge + projected balance sheet |
| **Cash Flow** | Integrated cash flow statement |
| **Checks** | Balance-sheet balance and cash-flow tie-out integrity checks |

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
| Total revenue | 19,124 | 29,194 | 41,994 | 56,232 | 71,303 | 87,520 |
| EBITDA | 1,347 | 2,560 | 6,364 | 11,453 | 17,715 | 24,421 |
| EBIT | (2,793) | (3,240) | 335 | 5,025 | 10,757 | 16,820 |
| Net income | (4,367) | (4,351) | 1,014 | 5,578 | 11,297 | 17,436 |
| Diluted EPS ($) | — | (2.04) | 0.48 | 2.61 | 5.30 | 8.17 |
| Ending cash | — | 62,292 | 64,127 | 69,246 | 78,823 | 93,660 |

Goodwill recognised on the deal: **~$45.3bn**; identifiable intangibles: **$15.0bn**.

## Running it

```bash
pip install -r requirements.txt
python build_model.py          # writes SpaceX_Cursor_Pro_Forma_Model.xlsx
python verify_model.py         # recalculates and prints integrity checks (optional)
```

`verify_model.py` recomputes the workbook with the `formulas` engine and confirms the
balance sheet balances and the cash flow ties to balance-sheet cash for every year.

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
