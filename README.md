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
| **Segments** | **Space**, **Starlink** and **xAI-Cursor**: revenue, EBITDA, and a full per-segment bridge to net income and **EPS contribution** |
| **Deal & PPA** | Consideration, purchase price allocation, goodwill, intangibles |
| **Schedules** | PP&E roll, intangible amortization & deferred tax, debt & spectrum, tax / NOL |
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
| EBIT | (540) | 748 | 6,275 | 11,888 | 18,540 | 26,185 |
| Net income | (2,114) | (402) | 6,860 | 12,544 | 17,689 | 22,411 |
| **Diluted EPS ($)** | **(0.99)** | **(0.19)** | **3.22** | **5.88** | **8.29** | **10.51** |

### EPS contribution by segment ($ / diluted share)

| Segment | 2025PF | 2026E | 2027E | 2028E | 2029E | 2030E |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Space | (1.04) | (0.74) | (0.17) | 0.26 | 0.59 | 0.91 |
| Starlink | 1.00 | 2.34 | 4.40 | 6.19 | 7.80 | 9.26 |
| xAI-Cursor | (0.95) | (1.79) | (1.02) | (0.57) | (0.10) | 0.34 |
| **Total** | **(0.99)** | **(0.19)** | **3.22** | **5.88** | **8.29** | **10.51** |

Starlink is the EPS engine; Space turns accretive around 2028 as Starship matures;
xAI-Cursor is the largest drag through the integration years (integration costs +
intangible amortization) before turning positive as the AI segment scales and
synergies build. Segment contributions **sum exactly to total diluted EPS** (see the
`Checks` tab). Goodwill recognised on the deal: **~$45.3bn**; identifiable
intangibles: **$15.0bn**.

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
