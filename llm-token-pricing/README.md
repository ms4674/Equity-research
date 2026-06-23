# LLM Token Pricing Indices & Indicators

An aggregated reference of the **Silicon Data LLM Token Expenditure Index
(SDLLMTK)** and comparable LLM token-pricing indices, trackers and open
datasets, compiled into a single spreadsheet.

This was put together to give an equity-research view of *who publishes a
"price of LLM tokens" number, what it actually measures, and how to access it.*

## Files

| File | What it is |
| --- | --- |
| `LLM_Token_Pricing_Indices.xlsx` | Formatted, multi-sheet workbook (open in Excel / Google Sheets) |
| `data/indicators.csv` | Main comparison matrix (one row per indicator) |
| `data/silicon_data_products.csv` | Detail on the Silicon Data product suite |
| `data/methodology.csv` | Side-by-side methodology / construction comparison |
| `data/sources.csv` | Reference links for every entry |
| `build_spreadsheet.py` | Generator script — the single source of truth for the data |

## Workbook sheets

1. **Indicators** — the headline matrix. Columns cover publisher, ticker/ID,
   type, what it measures, unit, blend ratio, aggregation method, weighting,
   model scope, coverage, update frequency, history, access (paid/free/open),
   data source and URL.
2. **Silicon Data Suite** — breakdown of Silicon Data's relevant products
   (SDLLMTK index, Token Index API, Token Marketpulse, Token Pricebook, Portal).
3. **Methodology** — how each index is constructed (single number vs per-model,
   aggregation, blend ratio, normalization, weighting, key distinction).
4. **Sources** — numbered list of every reference URL.

## What's covered

**Silicon Data**
- LLM Token Expenditure Index (`SDLLMTK`) — daily, usage/expenditure-weighted
  blended USD/1M-token benchmark; reads as marginal willingness-to-pay /
  "quality premium" of frontier over open-weight models.
- Token Index API, Token Marketpulse, Token Pricebook.

**Comparable indices / indicators**
- Artificial Analysis — Intelligence Index price series & cost-per-task
  (7:2:1 cache:input:output blend).
- MyTokenTracker — AI Cost Index (Frontier & Budget, 3:1 blend, equal-weighted).
- Pulse — Inference Token Index (open-weight commodity hosts, breadth-weighted
  median, 3:1 headline).
- GCPI — Global Compute Price Index (input-only, weighted geometric mean,
  2026-Q2 = 100).
- AIscending — AI CPI & Budget Index (open dataset, 75/25 blend, monthly).
- LLM-Stats — composite score with blended pricing.
- LLMRates, BenchGecko, OpenRouter, LiteLLM — pricing trackers / upstream
  source feeds.

## Key takeaways

- **No single standard.** Blend ratios vary widely (Silicon Data normalizes the
  in/out mix dynamically; Artificial Analysis uses 7:2:1; Pulse & MyTokenTracker
  use 3:1; AIscending 75/25; GCPI uses input-only).
- **Two distinct things get called an "index":** (1) *expenditure/usage-weighted*
  series like SDLLMTK that move with where demand concentrates, and (2)
  *fixed-basket* price series (MyTokenTracker, GCPI) that hold the basket constant
  to isolate price movement.
- **Access split:** Silicon Data and LLMRates are paid/freemium; Artificial
  Analysis, LLM-Stats, Pulse and GCPI are free on the web; AIscending,
  BenchGecko, OpenRouter and LiteLLM are fully open data.
- Most aggregators ultimately draw on **OpenRouter** and/or **LiteLLM** as the
  upstream price source.

## Regenerating

```bash
pip install openpyxl
python3 build_spreadsheet.py
```

To add a new indicator, append a row to the relevant list in
`build_spreadsheet.py` and re-run.

## Notes & caveats

- Figures such as "models tracked" and any index levels reflect each
  publisher's **public materials as of 2026-06-23** and are point-in-time
  references, not live values.
- Live index values from Silicon Data require a Portal subscription / API key
  (`POST /api/token-index/index`), so they are not embedded here.
- This is a reference compilation, not investment advice.
