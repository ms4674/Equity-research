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
| `data/frontier_price_time_series.csv` | Token price over time for frontier model versions |
| `data/silicon_data_products.csv` | Detail on the Silicon Data product suite |
| `data/methodology.csv` | Side-by-side methodology / construction comparison |
| `data/sources.csv` | Reference links for every entry |
| `build_spreadsheet.py` | Generator script — the single source of truth for the data |

## Workbook sheets

1. **Indicators** — the headline matrix. Columns cover publisher, ticker/ID,
   type, what it measures, unit, blend ratio, aggregation method, weighting,
   model scope, coverage, update frequency, history, access (paid/free/open),
   data source and URL.
2. **Frontier Price Time Series** — chronological published API token prices
   (USD/1M tokens) for frontier/flagship model versions from OpenAI, Anthropic,
   Google, xAI and DeepSeek, from GPT-4 (Mar 2023) through Claude Fable 5 (2026).
   Each launch and price change is its own row, with input, output and a 3:1
   blended price, context window and notes.
3. **Silicon Data Suite** — breakdown of Silicon Data's relevant products
   (SDLLMTK index, Token Index API, Token Marketpulse, Token Pricebook, Portal).
4. **Methodology** — how each index is constructed (single number vs per-model,
   aggregation, blend ratio, normalization, weighting, key distinction).
5. **Sources** — numbered list of every reference URL.

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
- **Frontier time series**: prices are standard synchronous API *list* rates
  (excludes Batch, cached-input and Fast-mode tiers). For models with
  context-tiered pricing (e.g. Gemini Pro) the row uses the standard <=200K
  tier and notes the long-context rate. The blended column uses a fixed 3:1
  input:output ratio, `(3*input + output) / 4`. A few recent release dates are
  approximate and flagged in the Notes column.
- This is a reference compilation, not investment advice.
