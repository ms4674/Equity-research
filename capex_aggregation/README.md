# Index capex aggregation

Bottom-up aggregation of annual capital expenditure for the constituents of
the **S&P 500** and **MSCI ACWI** (Bloomberg: MXWD) indices, written to
`../Index_Capex_Aggregation.xlsx`.

## Pipeline

1. `fetch_constituents.py` — downloads constituent lists:
   - S&P 500 from Wikipedia (`sp500_constituents.csv`)
   - MSCI ACWI proxied by iShares MSCI ACWI ETF holdings
     (`acwi_constituents.csv`), with local tickers mapped to Yahoo Finance
     symbols via an exchange-to-suffix table.
2. `fetch_capex.py` — pulls each company's annual cash-flow statement
   (Capital Expenditure line) and reporting currency from Yahoo Finance.
   Results are cached one JSON per ticker under `cache/` so the run is
   resumable. Yahoo rate-limits aggressively; re-run with `--retry-errors`
   (optionally `--skip-info` / `--info-only`) until coverage converges.
3. `build_workbook.py` — de-duplicates share classes, converts to USD using
   calendar-year average FX rates, aggregates by fiscal year, and writes the
   Excel workbook (summary + per-index detail + FX + methodology sheets).

## Requirements

```
pip install yfinance pandas openpyxl xlsxwriter lxml curl_cffi requests
```

See the workbook's Methodology sheet for sources and caveats.
