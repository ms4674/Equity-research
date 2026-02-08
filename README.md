# Daily Software Stock Memo Generator

An automated investment memo generator that connects to Bloomberg Terminal (or Yahoo Finance as a free alternative) and produces daily investment memos focused on the biggest movers in the software sector.

## Features

- **Bloomberg Terminal Integration** — connects via `blpapi` for institutional-grade real-time data
- **Yahoo Finance Fallback** — free data source for environments without Bloomberg access
- **Comprehensive Software Universe** — tracks 50+ software/tech stocks across sub-sectors:
  - Enterprise Software (CRM, NOW, WDAY, ADBE, ORCL, etc.)
  - Cloud Infrastructure (MSFT, GOOGL, AMZN, SNOW, NET, etc.)
  - Consumer Software (META, SHOP, UBER, ABNB, SPOT, etc.)
  - Semiconductors / AI Adjacent (NVDA, AMD, AVGO, ARM)
  - Emerging AI Software (PLTR, AI, PATH, GTLB, etc.)
- **Biggest Movers Detection** — identifies top gainers and losers with volume confirmation
- **Rich Memo Output** — generates professional memos in Markdown, HTML, or plain text
- **Market Breadth Analysis** — advancers vs. decliners, sector rotation insights
- **Volume Analysis** — flags high-volume moves indicating institutional activity
- **Terminal Dashboard** — colorful rich-formatted console output
- **Configurable** — YAML-based config for stock universe, thresholds, and output preferences

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Generate a Memo

```bash
# Using Yahoo Finance (default, free, no setup required)
python main.py generate

# Print memo to console as well
python main.py generate --stdout

# Generate HTML output
python main.py generate -f html

# Use Bloomberg Terminal (requires blpapi + Bloomberg license)
python main.py generate -p bloomberg

# Show top 10 movers instead of default 5
python main.py generate -n 10

# Verbose logging
python main.py -v generate
```

### 3. Other Commands

```bash
# Check data provider connectivity and config
python main.py check

# List all stocks in the universe
python main.py list-stocks
```

## Bloomberg Terminal Setup

To use Bloomberg as the data source:

1. Ensure Bloomberg Terminal is running on your machine (or you have SAPI access)
2. Install the Bloomberg Python SDK:
   ```bash
   pip install blpapi
   ```
3. Update `config.yaml`:
   ```yaml
   data_source:
     provider: "bloomberg"
     bloomberg:
       host: "localhost"
       port: 8194
   ```

If Bloomberg is unavailable, the application automatically falls back to Yahoo Finance.

## Configuration

All settings are in `config.yaml`:

| Setting | Description | Default |
|---------|-------------|---------|
| `data_source.provider` | Data source: `bloomberg` or `yahoo` | `yahoo` |
| `memo.top_movers_count` | Number of top gainers/losers to feature | `5` |
| `memo.min_move_threshold` | Min % change to be "notable" | `1.0` |
| `memo.output_format` | Output: `markdown`, `html`, or `text` | `markdown` |
| `memo.output_dir` | Directory for saved memos | `output` |
| `memo.include_sector_summary` | Include sub-sector breakdown | `true` |
| `memo.include_volume_analysis` | Include volume analysis section | `true` |

### Customizing the Stock Universe

Edit the `stock_universe` section in `config.yaml` to add or remove stocks:

```yaml
stock_universe:
  your_custom_sector:
    - ticker: "TICKER"
      name: "Company Name"
```

## Output

Memos are saved to the `output/` directory with the naming convention:

```
output/software_stock_memo_YYYY-MM-DD.md
```

### Memo Sections

1. **Executive Summary** — overall market tone, top mover highlights, breadth
2. **Market Breadth** — gainers vs. losers, average/median changes
3. **Top Gainers** — detailed profile with price, volume, valuation, commentary
4. **Top Losers** — detailed profile with price, volume, valuation, commentary
5. **Sub-Sector Performance** — comparative table across software sub-sectors
6. **Volume Analysis** — high-volume movers suggesting institutional activity
7. **Full Universe Performance** — complete table of all tracked stocks

## Project Structure

```
.
├── main.py                  # Application entry point
├── config.yaml              # Configuration file
├── requirements.txt         # Python dependencies
├── output/                  # Generated memos
└── src/
    ├── __init__.py
    ├── cli.py               # Click CLI interface
    ├── data_provider.py     # Bloomberg + Yahoo Finance data layer
    ├── stock_universe.py    # Stock universe & mover analysis
    ├── memo_generator.py    # Markdown/HTML/text memo generation
    └── console_display.py   # Rich terminal dashboard
```

## Requirements

- Python 3.9+
- Internet connection (for Yahoo Finance) or Bloomberg Terminal (for Bloomberg data)
- See `requirements.txt` for Python package dependencies
