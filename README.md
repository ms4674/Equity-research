# Equity-research

## Agentic AI: Revenue & Token Usage Across Sectors (August 2026)

Aggregated market data on revenue and token usage for agentic AI offerings, compiled August 23, 2026.

### Files

| File | Contents |
| --- | --- |
| `agentic_ai_revenue_token_usage_2026.xlsx` | Main workbook: Read Me (methodology & caveats), Offerings (29 rows, one per agentic offering), Sector Aggregates (live formulas), Token Usage (17 platform/ecosystem metrics), Sources (40 numbered references) |
| `data/offerings_revenue.csv` | Git-friendly mirror of the Offerings sheet |
| `data/sector_aggregates.csv` | Sector revenue totals (computed values) |
| `data/token_usage_metrics.csv` | Token usage metrics |
| `data/sources.csv` | Source list with URLs |
| `scripts/build_spreadsheet.py` | Single source of truth; regenerates all outputs (`pip install openpyxl`, then `python3 scripts/build_spreadsheet.py`) |

### Headline aggregates (annualized revenue, tracked offerings only)

| Sector | Revenue ($M) | Share |
| --- | ---: | ---: |
| Coding & software development | 30,367 | 86.9% |
| Enterprise platform agents | 2,500 | 7.2% |
| General-purpose & consumer agents | 690 | 2.0% |
| Healthcare | 506 | 1.4% |
| Customer experience & support | 300 | 0.9% |
| Legal | 300 | 0.9% |
| Financial services | 300 | 0.9% |
| **Total** | **~34,963** | 100% |

Memo (not summed, overlaps rows above): OpenAI company run-rate $40B (Aug 2026), Anthropic $65B (Jul 2026).

Key caveats: Claude Code ($15.1B) and Codex ($8.8B) are third-party *tracked* estimates, not audited disclosures; Microsoft 365 Copilot revenue is undisclosed and excluded (30M+ seats imply up to ~$10.8B at list price), so the enterprise-platform sector is materially understated. See the workbook Read Me sheet for full methodology.
