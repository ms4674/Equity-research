# Equity-research

## Agentic AI: Revenue & Token Usage Across Sectors (August 2026)

Aggregated market data on revenue and token usage for agentic AI offerings, compiled August 23, 2026.

### Files

| File | Contents |
| --- | --- |
| `agentic_ai_revenue_token_usage_2026.xlsx` | Main workbook: Read Me (methodology & caveats), Offerings (29 rows, one per agentic offering), Sector Aggregates (live formulas), Token Usage (17 platform/ecosystem metrics), Reasoning Tokens (13 metrics by agent type/workload), GitHub Commits (per-agent commit attribution census), Sources (50 numbered references) |
| `agentic_ai_time_series_2026.pptx` | 12-slide deck charting the workbook's key time series (ARR ramps, token growth, GitHub commits by agent, reasoning-token shares) |
| `data/offerings_revenue.csv` | Git-friendly mirror of the Offerings sheet |
| `data/sector_aggregates.csv` | Sector revenue totals (computed values) |
| `data/token_usage_metrics.csv` | Token usage metrics |
| `data/reasoning_tokens.csv` | Reasoning-token consumption by agent type/workload |
| `data/github_commits.csv` | GitHub commit attribution per coding agent |
| `data/sources.csv` | Source list with URLs |
| `scripts/build_spreadsheet.py` | Single source of truth; regenerates workbook + CSVs (`pip install openpyxl`, then `python3 scripts/build_spreadsheet.py`) |
| `scripts/build_presentation.py` | Regenerates the PPTX deck (`pip install python-pptx matplotlib`) |

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

### Reasoning tokens & GitHub commits highlights

- Reasoning-optimized models now serve **>50% of all OpenRouter tokens** (vs ~0% in early 2025); ~21.6% of agentic coding task tokens are reasoning tokens (input 53.9% / output 24.4%).
- Claude Code produced **886K attributed commits in Dec 2025–Apr 2026 — 50% of all AI-attributed commits** — and ~4% of all public GitHub commits; Codex is the mirror image (814K PRs, near-zero commit traces).
- Attribution undercounts silent agents: Cursor and IDE Copilot don't sign commits, so PR- and commit-based censuses capture nearly disjoint agent populations.

Key caveats: Claude Code ($15.1B) and Codex ($8.8B) are third-party *tracked* estimates, not audited disclosures; Microsoft 365 Copilot revenue is undisclosed and excluded (30M+ seats imply up to ~$10.8B at list price), so the enterprise-platform sector is materially understated. See the workbook Read Me sheet for full methodology.
