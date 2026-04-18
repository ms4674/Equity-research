# AI Foundation-Model & Agent Revenue + Token-Consumption Tracker

A reference dataset stitching together publicly disclosed revenue/ARR
time-series and monthly token-consumption figures for the leading
foundation-model labs and AI-agent vendors:

- **OpenAI** (foundation model lab)
- **Anthropic** (foundation model lab)
- **Harvey** (legal AI agent)
- **Sierra** (customer-experience AI agent)
- **Decagon** (customer-support AI agent)
- **Intercom Fin** (customer-support AI agent inside Intercom)

Two CSVs are provided:

| File | What it contains |
| --- | --- |
| `data/revenue_arr_timeseries.csv` | Date-stamped ARR / annualized run-rate observations for each company. |
| `data/monthly_token_consumption.csv` | API tokens-per-minute and converted tokens-per-month for OpenAI & Anthropic, with reference rows for Google, Microsoft Foundry, Together.ai and OpenRouter. |
| `ai_revenue_tokens.xlsx` | Multi-sheet Excel workbook with raw + pivoted data and embedded native bar charts (built by `build_spreadsheet.py`). |
| `charts/revenue_bar.png` | Grouped bar chart of ARR by company across snapshot dates. |
| `charts/tokens_bar.png` | Grouped bar chart of monthly token throughput by provider across snapshot dates. |
| `build_spreadsheet.py` | Generates the workbook and bar-chart PNGs from the CSVs. |
| `plot_timeseries.py` | Generates log-scale line charts from the CSVs. |

All figures are sourced from public disclosures (earnings calls,
DevDay keynotes, press releases, Reuters/The Information/Sacra
research) and are noted in the `source_note` column. Where a number
was not directly disclosed, it is marked as an interpolation/estimate
and the basis is noted.

## Headline numbers

### Revenue / ARR (USD millions, annualized run-rate unless noted)

| Date | OpenAI | Anthropic | Harvey | Sierra | Decagon | Intercom Fin |
|---|---:|---:|---:|---:|---:|---:|
| Dec 2023 | 1,600 | ~300 | ~10 | – | – | ~3 |
| Jun 2024 | 3,400 | ~600 | ~25 | – | ~5 | ~12 |
| Dec 2024 | 4,000 | 1,000 | 50 | ~25 | 10 | ~25 |
| Jun 2025 | 10,000 | ~3,000 | 75 | ~50 | ~20 | ~50 |
| Dec 2025 | 21,400 | 9,000 | 195 | 120 | ~40 | ~100 |
| Mar/Apr 2026 | 25,000–29,000 | 19,000–30,000 | 190 | 150+ | ~50 | ~150 |

Notes:
- OpenAI books revenue **net** of cloud-provider revenue share; Anthropic books **gross** customer payments. The two ARRs are not directly comparable.
- "Intercom Fin" ARR is the AI-agent product line, not the parent Intercom business (~$400M ARR Mar 2026).
- Several Anthropic, Sierra, Decagon, and Fin data points are interpolations between disclosed milestones; they are flagged in the CSV.

### Monthly API token consumption (foundation-model labs)

| Date | OpenAI tokens/min | OpenAI tokens/month | Anthropic tokens/min | Anthropic tokens/month |
|---|---:|---:|---:|---:|
| Oct 2023 | 300 M | ~13 T | – | – |
| Dec 2024 | ~2 B (est.) | ~88 T | – | – |
| Jun 2025 | ~4 B (est.) | ~175 T | ~1 B (est.) | ~44 T |
| Oct 2025 (DevDay) | **6 B** | **~263 T** | ~3 B (est.) | ~131 T |
| Mar 2026 | **15 B** | **~658 T** | ~5 B (est.) | ~219 T |

Tokens-per-month is computed from tokens-per-minute as
`TPM × 60 × 24 × 30.4375`, i.e. ~43,830 minutes per average month.

For market context the dataset also includes:
- **Google** — 480T (May 25), 980T (Jun 25), 1.3 quadrillion (Jul 25) tokens/month across Search, Gmail, Workspace, Gemini, Vertex etc.
- **Microsoft Foundry** — 1.7T (Apr 25) → 100T processed in a single quarter (Q2 FY26).
- **Together.ai** — ~2T tokens/day (Sep 25) of open-source inference.
- **OpenRouter** — ~100T tokens served across 2025.

## Sources

Primary public sources used to build the dataset (non-exhaustive):

- OpenAI: Reuters, The Information, CNBC, FT, OpenAI funding announcements,
  Sam Altman DevDay 2025 keynote, OpenAI "$122B raise" blog (Apr 2026).
- Anthropic: Reuters, Sherwood, Sacra, SaaStr, MegaOne AI, Bloomberg.
- Harvey: Sacra ("Harvey at $150M ARR", "Harvey at $195M ARR"),
  Harvey 2024 Year in Review, Pitchbook, ARR Club.
- Sierra: Sierra company blog ("100M ARR milestone", "Year two in review"),
  Sacra, Wikipedia.
- Decagon: Sacra, SiliconANGLE, Pitchbook, CB Insights.
- Intercom Fin: Mostly Metrics ("How Intercom reaccelerated growth"),
  Owen McCabe ("Fin Doubles Intercom's Growth Rate"), Intercom blog,
  Irish Times, NewsBreak.
- Token volumes: OpenAI DevDay 2025 (Altman), Tomasz Tunguz
  ("Beyond a Trillion: The Token Race"), Demirer/Fradkin et al.
  "The Emerging Market for Intelligence" (Dec 2025), Microsoft &
  Alphabet earnings calls, Adam Holter blog, Happycapy Guide,
  Anthropic public statements.

## Caveats

1. ARR is a **point-in-time annualization** of the most recent month's
   revenue, not trailing-twelve-months. Anthropic itself disclosed
   that its TTM revenue through Mar 2026 is materially below the
   $30B headline ARR.
2. Several private companies (Sierra, Decagon, Fin) only disclose
   round-number milestones; intermediate dates in the CSV are
   interpolations and labelled accordingly.
3. Token counts reflect inference token throughput (input + output +
   cached + reasoning where applicable). Provider definitions differ;
   Google in particular includes Search/Gmail/YouTube generative
   surfaces, while OpenAI/Anthropic figures are dominated by
   ChatGPT/Claude + their direct APIs.
4. All figures are denominated in current USD. No FX conversions
   applied (Intercom Fin is reported in USD by Intercom).
