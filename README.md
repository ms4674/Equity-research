# Equity-research

Equity / industry research artefacts in this repo.

## Agentic AI in banking & payments

A multi-sheet tracker of agentic-AI products, deployments, protocols and
funding events across banks, card networks, processors, neobanks and
fintech vendors.

- Workbook: [`data/agentic_ai_banking_payments.xlsx`](data/agentic_ai_banking_payments.xlsx)
- CSV mirrors (one per sheet) in [`data/csv/`](data/csv) for easy diffing
- Build script: [`scripts/build_spreadsheet.py`](scripts/build_spreadsheet.py)

### Sheets

| # | Sheet | Contents |
| - | ----- | -------- |
| 0 | Overview | Scope, definitions, sheet index |
| 1 | Bank Deployments | Agentic / generative AI rollouts at large incumbent banks (JPM, BofA, Citi, GS, MS, Wells, HSBC, Barclays, UBS, DB, Lloyds, NatWest, Santander, BBVA, ING, DBS, RBC, TD, Scotia, Capital One, US Bank, Truist, PNC) |
| 2 | Payments Networks & Processors | Visa, Mastercard, Amex, PayPal, Stripe, Fiserv, Adyen, FIS, Synchrony, Block |
| 3 | Fintechs & Neobanks | Klarna, Nubank, Revolut, Starling, Monzo, Brex, Ramp, Slash, Mercury |
| 4 | Vendors & Startups | Hebbia, Glean, Harvey, Cohere, Fenergo, Bretton, StackAI, Temenos, Dyna.Ai, Stacks, EnFi, Ralio, KX, Cognition, Skyfire, Crossmint |
| 5 | Enterprise Platforms | Microsoft, Google Cloud, OpenAI, Anthropic, NVIDIA, Salesforce, AWS, ServiceNow |
| 6 | Agentic Payments Protocols | ACP, Visa Trusted Agent Protocol, Mastercard Agent Pay, UCP, AP2, MPP, x402, MCP |
| 7 | Funding Rounds | Selected 2025-2026 venture rounds for agentic-AI fintech |
| 8 | Coding vs FS Contrast | Side-by-side comparison of coding agents vs banking/payments/FS agents on token consumption, task duration, cost, reliability, tool surface |
| 9 | Leading Companies | Pick of the leader(s) in each agent category (coding, banking incumbents, payments networks/issuers/processors, fintech, vendors, platforms, protocols) |
| 10 | Sources | URLs and publications used to compile the tracker |

### Rebuild

```bash
pip install openpyxl
python3 scripts/build_spreadsheet.py
```

The script regenerates both the `.xlsx` workbook and the per-sheet CSVs
from the curated data inside the script. Edit the script to add rows or
sources, then re-run.

### Role and token-consumption columns

Every row in sheets 1–4 is tagged with two extra dimensions:

- **Role (Front-end / Back-end)** — whether the agent faces external
  counterparties (customers, merchants, third-party AI surfaces) or runs
  inside the institution (employee productivity, ops, compliance,
  engineering, treasury). `Both` is used when one product line spans
  both. Network-level rails (Visa ICC, Mastercard Agent Pay, etc.) are
  classed as Front-end because their consumers are external AI agents.
- **Token consumption (tier)** — qualitative monthly LLM-token usage
  estimated from disclosed scale (users × interactions × avg context):
  Very High (>1B/mo), High (100M–1B), Medium (10M–100M), Low (<10M),
  N/A (pure infrastructure / non-LLM), N/D (not disclosed). These are
  illustrative — issuers and vendors rarely publish exact token numbers.

### Caveats

Data is hand-curated from public reporting (Q4 2025 - Q2 2026) including
press releases, vendor blogs, financial trade press and industry
conference coverage. Specific metrics (productivity gains, FTE
equivalents, cost savings) are reported as the issuing company / source
states them; they are not independently verified.
