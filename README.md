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
| 8 | Sources | URLs and publications used to compile the tracker |

### Rebuild

```bash
pip install openpyxl
python3 scripts/build_spreadsheet.py
```

The script regenerates both the `.xlsx` workbook and the per-sheet CSVs
from the curated data inside the script. Edit the script to add rows or
sources, then re-run.

### Caveats

Data is hand-curated from public reporting (Q4 2025 - Q2 2026) including
press releases, vendor blogs, financial trade press and industry
conference coverage. Specific metrics (productivity gains, FTE
equivalents, cost savings) are reported as the issuing company / source
states them; they are not independently verified.
