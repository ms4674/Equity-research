# Equity-research

A small repository of equity / thematic research notes.

## Notes

- [Impact of OpenAI and Anthropic (Claude) Co-Working AI Agents on Banking, Payments, and Financial Services](research/ai-agents-banking-payments-financial-services.md) — April 2026
- [Where Are AI Agents Most Effective in Banking and Financial Services?](research/where-ai-agents-are-most-effective-in-bfsi.md) — April 2026
  - Aggregated spreadsheet: [`data/bfsi_ai_agent_effectiveness.xlsx`](data/bfsi_ai_agent_effectiveness.xlsx)
  - Generator script: [`scripts/build_bfsi_ai_agents_workbook.py`](scripts/build_bfsi_ai_agents_workbook.py)

## Reproducing the spreadsheet

```bash
pip install openpyxl
python3 scripts/build_bfsi_ai_agents_workbook.py
# -> writes data/bfsi_ai_agent_effectiveness.xlsx
```

> All notes in this repository are for educational and discussion purposes only and are **not investment advice**.
