# Equity-research

A small repository of equity / thematic research notes.

## Notes

- [Impact of OpenAI and Anthropic (Claude) Co-Working AI Agents on Banking, Payments, and Financial Services](research/ai-agents-banking-payments-financial-services.md) — April 2026
- [Where Are AI Agents Most Effective in Banking and Financial Services?](research/where-ai-agents-are-most-effective-in-bfsi.md) — April 2026
  - Aggregated spreadsheet: [`data/bfsi_ai_agent_effectiveness.xlsx`](data/bfsi_ai_agent_effectiveness.xlsx)
  - Generator script: [`scripts/build_bfsi_ai_agents_workbook.py`](scripts/build_bfsi_ai_agents_workbook.py)
- [Verifiable Tasks: BFSI AI Agents vs General Coding Agents](research/verifiable-tasks-bfsi-vs-coding.md) — April 2026
  - Aggregated spreadsheet: [`data/bfsi_vs_coding_agent_verifiability.xlsx`](data/bfsi_vs_coding_agent_verifiability.xlsx)
  - Generator script: [`scripts/build_verifiability_workbook.py`](scripts/build_verifiability_workbook.py)
- [BFSI AI Agents: Token Consumption & Average Task Duration](research/bfsi-agent-token-economics.md) — April 2026
  - Aggregated spreadsheet: [`data/bfsi_agent_token_economics.xlsx`](data/bfsi_agent_token_economics.xlsx)
  - Generator script: [`scripts/build_token_economics_workbook.py`](scripts/build_token_economics_workbook.py)

## Reproducing the spreadsheets

```bash
pip install openpyxl
python3 scripts/build_bfsi_ai_agents_workbook.py
python3 scripts/build_verifiability_workbook.py
python3 scripts/build_token_economics_workbook.py
# -> writes data/bfsi_ai_agent_effectiveness.xlsx
#    writes data/bfsi_vs_coding_agent_verifiability.xlsx
#    writes data/bfsi_agent_token_economics.xlsx
```

> All notes in this repository are for educational and discussion purposes only and are **not investment advice**.
