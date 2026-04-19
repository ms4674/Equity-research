# Equity-research

## Returns to intelligence — extended with per-task agent economics

Extends the original 7-row "Returns to intelligence" ranking (the slide that
ranks domains where intelligence pays off most) with three additional
per-task columns:

- **Avg input tokens / task** — system + user prompt + retrieved context +
  tool outputs the agent ingests per task.
- **Avg reasoning tokens / task** — hidden chain-of-thought / "thinking"
  tokens (OpenAI o-series `reasoning_tokens`, Anthropic Claude
  extended-thinking budget).
- **Avg agent task duration** — wall-clock from agent kickoff to terminal
  action (PR opened, ticket closed, alert resolved, hypothesis logged,
  trade idea finalised).

### Files

- [`agent_token_economics_by_domain.md`](agent_token_economics_by_domain.md) — markdown table + per-domain notes and sources.
- [`agent_token_economics_by_domain.xlsx`](agent_token_economics_by_domain.xlsx) — workbook with three sheets:
  1. Extended ranking (low–high range + midpoint per column)
  2. Numeric (chartable midpoints)
  3. Methodology & sources
- [`Returns_to_Intelligence_Extended.pptx`](Returns_to_Intelligence_Extended.pptx) — slide that mirrors the original layout with the three new columns, plus a companion slide showing the LLM-agent layer + public anchors.
- [`build_returns_to_intelligence.py`](build_returns_to_intelligence.py) — generator (regenerate via `python3 build_returns_to_intelligence.py`).

### Data caveats

- Numbers are **order-of-magnitude midpoints** triangulated from public
  benchmarks (SWE-Bench Verified, Tau-Bench, GAIA, OSWorld, FinanceBench,
  PaperQA2) and vendor disclosures (Anthropic, OpenAI, MSFT Security
  Copilot, CrowdStrike Charlotte AI, ServiceNow Now Assist, FutureHouse,
  Sierra). Not single-vendor measurements.
- For domains whose **core models are not LLMs** (HFT execution, ads
  ranking, sub-100 ms fraud scoring) the numbers reflect the LLM-agent
  layer that wraps them (campaign briefs, research/strategy copilots,
  case-review agents), not the production scoring/execution path.
- Refresh quarterly — reasoning-token budgets are now configurable
  (GPT-5 Thinking minimal/medium/high, Claude 4 thinking-budget) and
  ranges will compress as orchestrators learn to spend CoT only when
  EV is positive.

### Reproduction

```bash
pip install openpyxl python-pptx
python3 build_returns_to_intelligence.py
```
