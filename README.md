# Agent Economy Data Model

Comprehensive data model covering the AI agent economy: top companies, task duration benchmarks, token consumption profiles, model pricing, and time-series projections (2024-2030).

## Deliverables

| File | Description |
|------|-------------|
| `agent_economy_data_model.xlsx` | Full Excel workbook with 12 sheets, charts, and time-series projections |
| `agent_economy_model.py` | Python script that generates the Excel workbook (re-run to regenerate) |

## Workbook Contents

### Dashboard
Summary of key metrics: 2026E market size ($35B), 2030E market size ($240B), CAGR, and navigation links to all sheets.

### Top Companies (20 companies)
Profiles of leading agent economy companies including OpenAI, Anthropic, Google DeepMind, Microsoft, Salesforce, Cursor, Cognition (Devin), LangChain, CrewAI, Adept AI, Replit, Cohere, Mistral, Perplexity, Hebbia, Sierra AI, and more. Includes valuations, revenue estimates, employee counts, and agent capabilities.

### Task Duration Benchmarks (21 task types)
Latency data across categories:
- Code generation (single function to full project scaffolding)
- Code review and debugging
- Document analysis and research
- Customer service
- Data analysis and reporting
- Email drafting and campaigns
- Workflow automation
- Image analysis and computer use

Includes avg, p50, p90, and p99 durations with bar charts.

### Token Consumption Profiles (24 task types)
Per-task token breakdowns:
- Input, output, and reasoning tokens
- Cost at frontier vs. mid-tier pricing
- Average tool calls per task
- Stacked bar chart for top consumers

### Model Pricing Reference (13 models)
Current API pricing for GPT-4o, Claude 4, Gemini 2.0, Mistral, Cohere, Llama, DeepSeek, and more.

### Time Series Projections (7 categories x 7 years)
Each category has its own sheet with data tables, YoY growth rates, CAGR calculations, and charts:

1. **Market Size** - Total agent market and segments (coding, enterprise, customer service, research)
2. **Daily Task Volume** - Global daily agent tasks by category (millions/day)
3. **Daily Token Consumption** - Global token usage (billions/day) split by input/output/reasoning
4. **Cost per Million Tokens** - Pricing trends for frontier, mid-tier, and open-source models
5. **Average Task Duration** - Speed improvements for simple/complex/agentic tasks
6. **Compute Infrastructure** - GPU hours, compute cost, agent share of inference
7. **Company Revenue Projections** - Revenue trajectories for OpenAI, Anthropic, Google, Microsoft, Salesforce, Cursor, and others

## Regenerating the Workbook

```bash
pip install xlsxwriter
python3 agent_economy_model.py
```

## Data Sources
Company filings, Crunchbase, PitchBook, press releases, IDC, Gartner, McKinsey, published API pricing, and research estimates. All forward-looking projections are illustrative.
