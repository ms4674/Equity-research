# MCP Ecosystem: GitHub Commits, SDK Downloads, Token Usage & Compute Split (2022–1Q26)

Quarterly time series tracking the Model Context Protocol (MCP) ecosystem from 2022 through Q1 2026, covering GitHub commit activity, SDK download volumes, server ecosystem growth, GitHub stars, industry-wide LLM token consumption, the training-to-inference compute shift, and hyperscaler capex — with correlation analysis across all metrics.

## Spreadsheet

**[`data/mcp_ecosystem_token_usage_correlation.xlsx`](data/mcp_ecosystem_token_usage_correlation.xlsx)** — Multi-tab Excel workbook with all data aggregated, charts, and a full correlation matrix. Tabs:

| Sheet | Contents |
|-------|----------|
| Aggregated Summary | All metrics side-by-side (commits, downloads, servers, stars, tokens, inference %, capex) with embedded charts |
| GitHub Commits Detail | Commits broken out by repo (servers, python-sdk, typescript-sdk, specification, inspector, registry, github-mcp-server) |
| SDK Downloads Detail | npm + PyPI quarterly downloads with QoQ growth rates |
| Server Package Downloads | Downloads for 9 individual MCP reference server npm packages |
| Token Usage & Compute Split | Daily token volume (T), quarterly volume, inference vs training share, hyperscaler capex breakdown |
| Correlation Analysis | Pearson correlation matrix across all active-quarter metrics with key findings |
| GitHub Stars | Star counts by repo |
| MCP Server Count | Server ecosystem growth with net additions and sources |
| Sources & Methodology | Data provenance, APIs used, coverage windows, and caveats |

## CSV Data Files

| File | Description |
|------|-------------|
| `data/mcp_ecosystem_quarterly_summary.csv` | Combined summary: commits, downloads, server count, stars, key events |
| `data/mcp_github_commits_quarterly.csv` | GitHub commits by repository (7 repos) |
| `data/mcp_sdk_downloads_quarterly.csv` | SDK downloads: npm (@modelcontextprotocol/sdk) and PyPI (mcp) |
| `data/mcp_npm_server_downloads_quarterly.csv` | npm downloads for individual MCP reference server packages |
| `data/mcp_server_count_quarterly.csv` | Public MCP server count with net additions per quarter |
| `data/mcp_github_stars_quarterly.csv` | GitHub star counts by repository |

## Key Metrics (as of Q1 2026)

| Metric | Value |
|--------|-------|
| Monthly SDK downloads (Mar 2026) | 97M combined |
| Quarterly SDK downloads (Q1 2026) | 589M (npm + PyPI) |
| Public MCP servers | 10,500+ |
| GitHub stars (all repos) | 169,528 |
| Total GitHub commits (cumulative) | 13,244 |
| Contributors (servers repo) | 440 |

## Growth Timeline

- **Sep 2024** — MCP repositories created; internal development begins
- **Nov 25, 2024** — MCP publicly open-sourced by Anthropic with 3 reference servers
- **Q4 2024** — 1,910 commits; 117K npm downloads; 50 servers; 12K stars
- **Q1 2025** — OpenAI announces MCP support (Mar 2025); Claude Desktop ships MCP
- **Q2 2025** — Peak commit activity (3,517); npm downloads 18x Q1; GitHub MCP server launches
- **Q3 2025** — Microsoft Copilot adds MCP; 77.5M quarterly npm downloads
- **Q4 2025** — AWS Bedrock + Google DeepMind add MCP; Anthropic donates MCP to Linux Foundation (Dec 2025); PyPI downloads surge to 165M
- **Q1 2026** — 97M monthly SDK downloads (Mar 2026); 10K+ public servers; MCP v2.0 alpha development; 4,750% growth in 16 months

## Correlation Findings

The **Correlation Analysis** sheet computes Pearson correlations across all MCP and macro metrics for the 6 active quarters (Q4 2024 – Q1 2026):

| Finding | Detail |
|---------|--------|
| MCP SDK downloads ↔ Daily token volume | **r > 0.95** — MCP adoption tracks industry inference scaling almost perfectly |
| MCP SDK downloads ↔ Inference % of compute | **r > 0.95** — MCP growth is tightly coupled with the training→inference shift |
| MCP SDK downloads ↔ Hyperscaler capex | **r > 0.90** — Infrastructure spending and MCP adoption ride the same wave |
| MCP GitHub commits ↔ Downloads | **Negative** — Development peaked Q2 2025 while adoption continued accelerating (typical maturation pattern) |
| MCP server count ↔ Inference % | **r > 0.95** — More servers deployed as inference dominates compute mix |

**Interpretation:** MCP is not driving token usage — it is being **pulled forward by the inference scaling wave**. As inference shifts from ~20% (2022) to ~75% (Q1 2026) of total AI compute, standardized tool integration via MCP becomes essential infrastructure for agentic and enterprise deployment.

## Sources

- **GitHub commits**: GitHub Stats API (participation endpoint + commits list API with date filters)
- **npm downloads**: npm Registry API (`api.npmjs.org/downloads`)
- **PyPI downloads**: PyPI Stats API (`pypistats.org/api`); data available from Oct 2025 onward
- **Server count**: Glama directory, Pulse MCP, Bloomberry analysis, DreamFactory tracking, Anthropic announcements
- **Stars**: GitHub API current values; historical estimates based on community tracking and Wayback Machine snapshots
- **Token volume**: a16z/OpenRouter 100T token study, Google earnings disclosures, OpenAI announcements, industry estimates
- **Training vs inference**: Oplexa, GPUnex, Deloitte, Epoch AI analyses of compute allocation
- **Hyperscaler capex**: Epoch AI (SEC EDGAR 10-Q/10-K filings for Amazon, Microsoft, Alphabet, Meta, Oracle)

## Notes

- MCP did not exist before Q3 2024. Rows for 2022-Q1 through 2024-Q2 are zero by design to provide a complete baseline for token/compute context.
- PyPI download data for the `mcp` package is only available from Oct 2025 onward in the PyPI Stats API. Earlier Python SDK usage existed but is not captured in the `mcp` PyPI package stats.
- The `servers` repo commit count in Q1 2026 (85) is lower than prior quarters because the community contribution model shifted—server implementations increasingly live in the MCP Registry and third-party repos rather than the monorepo.
- GitHub star estimates for quarters before Q1 2026 are interpolated from known milestones and repo age; only Q1 2026 values are exact API readings.
- Token volume estimates for 2022–2023 are order-of-magnitude approximations; granular daily data from providers is only available from mid-2024 onward.
- Training vs inference percentages are interpolated from annual published figures; actual quarterly variation may differ.
