# Business Model: Headless Data Store for Equity Research

## Core Thesis

Position the platform as **infrastructure**—not an application. The product is a headless data store that serves structured equity research data (filings, fundamentals, estimates, transcripts, alternative data) via API to downstream consumers: AI agents, applications, dashboards, and workflows.

No proprietary UI competes with customers. The value is in **data quality, normalization, freshness, and availability**—not in presentation.

---

## Why Headless

| Advantage | Detail |
|-----------|--------|
| **Larger addressable market** | Every app, agent, and workflow that needs equity data is a potential customer—not just human analysts using a terminal. |
| **Defensible moat** | Structured, normalized, low-latency data pipelines are expensive to replicate. |
| **Network effects** | More consumers → more feedback on coverage gaps → better data → more consumers. |
| **Agent-native** | AI agents don't need UIs. They need deterministic, well-documented APIs with predictable schemas. |

---

## Revenue Model: How to Charge for Agents

### Pricing Dimensions

Agents (and the developers/companies deploying them) can be charged along several axes:

| Dimension | Description | Pros | Cons |
|-----------|-------------|------|------|
| **Per-query / API call** | Metered usage; pay per request. | Simple to understand; scales with value. | Unpredictable bills can deter adoption. |
| **Per-record / per-entity** | Charge based on number of tickers, filings, or data points accessed. | Aligns cost with breadth of coverage consumed. | Hard to predict for customers doing exploratory work. |
| **Per-agent seat** | Flat fee per registered agent (analogous to per-seat SaaS). | Predictable revenue; easy to sell. | Doesn't scale with actual consumption; may feel unfair. |
| **Tiered subscriptions** | Bundles of usage at fixed price points (e.g., Starter / Growth / Enterprise). | Predictable for both sides; upsell path. | Requires careful tier design to avoid leaving money on the table. |
| **Compute + data hybrid** | Charge separately for raw data access vs. enriched/computed outputs (e.g., sentiment scores, summaries). | Captures value of transformation, not just storage. | More complex billing. |

### Recommended Approach: Usage-Based with Commitment Tiers

```
Free Tier          →  Limited calls/month, subset of data (hook developers)
Pay-as-you-go     →  Per-query pricing with volume discounts
Committed Plans   →  Monthly/annual commitments at discounted rates
Enterprise        →  Custom contracts, SLAs, dedicated capacity
```

**Why this works for agents:**

1. **Agents are bursty.** A research agent might make 10,000 calls in an hour during earnings season and zero calls the next day. Usage-based pricing captures this naturally.
2. **Developers start small.** Free/PAYG tiers lower friction. Once an agent is in production and generating value, the operator upgrades to committed plans for cost predictability.
3. **Value scales with usage.** An agent that queries 5,000 tickers daily is extracting more value than one querying 10. Pricing should reflect that.

---

## Charging Mechanisms for Agent-Specific Scenarios

### 1. Agent Authentication & Metering

Each agent (or agent operator) registers an API key. All usage is metered against that key.

```
Agent Key → Rate Limits → Usage Tracking → Invoice
```

### 2. Data Tiers by Depth

| Tier | Coverage | Example Use Case |
|------|----------|------------------|
| **Core** | Price, volume, basic fundamentals | Screening agents, portfolio trackers |
| **Professional** | Full financials, estimates, transcripts, filings | Research automation agents |
| **Premium** | Alternative data, proprietary scores, real-time | Alpha-generating agents, HFT systems |

### 3. Freshness Pricing

| Latency | Premium |
|---------|---------|
| End-of-day | Base price |
| 15-min delayed | 2–3x |
| Real-time | 5–10x |

Agents that need real-time data for trading pay a premium over those doing batch research overnight.

### 4. Transformation / Enrichment Add-Ons

Charge additionally for derived outputs:

- Summarized earnings call transcripts
- Sentiment scores on filings
- Comparable company mappings
- Custom calculated metrics

These can be priced per-call or as add-on subscriptions.

---

## Competitive Positioning

| Competitor | Model | Our Differentiation |
|------------|-------|---------------------|
| Bloomberg Terminal | Per-seat, bundled UI + data | We're headless—no UI tax; agent-friendly APIs. |
| Refinitiv/LSEG | Enterprise licenses | Lower entry point; self-serve for developers. |
| Polygon.io / Alpha Vantage | API-first, usage-based | Deeper fundamental/research data, not just market data. |
| SEC EDGAR direct | Free but raw | We normalize, structure, and enrich. |

---

## Key Metrics to Track

- **Monthly Active Agent Keys** — adoption signal
- **Queries per Agent** — engagement / stickiness
- **Revenue per Query** — unit economics
- **Net Revenue Retention** — are agents consuming more over time?
- **Data Coverage Requests** — demand signal for expansion

---

## Open Questions

1. **Should there be a marketplace?** Let third parties contribute data sets and take a revenue share?
2. **Caching policy** — do we charge for cache hits, or only origin fetches?
3. **Resale rights** — can an agent operator redistribute derived insights to their end users?
4. **Rate limiting strategy** — hard caps vs. throttling vs. overage fees?
5. **Multi-agent discount** — should an operator running 50 agents get volume pricing across all of them?
