# AI Model Token Usage: OpenAI vs Anthropic vs Google Gemini

## Training vs Inference Compute Split

*Last updated: March 2026*

---

## 1. Daily Inference Token Volume

| Provider | Daily Tokens Processed | Source / Date |
|----------|----------------------|---------------|
| **Google (Gemini)** | ~32.7 trillion | Tunguz, Jul 2025 |
| **OpenAI** | ~8.6 trillion | OpenRouter / a16z, Oct 2025 |
| **Anthropic** | ~5-10 trillion (est.) | Tunguz, 2025 |
| **Together.ai** (open-source) | ~2 trillion | Tunguz, Sep 2025 |
| **Microsoft Foundry** | ~0.057 trillion | Tunguz, Apr 2025 |

Google dominates raw inference volume at roughly 4x OpenAI's throughput, processing over **1.3 quadrillion tokens per month** as of Q3 2025 -- a 20x year-over-year increase. This is driven by Google's vertically integrated infrastructure (custom TPUs) and the embedding of Gemini across Search (AI Overviews reaching ~2B monthly users), Workspace, and Cloud.

OpenAI's token throughput is growing rapidly as well: API reasoning token consumption per enterprise organization increased **320x YoY**, and ChatGPT message volume grew **8x YoY** among enterprise customers.

Anthropic's exact daily volume is less precisely disclosed, but industry estimates place it in the 5-10 trillion tokens/day range, consistent with its ~$2.7B annual inference compute spend.

---

## 2. Compute Spend: Training vs Inference Breakdown

### OpenAI (2024 Actuals)

| Category | Spend | % of Compute |
|----------|-------|-------------|
| R&D Compute (training + research) | ~$5.0B | ~73% |
| - Final training runs (released models) | <$1.0B | <15% |
| - Experimental / derisking runs | ~$3.0B | ~43% |
| - Research compute (amortized) | ~$1.0B | ~15% |
| Inference compute | ~$1.8-2.0B | ~27% |
| **Total compute** | **~$7.0B** | **100%** |

Key insight: Most of OpenAI's 2024 compute went to **experiments and unreleased models**, not to final training runs or inference. The actual cost of training released models (GPT-4o, o3, etc.) was under $1B -- the bulk of the $5B R&D spend went to architectural exploration, scaling law validation, and models that never shipped (per Epoch AI analysis).

By 2025, OpenAI's inference costs are estimated to have grown to **~$7B annually** (out of ~$8.5B total OpEx), reflecting the shift as user adoption scaled dramatically (800M+ weekly active users, 2.5B daily queries).

### Anthropic (2025 Actuals)

| Category | Spend | % of OpEx |
|----------|-------|----------|
| R&D Compute (training) | ~$4.1B | ~42% |
| Inference compute | ~$2.7B | ~28% |
| Other (staff, overhead, etc.) | ~$2.9B | ~30% |
| **Total operating expenses** | **~$9.7B** | **100%** |

Anthropic's **training-to-inference ratio is roughly 60:40** (of compute-only spend), reflecting its heavier investment in frontier model development relative to its user base. For context, Anthropic spent $2.66B on AWS alone through September 2025 -- exceeding its September revenue of $2.55B.

### Google / DeepMind (2025 Estimated)

Google does not break out training vs inference compute costs separately. However:
- **2026 CapEx guidance**: $175-180B (total, not AI-only)
- **Inference efficiency**: 4.5x improvement in tokens per GPU-hour YoY; 78% reduction in serving costs
- **Token throughput**: 10B tokens/minute (up from 7B prior quarter), ~430 trillion tokens annualized
- Google's custom TPU infrastructure provides a **~4x cost advantage** over Nvidia GPU-based inference, giving it structural cost leadership

Google's vertically integrated model (owning the silicon, the cloud, and the applications) makes direct cost comparisons difficult, but its cost-per-token is estimated to be **70-80% lower than OpenAI or Anthropic** for equivalent workloads.

---

## 3. Training vs Inference: Industry-Wide Shift

| Year | Training Share | Inference Share | Source |
|------|---------------|----------------|--------|
| 2023 | ~67% | ~33% | Gartner / GPUnex |
| 2025 | ~50% | ~50% | Gartner / GPUnex |
| 2026 (projected) | ~33% | ~67% | Gartner / GPUnex / Deloitte |
| 2029 (projected) | ~35% | ~65%+ | Gartner |

The industry is undergoing a structural shift from training-dominated to inference-dominated compute:

- **Inference accounts for 80-90% of lifetime AI system costs** because training is a one-time event while inference runs continuously at scale
- The training cost for GPT-4 was ~$150M; OpenAI's 2024 inference spend was ~$2.3B -- a **15x ratio** that is projected to reach **118x by 2026**
- The inference market is projected to grow from **$106B (2025) to $255B (2030)** at a 19.2% CAGR

---

## 4. Company-Level Comparison Summary

| Metric | OpenAI | Anthropic | Google (Gemini) |
|--------|--------|-----------|----------------|
| **Daily inference tokens** | ~8.6T | ~5-10T (est.) | ~32.7T |
| **Monthly active users** | 800M+ weekly | Not disclosed | 750M+ monthly |
| **2025 annualized revenue** | ~$20B | ~$9B (end of 2025) | Not broken out (Cloud AI >200% YoY growth) |
| **2026 revenue target** | Not disclosed | $18-26B | Not broken out |
| **Inference compute spend (annual)** | ~$7B (2025 est.) | ~$2.7B (2025) | Not broken out (structural cost advantage via TPUs) |
| **Training compute spend (annual)** | ~$5B (2024) | ~$4.1B (2025) | Not broken out |
| **Training:Inference compute ratio** | ~73:27 (2024) shifting to ~45:55+ (2025) | ~60:40 (2025) | Not disclosed (likely inference-heavy given product breadth) |
| **API pricing (flagship, input/output per 1M tokens)** | GPT-5.4: $2.50/$15.00 | Claude Opus 4.6: $5.00/$25.00 | Gemini 3.1 Pro: $2.00/$12.00 |
| **Key infrastructure** | Azure / Oracle (Stargate) | AWS + GCP + Fluidstack ($95B+ committed) | Custom TPUs (v5e, v6e, v7) |
| **Inference cost advantage** | Baseline | ~Similar to OpenAI | ~70-80% lower per token (TPU advantage) |

---

## 5. Key Takeaways

1. **Google leads in raw token throughput** at ~32.7T tokens/day, roughly 4x OpenAI and 3-6x Anthropic. This is driven by TPU infrastructure advantages and Gemini's integration across Google's product ecosystem (Search, Workspace, Cloud).

2. **OpenAI's compute allocation shifted dramatically** from training-heavy (73% in 2024) toward inference-heavy in 2025, as user adoption scaled to 800M+ weekly active users and inference costs grew to an estimated $7B annually.

3. **Anthropic invests proportionally more in training** (60:40 training-to-inference ratio) than OpenAI, consistent with its strategy of pushing frontier model capabilities. It is also the only provider where compute costs still exceed revenue.

4. **The industry-wide trend is unmistakable**: inference is moving from one-third of compute (2023) to two-thirds (2026). Lifetime inference costs already exceed training costs by 15x or more.

5. **Google holds a structural cost advantage** of 70-80% lower per-token inference costs thanks to custom TPU silicon. This advantage compounds as the industry shifts toward inference-dominated workloads.

6. **All three providers operate at a loss on inference** today, pricing below cost to capture market share and drive ecosystem lock-in. This is a deliberate strategic choice, not an efficiency problem.

---

## 6. Token Usage: Coding Agents vs Other Use Cases

### 6.1 Token Intensity by Use Case

| Use Case | Tokens per Interaction | Typical Cost | Multiplier vs Chatbot |
|----------|----------------------|-------------|----------------------|
| **Simple chatbot Q&A** | ~200-500 tokens | ~$0.01 | 1x (baseline) |
| **General ChatGPT conversation** | ~500-1,500 tokens (avg 348 words, 1.7 messages) | ~$0.01-0.02 | 1-3x |
| **RAG-augmented query** | ~2,000-10,000 tokens | ~$0.02-0.10 | 10-50x |
| **Customer support agent** | ~5,000-15,000 tokens (avg 11 LLM calls per conversation) | ~$0.05-0.15 | 25-75x |
| **Coding agent - quick fix** | ~7,000 tokens (5K input + 2K output) | ~$0.05-0.23 | 35x |
| **Coding agent - feature build (1hr)** | ~100,000 tokens (80K input + 20K output) | ~$0.54-2.70 | 500x |
| **Coding agent - large refactor** | ~360,000 tokens (300K input + 60K output) | ~$1.80-9.00 | 1,800x |
| **Coding agent - full day heavy use** | ~950,000 tokens (800K input + 150K output) | ~$4.65-23.25 | 4,750x |
| **GitHub Copilot agentic workflow** | ~650K-860K tokens per run (avg) | Varies | ~4,000x |

Coding agents are by far the most token-intensive consumer use case -- a single feature-building session consumes **~500x more tokens** than a typical chatbot interaction, and a full-day coding session can reach **~5,000x** the baseline.

### 6.2 Why Coding Agents Consume So Many Tokens

The "agent tax" is driven by structural factors unique to coding workflows:

**Context re-reads (52% of total spend):** Every request re-sends the full conversation history, file contents, system instructions, and tool definitions. Claude Code data shows an input-to-output ratio of **166:1** -- for every 1 token of code generated, 166 tokens of context are read. Even a simple "Hello" message sends ~20,000 tokens of context overhead.

**Multi-turn reasoning loops:** Agentic workflows use ReAct-style loops (plan -> tool call -> observe -> reason -> repeat), with each cycle re-processing the full context. A 10-cycle loop can consume **50x** the tokens of a single-pass query.

**Tool call overhead:** Each tool invocation adds function schemas, call parameters, and result injection. A coding agent making 15-20 tool calls in a session accumulates substantial token overhead beyond the actual code being written.

**Retrieval inefficiency:** How the agent finds relevant code matters enormously. Structural code analysis uses ~8,500 tokens vs grep-based file search at ~117,000 tokens for the same task -- a **14x** difference.

**Prompt caching is critical:** Claude Code data shows 84% of input tokens are served from cache. Without caching, costs are ~$310 per 100M tokens; with caching, ~$82 -- a **74% reduction**. This is the single most important efficiency lever for coding agents.

### 6.3 Coding's Share of Total AI Token Volume

Coding has become the **dominant use case by token volume** across the AI industry:

| Period | Coding's Share of Total Tokens | Source |
|--------|-------------------------------|--------|
| Early 2025 | ~11% | OpenRouter / a16z |
| Late 2025 | **>50%** | OpenRouter / a16z |

This shift reflects the mainstream adoption of AI coding tools. On open-source model platforms, the distribution is different -- roleplay leads at ~52%, with coding second -- because closed-model providers (OpenAI, Anthropic) refuse roleplay content, concentrating that workload on open-source.

For Chinese open-source models specifically: coding + technology = ~39%, roleplay = ~33%.

### 6.4 Coding Agent Revenue vs Other AI Products

The AI coding tools market has rapidly scaled to become one of the largest AI product categories:

| Product | ARR (March 2026) | Notes |
|---------|-----------------|-------|
| **Claude Code** | ~$2.5B | Reached $1B ARR within months of launch; ~13% of Anthropic's total revenue |
| **Cursor** | ~$2.0B | Doubled from $1B in 3 months |
| **GitHub Copilot** | ~$2.0B | 46% of all developer-written code is Copilot-assisted |
| **Total AI coding tools market** | ~$8-8.5B (2025) | Projected to reach $91B by 2035 (27.6% CAGR) |

The three leaders (Copilot, Cursor, Claude Code) hold **70%+ combined market share**. 84% of developers use or plan to use AI coding tools, and 90% of Fortune 100 companies have adopted them.

### 6.5 The "Jevons Paradox" of AI Tokens

A counterintuitive dynamic is at play: **token prices have collapsed ~300x** since early 2023 (from ~$30/M to ~$0.10/M for Gemini 2.0 Flash input tokens), yet **total AI bills are rising** because:

1. Cheaper tokens enable more complex, multi-step agentic workflows that were previously cost-prohibitive
2. Coding agents consume 500-5,000x more tokens per session than simple chatbot queries
3. Reasoning models (now >50% of total token volume) use dramatically more tokens per task
4. Enterprise adoption is deepening -- structured workflow usage grew **19x** and reasoning token consumption grew **320x** YoY

The result: per-token costs are falling, but total inference demand is growing faster, driven overwhelmingly by coding and agentic use cases. This is the primary driver of the industry-wide shift from training-dominated to inference-dominated compute allocation described in Section 3.

### 6.6 Cost Comparison Across Providers for Coding

Average daily cost per developer using AI coding tools (API/pay-per-use pricing):

| Tool | Avg Daily Cost | Pricing Model |
|------|---------------|---------------|
| **Claude Code (API)** | ~$6/day (~$180/month) | Pay-per-token |
| **Claude Code (Max 20x sub)** | ~$6.67/day ($200/month) | Subscription, capped |
| **Cursor Pro** | ~$0.67/day ($20/month) | Subscription, usage-capped |
| **GitHub Copilot Pro** | ~$0.33/day ($10/month) | Subscription |
| **Devin** | ~$9/hour ($2-2.25/ACU) | Pay-per-compute-unit |

The wide pricing spread reflects fundamentally different architectures: subscription tools (Cursor, Copilot) cap usage and subsidize heavy users, while pay-per-token tools (Claude Code API, Devin) expose the true cost of agentic token consumption.

---

## Sources

- Epoch AI, "Most of OpenAI's 2024 compute went to experiments" (2025)
- Epoch AI, "Compute accounts for the majority of expenses of AI companies" (2025)
- Tomasz Tunguz, "Beyond a Trillion: The Token Race" (2025)
- a16z / OpenRouter, "State of AI: An Empirical 100 Trillion Token Study" (2025)
- AI2Work, "AI Inference Economics in 2025" (2025)
- Fintool / The Decoder, "Google's Gemini API requests more than double in five months" (2025)
- Reuters, "Anthropic aims to nearly triple annualized revenue in 2026" (Oct 2025)
- Reuters, "OpenAI CFO says annualized revenue crosses $20 billion" (Jan 2026)
- GPUnex, "AI Inference Economics: The 1,000x Cost Collapse Reshaping GPUs" (2026)
- Introl, "AI Inference vs Training Infrastructure" (2025)
- Deloitte, "More compute for AI, not less" (2026)
- Gartner, AI infrastructure spending projections (2025-2029)
- Sacra, "Anthropic" research report (2026)
- BSWEN, "Claude Code Token Usage: Real Data From 100M Tokens Tracked" (Mar 2026)
- BSWEN, "Claude Code Cost: Real Pricing Data From 100M Tokens" (Mar 2026)
- ClaudeCodePricing.com, Claude Code pricing plans and calculator (2026)
- Adam Holter, "OpenRouter's 100 Trillion Token Study" (2025)
- Adam Holter, "AI Costs in 2025: Cheaper Tokens, Pricier Workflows" (2025)
- Adam Holter, "Cheap AI Tokens, Expensive Tasks" (2025)
- Cash and Cache, "The Agent Tax: Why Your AI Workflow Costs 50x More Than You Think" (2026)
- Milind Nair, "Your AI Agent is Burning 10x the Tokens It Needs" (Mar 2026)
- Grislabs/AgentMeter, "Cost Anatomy of 1,127 Agent Runs" (2026)
- Zylos Research, "AI Agent Cost Optimization: Token Economics and FinOps in Production" (2026)
- IdeaPlan, "AI Coding Assistants Market Share 2026" (2026)
- OpenPR, "AI Code Tools Market Size Projected to Reach $91.09 Billion by 2035" (2026)
- WebFX, "How People Use ChatGPT: Stats From 13,252 Conversations" (2025)
- OpenAI, "The State of Enterprise AI 2025 Report" (2025)
