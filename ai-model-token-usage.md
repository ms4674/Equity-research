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

## 7. General-Purpose AI Agents: OpenClaw, Claude Cowork, Perplexity Computer

A new category of AI products is emerging beyond chatbots and coding agents: **general-purpose autonomous agents** that perform multi-step knowledge work, computer use, and real-world task execution. These products consume tokens at rates that rival or exceed coding agents, with distinct architectural patterns driving their costs.

### 7.1 Product Overview

| Product | Launched | What It Does | Pricing | Users |
|---------|----------|-------------|---------|-------|
| **OpenClaw** | Nov 2025 | Open-source AI agent platform; runs locally, connects to Discord/WhatsApp/Telegram; 5,700+ skills (email, calendar, GitHub, smart home); persistent memory | Free (self-hosted) + LLM API costs | 2M MAU, 335K GitHub stars |
| **Claude Cowork** | Jan 2026 | Anthropic's knowledge-work agent; operates on your computer via screen control, connectors (Gmail, Slack, Drive), and browser; handles docs, data, research tasks | $20-200/month (Pro to Max 20x) | Included in Claude subscriptions |
| **Perplexity Computer** | Feb 2026 | Orchestrates 19 AI models; decomposes goals into subtasks routed to specialist models (Claude Opus for reasoning, GPT-5.2 for long-context, Gemini for research); parallel cloud execution | $200/month (Max) + credits | Max subscribers |

### 7.2 Token Consumption Per Task

| Agent Type | Tokens per Interaction | Cost per Task | vs Chatbot Baseline |
|-----------|----------------------|---------------|-------------------|
| **Simple chatbot Q&A** | ~200-500 | ~$0.01 | 1x |
| **General ChatGPT conversation** | ~500-1,500 | ~$0.01-0.02 | 1-3x |
| **RAG-augmented query** | ~2,000-10,000 | ~$0.02-0.10 | 10-50x |
| **OpenClaw - single message** | ~8,000+ baseline | ~$0.01-0.03 | 40x |
| **OpenClaw - 5-turn conversation** | ~100,000+ (13x single turn) | ~$0.10-0.30 | 500x |
| **OpenClaw - web fetch task** | Variable (HTML processing) | ~$0.18/request | 900x |
| **Claude Cowork - light task** (file rename, draft email) | ~10,000-30,000 | ~$0.05-0.15 | 50-150x |
| **Claude Cowork - heavy task** (multi-source research, spreadsheet analysis) | ~100,000-500,000 | ~$0.50-5.00 | 500-2,500x |
| **Perplexity Computer - simple task** | Not disclosed (est. ~200-500 credits) | ~$2-5 | ~2,000-5,000x |
| **Perplexity Computer - complex workflow** (due diligence, multi-model project) | Not disclosed (500-2,000 credits) | ~$5-30+ | ~5,000-30,000x |
| **Coding agent - feature build (1hr)** | ~100,000 | ~$0.54-2.70 | 500x |
| **Coding agent - full day heavy use** | ~950,000 | ~$4.65-23.25 | 4,750x |

### 7.3 Why General-Purpose Agents Are So Token-Hungry

Each product category has distinct architectural drivers of token consumption:

**OpenClaw: Context compounding and baseline overhead**
- Every single message carries an **8,000-token baseline** payload (system prompt, memory, tool definitions) before any user content
- Multi-turn conversations suffer from **quadratic context growth**: a 5-turn chat costs **13x** a single turn because the entire history is re-sent each time
- Context compounding can cause sessions to balloon to **2.9MB** (~700K tokens) over just 35 messages as tool outputs are stored permanently in session files
- Gemini 2.5 Pro users report consuming **1.9M+ input tokens** in just a few dozen API calls due to the model's architecture
- Enabling reasoning/thinking mode adds **10-50x** token overhead

**Claude Cowork: Screen capture and multi-step execution loops**
- Computer use requires **iterative screenshot-action loops**: capture screen (~400-2,100 tokens per screenshot depending on resolution) -> analyze -> act -> repeat
- Each screenshot at 1920x1080 costs ~2,100 vision tokens; at typical web resolution (~1280x720), ~130-200 tokens
- A task requiring 50-100 screen interactions accumulates 10,000-210,000 tokens just from screenshots alone, before any reasoning
- Cowork shares the same token pool as Claude Chat and Claude Code, so heavy Cowork use directly cannibalizes chat/code capacity
- One engineer reported hitting Pro's session cap after only **~15 Cowork sessions**, implying each session consumes ~3x a standard chat conversation's allocation

**Perplexity Computer: Multi-model orchestration amplifier**
- The most token-intensive architecture: a single user request spawns **multiple sub-agents across 19 models** running in parallel
- The orchestrator (Claude Opus 4.6) decomposes tasks, then each sub-agent independently consumes tokens on its assigned model (GPT-5.2, Gemini, Grok, etc.)
- Total token consumption is the **sum across all models** -- a project touching 5 sub-agents effectively multiplies base token cost by 5x+
- Long-running tasks can execute for hours or months, continuously consuming tokens
- Token pricing varies dramatically by which sub-agent models are selected ($0.15/M for Flash vs $25/M output for Claude Opus), creating high variance in credit burn rates

### 7.4 Comparative Token Economics: Chatbot vs Coding Agent vs General Agent

| Dimension | Chatbot | Coding Agent | General-Purpose Agent |
|-----------|---------|-------------|----------------------|
| **Tokens per session** | 200-1,500 | 7,000-950,000 | 8,000-2,000,000+ |
| **Input:Output ratio** | ~3:1 | ~166:1 (Claude Code) | ~50-200:1 (varies) |
| **Primary token driver** | User prompt + response | Context re-reads (84% cached) | Context compounding, screenshots, multi-model fan-out |
| **Caching benefit** | Minimal | Critical (74% cost reduction) | Moderate-to-critical (varies by architecture) |
| **Session duration** | Seconds to minutes | Minutes to hours | Minutes to hours (or continuous) |
| **Tool calls per session** | 0-1 | 15-50+ | 10-100+ (including sub-agent spawns) |
| **Multi-model orchestration** | No | No (single model) | Yes (Perplexity: 19 models; OpenClaw: configurable) |
| **Computer vision tokens** | Rare | Rare | Frequent (Cowork: screen captures; Perplexity: browser automation) |
| **Cost predictability** | High | Moderate | Low (high variance by task complexity) |
| **Monthly cost (active user)** | $0-20 (subscription) | $20-200 (subscription) or ~$180 (API) | $20-200 (subscription) + potential overage |

### 7.5 Key Structural Differences

**Coding agents vs general-purpose agents operate at similar token volumes but for different reasons:**

Coding agents are dominated by **context re-reads** -- 99.4% of tokens are input, with 84% served from cache. The model reads large codebases repeatedly but generates relatively little output. Prompt caching is the primary cost lever, reducing costs by 74%.

General-purpose agents face **three compounding cost drivers** that coding agents largely avoid:
1. **Screen capture / vision tokens**: Computer use agents (Cowork, Perplexity) ingest screenshots at every step, adding 130-2,100 tokens per frame -- a cost channel that text-only coding agents don't have
2. **Multi-model fan-out**: Perplexity Computer's 19-model orchestration means one user request becomes 3-10+ parallel LLM calls across different providers, multiplicatively increasing total token consumption
3. **Unbounded session length**: Coding agents typically scope to a feature or bug fix; general-purpose agents can run for hours or continuously (Perplexity Computer can run for months), with no natural stopping point for token accumulation

**OpenClaw is structurally distinct** from both categories: as an open-source, self-hosted agent, the user directly bears API costs rather than the platform absorbing them via subscription margins. This makes token economics more transparent but also more volatile -- users report runaway consumption of 1-3M tokens within minutes without proper guardrails.

### 7.6 Market Context

| Product | Revenue / Ecosystem | Growth Trajectory |
|---------|-------------------|-------------------|
| **OpenClaw** | Ecosystem of 172 startups generating ~$360K/month; total spending est. $5-15M/month | 335K GitHub stars (surpassed React); 2M MAU with 92% retention; 27M monthly visitors (925% MoM growth, Feb-Mar 2026) |
| **Claude Cowork** | Part of Claude subscription revenue (~$19B ARR for Anthropic overall) | Expanded from Max-only to Pro tier in Jan 2026; computer use launched Mar 2026 |
| **Perplexity Computer** | Part of Perplexity Max ($200/month tier) | Launched Feb 2026; 19-model orchestration; early-stage adoption |

General-purpose agents represent the next frontier of token consumption growth. If coding agents drove tokens per session from hundreds to hundreds of thousands, general-purpose agents with multi-model orchestration and computer use are pushing toward **millions of tokens per session** -- a further 5-10x increase over coding agents for complex workflows.

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
- OpenClaw, "Token Usage & Cost Control Guide" (2026)
- BSWEN, "Why Does OpenClaw Burn Through Tokens So Fast?" (Mar 2026)
- Phala Network, "Understanding OpenClaw's Token Usage: A Data-Driven Deep Dive" (2026)
- OpenClaw VPS, "OpenClaw Statistics 2026: Growth, Users, Data" (2026)
- The Menon Lab, "From Fork to Industry: How OpenClaw Spawned a Market in Four Months" (Q1 2026)
- Anthropic, "Claude Cowork" product page (2026)
- TrySliq, "How Much Does Claude Cowork Actually Cost?" (2026)
- NYC Claw, "Claude Cowork Pricing: Plan Comparison" (2026)
- Anthropic, "Let Claude use your computer in Cowork" help article (Mar 2026)
- Perplexity, "Computer" product announcement (Feb 2026)
- BuildFastWithAI, "What Is Perplexity Computer? The 2026 AI Agent Explained" (2026)
- Ars Technica, "Perplexity announces Computer, an AI agent that assigns work to other AI agents" (Feb 2026)
- AICost, "Perplexity Computer: The 19-Model AI Digital Worker" (2026)
- TrySliq, "Perplexity Computer Pricing & Credits Explained" (2026)
- HackerNoon, "A Guide on How to Save Credits in Perplexity Computer" (2026)
- PageBolt, "Why screenshot MCPs cost 170x less than Playwright MCP" (2026)
- Anthropic, "Computer use tool" API documentation (2026)
