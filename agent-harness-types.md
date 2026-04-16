# AI Agent Harnesses: Taxonomy, Competitive Landscape & Investment Implications

## 1. Defining the Agent Harness Market

An **agent harness** (also called an agent orchestration framework) is the software layer that wraps around a large language model to give it the ability to plan, use tools, maintain state, and coordinate with other agents. The harness is distinct from the model itself — it is the scaffolding that turns a stateless text-completion API into an autonomous system capable of multi-step task execution.

The market has fragmented into five distinct harness categories, each targeting a different buyer persona and use case:

| Category | Definition | Target Buyer | Examples |
|----------|-----------|--------------|----------|
| Open-Source Framework | Standalone libraries providing agent primitives (loops, tools, memory) | Developers, AI engineers, startups | LangGraph, CrewAI, Strands Agents, LlamaIndex |
| Platform SDK | SDKs tightly coupled to a model provider's ecosystem | Developers building on a specific provider | OpenAI Agents SDK, Claude Agent SDK, Google ADK, Microsoft Agent Framework |
| Managed Platform / Runtime | Fully managed infra for deploying agents built with any framework | Enterprise DevOps, platform engineers | AWS Bedrock AgentCore, Google Vertex Agent Engine |
| Enterprise SaaS | End-to-end SaaS with built-in agent capabilities, low-code | Business users, CRM admins | Salesforce Agentforce, ServiceNow Agent, SAP Joule |
| Research / Experimental | Frameworks for academic research and prototyping | Researchers, academics | OpenAI Swarm (deprecated), CAMEL-AI, AgentBench |

This taxonomy matters for equity research because each category implies a different monetization model, competitive moat, and margin profile.

---

## 2. Market Sizing

The agent harness / orchestration market is nascent but growing at extraordinary rates. Analyst estimates vary due to differing category definitions:

| Source | 2025 | 2026E | 2030E | 2034E | CAGR |
|--------|------|-------|-------|-------|------|
| Marqstats (Agentic AI Enterprise Platform) | $4.35B | $7.03B | $47.8B | — | 61.5% |
| MarketIntelo (Agent Orchestration Platform) | $5.8B | — | — | $38.6B | 23.7% |
| Deloitte (Autonomous AI Agent Market) | — | $8.5B | $35–45B | — | ~35% |
| Gartner (Global AI Spending, all categories) | $1.75T | $2.52T | — | — | 44% YoY |

The wide range ($35B–$48B for 2030) reflects definitional ambiguity: narrow definitions capture only the orchestration framework layer, while broader definitions include the managed runtime, tooling, and observability stack. Regardless, even the most conservative estimate implies a ~$35B market within five years.

### 2.1 Enterprise Adoption Curve

Enterprise deployment is still early-stage, which means the growth runway is substantial:

| Metric | 2023 | 2024 | 2025 | 2026E |
|--------|------|------|------|-------|
| Enterprises deploying agentic AI at scale | <1% | ~1% | 2% | ~5% |
| Fortune 500 piloting multi-agent workflows | 18% | ~35% | 62% | ~75% |
| Apps embedding agentic capabilities (Gartner) | — | — | 12% | 40% |

The gap between "piloting" (62%) and "at scale" (2%) is the conversion opportunity. Camunda's 2026 survey found that 73% of decision-makers report a gap between their agentic AI vision and current reality, and 81% say agentic orchestration is essential for achieving a fully autonomous enterprise. Budget intentions are strong: 79% plan to increase automation spend, with budgets rising ~20% over the next two years.

---

## 3. Open-Source Framework Competitive Landscape

Open-source frameworks are the foundation of the agent harness ecosystem. They compete on developer adoption, ease-of-use, and production-readiness.

### 3.1 Adoption Metrics

| Framework | GitHub Stars | Monthly Downloads | Contributors | Total Funding | Latest Stable |
|-----------|-------------|-------------------|-------------|---------------|---------------|
| LangGraph (LangChain) | ~29,100 | ~90M (ecosystem) | 280 | $160M | v2.0 (Feb 2026) |
| CrewAI | ~48,000 | ~5M (PyPI) | 300+ | $18M | v4.x (2026) |
| AutoGen / AG2 | ~57,000 | Declining | 450+ | N/A (Microsoft) | v0.7.5 (Sep 2025) — maintenance mode |
| LlamaIndex | ~47,800 | ~25M | 1,500+ | ~$35M | v0.12.x (2026) |
| Strands Agents (AWS) | ~6,100 | N/A | 120 | N/A (AWS) | v1.35.0 (Apr 2026) |

**Key observations:**

- **LangChain/LangGraph** dominates on downloads (90M/mo) and is the de facto standard for production agent deployments. Its $160M in funding (Series B at $125M, Oct 2025, led by IVP with Sequoia, Benchmark, CapitalG) reflects investor conviction. FY2025 revenue was $16M, monetized via LangSmith (tracing, evaluation, deployment SaaS).

- **CrewAI** has the highest star-count growth trajectory and the lowest barrier to entry (~25 min to first agent). It reported 2B+ agentic executions in the trailing 12 months and claims nearly half the Fortune 500 as users. With only $18M raised (Series A led by Insight Partners), it is capital-efficient but faces monetization questions.

- **AutoGen** has the most GitHub stars (~57K) but is now in **maintenance mode** following Microsoft's October 2025 announcement merging it with Semantic Kernel into Microsoft Agent Framework. New users are directed away from AutoGen. This is a cautionary tale for open-source adoption: high star counts do not guarantee longevity when the corporate sponsor changes strategy.

- **LlamaIndex** excels at RAG-heavy use cases and document pipelines, with the largest contributor base (1,500+). It occupies a complementary niche — many production systems use LlamaIndex for retrieval alongside LangGraph for orchestration.

- **Strands Agents** is AWS's entry, launched May 2025, emphasizing minimal code and native AWS integrations. At ~6K stars it is early but benefits from AWS's distribution power via Bedrock AgentCore.

### 3.2 Architecture Pattern Comparison

The choice of architecture pattern is the most consequential technical decision when selecting a harness:

| Pattern | Framework | Control Flow | State Management | Best For | Key Limitation |
|---------|-----------|-------------|-----------------|---------|----------------|
| Graph-based state machines | LangGraph | Deterministic; explicit branching | Native checkpointing, time-travel | Complex conditional workflows; regulatory compliance | Verbose; steep learning curve |
| Role-based agent crews | CrewAI | Sequential / hierarchical delegation | Partial shared context | Rapid prototyping; content pipelines | Limited cyclic graph support |
| Conversational message passing | AutoGen | Dynamic; agents decide when to respond | Conversation history as implicit state | Research; multi-expert deliberation | Higher latency; token overhead |
| Handoff-based chains | OpenAI Agents SDK, Strands | Linear with delegated branches | Per-agent context with handoff state transfer | Customer service routing; voice agents | Less suited for deeply collaborative scenarios |
| Tool-use agentic loop | Claude Agent SDK | Model-driven tool selection | Conversation context + tool results | Coding agents; computer use | Primarily single-agent |
| Model-driven orchestration | Google ADK, Strands | Model plans and executes | Sessions + Memory Bank | Flexible workflows with unknown plans | Reliability depends on model quality |
| Declarative / low-code | Salesforce Agentforce | Template-driven; rule-based | Managed by platform | CRM automation; business users | Least flexible; vendor lock-in |

### 3.3 Performance Benchmarks

Standardized benchmarks on a research-task workload (source: agent-harness.ai, 2026):

| Metric | LangGraph | CrewAI | AutoGen |
|--------|-----------|--------|---------|
| Median end-to-end latency | 14.1s | 18.4s | 22.7s |
| Cost per 1,000 tasks | $41.70 | $48.20 | $67.40 |
| Integration complexity (1–10, lower = simpler) | 6.8 | 3.5 | 5.9 |
| Time to first agent | ~2 hours | ~25 min | ~1 hour |

LangGraph wins on latency and cost but loses on developer experience. CrewAI's integration complexity score of 3.5 vs. LangGraph's 6.8 explains its rapid star growth — faster onboarding drives community adoption, even if the framework is less capable for complex production scenarios.

---

## 4. Platform SDK Landscape

Model providers are increasingly shipping their own harnesses to capture more of the application stack and create switching costs.

### 4.1 OpenAI Agents SDK

- **Successor to Swarm** (deprecated experimental framework, ~21K stars)
- **~20,700 GitHub stars**, 240 contributors, 81 releases
- **Key features:** Handoffs, guardrails, tracing, realtime voice agents (`gpt-realtime-1.5`)
- **Monetization:** No SDK fee — revenue captured through OpenAI API consumption (per-token billing)
- **Strategic logic:** Reduce friction for developers building on OpenAI models; compete directly with LangGraph/CrewAI by offering a "batteries-included" experience that doesn't require a third-party framework

### 4.2 Claude Agent SDK (Anthropic)

- **Renamed from Claude Code SDK** in late 2025 to reflect expanded scope
- **Unique architecture:** Computer-like interface (file system, shell, web) rather than graph topologies or role assignments
- **9 built-in tools** (Read, Write, Edit, Bash, Glob, Grep, WebSearch, WebFetch, AskUserQuestion) — zero custom tool setup required
- **18 lifecycle hooks** for intercepting, blocking, logging, or transforming agent behavior
- **Monetization:** API consumption; Claude Pro/Team/Enterprise subscriptions
- **Strategic logic:** Anthropic positions the agent as an autonomous computer user, not a workflow node — differentiated from the graph/crew paradigms

### 4.3 Microsoft Agent Framework

- **v1.0 shipped April 2026** for .NET and Python
- **Merges AutoGen + Semantic Kernel** into a single framework with graph-based workflows, multi-agent orchestration, middleware hooks, and memory providers
- **Replaces two separate projects:** AutoGen (57K stars, now maintenance-only) and Semantic Kernel (now receiving only bug fixes)
- **Monetization:** Azure AI consumption; Copilot subscriptions; Visual Studio / GitHub Copilot ecosystem
- **Risk:** Forbes noted (April 2026) that Microsoft's agent stack "confuses developers while rivals simplify" — the consolidation from multiple frameworks creates migration burden and ecosystem fragmentation

### 4.4 Google ADK (Agent Development Kit)

- **Open-source** (Apache 2.0), model-agnostic but aligned with Vertex AI
- **Multi-language:** Python, Java (1.0.0 Mar 2026), Go, TypeScript
- **7M+ downloads** by November 2025
- **Differentiator:** Agent-to-Agent (A2A) protocol with 50+ partners (Salesforce, ServiceNow, UiPath) enabling cross-framework agent communication
- **Monetization:** Google Cloud consumption; Vertex AI Agent Engine runtime fees
- **Strategic logic:** A2A protocol is a standards play — if it becomes the default interop protocol, Google captures the coordination layer

---

## 5. Managed Platforms & Enterprise SaaS

### 5.1 AWS Bedrock AgentCore

- **Framework-agnostic runtime** supporting LangGraph, Strands, CrewAI, Google ADK, OpenAI Agents
- **MicroVM session isolation** (dedicated VM per user session) for enterprise security
- **Spring AI SDK** for Java developers (GA 2026)
- **Monetization:** Pay-per-use compute; no framework lock-in — AWS captures value at the infrastructure layer regardless of which harness the customer uses
- **Strategic position:** AWS is betting that the harness layer commoditizes and that managed runtime + infrastructure is the durable value capture point

### 5.2 Salesforce Agentforce

Agentforce is the largest revenue-generating agent harness today:

| Metric | Q3 FY26 | Q4 FY26 | Growth |
|--------|---------|---------|--------|
| Agentforce ARR | ~$540M | $800M | +48% QoQ, +169% YoY |
| Cumulative deals | ~19K | 29K | +50% QoQ |
| Production accounts | — | — | +70% QoQ |
| Tokens processed (cumulative) | — | 19T | 5× YoY |
| Q4 bookings from expansion | — | 60%+ | Existing customer upsell |

- **Total Salesforce revenue:** $41.5B (FY2026), 10% YoY
- **Monetization:** Per-conversation pricing + seat licenses
- **Strategic moat:** CRM data lock-in; Agentforce agents operate directly on customer records, cases, and workflows already in Salesforce — switching costs are extremely high
- **Signal:** 60%+ of Q4 bookings from expansion (not new logos) indicates genuine enterprise adoption, not pilot churn

---

## 6. Token Usage by Harness Type

Token consumption is the primary cost driver for any agent deployment and varies dramatically by harness architecture. A single agentic task consumes 3–100× more tokens than an equivalent one-shot API call, with the multiplier determined by the harness's orchestration pattern, context management strategy, and multi-agent topology.

### 6.1 The Token Multiplier Effect

Every agent harness introduces overhead above raw model API costs. The overhead compounds across loop iterations because each step appends tool results and reasoning to the context window, causing input tokens to grow quadratically rather than linearly:

| Agent Loop Step | Approx. Input Tokens (Cumulative) | Marginal Cost Behavior |
|----------------|-----------------------------------|----------------------|
| Step 1 | ~2,500 | Baseline: system prompt + user query |
| Step 3 | ~8,500 | Linear growth: tool results accumulating |
| Step 5 | ~15,000 | Acceleration: prior reasoning + results |
| Step 10 | ~26,500+ | Near-quadratic: full history re-sent each call |
| Step 20 | ~55,000+ | Context window pressure; compression needed |

The cost formula follows a triangular number pattern: a naive N-step loop accumulates input costs proportional to N(N+1)/2, not N. A 20-step task can consume 10× the tokens a simple per-step estimate would suggest.

### 6.2 Token Overhead by Framework

Benchmarked token overhead vs. equivalent raw API calls (source: agent-harness.ai, TokenMix, 2026):

| Framework | Token Overhead vs. Raw API | Primary Overhead Source | Cost per 1K Tasks (GPT-4o) |
|-----------|---------------------------|----------------------|--------------------------|
| LangGraph | +9% | State serialization; minimal wrapper | $41.70 |
| CrewAI | +18% | Role prompts; delegation instructions; crew coordination | $48.20 |
| AutoGen | +31% | Conversational message passing; bilateral agent-to-agent messages; negotiation rounds | $67.40 |
| OpenAI Agents SDK | ~+12% (est.) | Handoff context; guardrail checks; tracing metadata | Model-dependent |
| Claude Agent SDK | ~+8–15% (est.) | Tool-use loop; minimal wrapper; but tool results can be large (file contents, shell output) | Model-dependent |
| Strands Agents | ~+10% (est.) | Minimal orchestration; model-driven planning | Model-dependent |
| Salesforce Agentforce | Abstracted (per-conversation) | $2/conversation or $0.10/action via Flex Credits; token costs hidden from customer | $2.00/conversation |

**Key insight:** The 3.4× cost gap between LangGraph ($41.70) and AutoGen ($67.40) per 1,000 tasks is driven almost entirely by token overhead, not model costs. The underlying model (GPT-4o) is identical — the harness architecture determines how many tokens are consumed per task.

### 6.3 Where Tokens Go: Anatomy of an Agent Task

A typical 10-step agent task on a reasoning model breaks down as follows:

| Token Category | Tokens per Call | Calls (10-step) | Total Tokens | % of Total | Cost Driver |
|---------------|----------------|-----------------|-------------|-----------|------------|
| System prompt | 1,500–4,000 | 10 | 15,000–40,000 | 10–15% | Repeated every call; fixed cost per framework |
| Tool/function schemas | 2,000–5,000 | 10 | 20,000–50,000 | 12–18% | Grows with number of tools registered |
| Conversation history | 2,500→26,500 | 10 | ~145,000 | 40–50% | Quadratic growth; largest cost component |
| Tool results | 200–3,000 | 10 | 2,000–30,000 | 5–12% | Highly variable; file reads and web fetches can be massive |
| Thinking/reasoning tokens | 3–10× output | 10 | 30,000–100,000+ | 15–30% | Reasoning models (o3, o4-mini, Opus 4) only; invisible but billed |
| Output tokens | 200–800 | 10 | 2,000–8,000 | 3–5% | 3–8× more expensive per token than input |

**Conversation history is the dominant cost**, accounting for 40–50% of total tokens. This is where harness architecture matters most — frameworks that manage history aggressively (trimming, summarization, compression) can cut this component by 60–80%.

### 6.4 Multi-Agent Token Multipliers

Multi-agent architectures multiply token consumption because each agent maintains its own context and inter-agent communication consumes tokens bilaterally:

| Topology | Token Multiplier vs. Single Agent | Explanation |
|----------|----------------------------------|-------------|
| Single agent (tool loop) | 1× (baseline) | One context window, one loop |
| Sequential handoff (2 agents) | 1.5–2× | Handoff transfers condensed context; second agent builds new history |
| Parallel fan-out (3 agents) | 2.5–3.5× | Each agent runs independently; results merged by orchestrator |
| Crew with manager (4 agents) | 3–5× | Manager agent delegates + reviews; each worker has full context cycle |
| Conversational swarm (4+ agents) | 5–10× | All agents read all messages; bilateral token costs; negotiation rounds |
| Hierarchical sub-graphs (nested) | 3–8× | Sub-graphs encapsulate; parent only sees summaries; most token-efficient multi-agent pattern |

**CrewAI's role-based crews** typically land at 3–5× because the manager agent adds an orchestration layer that consumes tokens for delegation, status checks, and result synthesis. **AutoGen's conversational model** pushes to 5–10× because every agent reads every message in the shared conversation, creating bilateral token charges.

**LangGraph's sub-graph pattern** is the most token-efficient multi-agent approach because sub-graphs run in isolation with their own state, and only return compact results to the parent graph — avoiding the full-context-sharing problem.

### 6.5 Context Management Strategies by Framework

How each framework handles the context growth problem directly determines token efficiency:

| Framework | Primary Strategy | How It Works | Token Savings | Trade-off |
|-----------|-----------------|-------------|---------------|-----------|
| LangGraph 2.0 | Autonomous context compression | Model itself triggers compression at task boundaries; agent decides when to summarize | 60–80% | Requires reasoning model capable of self-assessing context quality |
| LangGraph | Trimming + checkpointing | `max_token` limit drops oldest messages; checkpointer persists full state externally | 40–60% | Lossy — old messages lost from context (but recoverable from checkpoint) |
| LangGraph | Summarization | Second LLM call condenses old messages into summary before context overflow | 50–70% | Extra LLM call costs tokens itself; summary quality varies |
| CrewAI | Shared context window | Crew members share condensed context; manager summarizes between handoffs | 20–40% | Less granular control; manager summarization adds overhead |
| AutoGen | Full conversation replay | All messages retained in shared conversation; no built-in compression | 0% (none) | Highest token usage; context window fills fastest |
| OpenAI Agents SDK | Per-agent context isolation | Each agent maintains its own context; handoffs transfer only relevant state | 30–50% | Information loss during handoff if context not carefully structured |
| Claude Agent SDK | Conversation + tool results | Tool results appended to conversation; no built-in compression; relies on model's context window | 0–20% | Large tool results (file reads) can dominate context |
| Google ADK | Sessions + Memory Bank | Short-term (session) and long-term (Memory Bank) tiers; session compaction automatic | 40–60% | Memory Bank retrieval adds latency; not all context recoverable |

### 6.6 Cost Modeling: Harness Choice Impact on Annual Spend

To illustrate how harness architecture affects real-world costs, consider an enterprise running 100,000 agent tasks per month on GPT-4o ($2.50/1M input, $10.00/1M output):

| Scenario | Avg. Steps/Task | Tokens/Task | Monthly Token Cost | Annual Cost | Harness Overhead |
|----------|----------------|------------|-------------------|-------------|-----------------|
| One-shot API (no agent) | 1 | ~3,000 | $750 | $9,000 | — (baseline) |
| LangGraph (single agent, compressed) | 8 | ~45,000 | $11,250 | $135,000 | +9% vs. raw agentic |
| OpenAI Agents SDK (2-agent handoff) | 10 | ~70,000 | $17,500 | $210,000 | ~+12% |
| CrewAI (3-agent crew) | 12 | ~120,000 | $30,000 | $360,000 | +18% |
| AutoGen (4-agent swarm) | 15 | ~250,000 | $62,500 | $750,000 | +31% |
| Salesforce Agentforce | N/A | Abstracted | $200,000 | $2,400,000 | Per-conversation; no token visibility |

**The harness decision is a 5–8× annual cost swing** between the most efficient (LangGraph, $135K) and least efficient (AutoGen, $750K) open-source options for the same task volume. Salesforce Agentforce's per-conversation pricing ($2/conversation) is even more expensive in absolute terms but trades cost for zero engineering effort.

### 6.7 Token Cost Optimization Levers

| Lever | Token Reduction | Applicable Harnesses | Implementation Complexity |
|-------|----------------|---------------------|--------------------------|
| Prompt caching (provider-side) | 50–90% on cached input | All (Anthropic, OpenAI, Google offer caching) | Low — enable via API flag |
| Model tiering (use cheaper model for simple steps) | 60–80% on routed steps | LangGraph, CrewAI, Strands (model-per-node) | Medium — requires step classification |
| Context window trimming | 40–60% | LangGraph, Google ADK | Low — configuration parameter |
| Rolling summarization | 50–70% | LangGraph 2.0 (autonomous), manual in others | Medium — extra LLM call per summary |
| Sub-agent delegation | 50–70% for multi-agent | LangGraph (sub-graphs), Strands (handoffs) | Medium — architecture redesign |
| Tool result truncation | 10–30% | All (custom implementation) | Low — truncate large tool outputs |
| Reasoning model → non-reasoning for simple steps | 70–90% on thinking tokens | All (model selection per step) | Medium — requires step classification |
| Hard token budget caps | Variable (prevents runaways) | All (custom implementation) | Low — but risks task failure |

**The single highest-impact lever is prompt caching**, which reduces repeated system prompt and schema costs by 50–90%. Anthropic and OpenAI both offer automatic prompt caching for recurring prefixes, and this benefit applies regardless of harness choice. The second-highest impact is **model tiering** — using a cheaper model (Haiku, GPT-4o-mini) for simple routing/classification steps and reserving frontier models for complex reasoning.

### 6.8 Implications for Model Providers

Token usage patterns by harness type have direct implications for model provider revenue:

- **Agentic workloads consume 3–100× more tokens per user interaction** than chat or one-shot API calls, making agent adoption the primary growth driver for inference revenue
- **Reasoning model adoption amplifies this further** — thinking tokens (3–10× output) are billed but invisible, creating a "hidden" revenue multiplier
- **Context management innovation (LangGraph 2.0 autonomous compression) works against provider revenue** by reducing token consumption per task — this creates a tension between framework efficiency and API revenue
- **Per-conversation pricing (Salesforce Agentforce at $2/conversation)** decouples customer costs from token consumption, allowing the platform to capture margin on token-efficient implementations while absorbing losses on token-heavy ones
- **Output token pricing premiums (3–8× vs. input)** disproportionately affect agent workloads because agents generate more structured output (tool calls, reasoning chains) per interaction than chat users

---

## 7. Investment Implications

### 6.1 Where Value Accrues

The agent harness stack has four layers, each with different competitive dynamics:

| Layer | Players | Moat | Margin Profile |
|-------|---------|------|---------------|
| **Model Provider** | OpenAI, Anthropic, Google, Meta | Model quality; training data; compute scale | Compressed (40–50% gross margin; inference costs rising) |
| **Harness / Framework** | LangGraph, CrewAI, OpenAI SDK, Claude SDK, Google ADK | Developer adoption; ecosystem; switching costs | Low (open-source monetization via SaaS tooling) |
| **Managed Runtime** | AWS Bedrock AgentCore, Google Vertex Agent Engine | Infrastructure scale; enterprise trust | High (cloud-like margins, 60%+) |
| **Application / SaaS** | Salesforce Agentforce, ServiceNow, SAP | Customer data lock-in; workflow integration | Highest ($800M ARR at SaaS margins) |

**Key insight:** Value is migrating to the extremes — the application layer (Salesforce) captures the most revenue today, while the runtime layer (AWS, Google) is positioned for durable margin. The harness/framework layer risks commoditization as model providers ship their own SDKs.

### 7.2 Bull and Bear Cases

**Bull case for open-source harnesses (LangChain, CrewAI):**
- Enterprise adoption is at 2% penetration → massive runway
- Multi-model, multi-cloud deployments favor vendor-neutral frameworks
- LangSmith-style observability/evaluation SaaS has strong retention
- 79% of enterprises plan to increase automation budgets

**Bear case for open-source harnesses:**
- Model providers (OpenAI, Anthropic, Google) ship increasingly capable SDKs, reducing the need for third-party frameworks
- Microsoft's consolidation of AutoGen demonstrates corporate sponsor risk
- Cloud providers (AWS Bedrock AgentCore) offer framework-agnostic runtimes that commoditize the harness layer
- Salesforce Agentforce shows that low-code/no-code platforms can capture enterprise revenue faster than developer-first frameworks

**Bear case for platform SDKs (OpenAI, Anthropic, Google):**
- Vendor lock-in concerns push enterprises toward open-source alternatives
- Interoperability protocols (MCP, A2A) may standardize the tool/agent interface, reducing SDK differentiation
- Revenue still captured at the model/API layer — SDKs are a cost center, not a profit center

### 7.3 Key Metrics to Track

| Metric | Why It Matters | Current Benchmark |
|--------|---------------|-------------------|
| Fortune 500 at-scale deployment rate | Conversion from pilot to production | 2% (2025) |
| LangChain ecosystem monthly downloads | Proxy for developer mindshare | 90M/mo |
| CrewAI agentic executions | Usage intensity, not just installs | 2B+ trailing 12mo |
| Salesforce Agentforce ARR | Enterprise willingness to pay | $800M (Q4 FY26) |
| MCP / A2A protocol adoption | Interoperability reduces switching costs | MCP: broad; A2A: 50+ partners |
| Automation budget growth (enterprise survey) | Leading indicator of market expansion | +20% over next 2 years |

### 7.4 Risks

1. **Open-source commoditization.** If harness frameworks become commodity infrastructure (like web frameworks), monetization will be extremely difficult. LangChain's $16M revenue on 90M monthly downloads implies low conversion.

2. **Model provider vertical integration.** OpenAI, Anthropic, and Google each shipping their own SDKs could marginalize third-party frameworks, especially for simpler single-model deployments.

3. **Enterprise inertia.** The 2% at-scale adoption rate could persist longer than expected. Agentic AI requires trust in autonomous systems — regulated industries (finance, healthcare) may adopt slowly.

4. **Protocol fragmentation.** MCP (Anthropic-led) and A2A (Google-led) are competing standards. If they don't converge, enterprises face integration complexity that slows adoption.

5. **Cost pressure.** AutoGen's $67.40 per 1K tasks vs. LangGraph's $41.70 shows that agent overhead adds materially to LLM inference costs. If base model prices continue dropping, the harness layer's relative cost becomes more visible.

---

## 8. Conclusion

The agent harness market is fragmenting into five distinct categories, each with different competitive dynamics and investment profiles. The open-source framework layer (LangGraph, CrewAI) has captured developer mindshare but faces monetization challenges and vertical integration pressure from model providers. Platform SDKs (OpenAI, Anthropic, Google) are strategic cost centers designed to drive API consumption rather than standalone revenue generators. Managed runtimes (AWS Bedrock AgentCore) are positioned for durable cloud-like margins. Enterprise SaaS platforms (Salesforce Agentforce at $800M ARR) are capturing the most revenue today by embedding agents into existing workflow lock-in.

For equity research purposes, the most investable thesis is likely at the application layer (Salesforce, ServiceNow) where customer data lock-in creates durable switching costs, and at the infrastructure layer (AWS, Google Cloud) where managed runtimes capture value regardless of which harness wins. The open-source harness layer is a high-growth, high-risk bet that depends on whether LangChain and CrewAI can build SaaS businesses on top of their developer communities before model providers absorb their functionality.

---

*Data as of April 15, 2026. Sources: GitHub, CB Insights, Tracxn, Crunchbase, Marqstats, MarketIntelo, Deloitte TMT Predictions 2026, Gartner, Camunda State of Agentic Orchestration 2026, Salesforce Q4 FY26 earnings, company disclosures. Estimates marked (est.) are author projections. This is for informational purposes and is not investment advice.*
