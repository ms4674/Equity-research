# Returns to intelligence — extended with per-task agent economics

Extends the original 7-row 'Returns to intelligence' ranking with three per-task columns:

- **Avg input tokens / task** — system + user prompt + retrieved context + tool outputs the agent ingests per task.
- **Avg reasoning tokens / task** — hidden chain-of-thought / 'thinking' tokens (OpenAI o-series `reasoning_tokens`, Anthropic Claude extended-thinking budget).
- **Avg agent task duration** — wall-clock from agent kickoff to terminal action (PR opened, ticket closed, alert resolved, hypothesis logged, trade idea finalised).

Numbers are order-of-magnitude midpoints triangulated from public benchmarks (SWE-Bench Verified, Tau-Bench, GAIA, OSWorld, FinanceBench, PaperQA2) and vendor disclosures (Anthropic, OpenAI, MSFT Security Copilot, CrowdStrike Charlotte AI, ServiceNow Now Assist, FutureHouse, Sierra). They are *not* single-vendor measurements — refresh quarterly.

## Extended ranking table

| Rank | Domain | Why returns are high | Avg input tokens / task | Avg reasoning tokens / task | Avg agent task duration |
|---:|---|---|---|---|---|
| 1 | **Trading / market-making / quant execution** | Extreme leverage + tight feedback + compounding | 40,000 – 400,000 (≈120,000) | 3,000 – 40,000 (≈12,000) | 20 s – 10 min (≈1.5 min) |
| 2 | **Ads ranking / recommender systems** | Massive scale + measurable uplift | 2,000 – 30,000 (≈8,000) | 200 – 6,000 (≈1,500) | 2 s – 1 min (≈10 s) |
| 3 | **Cybersecurity (SOC + response)** | High cost of failure + growing attack surface + automation thresholds | 20,000 – 200,000 (≈60,000) | 2,000 – 25,000 (≈8,000) | 15 s – 10 min (≈1.5 min) |
| 4 | **Fraud / credit / underwriting** | Leverage + clear outcomes (loss rate) | 8,000 – 120,000 (≈25,000) | 1,000 – 15,000 (≈5,000) | 10 s – 5 min (≈45 s) |
| 5 | **Drug discovery / biotech R&D** | Huge payoff for better hypotheses + time-to-market value | 100,000 – 1,500,000 (≈300,000) | 10,000 – 150,000 (≈40,000) | 2 min – 2 h (≈15 min) |
| 6 | **Software engineering agents** | Verifiable feedback + autonomy thresholds + huge labor spend | 50,000 – 600,000 (≈150,000) | 8,000 – 100,000 (≈30,000) | 1 min – 45 min (≈10 min) |
| 7 | **Complex enterprise ops (IT/SRE, finance ops, compliance)** | Repetitive workflows + measurable KPIs | 5,000 – 80,000 (≈20,000) | 500 – 12,000 (≈3,000) | 5 s – 4 min (≈30 s) |

## Where the LLM-agent layer actually sits

### 1. Trading / market-making / quant execution
- **Agent layer:** LLM agents sit on the *research/strategy/PM-copilot* layer (filings, transcripts, news, alt-data). HFT execution itself remains sub-millisecond RL/heuristic — no LLM in the hot path.
- **Public anchors:** BloombergGPT (50B-token finance LM); JPM IndexGPT & Athena; Bridgewater AIA Labs; Man Group / Two Sigma research copilots; FinanceBench (Patronus); FinGPT.
- **Notes:** Long context dominated by retrieved 10-Ks / transcripts (50-300k tokens). Reasoning budget mid-range (o3/Claude 4 thinking) because rec/PM workflows are evaluated against ground truth slowly (P&L) — agents err toward more context, less CoT.

### 2. Ads ranking / recommender systems
- **Agent layer:** Core ranking is non-LLM (DLRM/HSTU/transformers on feature vectors, <100 ms). LLM 'agents' here = campaign-orchestration copilots (Meta Advantage+, Google PMax, TikTok Symphony) and creative-generation agents.
- **Public anchors:** Meta Advantage+ Shopping Campaigns; Google Performance Max & Gemini-for-Ads; TikTok Symphony; Amazon Ads AI creative; Pinterest Performance+; Snap AR Genie.
- **Notes:** Per-impression: 0 LLM tokens. Per *campaign brief / creative iteration*: small input (advertiser brief + history), short CoT, sub-minute. Reasoning tokens low — most heavy lifting is in image/video diffusion models, not text CoT.

### 3. Cybersecurity (SOC + response)
- **Agent layer:** Triage / investigation / response agents: CrowdStrike Charlotte AI, MSFT Security Copilot, Google SecLM/Sec-Gemini, Palo Alto Cortex XSIAM AI, Wiz, SentinelOne Purple AI, Dropzone AI.
- **Public anchors:** MSFT Security Copilot per-investigation cost disclosures (FY24 earnings); CrowdStrike Charlotte AI public demos; Dropzone AI MTTR benchmarks; SOC.OS / Anvilogic; CyberSecEval (Meta).
- **Notes:** Heavy retrieval: alert + IOCs + endpoint timeline + asset context + threat-intel. Multi-step (enrich → correlate → hypothesize → contain). Tier-1 triage closer to low end; full incident response closer to high end.

### 4. Fraud / credit / underwriting
- **Agent layer:** Sub-100 ms scoring (XGBoost/GBM/RNN) is non-LLM. LLM agents wrap the *case-review / adverse-action / KYC narrative / manual-underwriting* layer: Stripe Radar Assistant, Sardine, Featurespace, Upstart manual review, Zest AI explainability.
- **Public anchors:** Upstart / Zest AI public model docs; Klarna AML agent (2024 letter); Stripe Radar; Plaid Beacon; FICO Falcon; CFPB adverse-action LLM pilots.
- **Notes:** Application packet (PDF statements + bureau pull + device + behaviour) drives the input token count. Case-narrative generation is the dominant CoT cost.

### 5. Drug discovery / biotech R&D
- **Agent layer:** Two layers: (a) structural / generative bio models (AlphaFold 3, Isomorphic, RFdiffusion, Recursion Phenom) — no LLM tokens; (b) LLM research agents: FutureHouse PaperQA2 / Crow / ChemCrow, Owkin K, Insilico Pharma.AI, BioNeMo agents.
- **Public anchors:** FutureHouse PaperQA2 paper (2024) — average ~80-200 retrieved abstracts per question; ChemCrow paper; Sakana AI Scientist; Insilico Pharma.AI white-paper; Recursion Phenom-1.
- **Notes:** Per *literature-grounded hypothesis* the corpus retrieved is huge (PubMed + patents + assays). Long-horizon: PaperQA2 / AI-Scientist runs reach 30-90 minutes. Highest reasoning-token budget of any domain because hypotheses must be defended.

### 6. Software engineering agents
- **Agent layer:** Coding agents and harnesses: Claude Code, OpenAI Codex / Codex CLI, Cursor Agent, Cognition Devin, GitHub Copilot Workspace, Replit Agent, Aider, OpenHands, SWE-agent.
- **Public anchors:** SWE-Bench Verified / Lite leaderboards (Anthropic, OpenAI, Cognition); Tau-Bench-Code; LiveCodeBench; OpenAI Codex 2025 blog (avg task ~7-12 min); Anthropic Claude 4 'Sonnet thinking' system card; Cursor agent traces.
- **Notes:** Repo + diff + test output drive input tokens. Reasoning budget is the *highest sustained CoT load* among production agents (plan → patch → run tests → repair). Bench task durations 5-45 min; production background tasks regularly hit 1-2 hours.

### 7. Complex enterprise ops (IT/SRE, finance ops, compliance)
- **Agent layer:** Ticket / incident / close-the-books / control-test agents: ServiceNow Now Assist, Workday Illuminate, MSFT Copilot for Service, PagerDuty AIOps, Glean, Moveworks, Trullion, Kira/ Harvey for compliance, BlackLine Studio360.
- **Public anchors:** Tau-Bench (Sierra) — telecom / retail / airline workflows; OSWorld; ServiceNow Now Assist deflection metrics (Q3-Q4 FY25); Workday Illuminate disclosures; PagerDuty AIOps; Glean usage data.
- **Notes:** High-volume, narrow-scope. Input dominated by ticket body + top-k KB articles + user/asset context. Reasoning short — policies are deterministic, agents mostly route + draft.

## Reading the numbers

- **Ads ranking & HFT show small token counts on purpose.** Core ranking and execution are not LLMs. The numbers reflect the *LLM-agent layer* that wraps them (campaign briefs, research/strategy copilots), not the production scoring or execution path.
- **Software-engineering agents carry the highest sustained CoT load.** Verifiable signal at every step (tests pass / fail) makes long plan→patch→test→repair loops economic. Anthropic Claude 4 and OpenAI o3/o4 system cards confirm code workloads dominate thinking-token consumption.
- **Drug-discovery research agents have the longest duration.** Literature-grounded hypothesis generation (PaperQA2, AI-Scientist, ChemCrow) retrieves 80-200+ documents and runs 30-90 min per question. Wet-lab feedback loops are *not* counted in these numbers.
- **Cybersecurity SOC agents are the highest-volume non-code use case** of multi-step agentic loops in production today (CrowdStrike Charlotte AI, MSFT Security Copilot). Token mix skews toward *retrieved context* (alert + IOCs + endpoint timeline) over CoT.
- **Caveats.** Reasoning-token budgets are now configurable (GPT-5 Thinking 'minimal'/'medium'/'high', Claude 4 thinking-budget). These ranges will compress as orchestrators learn to spend CoT only when EV is positive.
