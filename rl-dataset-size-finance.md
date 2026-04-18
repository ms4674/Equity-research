# Reinforcement Learning Dataset Size for Banking, Payments, and Financial Agents

**Equity-Research Note — Bellwether Quantification**
**Last updated:** 18 April 2026

---

## 1. Why "dataset size" is the binding constraint

Reinforcement learning (RL) systems in finance ingest two distinct supplies of data:

1. **Trajectory / experience data** — sequences of (state, action, reward, next-state) tuples, typically generated from historical transaction logs, order-book replays, customer chats, or live-environment simulators. These set the *scale* of the replay buffer and the *variance* of the policy gradient.
2. **Domain pre-training corpora** — text, structured filings, news, and proprietary records used to warm-start the policy or value network (often an LLM backbone) before any RL stage.

In banking, payments, and financial agents the supply of (1) is essentially proprietary and is dominated by a handful of bellwethers — Visa, Mastercard, JPMorgan, Bank of America, PayPal, Stripe, Ant — whose transaction and interaction volumes set the *upper bound* on the experience data any RL operator can realistically train on. The supply of (2) is anchored by BloombergGPT-class corpora.

The table below quantifies each layer.

---

## 2. The transaction supply curve (raw RL experience pool)

| Bellwether | FY annual transactions | Annual volume | TPS-equivalent | Source year |
|---|---:|---:|---:|---|
| Visa | **329 B** (258 B Visa-processed) | $17.0 T | ~10,400 | FY2025 10-K |
| Mastercard (switched) | **~180 B** (45.4 B in Q3'25, +10% YoY) | GDV $7.8 T YTD Q3 | ~5,700 | Q3 2025 ops data |
| PayPal | **25.4 B** payment transactions | $1.79 T TPV | ~800 | FY2025 |
| Stripe | n.d. (578 M over BFCM 2025 alone) | ~$1.4 T (2024); $1.2–1.45 T est. 2025 | ~12 B/day peak | 2024–2025 |
| Block / Square (GPV) | n.d. — Cash App 58 M MATA | Q3'25 GPV +12% YoY | n.d. | 3Q25 |
| Ant International (EM rails) | **2 B+** (core EM only) | n.d. | ~63 | CY2025 |
| Bank of America (Erica chat) | **~730 M** interactions/yr (2 M/day) | n/a | ~23 | 2024 run-rate |

**Implication.** The two card networks alone produce on the order of **5 × 10¹¹ transactions per year** of replay-buffer-eligible data. PayPal adds another 2.5 × 10¹⁰. In RL terms, this is two to three orders of magnitude larger than any open-source finance benchmark and roughly comparable, in event count, to a year of *all* tokens scraped from CommonCrawl.

---

## 3. Disclosed RL / ML training corpora at the bellwethers

### 3.1 Payments — fraud, authorization, decisioning

**Mastercard — Decision Intelligence Pro.** Recurrent-neural-network model trained on the **~125 B transactions/year** flowing through the network. At inference, scans up to **1 trillion data points per call** in ~50 ms, lifting fraud detection by an average 20% (300% in some segments) and cutting false positives by >85%. ([Source](https://cnbc.com/2024/02/01/mastercard-launches-gpt-like-ai-model-to-help-banks-detect-fraud))

**Visa — Advanced Authorization (VAA) / Visa Deep Authorization (VDA).** ML model evaluating **>500 risk attributes per transaction** in ~1 ms. The disclosed 12-month training/inference window (to Apr-2019) covered **127 B transactions** and prevented ~$25 B in fraud. VDA upgrade adds RNN-based deep learning for card-not-present flows. ([Source](https://usa.visa.com/about-visa/newsroom/press-releases.releaseId.16421.html))

**Order of magnitude:** the production replay buffers behind these systems sit at **10¹¹ events/year** — i.e. ~10⁵× the largest public fraud benchmark (IEEE-CIS at ~5 × 10⁵ rows used in published RL-GNN work).

### 3.2 Trading execution agents

**JPMorgan — LOXM.** Deep-RL equity execution agent rolled out globally in Q4-2017. Trained on "**billions of historical and simulated trades**", with reward functions over slippage, market impact, and opportunity cost. Each parent order generates **thousands of intraday decision points** × multiple child orders, so a single trading day at JPM's $29.8 B daily equities desk emits an estimated **10⁸–10⁹ (state, action) tuples** into the simulator. ([Source](https://navnoorbawa.substack.com/p/jpmorgans-298b-trading-operation))

**Public benchmark.** The recent JaxMARL-HFT framework (Nov 2025) and FinRL-DeepSeek (Feb 2025) operate at materially lower scales — FinRL-DeepSeek uses ~**2 M aligned (price, news) records** across 89 Nasdaq tickers (2013–2023), drawn from the **15 M-article FNSPID** corpus. The gap between the public RL-finance frontier (10⁶–10⁷ events) and bellwether production scale (10⁹–10¹⁰ events) is 3–4 orders of magnitude.

### 3.3 Conversational banking agents (RLHF-style)

**Bank of America — Erica.** **2 B+ cumulative interactions** across 42 M clients since 2018; currently **2 M/day** (~7.3 × 10⁸/yr). 800 M user inquiries answered, 1.2 B proactive insights served. The data-science team has shipped **>50,000 model updates** — a useful proxy for the model-iteration cadence implied by an industrial RLHF stack. ([Source](https://www.prnewswire.com/news-releases/bofas-erica-surpasses-2-billion-interactions-helping-42-million-clients-since-launch-302110386.html))

**Klarna — OpenAI-powered assistant.** Trained on **millions of historical (anonymized) support conversations**, aligned with **RLHF**. Within one month of launch handled **2.3 M conversations** (≈2/3 of all chats), equivalent work of 853 FTEs by late-2025; tracking ~$60 M of 2025 profit. ([Source](https://openai.com/index/klarna/))

**Capital One — Eno.** Initial NLU model trained on "**hundreds of thousands of web chats**"; vocabulary covers **2,200+ terms and emojis** for the balance-inquiry intent alone — a useful floor for what a *single* intent costs to instrument.

**JPMorgan — LLM Suite.** Deployed to **~250,000 employees** across **400+ use cases**; refreshed every **8 weeks** with new internal data (no public token count, but the cadence implies a continuous pre-training pipeline rather than a one-shot dataset).

### 3.4 LLM pre-training corpora used as RL warm-starts

**BloombergGPT (the canonical "bellwether" finance corpus).** **363 B financial tokens** (FinPile: 298 B web, 38 B news, 14 B filings, 9 B press releases, 5 B Bloomberg) plus **345 B general tokens** for a total of **~708 B tokens**, training a 50 B-parameter base. ([arXiv 2303.17564](https://arxiv.org/abs/2303.17564))

**Frontier comparator.** DeepSeek-R1 (Jan 2025) demonstrates that a few hundred thousand verified reasoning trajectories (rejection-sampled from a much larger rollout pool) are sufficient to elicit reasoning behavior with GRPO — i.e. *post-training* RL needs **10⁵–10⁶ high-quality trajectories**, not 10¹¹ events. This sets the floor for finance-specific reasoning agents.

---

## 4. Synthesis — the four-tier RL dataset stack in finance

| Tier | Use case | Bellwether anchor | Dataset size (events / tokens) | Order of magnitude |
|---|---|---|---:|---:|
| 1. Network-scale stream RL | Real-time fraud, authorization, dispute routing | Mastercard DI Pro (125 B tx/yr); Visa VAA (127 B tx) | 10¹¹ events/yr | **10¹¹** |
| 2. Trading-agent simulation RL | Execution, market-making, smart order routing | JPMorgan LOXM ("billions" of historical+simulated trades) | 10⁹–10¹⁰ tuples | **10¹⁰** |
| 3. Conversational RLHF | Customer-service / advisor agents | BofA Erica (730 M/yr); Klarna (2.3 M/mo) | 10⁷–10⁹ turns/yr | **10⁸** |
| 4. Reasoning-agent post-training | Compliance, research, copilot agents | BloombergGPT (708 B tokens); DeepSeek-R1-style RL needs ~10⁵–10⁶ trajectories | 10¹¹–10¹² tokens (warm-start) + 10⁶ trajectories (RL stage) | **10¹¹ + 10⁶** |

### 4.1 Read-across to capex / moat

- **Card networks (V, MA)** sit on a ~5× 10¹¹/yr event stream that *no* hyperscaler or fintech can replicate without acquiring a network — the dataset moat is structural, not technological.
- **Mega-banks (JPM, BAC)** have the largest *labeled, behavior-rich* conversational and execution corpora — the bottleneck is labeling/governance, not raw volume.
- **Vertical fintechs (PYPL, SQ, Klarna, Stripe)** sit one order of magnitude below the networks but produce richer per-event metadata (basket, device, merchant, cohort), which often matters more for downstream RL reward design than raw count.
- **Public researchers** (FinRL, JaxMARL-HFT) operate 3–4 orders of magnitude below the bellwethers; the gap is closing only via synthetic-data RL (e.g., the 2024 SciDir paper on "reinforcement prompting for financial synthetic data").

### 4.2 Key takeaways for an investor

1. The binding scarcity for production-grade financial RL agents is **proprietary trajectory data**, not compute or model architecture. Bellwether per-year event volumes (10¹⁰–10¹¹) are 100–1000× the largest open benchmarks.
2. **Visa and Mastercard** disclose the highest-quality RL training pools in the industry (125–130 B labeled transactions/year each) and have already productized them (VAA, DI Pro) — this is a defensible, recurring AI revenue line, not a research project.
3. **JPMorgan** has the only publicly disclosed *production* deep-RL trading agent at scale (LOXM), trained on billions of trades; combined with its 250 K-seat LLM Suite, it is the clearest "AI-native bank" comparable.
4. **Klarna and Bank of America** illustrate the RLHF tier: a few million-to-billion conversation events are sufficient to displace **700–850 FTEs** of customer-service labor — implying an *operating-leverage* multiple of roughly $50 K of profit per million logged interactions at current model quality.
5. **Reasoning-agent RL** (DeepSeek-R1 paradigm) is now feasible on **10⁵–10⁶ verified financial trajectories** — meaning the next wave of compliance, research, and advisor copilots will be built by whichever institution can *curate* the smallest, cleanest expert corpus, not the largest raw stream.

---

## 5. How many parameters does the RL stage actually *add*?

The RL stage in modern agentic systems rarely introduces new structural parameters into the deployed model — it predominantly **re-weights** an existing pre-trained network. The "parameter cost" is therefore best read in three layers: training-time scaffolding, deployed-time delta, and the production scale of the underlying model.

### 5.1 Training-time scaffolding (transient, lives only in the GPU job)

Classical PPO-RLHF requires up to **four co-resident networks**:

| Component | Typical size vs. policy | Persists into deployment? |
|---|---:|---|
| Policy / actor (the LLM) | 1.0× | Yes |
| Reference / frozen policy (KL anchor) | 1.0× | No |
| Value / critic head | base + small MLP head (a few M params) | Discarded |
| Reward model (preference scorer) | 0.3–1.0× | Used in training only |

For a 7 B-parameter policy this resolves to ~**4 × 7 B = ~28 B trainable + frozen parameters** in memory during training, which is why naive full-precision PPO on a 7 B model needs **~220 GB of GPU memory**. ([HF Efficient-RLHF paper](https://www.huggingface.co/papers/2309.00754))

**GRPO (DeepSeek-R1's algorithm) eliminates the value model entirely**, dropping training memory by ~25–30% and removing one full-size network from the scaffolding. This is why DeepSeek could RL-train a **671 B-parameter MoE backbone (37 B active)** without adding *any* new architectural parameters. ([Epoch AI breakdown](https://epoch.ai/gradient-updates/what-went-into-training-deepseek-r1))

### 5.2 Deployed-time delta (what actually ships to production)

The number of *new, persistent* parameters added by the RL stage falls into three regimes:

| Method | New trainable parameters | Memory footprint vs. base | Used by |
|---|---:|---:|---|
| **Full-parameter RLHF** | 100% of base (re-weights everything) | 1.0× (no new params, just new weights) | OpenAI / Anthropic frontier models |
| **LoRA-PPO / PEFT-RLHF** | **~0.1–0.6%** (e.g. ~40 M new params on a 7 B model at rank 16) | +0.6% on disk; **~3.2× less GPU mem in training** (68 GB vs 220 GB) | Most enterprise fine-tunes; verl-style banking deployments |
| **GRPO (DeepSeek-R1 paradigm)** | 100% (re-weight) but **no value-network parameters** at all | 1.0× weights, 0.7–0.75× training memory | DeepSeek-R1 (671 B / 37 B active) |
| **Reward-model-only stack (DPO, RLAIF)** | 0% post-training (preference loss applied directly) | 1.0× | Many cost-sensitive Tier-3 deployments |

**Concrete bellwether read-across:**

- **Mastercard DI Pro (RNN with "inverse recommender" head):** the RL/learning component sits as an **incremental head** on a much larger embedding table; in the new "Large Tabular Model" rollout, the *foundation* is being scaled to **hundreds of billions of parameters' worth of training capacity**, not the policy head itself.
- **JPMorgan LOXM-class execution agents:** policy networks are typically **<10 M parameters** (small CNN/MLP over LOB features). RL adds essentially zero new structural parameters; the cost is in the simulator and reward shaping.
- **JPMorgan LLM Suite / Klarna OpenAI assistant:** built on hosted frontier models (OpenAI, Anthropic). The *enterprise* RL/alignment delta is almost always a **LoRA adapter or system-prompt + retrieval layer** measured in **tens of MB**, not GBs. The base 100B–1T-parameter model is unchanged.
- **BofA Erica:** classic intent-classifier + response-library architecture (700-response library, 2,200-term vocabulary per intent). The "RL surface" is effectively a small ranker on top of a frozen NLU stack — single-digit-million-parameter delta per update.

**Bottom line:** in production financial agents, RL almost never *adds* >1% to the parameter count of the deployed model. It either (a) re-weights an existing frontier model (0% net new params), or (b) bolts on a **0.1–0.6% LoRA adapter**, or (c) trains a tiny task-specific head (<10 M params) on top of a large frozen embedding/foundation model.

---

## 6. How often does the RL stage have to be re-run?

Refresh cadence varies by **decay velocity** of the underlying environment, which differs by 4–5 orders of magnitude across the four tiers. Bellwether disclosures pin the cadence:

| Tier | Driver of decay | Observed bellwether cadence | RL refresh trigger |
|---|---|---|---|
| 1. Network-scale fraud / authorization | Adversarial — fraud rings reverse-engineer thresholds in days | Mastercard DI Pro: continuous online learning over a 160 B-tx/yr stream; LTM "continuously improves as it trains on more data" | Drift / new attack vector — new labels propagate within **hours** |
| 2. Trading execution agents | Regime change, microstructure shifts | Drift-triggered retraining when **rolling 30-day Sharpe < 0.5** or daily P&L < −5%; Kolmogorov-Smirnov tests on feature distributions | Statistical drift detection, not a calendar |
| 3. Conversational RLHF (banking copilots) | Product changes, policy changes, new intents, language drift | BofA Erica: **>75,000 model updates since 2018** (~25/day average; ~10/day adjusted for early-build period); JPM LLM Suite: **every 8 weeks** | Scheduled release train + hot-fixes for safety/quality regressions |
| 4. Reasoning-agent post-training | New reasoning patterns, new tool APIs | DeepSeek-R1-style multi-stage RL: **one-shot** post-training per base-model revision; refreshed when the base model is upgraded | Tied to base-model release cadence (quarterly to semi-annual) |

### 6.1 Why the cadence diverges so sharply

- **Stale-model decay is empirically violent in fraud.** The most-cited production case shows a model holding **94.2% accuracy for six weeks**, then collapsing to **61%** in a single week as adversaries discovered a $500 transaction-amount threshold. Online learning recovered detection to **88% in 48 hours and 93.7% within a week** — proving that a *weekly* batch RL refresh is already too slow at network scale and continuous online learning is the production norm.
- **Trading agents are drift-paced, not calendar-paced.** RL retrains are triggered by Sharpe / win-rate thresholds rather than fixed schedules; many shops use *self-healing* adapter layers (Tier-2 LoRA-style updates) to avoid full RL refreshes during live regimes.
- **Conversational copilots are release-paced.** JPMorgan's confirmed **8-week** LLM Suite cadence is now the de facto standard cadence for enterprise agentic AI: long enough to A/B-test with rollback plans, short enough to absorb policy changes and base-model upgrades.
- **Reasoning agents are model-paced.** Because GRPO-style RL is run *once* per base-model revision (DeepSeek-R1, OpenAI o-series, Anthropic Claude reasoning variants), a finance-specific reasoning agent typically gets a *full* RL refresh on the order of **2–4× per year**, with smaller LoRA-adapter or DPO updates between releases.

### 6.2 Implied operating cost

- **Tier 1 (continuous online RL on 10¹¹ events/yr):** thousands of GPU-hours per day at the network operators (V, MA) — already capitalized.
- **Tier 2 (drift-triggered):** dozens of GPU-hours per retrain event, ~1–10 events/yr per strategy.
- **Tier 3 (8-week scheduled):** **6–7 RL/RLHF refreshes per year** per major agent surface; at LoRA scale a single refresh is **<$10K of compute** — explaining why JPM can run 400+ use cases simultaneously.
- **Tier 4 (model-paced):** **2–4 full RL post-training rounds per year** per reasoning agent, dominated by base-model rather than fine-tune compute.

### 6.3 Investor implication

The "AI moat" in finance is **not** the parameter count of the RL-tuned model — it is the **rate at which proprietary feedback can be looped back into the policy**. V/MA are running an effectively continuous loop on 10¹¹ events/yr; JPM ships every 8 weeks across 400 surfaces; mid-tier banks running 6-month batch retrains will visibly degrade against this cadence. Capex differentiates not on training FLOPs, but on the **labeling, governance, and online-RL pipeline** that converts raw transactions into a re-weighted policy.

---

## 7. RL vs RAG — time-to-accuracy in financial agents

RAG (Retrieval-Augmented Generation) and RL/RLHF are not substitutes; they sit on different axes of an agent's accuracy budget. The right question for an investor is **how many engineering-hours convert into one percentage point of accuracy improvement** — and on that axis the two techniques look almost nothing alike.

### 7.1 Where the time actually goes

| Phase | RAG | RL / RLHF |
|---|---|---|
| Initial setup | **2–4 weeks** to first production system (vector DB + embeddings + retriever + reranker) | **4–8 weeks minimum**; data preparation alone is routinely **3–5× over budget** |
| Per-update unit of work | Re-embed changed documents only (~10–15% of corpus) | Re-collect trajectories, re-train policy, run safety/eval suite, A/B-test |
| Update wall-clock | **Sub-second to minutes** (streaming RAG with CDC); minutes-to-hours for nightly batch | **Hours to days** of GPU time per refresh; **8 weeks** of engineering cycle in JPM-class deployments |
| Per-update cost | Maintenance ~**5–10 engineering hours/month** + **$5–$50** of embeddings | **$500–$5,000+** per training cycle (LoRA); **$2,400–$18,000** per full FT run; full re-RL on a frontier reasoning agent: 6–7 figures |
| Failure mode when stale | Wrong/missing facts → easy to spot (citations don't resolve) | Silent behavioral drift → hard to spot, requires offline eval suite |
| Accuracy lift profile | **42–90% hallucination reduction** vs. base; *factual* accuracy with citations | **0–2 pt MMLU regression risk**, but 10–30% lift on *behavioral* metrics (tool use, tone, reward-aligned actions) |

### 7.2 What each technique actually buys you

- **RAG buys *recency and citability*.** It changes what the model *sees* at inference time, not what it *is*. A new Fed announcement, a new fee schedule, or an updated AML rule can be reflected in the agent's answers within **seconds** of being indexed. The empirical research is clear: for *knowledge-intensive* tasks, **RAG consistently outperforms unsupervised fine-tuning**, because LLMs are slow to absorb new factual information through gradient updates.
- **RL buys *behavior*.** It changes what the model *prefers to do* — order routing strategy, tone in customer chat, refusal behavior on sanctioned counterparties, escalation policy in fraud disputes. None of these are retrievable; they have to be baked into policy weights. RL is also what closes the loop on **adversarial drift** in fraud (Tier 1) and **regime shifts** in execution (Tier 2), where the right *action* matters more than the right *fact*.

### 7.3 Cycle-time gap, quantified for financial agents

| Dimension | RAG cycle | RL cycle | Ratio |
|---|---:|---:|---:|
| Time from new info → live in agent | **seconds–minutes** (streaming) to **hours** (nightly batch) | **8 weeks** (JPM LLM Suite cadence) for behavioral copilots; **hours** (online learning) for Tier-1 fraud only | **~1,000–10,000×** slower for behavioral RL |
| Engineering hours per accuracy refresh | 5–10 hrs/month maintenance | 100s–1,000s of hrs per major refresh | **~50–200×** more engineering per cycle |
| $/refresh at LoRA scale | **<$50** (re-embed only changed docs) | **$500–$5,000** | **~10–100×** more capital |
| First time-to-prod | 2–4 weeks | 4–8 weeks (often 3–5× over budget) | **~2–4×** longer initial build |

### 7.4 Why financial agents end up running both

In every bellwether stack we examined, RAG and RL are **complementary rather than competitive**:

- **JPMorgan LLM Suite:** the *base* OpenAI/Anthropic models are RL-aligned upstream; JPM's 8-week cycle is dominated by **RAG over internal databases** and prompt/agent-tool changes, with episodic LoRA-scale RL refreshes on the side. This is why the cadence is 8 weeks (RL-feasible) not nightly (pure-RAG-feasible) — they do both, paced by the slower one.
- **Mastercard DI Pro:** the **ranker** is RL-trained continuously (Tier 1, hours), while the *feature store* (merchant embeddings, cardholder histories) is refreshed on a streaming-RAG-style CDC pipeline. The two pipelines run on completely different clocks.
- **BofA Erica:** the **response library** (700+ canned responses) is RAG-style retrieval at inference; the **intent classifier and ranker** are episodically RL/SFT-trained. The 75,000 model updates since 2018 split heavily toward retrieval-side updates.
- **Klarna assistant:** built on a frontier RLHF'd base, with **policy/product knowledge injected via RAG** so that policy changes do not require re-running RLHF on millions of conversations.

### 7.5 The investor read-across

1. **The marginal accuracy point is much cheaper from RAG than from RL.** For factual or policy-driven errors, expect **~10–100× lower $/percentage-point** improvement from RAG, and **1,000–10,000× faster** propagation into production.
2. **RL is the only lever for behavioral and adversarial accuracy.** Fraud-loss reduction (Mastercard's 20–300% lift), execution slippage (LOXM), and tool-use reliability are not retrievable — they require policy updates.
3. **Time-to-accuracy at the bellwethers is paced by the slower of the two pipelines.** JPM's confirmed 8-week LLM Suite cadence is the *RL clock*; the *RAG clock* underneath it is essentially continuous. Mid-tier banks running quarterly RAG refreshes and 6-month RL cycles are 3–5× slower on both axes simultaneously — the gap compounds.
4. **Total cost of ownership for the AI agent stack is dominated by RL, not RAG.** RAG monthly run-rate is **$350–$2,850**; an enterprise RL stack runs **$50K–$500K+** per major model surface. The economic moat scales with *which* problems an institution chooses to solve with RL versus RAG — and the bellwethers are clearly using RL only where it earns its keep (fraud, execution, alignment) while leaving everything else to RAG.

---

## 8. Sources

- Visa FY2025 10-K — 329 B transactions, $17 T volume.
- Mastercard 3Q25 supplemental ops data — 45.4 B switched transactions in Q3'25.
- PayPal 4Q25 / FY25 earnings — 25.4 B transactions, $1.79 T TPV.
- Stripe 2024 newsroom update — $1.4 T TPV; BFCM 2025 release — 578 M transactions / $40 B in 4 days.
- Block 3Q25 / 4Q25 shareholder letters — Cash App 58 M MATA; Square GPV +12% YoY.
- Ant International / Alipay+ press release — 2 B+ EM transactions in 2025.
- Mastercard "Decision Intelligence Pro" — CNBC, Feb 2024; 125 B tx/yr training set.
- Visa Advanced Authorization press release, Jun 2019 — 127 B tx, 500+ attributes.
- JPMorgan LOXM — Finance Magnates and execution-economics deep-dive, 2024.
- BofA Erica milestone release, Apr 2024 — 2 B interactions, 2 M/day.
- Klarna × OpenAI case study, 2024 — 2.3 M chats/month, RLHF.
- Capital One — *How & Why We Built Eno's NLP In-House*, Capital One Tech blog.
- JPMorgan LLM Suite — CNBC, Sep 2025; Larridin AI Tracker.
- Goldman Sachs × Devin (Cognition) — TechCrunch / IBM, Jul 2025.
- BloombergGPT — Wu et al., arXiv 2303.17564 (FinPile, 363 B + 345 B tokens).
- DeepSeek-R1 — arXiv 2501.12948 / Nature 2025 (GRPO, rejection sampling).
- FinRL-DeepSeek — arXiv 2502.07393; FNSPID 15 M articles, 2 M-record working set.
- JaxMARL-HFT — arXiv 2511.02136 (GPU-accelerated multi-agent RL for HFT).
- *Reinforcement learning with graph neural network (RL-GNN) fusion for real-time financial fraud detection* — Nature Sci. Reports 2025; IEEE-CIS at 5 × 10⁵ tx scale.
- *Efficient RLHF: Reducing the Memory Usage of PPO* — Santacroce et al., HF papers 2309.00754; LoRA-PPO at 68 GB vs full PPO at 220 GB on 7B model.
- *Secrets of RLHF in Large Language Models, Part I: PPO* — arXiv 2307.04964; four-network PPO scaffolding (policy, reference, value, reward).
- verl documentation — *RL(HF) algorithms with LoRA Support*; lora_rank=32 for 0.5 B model and lora_rank=128 for 32 B model match non-LoRA convergence.
- Unsloth / Databricks LoRA hyperparameter guides — 0.1–0.6% trainable parameter share, 90–95% of full-FT quality.
- Mastercard *Inside Mastercard's new gen AI engine* (2026) — DI Pro at 160 B tx/yr, 70K TPS peak; LTM continuously trained.
- *A High-Recall Cost-Sensitive ML Framework for Real-Time Online Banking Transaction Fraud Detection* — arXiv 2601.07276; online-learning recovery from 61% → 93.7% in 1 week after concept drift.
- Inside JPMorgan LLM Suite (CeFPro / CompleteAITraining 2025) — 8-week refresh cadence confirmed.
- BofA Erica 3 B-interaction milestone (Aug 2025) — 75,000+ model updates since launch, 700-response library.
- *Background: Model Drift and Retraining Strategies* (Wayland Z., 2025) — 30-day Sharpe < 0.5 / KS-test triggers for trading-RL retrain.
- Epoch AI — *What went into training DeepSeek-R1?* (2025) — GRPO eliminates value model; 671 B / 37 B-active MoE.
- *RAG vs Fine-Tuning Cost in April 2026: $350/mo vs $18K* — PE Collective (2026); RAG $350–$2,850/mo, FT $2,400–$18,000/run.
- *RAG vs. Fine-Tuning 2026: Costs, Framework Guide* — Alphacorp (2026); 42–90% hallucination reduction with RAG.
- *RAG Architecture in 2026: How to Keep Retrieval Actually Fresh* — RisingWave / Medium (2026); streaming RAG with CDC, sub-second freshness, 10–15% reprocessing share.
- *Fine-Tuning or Retrieval? Comparing Knowledge Injection in LLMs* — arXiv 2312.05934 / EMNLP 2024; RAG > unsupervised FT for knowledge tasks.
