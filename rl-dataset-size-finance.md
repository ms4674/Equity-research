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

## 5. Sources

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
