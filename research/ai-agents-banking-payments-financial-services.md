# Equity Research Note

## Impact of OpenAI and Anthropic (Claude) Co-Working AI Agents on Banking, Payments, and Financial Services

**Sector:** Financial Services / Enterprise Software / Applied AI
**Authors:** Equity-research repository
**Date:** April 2026
**Document type:** Thematic research note
**Status:** For discussion — not investment advice

---

## 1. Executive summary

The rapid maturation of agentic AI platforms from OpenAI (GPT‑4.1 / GPT‑5 family, Assistants/Responses API, "Operator" computer-use agents) and Anthropic (Claude 3.5/3.7/4 Sonnet & Opus, Claude "Computer Use", Model Context Protocol) is moving generative AI in financial services from a productivity tool into an **execution layer**. Banks, card networks, payment processors, and asset managers are starting to deploy multi‑model "co‑working" agent stacks where OpenAI and Claude models are orchestrated together — typically routed by task, latency, cost, and risk profile — rather than chosen exclusively.

For BFSI (Banking, Financial Services & Insurance) we see five concrete near‑term impacts:

1. **Cost-to-serve compression** in retail banking contact centers and mid/back‑office operations of 20–40% within 24–36 months for early adopters.
2. **Re‑pricing of process labor** in KYC/AML, dispute & chargeback handling, payment exception management, and loan documentation — historically labor‑intensive cost centers.
3. **New agentic payment rails** (LLM‑driven checkout, "agent commerce" via Visa, Mastercard, Stripe, PayPal agent toolkits) creating both incremental transaction volume and a new class of fraud/authorization risk.
4. **Rising spend on AI infrastructure and governance** — model gateways, guardrails, observability, evaluation, model risk management (SR 11‑7 / EU AI Act compliance) — creating a durable software/services TAM.
5. **Strategic optionality** for incumbents that build a **multi‑model abstraction layer** (OpenAI + Anthropic + open‑weights) versus single‑vendor lock‑in.

We view this as a structurally positive theme for: AI infrastructure providers, hyperscalers (MSFT, AMZN, GOOG), card networks (V, MA), payment platforms (PYPL, ADYEN, FIS, FISV/FI), top‑tier banks with scale data assets (JPM, BAC, GS, MS), and AI‑native fintech enablers. We see structural pressure on: BPO/outsourcing labor arbitrage models, low‑differentiation neobanks, and fragmented mid‑market core banking vendors that fail to expose agent‑ready APIs.

---

## 2. What "OpenAI + Claude co‑working" actually means

The phrase covers several deployment patterns observed across Tier‑1 banks and large fintechs:

- **Multi‑model routing.** A request is dispatched to GPT‑class or Claude‑class models depending on cost, context‑window need (Claude's long context for 200k+ token documents), structured tool‑use reliability (OpenAI function calling), regulated‑content sensitivity, or latency.
- **Adversarial / dual‑model verification.** One model drafts (e.g., a credit memo, a SAR narrative, a customer email) and a second model from the *other* vendor critiques, redacts PII, or independently re‑derives the answer. This pattern is increasingly required by internal model risk management (MRM) committees because it materially reduces single‑vendor failure modes.
- **Agent‑to‑agent (A2A) workflows.** Anthropic's **Model Context Protocol (MCP)** and OpenAI's **Responses / Agents API** are converging on shared tool‑use semantics, enabling a Claude planner to call tools also exposed to a GPT executor and vice versa.
- **Computer‑use co‑pilots.** OpenAI "Operator" and Anthropic "Computer Use" can drive legacy thick‑client banking and treasury applications (FIS IBS, Fiserv DNA, Temenos T24, Murex, Calypso) that lack modern APIs — a critical unlock given how much BFSI workflow still lives inside green‑screen and Win32 software.

The strategic point: BFSI buyers are explicitly hedging vendor risk. None of the global systemically important banks (G‑SIBs) we have observed are willing to put their entire agentic stack on a single foundation‑model provider, both for resilience and for negotiating leverage.

---

## 3. Where agents are being deployed in BFSI

### 3.1 Retail and commercial banking

| Use case | Typical agent pattern | Quantified impact (early adopters) |
|---|---|---|
| Tier‑1/2 customer support | Voice + chat agent with retrieval over policy & account data; human handoff on low confidence | 30–55% containment, 20–35% AHT reduction |
| Dispute & chargeback intake | Agent gathers evidence, classifies under Reg E / Reg Z / scheme rules, drafts response | 40–60% cycle‑time reduction |
| KYC refresh & periodic review | Agent re‑pulls documents, screens sanctions/PEP, summarizes risk delta | 50–70% analyst time saved on low‑risk segments |
| Commercial loan origination | Agent extracts covenants from credit agreements, builds spreading models, drafts memo | 30–50% RM/credit analyst productivity uplift |
| Wealth client servicing | Agent prepares meeting prep packs, post‑meeting notes, suitability checks | 5–10 hours/week per advisor reclaimed |

### 3.2 Payments and card

- **Authorization optimization & fraud orchestration.** Models reason over enriched transaction context (merchant, device, behavioral) to challenge or step‑up; LLMs are now used as a *meta‑policy* layer above traditional gradient‑boosted fraud models.
- **Disputes and chargebacks at scale.** Visa, Mastercard, and large issuers (JPM, Citi, Capital One) are piloting agentic dispute resolution that drafts compelling evidence, files via VROL/MasterCom, and tracks lifecycle.
- **Agentic commerce / "agent‑initiated payments".** Visa Intelligent Commerce, Mastercard Agent Pay, PayPal's Agent Toolkit, and Stripe's Agent SDK expose tokenized payment credentials directly to OpenAI/Claude agents. Implication: a non‑trivial share of e‑commerce GMV (we model 5–15% by 2030 in developed markets) may be initiated by AI agents on behalf of consumers and SMBs, reshaping interchange, MDR, and consumer acquisition economics.
- **Cross‑border and B2B payments.** Agents reconcile invoices, classify SWIFT MT/MX messages, resolve sanctioned‑party hits, and trigger payment repair — a high‑labor area for SWIFT gpi participants and correspondents.

### 3.3 Capital markets and asset management

- **Research workflow.** Claude's long context is being used for multi‑filing 10‑K/Q comparison; OpenAI's tool‑use is preferred for code‑interpreter‑driven model building. Co‑working setups have analysts produce first‑pass notes in hours rather than days.
- **Trading desk co‑pilots.** Read‑only agents summarize order flow, news, and risk; execution authority remains with humans for now (regulatory and best‑execution constraints).
- **Operations.** Trade break investigation, corporate actions interpretation, and reconciliation across custodians.

### 3.4 Insurance

- **Underwriting triage** (especially commercial specialty), **claims FNOL** with computer‑use agents driving legacy claims systems, and **subrogation** are leading verticals.

### 3.5 Risk, compliance, and finance

- **AML transaction monitoring alert disposition.** Agents draft Level‑1 disposition narratives and SAR drafts; humans sign off. Productivity gains of 3–5x for L1 analysts are being reported in pilots at top‑tier banks.
- **Regulatory change management.** Continuous ingestion of CFPB, OCC, FCA, ECB, MAS, RBI publications with agentic mapping to internal control libraries.
- **Internal audit and SOX.** Agent‑generated walkthroughs and control test sampling.

---

## 4. Why the multi‑model "co‑work" pattern matters economically

1. **Negotiating leverage on inference cost.** Token prices have fallen ~70–90% YoY at the frontier tier. Banks that build a model‑agnostic gateway can re‑route workloads on every quarterly price update, capturing the savings rather than handing them to the vendor.
2. **Resilience under regulatory scrutiny.** US prudential regulators (OCC, Fed) under SR 11‑7 and the EU AI Act's high‑risk classification of credit scoring and biometric ID expect demonstrable model risk controls. Dual‑model setups give MRM teams a defensible challenger‑model story out of the box.
3. **Capability differentiation by task.** Empirically, OpenAI's o‑series reasoning models lead on structured numerical reasoning; Claude leads on long‑context document synthesis and tool‑use steerability for regulated content. A co‑working architecture exploits both.
4. **Data residency and deployment flexibility.** Both vendors are now available via Azure (OpenAI), AWS Bedrock and GCP Vertex (Claude), enabling the same workload to run in‑region across multiple clouds — important for GDPR, DORA, and APAC data‑sovereignty regimes.

---

## 5. Market sizing (directional)

We frame the BFSI AI agent opportunity in three layers:

- **L1 — Foundation model inference & fine‑tuning spend by BFSI:** ~$8–12bn in 2026, growing to $40–60bn by 2030 (35–45% CAGR).
- **L2 — Agent platforms, orchestration, governance, and observability software:** ~$6–9bn in 2026 → $30–45bn by 2030.
- **L3 — Services & systems integration (Big 4, Accenture, TCS, Infosys, Cognizant, EPAM, Capgemini, GlobalLogic):** ~$25–35bn in 2026 → $80–120bn by 2030, with margin pressure as agents automate the lower tiers of delivery labor.

These are top‑down estimates triangulated from public IT spend disclosures by the top 25 global banks, hyperscaler BFSI vertical commentary, and announced agent platform deal sizes. Sensitivity is high; treat as scenario inputs, not forecasts.

---

## 6. Investment implications

### 6.1 Likely beneficiaries

- **Hyperscalers / model distribution:** Microsoft (Azure OpenAI exclusivity for enterprise distribution; GitHub/Foundry agent tooling), Amazon (Bedrock as the dominant Claude distribution channel and primary Anthropic investor), Alphabet (Vertex + Gemini, hosting third‑party models including Claude).
- **Card networks:** Visa and Mastercard, via agent commerce rails, tokenization, and dispute services. Network effects deepen as agents standardize on their SDKs.
- **Payments platforms:** Stripe (private), Adyen, PayPal (Agent Toolkit + Braintree), Block. Differentiation via developer‑grade agent SDKs.
- **Core processing & issuer services:** Fiserv, FIS, Global Payments — to the extent they expose agent‑ready APIs ahead of disintermediation.
- **AI‑native infra:** model gateways (Kong, Cloudflare AI Gateway), vector / retrieval (MongoDB Atlas Vector, Elastic, Snowflake Cortex, Databricks), evaluation & observability (Datadog LLM Observability, Splunk, New Relic), guardrails (private: Guardrails AI, Lakera).
- **Tier‑1 banks with scale:** JPMorgan (LLM Suite, IndexGPT, Quest IndexGPT), Goldman Sachs (GS AI Platform), Morgan Stanley (AI @ Morgan Stanley with OpenAI), Bank of America (Erica + internal Copilot), Capital One (data‑native posture). Operating leverage from automation flows disproportionately to scale players.
- **Specialist BFSI software with proprietary data moats:** Intuit, S&P Global, MSCI, Moody's, FactSet, LSEG — agent demand for high‑quality structured data raises the value of curated datasets.

### 6.2 At‑risk

- **Labor‑arbitrage BPOs** in voice support, claims processing, and back‑office finance (Concentrix, Teleperformance, parts of WNS/Genpact) — partially offset for those that pivot to AI orchestration services.
- **Undifferentiated neobanks** whose only edge was a better app — agents commoditize the UX layer.
- **Mid‑market core banking vendors** without modern API surface area.
- **Standalone fraud point‑solutions** that cannot integrate as tools into agent stacks.

### 6.3 Wildcards

- **Anthropic's commercial trajectory** (Amazon's continued capital commitment, potential public listing path) and **OpenAI's restructuring/governance evolution** are the largest single‑name risks to the thesis.
- **Open‑weights models** (Meta Llama, Mistral, DeepSeek, Qwen) deployed on‑prem could compress L1 spend faster than expected, especially for EU and APAC banks with strict data‑sovereignty needs.

---

## 7. Risks

1. **Regulatory.** EU AI Act high‑risk obligations (Article 6 + Annex III), CFPB Section 1033 open banking rules, NYDFS Part 500 cybersecurity, OCC/Fed/FDIC interagency guidance on third‑party AI, FFIEC AIO, MAS FEAT, RBI AI framework. Non‑compliance carries fines up to 7% of global turnover under the EU AI Act.
2. **Model risk and hallucination in regulated outputs.** Particularly acute for credit decisioning, suitability, and disclosures.
3. **Fraud and prompt‑injection on agentic payments.** Agents holding payment credentials are a new attack surface. Expect issuer‑side liability debates analogous to early e‑commerce CNP fraud.
4. **Concentration risk.** OpenAI's dependency on Microsoft/Azure capacity and Anthropic's on AWS create correlated outage risk for "diversified" multi‑model stacks.
5. **Data leakage and IP.** Training‑data provenance lawsuits against frontier labs could reshape commercial terms and indemnities BFSI buyers rely on.
6. **Talent and change management.** The binding constraint for most banks is not model quality — it is the ability to refactor processes, retire legacy systems, and re‑skill staff fast enough to capture the savings.

---

## 8. What to watch (catalysts, next 12 months)

- Quarterly disclosure of agent‑driven productivity by JPM, BAC, GS, MS, C in earnings calls.
- Visa and Mastercard updates on agentic commerce volumes and authorization economics.
- EU AI Act high‑risk system enforcement timelines and the first material fine.
- Microsoft, Amazon, Google capex guidance and BFSI vertical commentary.
- Anthropic and OpenAI enterprise revenue disclosures, pricing actions, and any IPO/restructuring news.
- Emergence of standardized **agent identity, attestation, and consent** frameworks (work in progress at FIDO Alliance, W3C, OpenID Foundation, and the card networks).

---

## 9. Bottom line

OpenAI and Claude are no longer competing only for share of the foundation‑model market in BFSI; they are increasingly **co‑deployed** as complementary components of bank‑grade agentic systems. The combination accelerates automation in the most labor‑heavy BFSI processes, creates a new "agentic payments" growth vector for the card networks and payment platforms, and entrenches the hyperscalers as the primary distribution channel. We view the multi‑model agent thesis as one of the most concrete and near‑dated AI monetization stories in financial services, with the largest economic value accruing to (i) the rails that agents transact over and (ii) the institutions whose scale lets them industrialize the savings.

---

## 10. Disclaimer

This document is a thematic research note prepared for educational and discussion purposes within the `Equity-research` repository. It is **not investment advice**, not a solicitation to buy or sell any security, and does not constitute a recommendation. Forward‑looking statements reflect the author's views at the time of writing and may change without notice. Readers should perform their own due diligence and consult a licensed financial advisor before making investment decisions. Company and product names are trademarks of their respective owners.
