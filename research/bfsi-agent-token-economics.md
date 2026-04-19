# BFSI AI Agents: Token Consumption & Average Task Duration

**Companion to:** [`data/bfsi_agent_token_economics.xlsx`](../data/bfsi_agent_token_economics.xlsx)
**Generator:** [`scripts/build_token_economics_workbook.py`](../scripts/build_token_economics_workbook.py)
**Date:** April 2026
**Status:** Educational / discussion only — **not investment advice**

---

## 1. What this note answers

For each of five high‑value BFSI agent use cases, what is a reasonable planning estimate for:

- **Input tokens** (system prompt + RAG context + user message + history)
- **Reasoning tokens** (OpenAI o‑series `reasoning_tokens` / Claude extended‑thinking)
- **Output tokens** (visible response)
- **Average task duration** (end‑to‑end agent wall‑clock per task)
- **Cost per task** at current public list prices

Use cases:

1. Retail customer service virtual assistant
2. Document analysis for loans (credit agreement / covenant extraction / memo)
3. Insurance claims correspondence drafting
4. Wealth advisor knowledge assistant
5. Fraud detection (LLM overlay — *not* real‑time card authorization)

The full numeric table (Low / Typical / High per metric, plus cost per task and per 1k tasks) is in the spreadsheet's `Token_Economics` sheet. A long‑form per‑use‑case detail view is in `Per_UseCase_Detail`. Reference list prices are in `Model_Prices`. Citations are in `Sources`.

---

## 2. Method

### 2.1 What "one task" means

| Use case | Task unit |
|---|---|
| Retail customer service VA | One *resolved customer conversation* (avg 3–5 turns + RAG over policy/account) |
| Loan document analysis | One *credit agreement processed end‑to‑end* (extract covenants + structured fields + draft memo section) |
| Insurance claims correspondence | One *outbound letter* to claimant / counterparty (FNOL ack, status update, decision letter) |
| Wealth advisor knowledge assistant | One *advisor query answered* (RAG over internal research + portfolio context) |
| Fraud detection (LLM overlay) | One *flagged transaction reviewed by an LLM* (case review / explainability / dispute evidence) |

### 2.2 How tokens are billed

Both OpenAI and Anthropic bill *reasoning / extended‑thinking tokens at the output rate*. So for cost purposes:

```
cost = (InputTokens × InputPrice + (ReasoningTokens + OutputTokens) × OutputPrice) / 1e6
```

### 2.3 Reference list prices used

Public list prices, USD per 1M tokens (see `Model_Prices` sheet for full table):

| Tier | Examples | Input | Output |
|---|---|---|---|
| Small/cheap (chat, drafting, classification) | Claude Haiku 3.5, GPT‑4o mini, GPT‑4.1 mini | $0.80 | $4.00 |
| Mid (frontier non‑reasoning) | GPT‑4o, GPT‑4.1, Claude Sonnet 4 | $3.00 | $15.00 |
| Reasoning / extended thinking | o3 / o4‑mini, Claude Sonnet 4 with thinking | $3.00 | $15.00 |
| Large reasoning (premium) | Claude Opus 4 thinking, GPT‑5 high‑reasoning | $15.00 | $75.00 |

Production deployments typically negotiate ~30–60% off list at scale, so the cost‑per‑task figures below should be read as a list‑price upper bound.

### 2.4 What the duration figure includes

`DurationSec` is the **end‑to‑end agent wall‑clock per task**, including retrieval, all model calls in the agent loop, tool calls, and any human‑in‑the‑loop wait baked into a "completed task." It is *not* raw model time‑to‑first‑token. For real‑time card authorization (sub‑100ms latency budget), LLMs are not used in the hot path; the fraud row covers the *post‑auth* LLM overlay.

---

## 3. Headline table (Typical values)

| # | Use case | Input | Reasoning | Output | Total tokens | Avg duration | $ / task | $ / 1k tasks |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| 1 | Retail customer service VA | 8,000 | 200 | 400 | 8,600 | 60s | $0.0088 | $8.80 |
| 2 | Document analysis for loans (per credit agreement) | 120,000 | 8,000 | 5,000 | 133,000 | 180s | $0.5550 | $555.00 |
| 3 | Insurance claims correspondence (per letter) | 8,000 | 500 | 600 | 9,100 | 30s | $0.0108 | $10.80 |
| 4 | Wealth advisor knowledge assistant (per query) | 20,000 | 2,000 | 800 | 22,800 | 30s | $0.1020 | $102.00 |
| 5 | Fraud detection LLM overlay (per case review) | 3,000 | 300 | 200 | 3,500 | 5s | $0.0165 | $16.50 |

The full Low / Typical / High range for every cell, plus duration ranges, is in the `Token_Economics` sheet of the workbook.

---

## 4. Per‑use‑case ranges (Low → High)

### 4.1 Retail customer service virtual assistant

| Metric | Low | Typical | High |
|---|---:|---:|---:|
| Input tokens | 2,000 | 8,000 | 25,000 |
| Reasoning tokens | 0 | 200 | 1,000 |
| Output tokens | 100 | 400 | 1,500 |
| Duration (sec) | 5 | 60 | 300 |
| **Cost / task ($)** | $0.0020 | $0.0088 | $0.0300 |

- **Anchors:** Klarna disclosed avg resolution time fell from 11 min to <2 min (~120s). BofA Erica >2bn cumulative interactions. Wells Fargo Fargo ~245m interactions in 2024.
- **Why input is much larger than output:** RAG over policy KB + customer account context dominates. Each retrieved chunk is typically 1–3k tokens; agents pull 2–6 chunks per turn × 3–5 turns.
- **Cheap model tier (Haiku 3.5 / GPT‑4o mini class) is sufficient** because (i) responses are short and bounded, (ii) hard cases escalate to humans, (iii) volume requires low unit cost.

### 4.2 Document analysis for loans (credit agreement)

| Metric | Low | Typical | High |
|---|---:|---:|---:|
| Input tokens | 30,000 | 120,000 | 400,000 |
| Reasoning tokens | 1,000 | 8,000 | 30,000 |
| Output tokens | 1,000 | 5,000 | 20,000 |
| Duration (sec) | 30 | 180 | 900 |
| **Cost / task ($)** | $0.12 | $0.555 | $1.95 |

- **Anchors:** Syndicated credit agreements are ~80–200 pages (~80–150k tokens). With amendments and ISDA schedules they can exceed 400k. JPMorgan COiN historically saved ~360,000 lawyer‑hours/yr on commercial loan agreements (pre‑LLM).
- **By 2–3 orders of magnitude the most expensive task per unit** in this analysis. Input dominates cost (~75–80%); reasoning tokens add another 20–40% on top of output.
- Long‑context models (Claude 200k+, Gemini 1M+) make single‑shot ingest practical and *reduce* total cost vs. chunk‑and‑map approaches because system‑prompt / instruction tokens stop being amortized across many calls.

### 4.3 Insurance claims correspondence drafting

| Metric | Low | Typical | High |
|---|---:|---:|---:|
| Input tokens | 2,000 | 8,000 | 30,000 |
| Reasoning tokens | 0 | 500 | 2,000 |
| Output tokens | 200 | 600 | 2,000 |
| Duration (sec) | 5 | 30 | 120 |
| **Cost / task ($)** | $0.0024 | $0.0108 | $0.0400 |

- **Anchors:** Allstate publicly disclosed ~50,000 GenAI‑drafted claim letters per day.
- Letters are highly templated (regulatory boilerplate + per‑claim variables). Input includes claim file extract + customer profile + applicable state‑law template; output is a single bounded letter.
- At ~$0.01/letter list, 50,000 letters/day ≈ **~$540/day or ~$200k/yr** in raw inference for the full Allstate disclosed volume — a rounding error vs. the labor it replaces.

### 4.4 Wealth advisor knowledge assistant

| Metric | Low | Typical | High |
|---|---:|---:|---:|
| Input tokens | 4,000 | 20,000 | 80,000 |
| Reasoning tokens | 200 | 2,000 | 10,000 |
| Output tokens | 300 | 800 | 3,000 |
| Duration (sec) | 5 | 30 | 120 |
| **Cost / task ($)** | $0.0173 | $0.102 | $0.435 |

- **Anchors:** Morgan Stanley AI @ MS Assistant (OpenAI) covers ~100k+ research documents; ~98% advisor‑team adoption. Typical query retrieves 5–15 chunks of 1–3k tokens each. Multi‑document synthesis ("compare two house views") drives the high end.
- Mid‑tier model (GPT‑4o / Claude Sonnet 4 class) is the right default — quality matters more than per‑query cost because the per‑user query rate is low (advisors, not consumers).

### 4.5 Fraud detection (LLM overlay)

| Metric | Low | Typical | High |
|---|---:|---:|---:|
| Input tokens | 500 | 3,000 | 15,000 |
| Reasoning tokens | 0 | 300 | 2,000 |
| Output tokens | 50 | 200 | 1,500 |
| Duration (sec) | 1 | 5 | 30 |
| **Cost / task ($)** | $0.0023 | $0.0165 | $0.0975 |

- **Real‑time card authorization is gradient‑boosted ML**, not LLMs (sub‑100ms latency budget). LLMs are used for the **post‑auth overlay**: case review, analyst‑facing explainability, dispute / chargeback evidence drafting, and scam‑message classification.
- Mastercard Decision Intelligence Pro and Visa AI use generative models in this overlay role.
- A bank running 10m post‑auth LLM reviews/month at typical settings ≈ **~$165k/month** in raw inference at list — a cost line that fits inside any large issuer's existing fraud‑ops budget.

---

## 5. Patterns that fall out of the numbers

1. **Input tokens dominate cost everywhere.** In every use case other than fraud‑overlay, input is 70–95% of the total tokens. That makes **prompt/RAG efficiency** the highest‑leverage cost lever, well ahead of model selection.
2. **Reasoning tokens are a meaningful but bounded surcharge.** They typically add 5–40% on top of output for hard tasks. They are *not* an order‑of‑magnitude cost item except on the very hardest reasoning loops (e.g., complex covenant extraction with cross‑references).
3. **Cost per task spans ~2.5 orders of magnitude across BFSI use cases**, from ~$0.009 (retail VA) to ~$0.55 (loan document analysis). Aggregating "AI inference cost" across BFSI without segmenting by use case will give a meaningless answer.
4. **Duration is decoupled from cost.** A 5‑second fraud overlay can cost more per task than a 60‑second customer‑service conversation, because model‑tier choice and input size matter more than wall‑clock.
5. **At list price, BFSI inference cost is typically 0.1–2% of the labor it displaces** for the use cases above. The economic constraint on BFSI agent rollout is rarely token cost; it is governance, integration, and change management.

---

## 6. What is *not* in this analysis

- **Real‑time card authorization scoring** is excluded from the LLM cost view because LLMs are not used in that path.
- **Embedding / vector index cost** for RAG corpora is excluded; it is one‑time and small per query (typically <5% of inference cost).
- **Fine‑tuning cost** is excluded; most BFSI deployments are RAG + prompting, not custom fine‑tunes.
- **Voice / ASR / TTS cost** for voice channel customer service is excluded; add roughly $0.005–$0.015 per minute of audio for transcription + synthesis on top of the LLM cost.
- **Discounted enterprise pricing** is excluded; Tier‑1 buyers typically negotiate 30–60% off list at committed volume.

---

## 7. Reproducing the workbook

```bash
pip install openpyxl
python3 scripts/build_token_economics_workbook.py
# -> writes data/bfsi_agent_token_economics.xlsx
```

Edit `USE_CASES` or `MODEL_PRICES` in the script to change the assumptions; cost columns recompute on the next run.

---

## 8. Caveats

- All token counts and durations are **modelled estimates** triangulated from public vendor disclosures, customer case studies, and observed bank pilots. They are intended as order‑of‑magnitude planning numbers, not contractual commitments.
- Public disclosures rarely give per‑task token counts directly; the values here are derived from disclosed *workflow shapes* (e.g., "AI handled 2/3 of chats with 2‑min average resolution") plus reasonable assumptions about retrieval depth, system prompt size, and response length.
- Prices change monthly. The `Model_Prices` sheet captures the prices used at the time of writing (early–mid 2026 era list prices); rerun the script with updated prices to refresh.
