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
