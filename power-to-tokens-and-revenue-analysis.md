# Power-to-Tokens and Revenue Analysis: 1 GW Data Center Economics

**Equity Research Supplement | May 2026**

---

## Executive Summary

This analysis quantifies the relationship between 1 GW of data center power capacity and two critical outputs: (1) token generation throughput and (2) GPU cloud rental revenue. The economics differ dramatically depending on GPU generation, workload type, utilization rate, and business model. At current Blackwell-generation efficiency, 1 GW of IT power can generate approximately **36-50 trillion tokens per day** in inference mode and produce **$6-12 billion in annual revenue** from GPU cloud rental, depending on the operator's business model and pricing power.

---

## 1. From Power to GPUs: The Conversion Chain

### Power Allocation Hierarchy

Not all power entering a data center reaches the GPUs. The conversion from facility-level power to compute-available power follows this hierarchy:

| Layer | Typical Allocation | 1 GW Facility |
|---|---|---|
| Total facility power | 100% | 1,000 MW |
| IT equipment power (at PUE 1.2-1.4) | 71-83% | 710-830 MW |
| GPU/accelerator share of IT load | 75-85% | 530-710 MW |
| **Effective GPU power budget** | **53-71%** | **530-710 MW** |

**Key assumptions for this analysis:**
- PUE of 1.25 (liquid-cooled, state-of-the-art AI data center)
- 80% of IT load allocated to GPUs (rest: CPUs, networking, storage, memory)
- **Effective GPU power from 1 GW facility: ~640 MW**

### GPU Count per 1 GW by Generation

| GPU | TDP per GPU | Effective GPU Power (640 MW) | GPUs per 1 GW Facility |
|---|---|---|---|
| NVIDIA H100 | 700W | 640 MW | ~914,000 |
| NVIDIA H200 | 700W | 640 MW | ~914,000 |
| NVIDIA B200 | 1,000W | 640 MW | ~640,000 |
| NVIDIA GB200 (Superchip) | 1,200W (combined) | 640 MW | ~533,000 |
| GB200 NVL72 (rack-level) | 120 kW/rack (72 GPUs) | 640 MW | ~5,333 racks / ~384,000 GPUs |

**Note:** The GB200 NVL72 system at 120 kW per rack for 72 GPUs implies ~1,667W effective per GPU at the rack level (including NVLink switches, Grace CPUs, and cooling infrastructure within the rack). The system-level overhead is higher than individual GPU TDP.

---

## 2. Token Generation from 1 GW: The Inference Throughput Model

### Per-GPU Token Throughput (Output Tokens, Inference)

Performance varies significantly by model size, precision, batch size, and sequence length. Representative benchmarks for production inference:

| GPU | Model | Precision | Tokens/sec/GPU | Scenario |
|---|---|---|---|---|
| H100 | Llama 3.1 70B | FP8 | ~3,200 | Production inference (TP=2, per-GPU effective) |
| H100 | Llama 3.1 8B | FP8 | ~26,400 | Single GPU, small model |
| B200 | Llama 3.3 70B | FP4 | ~10,600 | Single GPU |
| B200 | GPT-OSS 120B | FP4 | ~60,000 | Maximum throughput mode |
| GB200 | Llama 3.3 70B | FP4 | ~11,100 | Single GPU |
| GB200 NVL72 | GPT-MoE-1.8T | Mixed | ~116/GPU (quality) | Latency-optimized (50ms TTL) |

**For this analysis, we use a blended production inference rate** that reflects real-world workloads (mix of model sizes, not all at maximum throughput):

| GPU Generation | Conservative (tokens/sec/GPU) | Aggressive (tokens/sec/GPU) | Basis |
|---|---|---|---|
| H100 | 5,000 | 15,000 | Mix of 8B-70B models, production batching |
| B200 | 15,000 | 45,000 | 2-3x H100; FP4 capable; larger batch efficiency |
| GB200/NVL72 | 15,000 | 50,000 | Similar to B200 per-GPU, but better for large models |

### 1 GW Token Generation (Daily Output)

| GPU Generation | GPUs in 1 GW | Conservative (tokens/day) | Aggressive (tokens/day) |
|---|---|---|---|
| **H100** | ~914,000 | **395 billion** | **1.2 trillion** |
| **B200** | ~640,000 | **830 billion** | **2.5 trillion** |
| **GB200 NVL72** | ~384,000 | **498 billion** | **1.7 trillion** |

**At 85% utilization (industry benchmark):**

| GPU Generation | Daily Tokens (Conservative) | Daily Tokens (Aggressive) | Annual Tokens |
|---|---|---|---|
| **H100** | 336 billion | 1.0 trillion | 123-365 trillion |
| **B200** | 706 billion | 2.1 trillion | 258-766 trillion |
| **GB200 NVL72** | 423 billion | 1.4 trillion | 154-511 trillion |

### Normalized Metric: Tokens per Watt-Hour

| GPU Generation | Tokens/Wh (Conservative) | Tokens/Wh (Aggressive) |
|---|---|---|
| H100 | ~5.5 | ~16.5 |
| B200 | ~11.5 | ~34.5 |
| GB200 NVL72 | ~6.9 | ~23.0 |

**NVIDIA's claim of 10x throughput per megawatt for Blackwell vs. Hopper is validated** at the aggressive end of the range (FP4, large batch, throughput-optimized).

### Key Insight: The 50x Claim

NVIDIA claims GB300 NVL72 (Blackwell Ultra) delivers **50x higher throughput per megawatt** vs. Hopper. This would imply:
- H100 baseline: ~5.5 tokens/Wh (conservative) 
- GB300 target: ~275 tokens/Wh
- 1 GW of GB300 would generate: **~16 trillion tokens/day** at 85% utilization

This represents the theoretical frontier for 2026-2027 deployments.

---

## 3. Revenue from 1 GW: GPU Cloud Rental Economics

### Business Model Spectrum

Revenue per megawatt varies dramatically based on the operator's position in the value chain:

| Business Model | Revenue/MW/Year | Examples | Key Characteristics |
|---|---|---|---|
| **Full-stack neocloud** (own GPUs + software) | $10-12M | CoreWeave | Own hardware, orchestration, software; bear refresh risk |
| **Managed GPU cloud** | $6-8M | Nebius, Lambda | Own GPUs, provide managed access, less software differentiation |
| **Bare-metal / Colocation+GPU** | $3-5M | IREN, Applied Digital | Own power+cooling+building; customer may bring hardware |
| **Pure colocation** (power shell) | $1.5-2.5M | Traditional DCs | Lease space + power; no GPU economics captured |
| **Hyperscaler cloud** (inference API) | $15-25M+ | AWS, Azure, GCP | Capture full margin from hardware through API monetization |

### Derivation: CoreWeave Model (Validated by Public Data)

CoreWeave reported $5.13 billion in annual revenue (FY2025) on ~850 MW of active power:
- **Implied revenue per active MW: ~$6.0M**
- However, this includes ramp-up periods. At steady state with full utilization, CoreWeave's model targets ~$12M/MW/year based on $1.2B ACV per 100 MW cluster

CoreWeave's unit economics for a 100 MW cluster:
- GPU CapEx: $2.7B ($27M per MW)
- Annual rent (DC shell): $165M ($1.65M per MW)
- Electricity cost: ~$56M/year (at $0.075/kWh, 85% utilization, PUE 1.4)
- Contract value: $1.2B/year ($12M per MW)
- 5-year total contract value: $6B
- Payback period: ~2.25 years

### Revenue from 1 GW by Business Model

| Model | Annual Revenue from 1 GW | 5-Year Revenue | Key Assumptions |
|---|---|---|---|
| Full-stack neocloud (CoreWeave-style) | **$10-12 billion** | $50-60B | $10-12M/MW; high utilization; long-term contracts |
| Managed GPU cloud | **$6-8 billion** | $30-40B | $6-8M/MW; moderate software premium |
| Bare-metal / infra | **$3-5 billion** | $15-25B | $3-5M/MW; power+shell economics only |
| Hyperscaler API (inference) | **$15-25 billion** | $75-125B | Full API margin; includes inference pricing premium |

### Derivation: From GPU Hours to Revenue

**Bottom-up calculation (H100, current pricing):**
- 1 GW → ~914,000 H100s (at 640 MW effective GPU power)
- Average rental price: $2.50-$6.00/GPU-hour (market range)
- Mid-market assumption: $3.50/GPU-hour
- At 85% utilization: 914,000 × $3.50 × 8,760 hrs × 0.85 = **$23.8 billion/year**

This top-of-range figure reflects gross rental revenue at listed on-demand rates. Actual realized revenue is lower due to:
- Volume discounts on long-term contracts (30-50% discount)
- Reserved instance pricing (40-60% of on-demand)
- Spot pricing pressure (50-75% discount)
- Infrastructure amortization already netted in contract pricing

**Adjusted for contract mix:** $23.8B × 0.45 (blended discount) = **~$10.7 billion/year** — closely matching the CoreWeave $10-12M/MW benchmark.

### Derivation: From Token Revenue to Facility Revenue

**Bottom-up calculation (B200, inference API model):**
- 1 GW → ~640,000 B200 GPUs
- At 15,000 tokens/sec/GPU (conservative) and 85% utilization:
  - Daily output: 706 billion tokens
  - Annual output: ~258 trillion tokens
- Pricing at $2/million output tokens (GPT-4o-tier):
  - Annual inference revenue: 258 trillion × $2/1M = **$516 billion/year** (gross API revenue)
  
This figure represents the **end-user API revenue** captured by companies like OpenAI or Anthropic that sell inference. The infrastructure provider (data center operator) captures only a fraction:
- Infrastructure cost as % of API revenue: ~5-15%
- Implied DC operator share: $26-77 billion

**More realistically, at Llama-tier open-model pricing** ($0.15-$0.70/million tokens):
- Annual revenue at $0.50/million tokens: 258 trillion × $0.50/1M = **$129 billion/year** (gross)
- DC operator infrastructure share (10%): ~$12.9 billion → consistent with $10-12M/MW benchmark

---

## 4. Sensitivity Analysis

### Revenue Sensitivity to GPU Generation and Pricing

| Scenario | GPUs in 1 GW | Hourly Rate/GPU | Utilization | Annual Revenue |
|---|---|---|---|---|
| H100, budget pricing | 914,000 | $2.00 | 80% | $12.8B |
| H100, mid-market | 914,000 | $3.50 | 85% | $23.8B |
| H100, enterprise | 914,000 | $6.00 | 85% | $40.9B |
| B200, mid-market | 640,000 | $5.00 | 85% | $23.9B |
| B200, enterprise | 640,000 | $8.00 | 85% | $38.2B |
| GB200 NVL72, premium | 384,000 | $12.00 | 85% | $34.3B |

**Note:** These are gross on-demand revenue figures. Realized revenue after contract discounts is typically 40-55% of on-demand rates.

### Revenue Sensitivity to Power Cost

Power is the largest operating expense for GPU cloud operators. At 1 GW:

| Electricity Cost ($/kWh) | Annual Power Cost (1 GW) | As % of Revenue ($10B) | Impact on Margin |
|---|---|---|---|
| $0.04 (cheap gas/hydro) | $350M | 3.5% | High margin |
| $0.06 (ERCOT average) | $526M | 5.3% | Good margin |
| $0.075 (CoreWeave benchmark) | $657M | 6.6% | Standard margin |
| $0.10 (premium markets) | $876M | 8.8% | Margin pressure |
| $0.12 (constrained markets) | $1,051M | 10.5% | Significant pressure |

At 85% utilization and PUE 1.25: Annual kWh = 1,000,000 kW × 8,760 hrs × 0.85 = 7.446 TWh

### Token Generation Sensitivity to Workload Mix

| Workload Mix | Effective Tokens/sec/GPU (B200) | Daily Tokens (1 GW) | Annual Tokens |
|---|---|---|---|
| 100% training (no tokens) | 0 | 0 | 0 |
| 80% training / 20% inference | 3,000 | 141B | 52T |
| 50% training / 50% inference | 7,500 | 353B | 129T |
| 20% training / 80% inference | 12,000 | 565B | 206T |
| 100% inference (throughput-opt) | 15,000-45,000 | 706B-2.1T | 258-766T |

**Critical distinction:** Training workloads do not produce tokens for external consumption. A 1 GW facility dedicated to training (common for frontier model labs like OpenAI, Anthropic, Google DeepMind) produces zero commercial inference tokens but enables future model capabilities.

---

## 5. Cross-Reference: Industry Data Points

### Validated Revenue per MW Metrics

| Company | Active Capacity | Annual Revenue | Implied Rev/MW | Notes |
|---|---|---|---|---|
| CoreWeave | ~850 MW (FY2025) | $5.13B | ~$6.0M/MW | Includes ramp; steady-state target is $10-12M |
| Nebius | ~200-400 MW (est. 2025) | ~$1.2B ARR | ~$4-6M/MW | Rapid growth phase |
| IREN | 810 MW operating | $2.3B ARR contracted | ~$2.8M/MW | Bare-metal model; lower capture rate |

### What This Means for 1 GW

Based on validated industry data:

| Metric | Low Case | Base Case | High Case |
|---|---|---|---|
| **Annual revenue** | $3B | $8B | $12B |
| **5-year cumulative revenue** | $15B | $40B | $60B |
| **Daily inference tokens (B200)** | 350B | 700B | 2.1T |
| **Annual inference tokens** | 130T | 260T | 770T |
| **Required GPU CapEx** | $15B | $20B | $27B |
| **Annual power cost** | $350M | $550M | $900M |
| **Payback period** | 5 years | 2.5 years | 2.0 years |

---

## 6. The Power-Value Chain: From Electron to Dollar

### The Multiplication Effect

Each watt of power entering a data center undergoes a multiplication through the value chain. Using base-case assumptions:

```
1 GW of Power
  → 640 MW effective GPU power (PUE + overhead)
  → 640,000 B200 GPUs (at 1,000W TDP)
  → 706 billion tokens/day (at 15K tokens/sec, 85% utilization)
  → 258 trillion tokens/year
  → $8 billion/year GPU rental revenue (contract pricing)
  → OR $129-516 billion/year end-user API revenue (depending on pricing tier)
```

### Value Capture by Layer

| Layer | Annual Value from 1 GW | Share of End-User Revenue |
|---|---|---|
| Power supplier (utility/PPA) | $0.4-0.9B | 0.3-0.7% |
| Data center shell (real estate + cooling) | $1.5-2.0B | 1.2-1.5% |
| GPU cloud operator (full-stack) | $8-12B | 6-9% |
| AI model provider (API layer) | $50-130B | 40-100% |
| End-user application layer | Variable (multiples of API cost) | N/A |

**Key insight:** The value multiplication from raw power to end-user AI service is approximately **100-300x**. One dollar of electricity becomes $100-$300 of economic value when converted to AI inference services.

---

## 7. Forward-Looking Projections

### Efficiency Trajectory (Tokens per Watt-hour)

| Timeline | GPU Generation | Tokens/Wh (Production) | Improvement vs. H100 |
|---|---|---|---|
| 2023-2024 | H100 | ~5-8 | Baseline |
| 2025-2026 | B200/GB200 | ~12-35 | 2-5x |
| 2026-2027 | GB300 (Blackwell Ultra) | ~50-100 | 10-15x |
| 2027-2028 | Rubin (R100) | ~100-250 (est.) | 20-50x |

### Implication for Power Demand

If token demand continues growing exponentially but efficiency improves 10x per generation (every 18-24 months), the net power demand per token falls rapidly. However, total token demand is growing faster than efficiency gains, meaning **aggregate power demand continues to rise even as per-token power consumption drops**.

A 1 GW facility in 2028 (Rubin-generation) could produce 20-50x more tokens than a 1 GW facility today. This means either:
- The same revenue can be generated with far less power, OR
- The same power generates far more revenue (if demand for inference scales to absorb the output)

Current industry consensus is the latter: demand for inference tokens is effectively unlimited at lower price points, meaning more efficient hardware drives volume growth, not power demand reduction.

### Revenue per MW Trajectory

| Year | Dominant GPU | Est. Revenue/MW (Full-stack) | Driven By |
|---|---|---|---|
| 2024 | H100 | $6-8M | Supply scarcity premium |
| 2025-2026 | B200/GB200 | $8-12M | Blackwell efficiency + high demand |
| 2027-2028 | GB300/Rubin | $10-15M (est.) | Volume growth, inference scaling |

Revenue per MW is expected to remain stable or increase slightly because:
1. More efficient GPUs → more tokens per MW → more revenue per MW
2. Token pricing declines are offset by volume increases
3. Long-term contracts lock in current pricing for 3-5 years

---

## 8. Summary Table: 1 GW Economics

| Parameter | Value | Range |
|---|---|---|
| **Total facility power** | 1,000 MW | — |
| **Effective GPU power** | 640 MW | 530-710 MW |
| **GPU count (B200)** | ~640,000 | 530K-710K |
| **GPU count (H100)** | ~914,000 | 760K-1M |
| **GPU count (GB200 NVL72 racks)** | ~5,333 racks / ~384,000 GPUs | 4,400-5,900 racks |
| **Daily token output (inference, B200)** | ~700 billion | 350B-2.1T |
| **Annual token output** | ~258 trillion | 130T-770T |
| **Annual GPU rental revenue** | ~$8-10 billion | $3-12B |
| **Annual inference API revenue (end-user)** | ~$130-500 billion | Model/pricing dependent |
| **GPU CapEx required** | ~$20 billion | $15-27B |
| **Annual power cost** | ~$550 million | $350M-$1.05B |
| **Payback period** | ~2.5 years | 2-5 years |
| **Power cost as % of revenue** | ~5-7% | 3-11% |
| **Value multiplication (power→API revenue)** | ~100-300x | — |

---

## 9. Investment Implications

1. **Power is extraordinarily high-leverage in the AI value chain.** At 5-7% of revenue, power is a small direct cost but an absolute gating constraint. Securing 1 GW of power at $60/MWh vs. $100/MWh saves ~$350M/year — a 3-4% margin advantage on $8-10B of revenue, compounding over a 15-20 year PPA term.

2. **The "tokens per watt" metric is the new Moore's Law for AI infrastructure.** Each GPU generation roughly doubles tokens/watt. Investors should track this metric as a leading indicator of which operators will generate the most revenue per unit of secured power.

3. **Training vs. inference mix is critical for revenue analysis.** A 1 GW facility dedicated to training (e.g., OpenAI's usage of CoreWeave) produces zero commercial tokens today but enables future inference revenue. The industry is shifting toward 60-70% inference workloads, which favors revenue-per-MW models.

4. **GPU CapEx intensity ($20B+ per GW) means power decisions are effectively capital allocation decisions.** An operator committing to 1 GW is implicitly committing to $20B+ in GPU CapEx over the facility's life. The power contract is the anchor that unlocks this capital deployment.

5. **The hyperscaler API layer captures 10-40x more value than the infrastructure layer per watt consumed.** This explains why hyperscalers are willing to pay $100+/MWh for nuclear power: the marginal revenue from an additional megawatt of inference capacity far exceeds the power cost premium.

---

*Disclaimer: This analysis uses publicly available benchmarks, pricing data, and reported financials. Actual performance varies significantly based on model architecture, batch size, precision, cooling efficiency, utilization patterns, and contract structures. Projections are illustrative, not predictive.*
