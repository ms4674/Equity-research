# KV Cache & Prefill Usage for Frontier LLMs

## Equity Research Report — March 2026

---

## Executive Summary

KV (Key-Value) cache has emerged as the single most important infrastructure constraint for frontier AI inference. As large language models scale to 128K–1M token contexts and agent workloads dominate production deployments, KV cache memory now consumes up to 82% of total GPU VRAM — far exceeding model weights. This report analyzes KV cache economics across 10 frontier models, 3 agent platforms, 4 inference frameworks, and quantifies the downstream impact on HBM memory demand, GPU utilization, and infrastructure costs.

**Key thesis**: The "inference flip" (inference now at 67% of AI compute, up from 20% in 2022) combined with expanding context windows means KV cache — not model weights, not training compute — is the marginal driver of GPU memory demand and the binding constraint on AI infrastructure scaling.

---

## 1. What Is KV Cache and Why It Matters

During LLM inference, the transformer's attention mechanism needs to access Key and Value vectors from all prior tokens at every generation step. Without caching, computation grows quadratically with sequence length. The KV cache stores these pre-computed vectors so only new tokens require fresh computation.

**Memory formula**: `2 × layers × KV_heads × head_dim × seq_length × bytes_per_element`

**Rule of thumb**: ~0.1–1.0 MB per token for modern models, depending on architecture.

### Training vs Inference: A Fundamental Asymmetry

| Dimension | Training | Inference |
|-----------|----------|-----------|
| **KV Cache** | NOT stored — recomputed via FlashAttention | PERSISTENT — stored for entire generation |
| **Why?** | All tokens known upfront; full parallel attention | Tokens generated sequentially; must reuse past K/V |
| **Memory Bottleneck** | Optimizer states (2× weights) + gradients (1× weights) | KV cache (grows with context × concurrent users) |
| **KV Memory (70B, 128K)** | 0 GB (FlashAttention recomputes) | 42 GB per user (8 users = 336 GB) |
| **Compute Share (2026)** | 33% of total AI compute | 67% of total AI compute |
| **Cost Characteristic** | Discrete event (one-time) | Continuous (billions of queries/day) |

Training avoids the KV cache bottleneck entirely because all tokens are known upfront and attention can be computed in parallel using FlashAttention, which recomputes K/V on-the-fly rather than storing them. The dominant training memory consumers are optimizer states (2× model size for Adam), gradients (1× model size), and activations — not KV cache.

Inference is fundamentally different. Tokens are generated one-by-one, and the KV cache must be maintained for the entire conversation to avoid re-processing every prior token. As inference now accounts for 67% of all AI compute spending, KV cache has become the binding infrastructure constraint.

---

## 2. KV Cache Across Frontier Model Architectures

### Per-Token KV Cache by Model

| Model | Params | Layers | KV Heads | Attention | KV (MB/tok, FP16) | KV @ 128K (GB) |
|-------|--------|--------|----------|-----------|-------------------|----------------|
| Llama 3.1 8B | 8B | 32 | 8 | GQA | 0.125 | 15.6 |
| Qwen3 8B | 8B | 36 | 4 | GQA-4 | 0.074 | 9.2 |
| Llama 3.1 70B | 70B | 80 | 8 | GQA | 0.328 | 41.0 |
| Qwen3-235B | 235B (22B active) | 94 | 4 | GQA-4 | 0.154 | 19.2 |
| Llama 3.1 405B | 405B | 126 | 8 | GQA | 0.516 | 64.5 |
| Llama 4 Maverick | 400B (17B active) | 48 | 8 | iRoPE+GQA | 0.197 | 24.6 |
| DeepSeek V3.2 | 685B (37B active) | 60 | MLA | MLA | 0.018 | 2.3 |
| Claude Sonnet 4.6 | Undisclosed | ~80 | ~8 | GQA (est.) | ~0.328 | ~41.0 |
| GPT-4.1 | Undisclosed | ~128 | ~8 | GQA (est.) | ~0.524 | ~65.5 |
| Gemini 2.5 Pro | ~2-4T (MoE) | ~128 | ~16 | MHA | ~1.049 | ~131.1 |

### Attention Mechanism Impact

| Mechanism | Example | KV Compression vs MHA | Per-Token KV |
|-----------|---------|----------------------|--------------|
| MHA (Multi-Head) | Gemini 2.5 Pro | 1.0× (baseline) | ~1.05 MB |
| GQA-8 (Grouped Query) | Llama 3/4 | ~8–16× | ~0.13–0.52 MB |
| GQA-4 (Aggressive) | Qwen3 | ~32× | ~0.07–0.15 MB |
| MLA (Multi-Head Latent) | DeepSeek V3.2 | ~55× | ~0.02 MB |

**Key insight**: DeepSeek V3.2 (685B total params) has smaller per-token KV cache than Llama 3.1 8B (8B params). Architecture dominates parameter count.

---

## 3. Parameter Scaling: A Non-Linear Relationship

KV cache does NOT scale linearly with model parameters. The relationship is governed by `layers × KV_heads × head_dim`, not total parameter count.

| Scaling | Llama 8B → 70B | Llama 8B → 405B | Naive expectation |
|---------|----------------|-----------------|-------------------|
| Parameter growth | 8.75× | 50.6× | — |
| KV cache growth | 2.6× | 4.1× | 8.75× / 50.6× |
| Layer growth | 2.5× (32→80) | 3.9× (32→126) | — |
| KV heads | 8 → 8 (unchanged) | 8 → 8 (unchanged) | — |

**Layer count is the primary KV scaling knob.** KV heads remain fixed at 8 across the entire Llama family, and head_dim is universally 128. The only dimension that changes is depth.

**MoE further decouples KV from parameters.** Llama 4 Maverick (400B total, 17B active) has lower per-token KV cache than Llama 3.1 70B (dense) because it uses only 48 layers vs 80.

---

## 4. Training vs Inference KV Cache Demand

### Memory Breakdown: Llama 70B

| Component | Training (FP16) | Inference (FP8, 8 users × 128K) |
|-----------|----------------|----------------------------------|
| Model Weights | 140 GB | 70 GB |
| Gradients | 140 GB | 0 |
| Optimizer States (Adam) | 280 GB | 0 |
| Activations | 32 GB | 2 GB |
| **KV Cache** | **0 GB** | **336 GB** |
| **Total** | **592 GB** | **408 GB** |
| **KV as % of total** | **0%** | **82%** |

### The Inference Flip: Compute Spending Over Time

| Year | Training Share | Inference Share | Total AI Spend | Inference Spend |
|------|---------------|-----------------|----------------|-----------------|
| 2022 | 80% | 20% | $50B | $10B |
| 2023 | 67% | 33% | $80B | $26B |
| 2024 | 50% | 50% | $130B | $65B |
| 2025 | 40% | 60% | $200B | $120B |
| 2026 | 33% | 67% | $300B | $201B |
| 2028E | 25% | 75% | $500B | $375B |

### KV Cache Share of Inference Memory Over Context Eras

| Context Era | KV Cache (1 user, 70B) | KV as % of Inference Memory |
|-------------|------------------------|----------------------------|
| 4K (2023) | 1.3 GB | 2% |
| 32K (2024) | 10.5 GB | 13% |
| 128K (2025) | 42.0 GB | 37% |
| 1M (2026) | 328.0 GB | 82% |

**Bottom line**: KV cache was negligible in the 4K-context era. At 128K+ contexts with multiple concurrent users, it utterly dominates GPU memory and is the primary driver of infrastructure demand.

---

## 5. Prompt Caching Economics

### Provider Pricing Comparison

| Model | Input $/M | Cached $/M | Cache Discount | TTL |
|-------|-----------|------------|----------------|-----|
| GPT-4.1 | $2.00 | $1.00 | 50% | 5-10 min |
| GPT-4o | $2.50 | $1.25 | 50% | 5-10 min |
| Claude Sonnet 4.6 | $3.00 | $0.30 | 90% | 5 min |
| Claude Opus 4.6 | $5.00 | $0.50 | 90% | 5 min |
| Gemini 2.5 Pro | $1.25 | $0.125 | 90% | 1 hr (explicit) |
| DeepSeek V3.2 | $0.28 | $0.028 | 90% | ~5 min |

### Cost Sensitivity to Cache Hit Rate

At 50K input + 500 output tokens per request:

| Cache Hit Rate | Claude Sonnet | GPT-4.1 | Gemini 2.5 Pro | DeepSeek V3.2 |
|---------------|--------------|---------|----------------|---------------|
| 0% (cold) | $0.158 | $0.104 | $0.068 | $0.014 |
| 50% | $0.090 | $0.079 | $0.037 | $0.008 |
| 84% (Claude Code avg) | $0.047 | $0.063 | $0.018 | $0.004 |
| 95% | $0.031 | $0.055 | $0.011 | $0.003 |

**Anthropic and Google's 90% cache discount** makes cache hit rate optimization 2× more impactful than on OpenAI (50% discount).

---

## 6. Agent Platform Comparison: OpenClaw vs Claude Cowork vs Perplexity Computer

### KV Cache Architecture Comparison

| Dimension | OpenClaw | Claude Cowork | Perplexity Computer |
|-----------|----------|---------------|---------------------|
| Architecture | Single model, self-hosted or API | Single model (Claude), cloud-only | 19-model orchestrator, cloud-only |
| System Prompt | ~40K tokens (bootstrap files) | ~78K tokens (system + tools + history) | Varies per sub-model; task-scoped |
| Cache Hit Rate | ~35% (known bug: dynamic prompts) | 84–97% (best-in-class) | ~15% (model switching = cold start) |
| Primary Risk | Timestamps/metadata in prefix | Byte-exact match fragility | Every model switch = full miss |
| Self-Hosting | Yes (full KV cache control) | No | No |
| Cost Model | Per-token (user's provider) | Per-token (Anthropic) | Credits ($200/mo, 10K credits) |

### Multi-Turn Session Cost (10 turns, 50K input baseline)

| Platform | Total 10-Turn Cost | vs Claude Cowork |
|----------|--------------------|------------------|
| Claude Cowork | ~$0.35 | — |
| OpenClaw | ~$0.88 | 2.5× more expensive |
| Perplexity Computer | ~$1.15 | 3.3× more expensive |

Claude Cowork's superior cache hit rate (84%) compounds across turns, producing dramatically lower costs in multi-turn agent sessions. OpenClaw's cache invalidation bug (#45110, #40256) causes full re-prefill on most turns. Perplexity Computer pays the highest price because every model switch across its 19-model roster destroys KV cache at each provider boundary.

---

## 7. TurboQuant: Changing the Memory Calculus

Google's TurboQuant (ICLR 2026) is a data-oblivious quantization method that compresses KV cache to ~3 bits per channel.

### TurboQuant vs Other Quantization Methods

| Method | Bits | Compression | Throughput Gain | Accuracy Loss | Year |
|--------|------|-------------|-----------------|---------------|------|
| FP16 (baseline) | 16 | 1.0× | 1.0× | None | — |
| FP8 | 8 | 2.0× | ~1.5× | Negligible | 2023 |
| KIVI | 2 | 2.6× | 2.35–3.47× | <1 ppl | 2024 |
| RotateKV | 2 | 3.97× | 2.32× | <0.3 ppl | 2025 |
| **TurboQuant** | **3** | **6.0×** | **8.0× (H100)** | **Zero** | **2026** |
| KV-Compress | Variable | 8–64× | 5.18× | >90% at 64× | 2025 |

### Impact on Concurrent User Capacity (8×H100, 640 GB, 128K context)

| Model | FP16 KV Users | TurboQuant Users | Improvement |
|-------|---------------|------------------|-------------|
| Llama 3.1 8B | 40 | 243 | 6.1× |
| Llama 3.1 70B | 40 | 243 | 6.1× |
| Llama 3.1 405B | 3 | 22 | 7.3× |

### Impact on Global HBM Demand

| Scenario | KV Cache HBM (2026 est.) | Reduction |
|----------|--------------------------|-----------|
| FP16 (current) | 1.40 EB | — |
| FP8 (common) | 0.70 EB | 50% |
| KIVI (2-bit) | 0.54 EB | 61% |
| TurboQuant (3-bit) | 0.23 EB | 83% |

If universally adopted, TurboQuant would reduce KV cache HBM demand from 1.4 EB to 0.23 EB — potentially alleviating the HBM supply shortage years before new fab capacity arrives.

---

## 8. HBM Memory Supply-Demand Mismatch

### The Structural Shortage

| Year | HBM Supply (B GB) | HBM Demand (B GB) | Gap | Shortage % |
|------|-------------------|-------------------|-----|-----------|
| 2023 | 1.5 | 1.4 | -0.1 | Balanced |
| 2024 | 2.8 | 3.5 | 0.7 | 20% |
| 2025 | 7.6 | 9.5 | 1.9 | 20% |
| 2026 | 12.0 | 18.0 | 6.0 | 33% |
| 2027E | 20.0 | 28.0 | 8.0 | 29% |
| 2028E | 35.0 | 40.0 | 5.0 | 13% |

### HBM Per GPU Escalation

| GPU | HBM Capacity | HBM Type | HBM as % of GPU Cost |
|-----|-------------|----------|---------------------|
| A100 (2020) | 80 GB | HBM2e | 15% |
| H100 (2022) | 80 GB | HBM3 | 20% |
| H200 (2024) | 141 GB | HBM3e | 28% |
| B200 (2025) | 192 GB | HBM3e | 35% |
| B300 (2025) | 288 GB | HBM3e | 38% |

### KV Cache as Demand Multiplier

At 128K context with 8 concurrent users (Llama 70B, FP16):
- Model weights (FP8): 70 GB
- KV cache: 336 GB
- **KV cache is 82% of total VRAM**

The HBM shortage is not merely a manufacturing problem — KV cache requirements are a demand-side multiplier that amplifies total memory needs far beyond model weights alone.

### Vendor Concentration Risk

| Vendor | 2025 Share | 2026E Share |
|--------|-----------|-------------|
| SK Hynix | 62% | 54% |
| Samsung | 22% | 28% |
| Micron | 16% | 18% |

SK Hynix and Micron have announced their entire 2026 HBM production is sold out. Samsung has struggled to meet NVIDIA qualification standards. This near-duopoly creates extreme supply vulnerability.

---

## 9. Inference Framework Landscape

| Framework | KV Management | Prefix Caching | Throughput (8B, H100) | Best For |
|-----------|---------------|----------------|----------------------|----------|
| vLLM v1 | PagedAttention | Zero-overhead | 12,553 tok/s | Multi-modal, high-concurrency |
| SGLang | RadixAttention | Automatic (radix tree) | 16,215 tok/s | Shared prefix, agents, DeepSeek |
| TensorRT-LLM | Paged + aggressive reuse | Yes | ~14,000 tok/s | Batch processing, long-context |
| NVIDIA Dynamo | KV-aware routing | Cache pinning + speculative | Orchestrator | Agent workloads, multi-node |

SGLang's RadixAttention delivers 29% higher throughput than vLLM for prefix-heavy agent workloads by automatically discovering and reusing shared KV cache prefixes via a radix tree with LRU eviction.

---

## 10. Key Takeaways

### Architecture & Memory
1. **MLA is transformational**: DeepSeek V3.2's MLA reduces per-token KV cache by ~55× vs MHA
2. **GQA is the pragmatic default**: 8–32× compression with negligible quality loss
3. **Parameter scaling is sublinear for KV**: 70B model has 2.6× the KV of 8B (not 8.75×)
4. **Architecture matters more than size**: DeepSeek 685B has smaller KV than Llama 8B

### Training vs Inference
5. **Training does NOT store KV cache**: FlashAttention recomputes attention; KV = 0 bytes
6. **Inference KV cache dominates at scale**: 82% of total memory at 128K × 8 users
7. **The inference flip amplifies KV demand**: 67% of AI compute is inference (2026)
8. **Context expansion is the multiplier**: KV share grew from 2% (4K) to 82% (128K)

### Prompt Caching
9. **90% cache discounts** from Anthropic/Google vs 50% from OpenAI
10. **Cache hit rate is the #1 cost lever**: 84% hit rate = 74% cost reduction
11. **Claude Cowork leads at 84–97%** cache hit rate; OpenClaw at ~35%; Perplexity at ~15%

### TurboQuant
12. **6× KV compression at zero accuracy loss** (ICLR 2026, data-oblivious, training-free)
13. **Concurrent users 6× on same hardware**: 40 → 243 users (Llama 70B, 128K, 8×H100)
14. **Global HBM relief**: 1.4 EB → 0.23 EB if universally adopted

### Memory Supply-Demand
15. **HBM demand exceeds supply by 33% in 2026**; gap persists to 2028–2029
16. **HBM per GPU grew 3.6× in 5 years** (80 GB → 288 GB); now 38% of GPU cost
17. **Near-duopoly risk**: SK Hynix ~60% share; only 2 reliable HBM suppliers

---

## Data Sources

- Model architectures: DeepSeek V3 technical report, Llama 4 release blog, Qwen3 model card, HuggingFace config files
- Pricing: OpenAI, Anthropic, Google, DeepSeek official API pricing pages (March 2026)
- Inference benchmarks: vLLM/SGLang/TensorRT-LLM published comparisons, NVIDIA DGX Spark specs
- Agent data: Claude Code usage statistics, OpenClaw GitHub issues, Perplexity Computer reverse engineering analysis
- TurboQuant: Google Research (ICLR 2026, arXiv:2504.19874)
- HBM market: SK Hynix/Samsung/Micron financials, TrendForce, Yole Group, CNBC
- GPU installed base: Epoch AI (15M H100-equiv.), NVIDIA GTC 2026
- Compute split: Deloitte, SemiAnalysis, GPUnex, TechTicker estimates

---

*Analysis notebook, Excel data (18 sheets), 19 static charts, and interactive dashboard available in the repository.*
