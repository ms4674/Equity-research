# The Two Axes of LLM Inference Memory: Context Length vs. Parameter Scaling

*A companion note to `LLM_Parametric_vs_Memory_Scaling.xlsx`. All illustrative
numbers below are reproducible from that workbook (FP16 weights, FP16 KV,
batch = 1 unless stated).*

---

## 1. The setup: inference memory has two independent growth vectors

Serving a large language model consumes device memory (HBM) in three buckets:

1. **Parametric memory (weights)** — fixed once the model is trained.
   `weights_bytes = params x bytes_per_param`. It scales **linearly with
   parameter count** and is set entirely by precision (FP16 = 2 B/param,
   FP8 = 1 B, INT4 = 0.5 B).
2. **KV cache** — the running store of per-token key/value projections that
   lets decode avoid recomputing attention over the prompt.
   `KV_bytes = 2 x layers x kv_heads x head_dim x seq_len x batch x kv_bytes`.
   It scales **linearly with context length AND with batch (concurrency)**.
3. **Activations / runtime overhead** — transient, comparatively small at
   decode time.

The central question — *which axis drives memory growth, and specifically
the KV cache?* — matters because the two axes have diverged sharply over the
last five years, and the hardware/architecture response to each has been
different.

---

## 2. How each axis has grown historically

### 2.1 Parameter scaling (weights)

| Era | Representative model | Params | FP16 weights |
|-----|----------------------|--------|--------------|
| 2019 | GPT-2 | 1.5 B | ~3 GB |
| 2020 | GPT-3 | 175 B | ~350 GB |
| 2022 | PaLM | 540 B | ~1,080 GB |
| 2023 | Llama-2 | 7–70 B | 14–138 GB |
| 2024 | Llama-3.1 | 8–405 B | 16–810 GB |
| 2024+ | Frontier MoE (e.g. DeepSeek-V3 671B, GPT-4-class) | 0.5–1.8 T total | often 1 TB+ *total*, but only a fraction **active** per token |

Two things happened to *decouple* headline parameter growth from the memory
that actually has to move during decode:

- **Quantization.** The effective bytes/param fell from 4 (FP32) → 2 (FP16/BF16)
  → 1 (FP8, native on Hopper/Blackwell) → 0.5 (INT4/NF4 weight-only). A 70 B
  model is ~138 GB in FP16 but ~35 GB in INT4 — a 4x compression on the same
  parameter count.
- **Mixture-of-Experts (MoE).** Total parameters kept climbing (into the
  trillions), but only a small subset of experts activate per token, so the
  memory *bandwidth* cost per decode step grew far slower than the headline
  count. MoE shifts the binding constraint from bandwidth back toward
  **capacity** (all experts must still be resident in HBM).

**Net:** weights memory per model has grown roughly 2–3 orders of magnitude
since GPT-3, but precision and MoE have blunted how much of that must be *read
per token*.

### 2.2 Context length

Context windows have grown far faster than parameter counts:

| Year | Typical usable context |
|------|------------------------|
| 2020 (GPT-3) | 2 K |
| 2023 (GPT-4, Llama-2) | 4 K – 32 K |
| late 2023–2024 (GPT-4 Turbo, Claude, Llama-3.1) | 128 K |
| 2024–2025 (Gemini 1.5/2, Llama-4-class) | 1 M – 10 M |

That is roughly a **3–4 orders-of-magnitude** increase (2 K → 1 M+) in about
five years — steeper than the ~2–3 orders of parameter growth, and critically,
**each additional context token adds KV cache linearly with no offsetting
compression from the "scaling" itself.**

---

## 3. Which is the bigger driver of the KV cache?

The KV cache formula makes the dependency explicit:

```
KV_bytes = 2 x layers x kv_heads x head_dim x seq_len x batch x kv_bytes
           \-------- set by architecture --------/   \--- runtime ---/
```

- **Parameter scaling** enters only through `layers x kv_heads x head_dim`
  (model depth and attention width). All else equal, a bigger model does have
  a larger *per-token* KV footprint.
- **Context length** (`seq_len`) and **batch** enter as pure linear
  multipliers that the runtime controls.

### 3.1 The empirical answer: context length dominates — *because architecture broke the parameter link*

The decisive fact is that **the industry deliberately severed the coupling
between parameter scaling and KV-cache size**, while doing nothing to slow
context growth. Grouped-Query Attention (GQA) and Multi-head Latent Attention
(MLA) cut `kv_heads` (or compress KV entirely), so per-token KV has actually
*fallen* even as models got bigger:

| Model | Params | Attention | KV / token | Comment |
|-------|--------|-----------|-----------|---------|
| Llama-2-7B | 7 B | MHA (32 KV heads) | **512 KB** | Multi-head: KV scales with all heads |
| Llama-3.1-8B | 8 B | GQA (8 KV heads) | **128 KB** | 4x smaller KV on a *larger* model |
| GPT-3 175B | 175 B | MHA (96 KV heads) | **4,608 KB** | Old MHA at scale is brutal |
| Llama-3.1-70B | 70 B | GQA (8 KV heads) | **320 KB** | 14x smaller KV/token than GPT-3 despite 40% the size |
| Llama-3.1-405B | 405 B | GQA (8 KV heads) | **504 KB** | KV/token barely above Llama-2-7B |

Read the table carefully: a **405 B** GQA model carries less KV per token than
a **7 B** MHA model does. Parameter scaling, once GQA/MLA are applied, is a
*weak* driver of KV size. Context length is the strong one.

### 3.2 The crossover: at long context, KV cache overtakes weights entirely

For a fixed model, weights are constant while KV grows linearly with context ×
batch. The two curves cross:

- **Llama-3.1-8B**, FP16: weights ≈ **16 GB**.
  - KV at 8 K context (batch 1) ≈ 1 GB — weights dominate.
  - KV at 128 K context (batch 1) ≈ **17 GB** — KV now *exceeds* the weights.
  - KV at 128 K context, **batch 8** ≈ **137 GB** — KV is ~9x the weights and
    the model no longer fits on a single 80 GB GPU.

So the practical driver of *incremental* inference memory in production today
is **context length multiplied by concurrency (batch)**, not raw parameter
count. Parameter scaling sets the fixed weights floor; context/batch set the
slope — and the slope is what blows past HBM capacity in long-context, high-QPS
serving.

### 3.3 Caveat — the honest nuance

This is a statement about *today's architectures and workloads*, not a law:

- If a lab ships a large model with **full MHA** and many KV heads, parameter
  scaling roars back as a KV driver (see GPT-3's 4.6 MB/token).
- Parameter scaling still dominates the **weights floor** and therefore the
  minimum GPU count to load a model at all (a 405 B model needs ~12x H100 just
  for FP16 weights, before a single KV byte).
- MoE inverts the constraint again: enormous *total* params (capacity-bound)
  with modest active params (bandwidth-bound).

The clean summary: **across models, weights are driven by parameter scaling;
within a deployed model, the KV cache is driven overwhelmingly by context
length × batch, with per-token KV held down by GQA/MLA.**

---

## 4. How GPU / ASIC architecture has responded

Because the two memory axes grew this way, the hardware and systems response
split into a **capacity** track and a **bandwidth** track, plus algorithmic
compression of the KV cache itself.

### 4.1 HBM capacity (chasing weights + long-context KV)

| Accelerator | HBM | Bandwidth | Notable |
|-------------|-----|-----------|---------|
| A100 (2020) | 40 / 80 GB HBM2e | 1.6 / 2.0 TB/s | Ampere |
| H100 (2022) | 80 GB HBM3 | 3.35 TB/s | FP8 tensor cores |
| H200 (2024) | 141 GB HBM3e | 4.8 TB/s | Same compute, +76% memory for KV |
| B200 (2024–25) | 192 GB HBM3e | ~8 TB/s | Blackwell, FP4/FP6 |
| B300 / GB300 | 288 GB HBM3e | higher still | Memory-first refresh |
| AMD MI300X | 192 GB HBM3 | 5.3 TB/s | Capacity-led positioning |

The through-line: memory **capacity and bandwidth have become the marketed
spec**, often ahead of FLOPS. H200 is literally an H100 with more/faster HBM —
a direct response to KV-cache and long-context pressure. The HBM roadmap
(HBM3 → HBM3e → HBM4) and per-package stack counts are now first-order to
inference TCO, which is why HBM suppliers (SK hynix, Samsung, Micron) sit on
the critical path.

### 4.2 Numeric formats (attacking bytes-per-element on both axes)

- **Weights:** FP16 → FP8 (Hopper) → FP4/FP6 (Blackwell) roughly halve, then
  halve again, the weights footprint and the bytes moved per decode step.
- **KV cache:** FP8 and INT4 **KV-cache quantization** directly shrink the
  fastest-growing bucket. Blackwell's native low-precision path makes FP8 KV
  close to free in quality terms for many workloads.

### 4.3 Attention architecture (attacking KV at the source)

- **MHA → MQA → GQA → MLA.** Grouped-Query Attention (Llama-2-70B onward) is
  now the default; Multi-head Latent Attention (DeepSeek-V2/V3) compresses KV
  into a low-rank latent, cutting KV memory by an order of magnitude. This is
  *the* reason large models remain long-context-serviceable.
- **Sliding-window / local / sparse attention** (Mistral, others) caps the KV
  that must be retained, bounding growth for very long inputs.

### 4.4 Serving-systems innovations (making KV memory efficient)

- **PagedAttention / vLLM** — virtual-memory-style KV paging removes
  fragmentation and lifts effective batch size (higher tokens/GPU).
- **Prefix / prompt caching** — reuse KV for shared system prompts across
  requests.
- **Disaggregated prefill/decode** — split the compute-bound prefill from the
  bandwidth-bound decode onto differently provisioned hardware.
- **KV offload & tiering** — spill KV to CPU DRAM / NVMe for very long context,
  trading bandwidth for capacity.

### 4.5 Scale-up fabric (pooling memory across chips)

- **NVLink / NVSwitch domains** (e.g. GB200 NVL72 pooling 72 GPUs) let weights
  and KV span many accelerators at near-local bandwidth — the systems answer to
  trillion-param MoE weights and multi-hundred-GB KV working sets.

### 4.6 The radical end: SRAM-centric ASICs

- **Groq LPU, Cerebras WSE, SambaNova** move the model into massive on-chip
  **SRAM**, sidestepping HBM bandwidth limits for decode. The trade-off flips:
  extreme bandwidth/latency, but tight capacity — which puts a hard premium on
  KV compression and small footprints. Google **TPU** and AWS **Trainium/
  Inferentia** pursue the more conventional HBM path with custom interconnect.

---

## 5. Investment / strategic takeaways

1. **Context length, not parameter count, is the marginal memory driver in
   production inference.** GQA/MLA capped per-token KV even as models grew, but
   context (2 K → 1 M+) and concurrency multiply KV without relief. Long-context
   × high-QPS is what breaches HBM capacity.
2. **Parameter scaling still sets the weights floor** (minimum GPU count to load
   a model) and, via MoE, drives *capacity* demand even when bandwidth demand is
   contained.
3. **HBM has become the binding constraint and the marketed spec.** Capacity
   (H200/B300/MI300X) and bandwidth (HBM3e → HBM4) upgrades track the KV-cache
   problem directly — favorable for HBM suppliers and memory-rich accelerators.
4. **The efficient frontier is increasingly algorithmic:** FP8/FP4 KV
   quantization, MLA, sliding-window attention, PagedAttention, prefix caching,
   and prefill/decode disaggregation collectively bend the KV curve more cheaply
   than buying HBM — a key margin lever for inference providers.
5. **Architectural bifurcation:** HBM-centric GPUs/ASICs (NVIDIA, AMD, TPU,
   Trainium) optimize for capacity + interconnect to hold big weights and long
   KV; SRAM-centric ASICs (Groq, Cerebras) optimize for decode bandwidth at the
   cost of capacity, betting on aggressive KV/weight compression.

*See `LLM_Parametric_vs_Memory_Scaling.xlsx` (tabs: Weights Memory, KV Cache
Scaling, Decode Economics, Parametric Scaling) to flip precision, batch, GPU,
and context and reproduce every figure cited here.*
