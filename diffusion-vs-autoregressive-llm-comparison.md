# Diffusion vs. Autoregressive LLMs: Token Usage, Leading Models, Training Data & Reinforcement Learning

## Executive Summary

Large language models (LLMs) have historically been dominated by autoregressive (AR) architectures that generate tokens sequentially via next-token prediction. A newer paradigm — **diffusion language models (dLLMs)** — generates tokens in parallel through iterative denoising, analogous to image diffusion models applied to discrete text. As of early 2026, autoregressive models remain the dominant production paradigm, but diffusion LLMs are advancing rapidly with promising speed and quality trade-offs.

---

## 1. Architecture Overview

### Autoregressive (AR) Models

- Generate tokens **one at a time**, left to right
- Each token is conditioned on all previously generated tokens
- Training objective: maximize log-likelihood of next token given preceding context
- Inference is inherently sequential, limiting throughput per sequence

### Diffusion Language Models (dLLMs)

- Generate **all tokens in parallel** through iterative refinement
- Start from fully masked (or noisy) sequences and progressively denoise
- Training objective: predict masked tokens at varying noise levels (masked diffusion) or learn a reverse diffusion process
- Inference can leverage parallelism for higher throughput per sequence

---

## 2. Leading Autoregressive Models

| Model | Developer | Parameters | Training Tokens | Architecture | Release |
|-------|-----------|-----------|----------------|--------------|---------|
| **GPT-4 / GPT-4o** | OpenAI | Undisclosed (est. ~1.8T MoE) | ~13T tokens | Dense / MoE Transformer | Mar 2023 / May 2024 |
| **Claude 3.5 / 4** | Anthropic | Undisclosed | Undisclosed (est. 15-20T+) | Dense Transformer | Jun 2024 / Mar 2025 |
| **Gemini 2.5 Pro** | Google DeepMind | Undisclosed | Undisclosed | Dense Transformer | Mar 2025 |
| **DeepSeek-V3** | DeepSeek | 671B total / 37B active | **14.8T tokens** | MoE Transformer | Dec 2024 |
| **DeepSeek-R1** | DeepSeek | 671B total / 37B active | 14.8T (base) + RL | MoE Transformer | Jan 2025 |
| **Llama 3.1 405B** | Meta | 405B | ~15T tokens | Dense Transformer | Jul 2024 |
| **Llama 4 Maverick** | Meta | 400B total / 17B active (128 experts) | Undisclosed | MoE Transformer | Apr 2025 |
| **Llama 4 Scout** | Meta | 109B total / 17B active (16 experts) | Undisclosed | MoE Transformer | Apr 2025 |
| **Qwen 2.5** | Alibaba | 0.5B–72B (dense) | **18T tokens** | Dense Transformer | Sep 2024 |
| **Mistral Large 3** | Mistral AI | 675B total / 41B active | Undisclosed | Granular MoE | Dec 2025 |
| **OLMo 3** | AI2 | 7B, 32B | Undisclosed | Dense Transformer | Dec 2025 |

### Training Data Scale Trends

- **2020 (GPT-3):** 300 billion tokens
- **2023 (GPT-4):** ~13 trillion tokens
- **2024 (Llama 3.1):** ~15 trillion tokens
- **2024 (Qwen 2.5):** 18 trillion tokens
- **2025–2026 frontier models:** Estimated 20-30T+ tokens, increasingly using synthetic data and multi-epoch training due to the "data wall" (high-quality public internet text estimated at only 15–20T tokens total)

---

## 3. Leading Diffusion Language Models

| Model | Developer | Parameters | Training Tokens | Architecture | Release |
|-------|-----------|-----------|----------------|--------------|---------|
| **Mercury 2** | Inception Labs | Undisclosed | Undisclosed | Diffusion Transformer | Feb 2026 |
| **Mercury Coder** | Inception Labs | Undisclosed | Undisclosed | Diffusion Transformer | Feb 2025 |
| **LLaDA 8B** | Peking University / GSAI | 8B | ~2T tokens (est.) | Masked Diffusion Transformer | Feb 2025 |
| **LLaDA 2.0-mini** | InclusionAI | 16B | Converted from AR base | Masked Diffusion Transformer | Dec 2025 |
| **LLaDA 2.0-flash** | InclusionAI | 100B (MoE) | Converted from AR base | Masked Diffusion MoE | Dec 2025 |
| **LLaDA 2.1** | InclusionAI | 16B+ | Converted from AR base | Masked Diffusion (Token Editing) | Feb 2026 |
| **MMaDA** | Multi-institution | 8B | Undisclosed | Multimodal Diffusion | May 2025 |
| **Evo** | Research | 8B | Undisclosed | AR-Diffusion Hybrid | Feb 2026 |
| **MDLM** | Research | <1B | Undisclosed | Masked Discrete Diffusion | 2024 |

### Notable Characteristics

- **Mercury** is the first commercially deployed dLLM, achieving 1,000+ tokens/sec on H100/Blackwell GPUs — up to **10x faster** than speed-optimized AR frontier models
- **LLaDA 2.0** demonstrated that AR models can be **converted into diffusion models** via a 3-phase block-level training scheme, avoiding the need to pretrain from scratch
- **LLaDA 2.0-flash-CAP** achieves 2.1x inference acceleration with parallel decoding at up to 535 tokens/sec
- Diffusion models remain smaller in scale (8B–100B) compared to frontier AR models (hundreds of billions to trillions of parameters)

---

## 4. Token Usage: Inference Efficiency Comparison

| Metric | Autoregressive | Diffusion |
|--------|---------------|-----------|
| **Generation pattern** | Sequential (1 token/step) | Parallel (many tokens/step) |
| **Tokens/sec (typical)** | 50–200 tok/s (frontier models) | 500–1,100 tok/s (Mercury family) |
| **Compute per token** | One forward pass per token | Amortized across denoising steps |
| **Latency to first token** | Low (streaming) | Higher (batch generation) |
| **Total latency for long outputs** | Scales linearly with length | Sub-linear scaling potential |
| **KV-cache memory** | Grows linearly with context | Not required (no causal dependency) |
| **Speculative decoding compatible** | Yes (well-studied) | Inherent parallelism instead |

### Key Insight on Token Efficiency

Diffusion models amortize compute across multiple tokens per denoising step, making them more efficient for bulk generation. AR models excel at streaming use cases where partial output is valuable immediately. The throughput advantage of dLLMs becomes most pronounced for longer outputs.

---

## 5. Training Data Comparison

### Autoregressive Models — Disclosed Training Data

| Model | Pretraining Tokens | Data Sources |
|-------|-------------------|--------------|
| GPT-3 (2020) | 300B | Common Crawl, WebText2, Books, Wikipedia |
| Llama 2 (2023) | 2T | Publicly available web data |
| Llama 3.1 (2024) | ~15T | Web data, multilingual (8 languages) |
| DeepSeek-V3 (2024) | 14.8T | Diverse web + code + math data |
| Qwen 2.5 (2024) | 18T | Web, code, math, multilingual |
| OLMo 2 (2024) | Dolma dataset (~3T) | Fully open-source data |

### Diffusion Models — Training Data

| Model | Pretraining Tokens | Approach |
|-------|-------------------|----------|
| LLaDA 8B | ~2T (estimated) | Trained from scratch on web data |
| LLaDA 2.0 (16B/100B) | Inherited from AR base | Converted from pretrained AR models; 3-phase fine-tuning |
| Mercury | Undisclosed | Proprietary training pipeline |
| Evo 8B | Undisclosed | End-to-end hybrid AR-diffusion training |

**Key difference:** LLaDA 2.0 pioneered **knowledge inheritance** — converting already-trained AR models into diffusion models rather than pretraining from scratch. This dramatically reduces the compute needed to create competitive dLLMs.

---

## 6. Reinforcement Learning & Post-Training Alignment

### Autoregressive Models

| Model | RL / Alignment Method | Details |
|-------|----------------------|---------|
| **GPT-4 / ChatGPT** | RLHF (PPO) | Reward model trained on human preferences; PPO fine-tuning of policy model |
| **Claude 3.5 / 4** | RLHF + Constitutional AI (CAI) | Self-critique and revision using principles; RLAIF (RL from AI Feedback) |
| **Gemini 2.5** | RLHF | Details undisclosed; multi-stage alignment |
| **DeepSeek-R1** | **Pure RL (GRPO)** | Group Relative Policy Optimization without initial SFT; 2 RL stages + 2 SFT stages in full pipeline |
| **DeepSeek-R1-Zero** | RL only (no SFT) | Demonstrated reasoning emergent from RL alone |
| **Llama 3.1** | RLHF (PPO + DPO) | Multi-round alignment with rejection sampling |
| **Llama 4** | RLHF | Details in withdrawn technical report |
| **Qwen 2.5** | SFT + DPO + Online RL | Multi-stage post-training pipeline |
| **OLMo 2/3** | RLVR (RL with Verifiable Rewards) | Domain-specific RL-Zero variants for math, code |
| **Mistral Large 3** | DPO + SFT | Preference optimization |

#### Notable AR Reinforcement Learning Advances

- **DeepSeek-R1-Zero** showed that **reasoning can emerge purely from RL** without supervised fine-tuning, a landmark finding
- Google DeepMind's efficient exploration for RLHF achieves **10x data efficiency** (matching 200K-label performance with <20K labels)
- **Satori** (7B) internalizes self-reflection and autoregressive search via RL with Chain-of-Action-Thought (COAT)
- **ReflexiCoder** (Mar 2026) uses RL to achieve 40% better token efficiency through self-correction at inference time

### Diffusion Language Models

| Model | RL / Alignment Method | Details |
|-------|----------------------|---------|
| **LLaDA** | SFT + DPO | Standard preference optimization adapted for diffusion |
| **LLaDA 2.0** | SFT + DPO | Post-training alignment after AR-to-diffusion conversion |
| **Mercury 2** | Undisclosed | Tunable reasoning suggests RL-based training |
| **Research: GDPO** | Group Diffusion Policy Optimization | Sequence-level ELBO with semi-deterministic Monte Carlo; outperforms diffu-GRPO |
| **Research: DiFFPO** | Diffusion Fast and Furious PO | Off-policy RL with surrogate policies + joint efficient sampler training |
| **Research: dTRPO** | Trajectory Reduction in PO | Single-forward-pass design; +9.6% on STEM, +4.3% on coding |

#### Notable dLLM Reinforcement Learning Advances

- **GDPO** addresses the intractable likelihood problem in diffusion models by using ELBO as a surrogate for sequence-level likelihoods
- **DiFFPO** jointly trains efficient samplers that adaptively allocate inference-time compute, improving both accuracy and speed
- **dTRPO** reduces trajectory probability calculation costs, enabling offline RL training at scale for diffusion models
- Standard RLHF/DPO methods require adaptation for diffusion models because likelihoods are not directly available (must be approximated via ELBO)

---

## 7. Comparative Summary

| Dimension | Autoregressive LLMs | Diffusion LLMs |
|-----------|--------------------|--------------------|
| **Maturity** | Production-proven, dominant paradigm | Emerging, early commercial deployment |
| **Largest model** | 675B+ total params (Mistral Large 3) | 100B MoE (LLaDA 2.0-flash) |
| **Max disclosed training tokens** | 18T (Qwen 2.5) | ~2T (LLaDA 8B, est.) |
| **Inference speed** | 50–200 tok/s (frontier) | 500–1,100 tok/s (Mercury) |
| **Streaming support** | Native (token-by-token) | Limited (batch refinement) |
| **RL alignment** | Mature (RLHF, DPO, GRPO, RLVR) | Active research (GDPO, DiFFPO, dTRPO) |
| **Reasoning capability** | Strong (DeepSeek-R1, o1, Gemini 2.5) | Emerging (Mercury 2 tunable reasoning) |
| **Multimodal** | Common (Gemini, GPT-4o, Llama 4) | Early (MMaDA) |
| **Reversal curse** | Affected (unidirectional bias) | Naturally addressed (bidirectional) |
| **Open-source availability** | Extensive (Llama, Qwen, OLMo, DeepSeek) | Growing (LLaDA family) |

---

## 8. Key Takeaways

1. **Autoregressive models lead in scale and maturity**: The largest AR models train on 15–18T+ tokens with hundreds of billions of parameters. They have mature RL alignment pipelines (RLHF, DPO, GRPO) and strong reasoning capabilities.

2. **Diffusion models offer a speed advantage**: Mercury achieves 5–10x faster inference than comparable AR models by generating tokens in parallel. This throughput advantage is most relevant for latency-sensitive applications.

3. **Knowledge inheritance bridges the gap**: LLaDA 2.0 showed that pretrained AR model weights can be converted into diffusion models, avoiding the massive compute cost of pretraining dLLMs from scratch.

4. **RL for diffusion models is rapidly maturing**: Novel algorithms (GDPO, DiFFPO, dTRPO) address the unique challenge of intractable likelihoods in diffusion models, enabling preference-based alignment comparable to AR methods.

5. **The data wall affects both paradigms**: High-quality public text is estimated at only 15–20T tokens. Both AR and diffusion models increasingly rely on synthetic data, multi-epoch training, and data-efficient RL techniques.

6. **Hybrid approaches are emerging**: Models like Evo combine AR and diffusion paradigms, suggesting the future may not be either/or but a continuum of generation strategies.

---

*Sources: arXiv papers (2502.09992, 2512.15745, 2501.12948, 2510.08554, 2510.02212, 2603.18806, 2603.06617), Inception Labs blog posts, Meta AI blog, Qwen technical report, DeepSeek V3 technical report, Google DeepMind publications, AI2 OLMo documentation.*
