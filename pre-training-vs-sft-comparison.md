# Pre-Training vs Supervised Fine-Tuning: Gemini, OpenAI, and Chinese Open-Source Models

## Executive Summary

The LLM industry is undergoing a structural shift in how training compute is allocated. From 2023 to 2026, the ratio has moved from roughly 75% pre-training / 25% post-training to approximately 45% pre-training / 55% post-training. Each major model family — OpenAI (GPT), Google (Gemini), and leading Chinese open-source labs (DeepSeek, Qwen/Alibaba, Yi/01.AI) — navigates this transition differently, with meaningful implications for cost structure, capability differentiation, and ecosystem dynamics.

---

## 1. Industry-Wide Context: The Shift from Pre-Training to Post-Training

| Dimension | 2023 Era | 2025-2026 Era |
|---|---|---|
| Compute allocation | ~75% pre-training | ~45% pre-training |
| Primary scaling lever | More tokens, bigger models | RL, SFT, inference-time compute |
| Binding constraint | GPU supply | High-quality data exhaustion (projected 2026-2028) |
| Cost driver | Training FLOPS | Inference cost per query |
| Competitive moat | Training data volume | Post-training data quality + RL techniques |

Key post-training efficiency gains driving this reallocation:
- DPO delivers RLHF-equivalent performance with ~40% less compute
- Expert data curation yields 25-35% accuracy improvement over raw-data baselines
- RLHF doubled GPT-4's accuracy on adversarial questions
- DeepSeek-R1 approaches GPT-4 performance via pure RL with test-time compute rather than massive pre-training

---

## 2. OpenAI (GPT Series)

### 2.1 Pre-Training

| Model | Pre-Training Tokens | Parameters | Estimated Pre-Training Compute | Estimated Cost |
|---|---|---|---|---|
| GPT-4 (Mar 2023) | ~13T tokens (2 epochs text, 4 epochs code) | ~1.8T (MoE, 120 layers) | ~2.15 x 10^25 FLOPS | ~$63M (A100); ~$21.5M repriced to H100 |
| GPT-4.5 (Feb 2025) | Not disclosed | Not disclosed | Larger than GPT-4 | ~$200M (estimated) |
| GPT-5 (Aug 2025) | ~70T cleaned tokens | Not disclosed | **Less than GPT-4.5** (~10x reduction) | ~$20M pre-training (but $500M+ per full training run including post-training) |

**Critical finding:** GPT-5 used *less* pre-training compute than GPT-4.5, breaking the generational trend. OpenAI shifted investment into post-training, finding that new reasoning techniques developed around September 2024 made it possible to reduce pre-training compute by roughly 10x while maintaining equivalent performance. GPT-6 is expected to return to higher pre-training compute as post-training techniques mature at scale.

### 2.2 Post-Training / SFT Approach

OpenAI's post-training pipeline has evolved significantly:

- **GPT-4 era:** SFT + RLHF (standard InstructGPT methodology)
- **o1/o3 era (late 2024):** Introduction of reasoning-focused RL, chain-of-thought training
- **GPT-5 (Aug 2025):** Unified system architecture combining:
  - A fast main model (`gpt-5-main`) for routine queries
  - A deep reasoning model (`gpt-5-thinking`) trained via RL for complex problems
  - A real-time router (continuously trained on user preferences, correctness signals)
  - "Scalable techniques that enable training larger models with data derived from smaller models"

### 2.3 Fine-Tuning-as-a-Service

OpenAI offered supervised fine-tuning via API for enterprise customers on models including GPT-4.1, GPT-4.1-mini, and GPT-4.1-nano. Notable results include Cosine achieving SOTA on SWE-bench (43.8%) with fine-tuned GPT-4o. **However, as of May 8, 2026, OpenAI is winding down the fine-tuning platform** — no longer accessible to new users. This signals a strategic move away from customer-driven SFT toward OpenAI-controlled post-training.

### 2.4 Investment Context

The GPT-5 project (codenamed Orion) consumed ~$19.6B over two years. Individual six-month training runs cost over $500M in compute alone, with at least two large runs that fell short of expectations. The two-year struggle to ship GPT-5 was a key catalyst for the pivot from pre-training scaling to post-training and reasoning scaling.

---

## 3. Google DeepMind (Gemini Series)

### 3.1 Pre-Training

| Model | Architecture | Pre-Training Data | Context Window | Infrastructure |
|---|---|---|---|---|
| Gemini 1.5 (Feb 2024) | Sparse MoE Transformer | Multimodal (text, code, images, audio, video) | 1M tokens | TPU v5p pods |
| Gemini 2.0 (Dec 2024) | Sparse MoE Transformer | Expanded multimodal | 1M tokens | TPU v5p/v6 pods |
| Gemini 2.5 Pro (Mar 2025) | Sparse MoE Transformer | Large-scale diverse data: web docs, code, images, audio, video | 1M tokens input / 64K output | TPU pods |

Google has not disclosed specific token counts or compute costs for Gemini models. The sparse MoE architecture is a key design choice: it decouples total model capacity from per-token computation cost, allowing larger effective models with lower inference costs.

### 3.2 Post-Training / SFT Approach

Google employs a two-stage post-training pipeline:

1. **Supervised Fine-Tuning (SFT):** Curated instruction tuning data with multimodal paired instructions and responses
2. **Preference Tuning (RLHF-equivalent):** Learns from human feedback data using prompts paired with preferred/dispreferred responses

Google's recommended workflow: apply SFT on preferred-response data first, then continue with preference tuning to increase the likelihood gap between good and bad outputs. Post-training also incorporates tool-use data for agentic capabilities.

### 3.3 Fine-Tuning-as-a-Service (Vertex AI)

Google offers a fully managed SFT experience via Vertex AI for Gemini 2.5 Pro, Flash, and Flash-Lite:

- Multimodal fine-tuning support (text, image, audio, video, document)
- No GPU provisioning required — fully managed
- Dataset: minimum 100 examples recommended; up to 10M text-only or 300K multimodal examples
- LoRA-based adapter sizes (1, 2, 4, 8, or 16)
- Preference tuning also available as a separate service

**Strategic difference from OpenAI:** Google continues to invest in customer-facing fine-tuning infrastructure as a core Vertex AI offering, while OpenAI is pulling back.

### 3.4 Key Differentiators

- **Native multimodality** in both pre-training and fine-tuning (images, audio, video from the ground up)
- **TPU infrastructure** provides cost advantages vs NVIDIA GPU pricing
- **Opacity on pre-training details** makes compute/data scale comparison difficult vs. more transparent Chinese labs

---

## 4. Chinese Open-Source Models

### 4.1 DeepSeek (DeepSeek AI)

#### Pre-Training

| Model | Pre-Training Tokens | Parameters (Total / Active) | Compute Cost | Infrastructure |
|---|---|---|---|---|
| DeepSeek-V3 (Dec 2024) | 14.8T tokens | 671B / 37B (MoE) | 2.788M H800 GPU-hours (~$5.6M) | 2,048 H800 GPUs |
| DeepSeek-R1 (Jan 2025) | Built on V3 base | 671B / 37B (MoE) | Not separately disclosed (post-training only on V3 base) | Not disclosed |

**Pre-training cost of $5.6M for a 671B-parameter model is a landmark result**, roughly 10-100x cheaper than comparable Western frontier models. Key innovations enabling this:
- Auxiliary-loss-free load balancing for MoE
- Multi-token prediction training objective
- FP8 mixed-precision training
- Zero irrecoverable loss spikes or rollbacks during the entire training run

Pre-training consumed 2.664M GPU-hours (95.5%), context extension 119K (4.3%), and post-training only 5K (0.2%).

#### Post-Training / SFT Approach

DeepSeek's post-training is the most technically innovative among the three groups:

**DeepSeek-R1-Zero (Pure RL, No SFT):**
- Trained via large-scale RL using Group Relative Policy Optimization (GRPO) directly on DeepSeek-V3-Base
- No supervised fine-tuning at all as a preliminary step
- Reward signal: purely correctness of final answers against ground truth
- Emergent behaviors: self-reflection, verification, dynamic strategy adaptation — all without explicit instruction
- AIME 2024 pass@1: 15.6% → 71.0% (86.7% with majority voting)

**DeepSeek-R1 (Multi-stage):**
- Cold-start SFT data → Reasoning-based RL (GRPO) → Rejection sampling SFT → General RL
- Addresses R1-Zero's readability and language-mixing issues
- Performance comparable to OpenAI o1

**DeepSeek-R1 Distill Models:**
- Fine-tuned smaller open-source models (Qwen, Llama bases) on samples generated by DeepSeek-R1
- Available in 1.5B, 7B, 8B, 14B, 32B, and 70B variants
- 621+ community fine-tuned variants on Hugging Face

#### Open-Source Ecosystem

All model weights released openly. The cross-pollination between DeepSeek and Qwen bases (e.g., DeepSeek-R1-Distill-Qwen models) creates a synergistic open-source ecosystem that is absent in the OpenAI/Google world.

### 4.2 Qwen (Alibaba Cloud)

#### Pre-Training

| Model | Pre-Training Tokens | Parameters (Total / Active) | Languages |
|---|---|---|---|
| Qwen2 (mid 2024) | 7T tokens | Up to 72B | 29 languages |
| Qwen2.5 (Sep 2024) | 18T tokens | Up to 72B | 29 languages |
| Qwen3 (Apr 2025) | 36T tokens | 235B / 22B (MoE flagship) | 119 languages |

Qwen shows aggressive token-count scaling: 7T → 18T → 36T across three generations in under a year. The Qwen3 family spans 0.6B to 235B parameters.

#### Data Pipeline Innovation

Qwen3's pre-training pipeline uses earlier Qwen models (Qwen2.5-VL, Qwen2.5-Math, Qwen2.5-Coder) as a "data factory" to generate synthetic training data at scale through:
- Text recognition and extraction
- Mathematical problem synthesis
- Code generation

A sophisticated multilingual annotation system labels 30T of the 36T tokens by educational value, domain, and safety to optimize data mixture.

#### Post-Training / SFT Approach

Qwen3's post-training uses a four-stage pipeline:
1. Long Chain-of-Thought (CoT) cold start SFT
2. Reasoning-based RL
3. Thinking/non-thinking mode training (two curated distributions, not simple switches)
4. Long-CoT data filtering around verifiable answers

**Key distinction vs DeepSeek:** Qwen relies more heavily on SFT and domain-specific pre-training (code, math data), with RL as a secondary optimization step. DeepSeek emphasizes RL as the primary capability driver (SFT → RL → SFT → RL cycles).

### 4.3 Yi (01.AI)

Yi-Lightning uses an enhanced MoE architecture with advanced expert segmentation and routing. The training pipeline includes comprehensive pre-training, SFT, and RLHF with multi-stage training and synthetic data construction. Specific pre-training scale details are not publicly disclosed.

---

## 5. Comparative Analysis

### 5.1 Pre-Training Scale and Approach

| Dimension | OpenAI (GPT-5) | Google (Gemini 2.5) | DeepSeek V3 | Qwen3 |
|---|---|---|---|---|
| Pre-training tokens | ~70T | Not disclosed | 14.8T | 36T |
| Parameter count | Not disclosed | Not disclosed (MoE) | 671B / 37B active | 235B / 22B active |
| Architecture | Dense → Router ensemble | Sparse MoE | Sparse MoE (MLA) | Sparse MoE |
| Pre-training cost | $500M+ per run (full) | Not disclosed (TPU) | ~$5.6M | Not disclosed |
| Hardware | NVIDIA H100/H200 | Google TPU v5p/v6 | NVIDIA H800 | NVIDIA GPUs |
| Data openness | Closed | Closed | Detailed technical report | Detailed technical report |

### 5.2 Post-Training Philosophy

| Dimension | OpenAI | Google | DeepSeek | Qwen |
|---|---|---|---|---|
| Primary method | SFT + RLHF + RL for reasoning | SFT → Preference Tuning | **Pure RL (GRPO) pioneer**; multi-stage SFT+RL | SFT-heavy + RL secondary |
| Innovation focus | Reasoning routing, model ensembles | Multimodal preference tuning | RL without SFT (R1-Zero); emergent reasoning | Synthetic data factory; multilingual scale |
| RL intensity | High (o-series, GPT-5 thinking) | Moderate | Very High | Moderate |
| SFT role | Foundation for instruction following | Core post-training stage | Cold-start correction tool | Primary post-training method |
| Customer fine-tuning | **Winding down** (May 2026) | Active (Vertex AI, fully managed) | Open weights — community fine-tunes | Open weights — community fine-tunes |

### 5.3 Strategic Positioning

| | OpenAI | Google | Chinese Open-Source |
|---|---|---|---|
| **Business model** | Closed API, subscription | Cloud platform integration | Open weights + API services |
| **Pre-training moat** | Scale + data deals | TPU cost advantage + multimodal data | Cost efficiency (10-100x cheaper) |
| **Post-training moat** | Proprietary RL + reasoning techniques | Multimodal preference data | Published techniques, community iteration |
| **Fine-tuning strategy** | Centralizing (shutting down customer SFT) | Expanding (Vertex AI) | Decentralized (open weights) |
| **Risk** | Capital intensity; $500M+ per training run | Slower iteration vs. nimble labs | US export controls on advanced GPUs |

---

## 6. Key Takeaways for Equity Research

1. **Pre-training cost asymmetry is extreme.** DeepSeek trained a frontier-competitive 671B model for ~$5.6M. OpenAI's GPT-5 runs cost $500M+. This 100x gap challenges the assumption that massive capital expenditure is required for frontier AI.

2. **The post-training revolution favors capital-efficient players.** As the industry shifts from pre-training to post-training as the primary scaling lever, the compute advantage of well-funded Western labs narrows. DeepSeek's pure-RL approach (R1-Zero) demonstrated that reasoning capabilities can emerge without expensive human-labeled SFT data.

3. **OpenAI is consolidating while Google is expanding fine-tuning access.** OpenAI's decision to wind down its fine-tuning platform (May 2026) suggests a strategic shift toward controlling the full training pipeline. Google is moving in the opposite direction, making fine-tuning a core Vertex AI offering. Chinese labs take a third path: open weights enabling unrestricted community fine-tuning.

4. **Data exhaustion is the binding constraint.** High-quality human-generated text data is projected to be exhausted by 2026-2028. All three groups are investing in synthetic data generation, but Qwen's "data factory" approach (using prior models to generate training data) and DeepSeek's pure-RL approach (reducing dependence on labeled data) represent distinct strategic responses.

5. **Hardware matters for cost structure.** Google's TPU advantage creates opacity but likely cost advantages. Chinese labs' demonstrated efficiency on H800s (a reduced-capability export-compliant GPU) is remarkable. US export restrictions on advanced chips remain a tail risk for Chinese AI labs.

6. **Open-source creates ecosystem moats.** The DeepSeek-Qwen cross-pollination (e.g., R1-Distill-Qwen models with 621+ community fine-tunes) generates compounding network effects absent in the closed OpenAI/Google ecosystems. This community flywheel acts as a distribution and improvement mechanism that closed models cannot replicate.

7. **Inference cost, not training cost, is becoming the key economic metric.** As post-training and inference-time compute (e.g., chain-of-thought reasoning at serve time) grow, the marginal cost per query displaces the fixed cost of training as the primary concern for operators and investors.

---

*Last updated: May 2026*
*Sources: OpenAI system cards and blog posts, Google Gemini model cards and Vertex AI documentation, DeepSeek technical reports (arXiv), Qwen3 technical report (arXiv), Epoch AI analysis, industry reports.*
