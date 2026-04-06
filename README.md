# KV Cache & Prefill Usage for Frontier LLMs

## Memory, Cost, and Performance Analysis (2024–2026)

Comprehensive analysis of KV cache memory requirements, prefill throughput characteristics, prompt caching economics, and optimization strategies across frontier large language models.

### Key Findings

| Metric | Finding | Impact |
|--------|---------|--------|
| MLA compression (DeepSeek V3.2) | ~55× KV cache reduction vs MHA | 2.3 GB at 128K vs 131 GB equivalent |
| GQA adoption | 8–32× compression; universal in Llama, Qwen, likely Claude/GPT | Pragmatic default for all frontier models |
| Prefill/decode asymmetry | Prefill is 100–400× faster than decode | TTFT scales well; long-context still multi-second |
| Anthropic/Google cache discount | 90% off cached input tokens | 2× more impactful than OpenAI's 50% discount |
| Claude Code cache hit rate | 84% average; 97.2% in agent teams | 74% cost reduction per request |
| Agent I/O ratio | 166:1 (Claude Code); >99% of tokens are input | KV cache = dominant cost driver for agents |
| DeepSeek V3.2 absolute cost | $0.28/M input; 7–18× cheaper than alternatives | Cost leader even without caching |
| TurboQuant (Google, March 2026) | 6× KV cache compression, zero quality loss | Potential breakthrough for self-hosted inference |
| SGLang RadixAttention | 29% higher throughput vs vLLM for prefix workloads | Best framework for agent/shared-prefix serving |
| Consumer GPU prefill | ~4.5 tok/s on modest GPU | 300s to prefill 80K tokens — prohibitive |
| **HBM supply-demand gap** | Demand exceeds supply by ~33% in 2026 | SK Hynix & Micron sold out through 2026; gap persists to 2028–2029 |
| **HBM per GPU: 3.6× growth** | 80 GB (A100) → 288 GB (B300) in 5 years | HBM now 35–38% of GPU manufacturing cost |
| **KV cache = demand multiplier** | 83% of VRAM at 128K × 8 users | Turns memory shortage into memory crisis |
| **Inference flip** | 67% of all AI compute is inference (2026) | Inference is KV-cache-dominated; compounds memory demand |
| **$54.6B HBM market (2026)** | Growing to $100B+ by 2028 | New fab capacity won't close gap until 2028–2029 |
| **Near-duopoly risk** | SK Hynix ~60% share; Samsung quality issues | Only 2 reliable suppliers for most critical AI component |

### Models Covered

GPT-4.1, GPT-4o, Claude Sonnet 4.6, Claude Opus 4.6, Gemini 2.5 Pro, DeepSeek V3.2, Llama 4 Maverick, Llama 4 Behemoth, Qwen3-235B, Llama 3.3 70B

### Files

| File | Description |
|------|-------------|
| `KV_Cache_Prefill_Report.md` | **Standalone report** capturing all key findings, training vs inference KV demand, and investment implications |
| `kv_cache_prefill_analysis.ipynb` | Full analysis notebook with 19 visualizations + interactive dashboard |
| `kv_cache_prefill_data.xlsx` | All data tables (18 sheets: architectures, pricing, benchmarks, optimizations, frameworks, agents, GPU budget, HBM supply-demand, GPU HBM requirements, installed base memory, KV cache scenarios, agent KV comparison, KV quantization methods, parameter scaling, training vs inference memory, compute split timeline, multi-turn cost model, cost sensitivity) |
| `01_kv_cache_memory_scaling.png` | KV cache memory scaling by context length across models |
| `02_attention_mechanism_kv_impact.png` | MHA vs GQA vs MLA per-token KV cache comparison |
| `03_prefill_vs_decode_throughput.png` | Prefill and decode throughput by model and framework |
| `04_ttft_sensitivity_input_length.png` | Time-to-first-token sensitivity (with/without cache) |
| `05_prompt_caching_economics.png` | Fresh vs cached pricing and discount rates by provider |
| `06_cost_sensitivity_cache_hit_rate.png` | Cost per request across cache hit rates |
| `07_agent_workload_profile.png` | Agent I/O ratios and cache hit rate scatter |
| `08_kv_cache_optimization_timeline.png` | Optimization technique timeline by stack layer |
| `09_monthly_cost_model.png` | Monthly cost projections at various usage levels |
| `10_gpu_memory_budget.png` | Model weights vs KV cache memory breakdown |
| `11_concurrent_users_vs_gpu.png` | Concurrent user capacity by GPU configuration |
| `12_hbm_supply_demand_mismatch.png` | HBM supply vs demand, revenue, per-GPU requirements, vendor share |
| `12_interactive_dashboard.html` | Interactive Plotly dashboard (4-panel) |
| `13_kv_cache_memory_demand_multiplier.png` | KV cache as VRAM demand multiplier + inference flip |
| `14_aggregate_memory_gap.png` | Global installed base memory breakdown and KV cache share |
| `15_openclaw_cowork_perplexity_comparison.png` | OpenClaw vs Claude Cowork vs Perplexity Computer: prompt size, cache rates, costs, invalidation risks, scorecard |
| `16_multi_turn_cost_comparison.png` | Multi-turn session cost escalation across the three platforms |
| `17_turboquant_memory_impact.png` | TurboQuant impact: KV reduction, concurrent users, quantization comparison, HBM demand |
| `18_parameter_scaling_kv_cache.png` | Parameter scaling vs KV cache: log-log scatter, KV share, decomposition, FP16 vs TurboQuant |
| `19_training_vs_inference_kv_demand.png` | Training vs inference memory breakdown, compute spending split, KV share over context eras |

### Data Sources

- Model architecture specifications: DeepSeek V3 technical report, Llama 4 release blog, Qwen3 model card, published benchmarks
- Pricing: OpenAI, Anthropic, Google, DeepSeek official API pricing pages (March 2026)
- Inference benchmarks: vLLM/SGLang/TensorRT-LLM published comparisons, NVIDIA DGX Spark specs, Artificial Analysis
- Agent workload data: Claude Code usage statistics, GitHub Copilot engineering blog, Cursor technical posts
- Optimization research: TurboQuant (Google, March 2026), KV-Compress (2025), PagedAttention (vLLM), RadixAttention (SGLang), NVIDIA Dynamo docs
- HBM market data: SK Hynix/Samsung/Micron financials, TrendForce, Yole Group, CNBC, Epoch AI GPU tracker
- GPU installed base: Epoch AI (15M H100-equivalents), NVIDIA GTC 2026 disclosures, SemiAnalysis estimates

### Requirements

```
pandas
numpy
matplotlib
seaborn
plotly
openpyxl
kaleido
```
