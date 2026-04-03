# Equity-research

## GPU & TPU Pareto Frontier: Rental Cost vs Output Token Cost

Analysis of GPU and TPU accelerator rental costs per hour across cloud providers, mapped against LLM inference output token costs to identify the Pareto frontier of cost-efficient accelerators.

### Files

| File | Description |
|------|-------------|
| `gpu_tpu_pareto_frontier.xlsx` | Multi-sheet workbook with rental costs, token economics, Pareto frontier, and sources |
| `pareto_frontier_chart.png` | Visualization of the Pareto frontier |
| `gpu_tpu_pareto_analysis.py` | Python script to regenerate the analysis |

### Accelerators Covered

**NVIDIA GPUs:** T4, L4, L40S, A100 40GB, A100 80GB, H100 SXM, H200 SXM, B200 SXM
**AMD GPUs:** MI300X
**Google TPUs:** v5e, v5p, v6e (Trillium)

### Pareto Frontier (April 2026)

The Pareto frontier identifies accelerators where no other option is both cheaper per hour AND cheaper per output token:

| SKU | $/hr | $/M Output Tokens |
|-----|------|-------------------|
| NVIDIA T4 | $0.35 | $3.24 |
| Google TPU v5e | $0.60 | $2.22 |
| NVIDIA L40S | $0.89 | $2.06 |
| Google TPU v6e | $1.35 | $1.72 |
| Google TPU v5p | $2.10 | $1.46 |
| AMD MI300X | $2.50 | $1.34 |
| NVIDIA H100 SXM | $2.89 | $0.92 |
| NVIDIA H200 SXM | $3.59 | $0.64 |
| NVIDIA B200 SXM | $4.81 | $0.34 |

Benchmark model: Llama 70B-class (FP8/INT8 where available)
