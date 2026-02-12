# LLM Token Consumption: Open Source vs Proprietary

## US and Rest of World — Time Series Analysis (Q1 2023 – Q1 2026)

This analysis estimates quarterly token consumption across **open-source** and **proprietary** large language models, broken down by **US** vs **Rest of World** (with sub-regional granularity for China/East Asia, Europe, and emerging markets).

### Key Findings

| Metric | Q1 2023 | Q1 2026 | Trend |
|--------|---------|---------|-------|
| Global token consumption | ~3.5T tokens/quarter | ~99T tokens/quarter | ~28x growth |
| Open-source global share | ~14% | ~56% | Crossed 50% in H2 2025 |
| US open-source share | ~14% | ~44% | Accelerating rapidly |
| China/Asia open-source share | ~15% | ~83% | Dominant since mid-2024 |
| US share of global consumption | ~60% | ~52% | RoW catching up |

### Structural Shifts

1. **Open source has won the volume game globally** — more than half of all LLM tokens are now processed by open-weight models
2. **The US remains the proprietary stronghold**, but enterprise open-source deployment is accelerating
3. **China/Asia is overwhelmingly open source** — driven by geopolitical constraints and strong domestic models (Qwen, DeepSeek)
4. **GPT family dominance has ended** — from ~62% to ~15% share, despite growing absolute volumes
5. **Both segments still growing** — proprietary ~14.5x, open source ~111x since Q1 2023

### Files

| File | Description |
|------|-------------|
| `llm_token_consumption_analysis.ipynb` | Full analysis notebook with all visualizations |
| `llm_token_consumption_data.xlsx` | Raw data export (quarterly data, model families, heatmap) |
| `01_global_token_consumption_overview.png` | Global absolute + share view |
| `02_us_vs_row_trajectories.png` | US vs RoW volume and share |
| `03_open_source_share_by_region.png` | Open-source adoption by region |
| `04_four_segment_stacked.png` | Four-segment breakdown (US/RoW × Prop/OSS) |
| `05_regional_deep_dive.png` | Sub-regional analysis (US, China, Europe, Other) |
| `06_model_family_market_share.png` | Model family evolution (GPT, Llama, DeepSeek, etc.) |
| `07_growth_rate_analysis.png` | QoQ growth rates and cumulative growth index |
| `08_interactive_dashboard.html` | Interactive Plotly dashboard |
| `09_heatmap_oss_share.png` | Heatmap of open-source share by region and quarter |

### Data Sources

Estimates synthesized from: Similarweb traffic data, Hugging Face download statistics, a16z/Menlo Ventures AI surveys, SemiAnalysis GPU deployment estimates, public financial disclosures (OpenAI ARR, Microsoft/Google AI revenue), Meta Llama download announcements, DeepSeek API traffic data, Artificial Analysis API benchmarks, and Stanford HAI AI Index 2025.

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
