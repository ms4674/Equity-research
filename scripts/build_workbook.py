#!/usr/bin/env python3
"""Build the AI lab compute-economics workbook.

Aggregates estimated quarterly time series (Q1 2024 - Q3 2026) for frontier
(closed) and open-weight AI labs across four metrics:
  * operational GW compute capacity (owned + leased, end of quarter)
  * revenue (quarterly booked + annualized run rate)
  * compute cost (estimated from industry $/GW benchmarks)
  * token volume (trillions of tokens processed per quarter)
plus derived revenue-per-GW and cost-per-GW metrics.

Values marked as disclosed in the Sources sheet come from public reporting;
everything else is an estimate interpolated between disclosed anchors.

Output: ai_lab_compute_economics_quarterly.xlsx (repo root).
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart import Series as ChartSeries
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Data (see Sources & Methodology sheet for anchors behind every series)
# ---------------------------------------------------------------------------

QUARTERS = [
    "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024",
    "Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025",
    "Q1 2026", "Q2 2026", "Q3 2026 (part.)",
]
N_Q = len(QUARTERS)
LAST_FULL_Q_IDX = 9  # Q2 2026 — last complete quarter

LABS = [
    ("OpenAI", "Frontier (closed)"),
    ("Anthropic", "Frontier (closed)"),
    ("Google DeepMind", "Frontier (closed)"),
    ("xAI", "Frontier (closed)"),
    ("Meta (Llama)", "Open-weight"),
    ("DeepSeek", "Open-weight"),
    ("Alibaba (Qwen)", "Open-weight"),
    ("Mistral AI", "Open-weight"),
]

# Operational GW of AI compute serving each lab (owned + leased IT power,
# end of quarter, estimated).
GW = {
    "OpenAI":          [0.25, 0.30, 0.35, 0.45, 0.55, 0.65, 0.80, 1.00, 1.20, 1.40, 1.70],
    "Anthropic":       [0.08, 0.10, 0.12, 0.15, 0.20, 0.30, 0.45, 0.70, 1.20, 1.80, 2.20],
    "Google DeepMind": [0.80, 0.90, 1.00, 1.20, 1.40, 1.60, 1.90, 2.20, 2.50, 2.80, 3.20],
    "xAI":             [0.02, 0.05, 0.15, 0.20, 0.25, 0.30, 0.40, 0.60, 0.90, 1.20, 1.40],
    "Meta (Llama)":    [1.00, 1.15, 1.30, 1.50, 1.70, 1.95, 2.20, 2.50, 3.00, 3.60, 4.20],
    "DeepSeek":        [0.010, 0.012, 0.014, 0.016, 0.018, 0.020, 0.022, 0.024, 0.026, 0.030, 0.035],
    "Alibaba (Qwen)":  [0.25, 0.30, 0.35, 0.40, 0.50, 0.60, 0.70, 0.85, 1.00, 1.20, 1.40],
    "Mistral AI":      [0.005, 0.006, 0.008, 0.010, 0.012, 0.015, 0.018, 0.025, 0.032, 0.040, 0.050],
}

# Quarterly booked revenue, $B. None = not applicable / not separable.
REV_Q = {
    "OpenAI":          [0.60, 0.80, 1.00, 1.30, 2.20, 2.90, 3.60, 4.40, 5.90, 7.20, 9.50],
    "Anthropic":       [0.15, 0.20, 0.28, 0.37, 0.45, 0.79, 1.30, 2.00, 4.80, 11.50, 16.00],
    "Google DeepMind": [0.90, 1.10, 1.40, 1.60, 2.60, 3.30, 4.10, 5.00, 6.20, 7.50, 8.80],
    "xAI":             [0.01, 0.02, 0.03, 0.04, 0.06, 0.09, 0.11, 0.14, 0.82, 0.45, 0.50],
    "Meta (Llama)":    [None] * 11,
    "DeepSeek":        [0.001, 0.002, 0.003, 0.004, 0.02, 0.04, 0.06, 0.08, 0.12, 0.15, 0.17],
    "Alibaba (Qwen)":  [0.28, 0.35, 0.44, 0.55, 0.68, 0.85, 1.05, 1.25, 1.50, 1.82, 2.20],
    "Mistral AI":      [0.004, 0.006, 0.008, 0.012, 0.02, 0.035, 0.055, 0.075, 0.11, 0.20, 0.26],
}

# Annualized revenue run rate at end of quarter, $B.
ARR = {
    "OpenAI":          [2.8, 3.4, 4.4, 5.5, 9.5, 12.0, 16.0, 21.5, 25.0, 31.0, 40.0],
    "Anthropic":       [0.6, 0.85, 1.2, 1.6, 2.2, 3.1, 5.5, 9.0, 22.0, 47.0, 65.0],
    "Google DeepMind": [3.8, 4.6, 5.6, 6.8, 9.0, 12.0, 15.0, 18.0, 24.0, 30.0, 34.0],
    "xAI":             [0.03, 0.06, 0.10, 0.16, 0.20, 0.30, 0.40, 0.50, 1.00, 1.50, 1.80],
    "Meta (Llama)":    [None] * 11,
    "DeepSeek":        [0.004, 0.008, 0.012, 0.016, 0.08, 0.16, 0.24, 0.32, 0.50, 0.70, 0.80],
    "Alibaba (Qwen)":  [1.1, 1.4, 1.8, 2.2, 2.7, 3.4, 4.2, 5.0, 6.0, 7.3, 8.8],
    "Mistral AI":      [0.008, 0.012, 0.016, 0.020, 0.06, 0.12, 0.20, 0.312, 0.50, 1.00, 1.20],
}

# Tokens processed per quarter across all surfaces, trillions (estimated).
TOKENS = {
    "OpenAI":          [80, 120, 170, 260, 400, 640, 1060, 1650, 2590, 3890, 4950],
    "Anthropic":       [15, 20, 28, 37, 45, 80, 130, 220, 500, 1200, 1700],
    "Google DeepMind": [17, 30, 80, 210, 560, 1380, 2340, 3800, 6180, 10050, 15300],
    "xAI":             [1, 2, 3, 5, 8, 15, 25, 40, 60, 90, 120],
    "Meta (Llama)":    [10, 20, 35, 55, 80, 120, 160, 210, 280, 360, 430],
    "DeepSeek":        [1, 2, 4, 8, 40, 80, 120, 200, 350, 600, 800],
    "Alibaba (Qwen)":  [2, 4, 8, 15, 30, 60, 110, 180, 300, 500, 700],
    "Mistral AI":      [1, 2, 3, 5, 8, 12, 18, 25, 35, 50, 65],
}

# Industry all-in build capex per GW ($B/GW), by GPU generation in deployment.
BUILD_COST_PER_GW = [32, 32, 33, 34, 35, 36, 38, 40.5, 42, 45, 47]

# Annualized total cost of ownership per operational GW ($B/GW/yr, Epoch AI).
TCO_PER_GW_YR = 8.5

LAB_NOTES = {
    "OpenAI": "Compute mostly leased (Microsoft Azure, Oracle, Stargate JV). Abilene ~0.3 GW live mid-2026; 9+ GW contracted through 2029.",
    "Anthropic": "Compute leased from AWS (Project Rainier / New Carlisle ~1 GW live Mar 2026), Google TPU (1 GW+ landing 2026), SpaceX (300 MW). Best-documented revenue ramp.",
    "Google DeepMind": "TPU fleet allocated to Gemini training + inference (subset of Alphabet's total fleet). Revenue = Gemini/Vertex/Workspace-AI attributable estimate, not an Alphabet segment.",
    "xAI": "Self-built Colossus 1 (~0.3 GW, Memphis) + Colossus 2 ramping toward 2 GW. Standalone AI revenue only (excludes X ads/subscriptions).",
    "Meta (Llama)": "AI accelerator fleet across all workloads (GenAI + ranking/recsys). No direct model revenue - monetized via ads; Llama licensed free. Plans 7 GW deployed during 2026, 14 GW by 2027.",
    "DeepSeek": "~20k H100-equivalent fleet (mid-2026 disclosure); extreme efficiency focus; 1 GW Ulanqab site planned for late 2027/28. Revenue = API hosting only; open weights are free.",
    "Alibaba (Qwen)": "Revenue = Alibaba's disclosed 'AI-related product revenue' (calendar-quarter mapped). GW = estimated AI portion of Alibaba Cloud fleet. Qwen weights free (3B+ downloads).",
    "Mistral AI": "European open-weight lab. Small self + leased footprint (Eclairion France DC, Sweden DC committed). ARR $312M Dec 2025 -> ~$1B May 2026.",
}

# Selected disclosed anchor cells to render in bold (metric key -> lab -> quarter indices).
DISCLOSED = {
    "REV_Q": {
        "Anthropic": [5, 9],           # Q2'25 $787M, Q2'26 $11.5B (both disclosed)
        "Alibaba (Qwen)": [9],         # Q2'26 $1.82B disclosed (SEC filing)
        "xAI": [8],                    # Q1'26 $818M reported
    },
    "ARR": {
        "OpenAI": [7, 8, 10],          # end-2025 >$20B, Q1'26 $25B, Aug'26 $40B
        "Anthropic": [7, 9, 10],       # late-2025 >$9B, May'26 $47B, Jul'26 $65B
        "Mistral AI": [7, 9],          # Dec'25 $312M, May'26 ~$1B
        "DeepSeek": [10],              # mid-2026 ~$500-700M annualized
    },
}

SOURCES = [
    # (Metric, Lab/Scope, Data point, As of, Source, URL)
    ("Revenue", "Anthropic", "Annualized run rate $65B (7x YoY); Q2 2026 revenue $11.5B prelim (14x YoY vs $787M Q2 2025); positive adjusted operating income in Q2 2026", "Jul-Aug 2026", "CNBC / Bloomberg", "https://www.cnbc.com/2026/08/17/anthropic-says-annualized-revenue-climbed-to-65-billion-in-july.html"),
    ("Revenue", "Anthropic", "Run rate >$47B (May 2026), ~$30B (Apr 2026), >$9B (late 2025); ~$10B revenue for all of 2025 (gross basis)", "2025-2026", "CNBC / Bloomberg / Yahoo Finance", "https://finance.yahoo.com/technology/ai/articles/anthropic-revenue-run-rate-surpasses-193745178.html"),
    ("Revenue", "OpenAI", "Annualized run rate >$40B (Aug 2026), ~2x from end-2025; monthly run rate +20% in July alone", "Aug 2026", "Bloomberg", "https://www.bloomberg.com/news/articles/2026-08-13/openai-s-revenue-run-rate-tops-40-billion-ahead-of-ipo"),
    ("Revenue", "OpenAI", "2024 revenue $3.7B; 2025 booked revenue $13.07B (audited, leaked) with $20.9B operating loss; ARR $21.5B end-2025, $25B Q1 2026 ($2B/month); enterprise >40% of revenue Apr 2026", "2024-2026", "ValueAdd VC / Happycapy compilation of leaked financials", "https://valueaddvc.com/blog/openai-revenue-2026-20b-arr-4b-month-path-to-profitability"),
    ("Revenue", "Google DeepMind", "AI-attributable revenue est.: ~$5B 2024, ~$15B 2025, ~$28B ARR May 2026 (not an Alphabet reporting segment)", "May 2026", "Presenc AI comparison table", "https://presenc.ai/research/ai-lab-revenue-and-valuation-2026"),
    ("Revenue", "xAI", "Standalone AI revenue ~$100M 2024, ~$350-500M 2025; Q1 2026 reported $818M revenue / $2.47B operating loss; mid-2026 ARR estimates $0.5B (Bloomberg/Sacra) to $1.5B (Presenc); 2026 target $2B", "2024-2026", "ValueAdd VC / Sacra / Presenc AI", "https://valueaddvc.com/blog/xai-revenue-how-grok-and-the-api-actually-make-money"),
    ("Revenue", "DeepSeek", "Annualized revenue approaching $500M (The Information, mid-2026); ~$700M annualized May 2026 (Presenc est.); ~$200M 2025; 70-80% gross margins at $2.19/M tokens; $52B valuation Jul 2026", "mid-2026", "AI Weekly / Caixin / Presenc AI", "https://aiweekly.co/alerts/deepseek-nears-500m-annualized-revenue-plans-2027-ipo"),
    ("Revenue", "Alibaba (Qwen)", "AI-related product revenue RMB 12,376M (US$1,824M) quarter ended Jun 30 2026, 12th consecutive quarter of triple-digit YoY growth; AI Cloud & Compute Services $7.1B (+45% YoY); MaaS ARR >RMB 16B", "Jun 2026", "Alibaba Group earnings / SEC 6-K exhibit", "https://www.sec.gov/Archives/edgar/data/1577552/000110465926099220/tm2623667d1_ex99-1.htm"),
    ("Revenue", "Mistral AI", "ARR ~$16-20M end-2024, $312M Dec 2025, $400M Jan 2026, ~$1.0B May 2026; 2026 target $1.1-1.2B", "2024-2026", "Sacra / Klover / Presenc AI", "https://sacra.com/c/mistral/"),
    ("Revenue", "Other open-weight labs (context)", "May 2026 ARR: Moonshot AI (Kimi) ~$500M, Zhipu/Z.AI (GLM) ~$480M, Cohere ~$500M (not modeled as separate rows)", "May 2026", "Presenc AI", "https://presenc.ai/research/ai-lab-revenue-and-valuation-2026"),
    ("GW capacity", "OpenAI", "Stargate Abilene ~0.3 GW operational mid-2026 (capped at 1.2 GW); 7 sites >9 GW planned ($400B+); 4.5 GW Oracle lease; balance of serving on Azure", "mid-2026", "Founder OS / Momoview / TCDEV", "https://fos.kentlangley.com/b/ai-lab-compute/"),
    ("GW capacity", "Anthropic", "AWS New Carlisle ~1 GW live by Mar 2026 (2.2 GW campus); Project Rainier 500k -> 1M+ Trainium2; Google TPU >1 GW landing 2026 (up to 1M TPUs); SpaceX 300 MW / 220k GPUs (May 2026); 5 GW AWS + 3.5 GW Broadcom contracted", "2025-2026", "Momoview / TCDEV / Presenc AI / Adam N.", "https://momoview.com/blog/en/posts/ai-compute-landscape-2026-operational-vs-under-construction-hyperscalers-ai-labs-neoclouds-gigawatt-map/"),
    ("GW capacity", "xAI", "Colossus 1 ~300 MW / ~230k GPUs (mid-2025, verified); ~555k GPUs total Jan 2026 (~$18B hardware); Colossus 2 ramping through 1 GW toward 2 GW in 2026; 1M GPU roadmap end-2026", "2024-2026", "Founder OS / TCDEV / Momoview", "https://www.tcdev.de/blog/where-are-claudes-data-centers/"),
    ("GW capacity", "Meta (Llama)", "Plans to deploy 7 GW of compute during 2026 and reach 14 GW total by 2027 (internal memo via Reuters); Prometheus ~1 GW (Ohio) online 2026; Hyperion (LA) scaled to 5 GW; ~600k+ GPUs; MTIA 'Iris' chip Sept 2026; AMD deal up to 6 GW", "Jul 2026", "Reuters / StockScreener / Presenc AI", "https://www.reuters.com/world/asia-pacific/meta-put-ai-chip-into-production-september-it-looks-double-computing-capacity-2026-07-09/"),
    ("GW capacity", "Google DeepMind", "TPU fleet (v6/v7 Ironwood) 1M+ accelerators; '>1 GW of new compute online through 2026'; Alphabet capex ~$175-205B 2026 - Gemini-allocated share estimated here", "2026", "Momoview / Presenc AI / CNBC", "https://www.cnbc.com/2026/07/29/zuckerberg-metas-ai-capacity-dilemma-what-to-sell-vs-what-to-keep.html"),
    ("GW capacity", "DeepSeek", "~20,000 H100-equivalent GPU fleet (founder briefing, mid-2026, 'most arrived in the past month or two'); 1 GW Ulanqab data center planned online late 2027-early 2028", "mid-2026", "XenoSpectrum / StartupHub", "https://xenospectrum.com/en/deepseek-forced-price-hike-demand-overwhelms-20000-gpu-fleet/"),
    ("GW capacity", "Alibaba (Qwen)", "Quarterly capex US$9.975B (quarter ended Jun 2026, +75% YoY) largely AI infrastructure; 100-day hyperscale AIDC delivery; Zhenwu M890 supernode at commercial scale - AI-dedicated GW estimated here", "Jun 2026", "Alibaba earnings / SEC 6-K", "https://www.alibabagroup.com/en-US/document-2027233133950140416"),
    ("GW capacity", "Industry total", "Top-10 AI models consume ~7.8 GW operational (2026), projected ~24 GW by end-2027; global operational AI compute >15M H100e / >10 GW (Epoch AI)", "2026", "jasperbernaers.com / Epoch AI via tokensperday", "https://jasperbernaers.com/ai-gigawatts/"),
    ("Cost per GW", "Industry benchmark", "1 GW AI datacenter: $38B upfront capex (GB200-class, ~60% servers), $0.9B/yr opex, $8.5B/yr annualized TCO (5-yr IT / 14-yr facility lives)", "May 2026", "Epoch AI", "https://epoch.ai/data-insights/ai-datacenter-cost-breakdown"),
    ("Cost per GW", "Industry benchmark", "Vera Rubin era ~$47B/GW all-in (Foxconn chairman; Bernstein $47.3B); Blackwell ~$40.5B; earlier Hopper ~$35B; NVIDIA guidance $50-60B; Turner & Townsend $45-55B/GW campus", "2026", "GeoTechNexus / IT Voice", "https://geotechnexus.com/the-47-billion-gigawatt-what-vera-rubins-price-tag-tells-us-about-the-ai-build-out/"),
    ("Cost per GW", "Industry benchmark", "A lab needs ~$31B/yr of end-user inference revenue per GW to clear base-case economics ($3.30/GPU-hr, 90% util.); $8.5B/yr annualized cost per GW", "Jun 2026", "YK Research", "https://investments.chiayong.com/ai-buildout/economics-2026-06-18/"),
    ("Cost (lab-reported)", "OpenAI", "2025 operating loss $20.9B on $13.07B revenue; ~$17B/yr cash burn (2026); $1.4T+ datacenter commitments", "2025-2026", "ValueAdd VC / Happycapy", "https://happycapyguide.com/blog/openai-15-billion-tokens-per-minute-25-billion-revenue-ai-scale-2026"),
    ("Cost (lab-reported)", "xAI", "~$1B/month infrastructure + training burn; Q1 2026 operating loss $2.47B", "2026", "ValueAdd VC / Bloomberg", "https://valueaddvc.com/blog/xai-revenue-how-grok-and-the-api-actually-make-money"),
    ("Cost (lab-reported)", "Meta (Llama)", "2026 capex guidance $130-145B (company-wide, AI-driven); JV structures with BlackRock / Blue Owl for El Paso and Louisiana sites", "Jul 2026", "CNBC / InsideAI", "https://insideai.news/news/ai-hardware-infrastructure/meta-raises-2026-capex-floor-to-130-billion-for-ai-data-centers/6493/"),
    ("Cost (lab-reported)", "Anthropic", "Positive adjusted operating income in Q2 2026 (implies compute cost below $11.5B quarterly revenue); $50B own-datacenter program with Fluidstack (TX, NY)", "Q2 2026", "Bloomberg / TCDEV", "https://www.tcdev.de/blog/where-are-claudes-data-centers/"),
    ("Token volume", "Google DeepMind", "9.7T tokens/month May 2024 -> ~480T May 2025 -> 3.2 quadrillion May 2026 (all surfaces, disclosed at I/O); model APIs 19B tokens/min May 2026, 22B/min Q2 2026 earnings (~1Q/month API alone); 10B/min Q4 2025", "May-Jul 2026", "Google I/O 2026 keynote / io-fund", "https://blog.google/innovation-and-ai/sundar-pichai-io-2026/"),
    ("Token volume", "OpenAI", "API 6B tokens/min at DevDay 2025 (Oct) -> 15B+ tokens/min Q1-2026 (CRO memo Apr 8 2026); 900M+ weekly ChatGPT users (first-party traffic est. on top of API)", "2025-2026", "Shuai Guan compilation / Casky", "https://shuaiguan.io/blog/openai-statistics"),
    ("Token volume", "Anthropic", "No disclosed aggregate. OpenRouter slice: ~24.7T/month routed, 11.8% of routed volume but 58.7% of routed spend (Jul 2026). Series here estimated from revenue / blended ~$10 per M tokens", "Jul 2026", "Fabryka OpenRouter memo / Dirac", "https://ort.fabryka.ai/research.html"),
    ("Token volume", "DeepSeek", "22.6% of OpenRouter routed tokens (~47.3T/month, Jul 2026), largest single lab; ~23% of Vercel tokens Jun 2026; total incl. Chinese clouds + self-hosted estimated here", "Jun-Jul 2026", "Capital & Compute / Fabryka / AI Weekly", "https://capitalandcompute.net/blog/open-source-llms-overtake-2026/"),
    ("Token volume", "Market structure", "Open-weight models overtook proprietary on OpenRouter: 40/60 Mar 2026 -> 60/40 Jun 2026 (~6T tokens/day routed); Chinese-origin models ~46-61% of routed tokens; Meta Llama <1% of routed volume", "Jun 2026", "Dirac / Capital & Compute / DataGravity", "https://dirac.run/labs-market-share"),
    ("Token volume", "Global anchors", "Global inference ~360-430T tokens/day mid-2026 (~11-13 quadrillion/month); Goldman: 1.7Q/month mid-2025 -> 47Q/month 2028F; OpenRouter on pace for >1Q routed tokens in 2026", "mid-2026", "tokensperday.com / Epoch AI / io-fund", "https://tokensperday.com/"),
    ("Context", "Valuations", "OpenAI $852B (Mar 2026, $122B round); Anthropic $380B Series G (secondaries ~$1T); xAI $230B Series E (SpaceX merger ~$1.25T combined); DeepSeek $52B; Mistral ~$15B", "2026", "Bloomberg / Presenc AI / Caixin", "https://presenc.ai/research/ai-lab-revenue-and-valuation-2026"),
]

METHODOLOGY = [
    ("GW capacity", "Operational IT power (owned + leased) serving each lab at quarter end, in gigawatts. Interpolated between disclosed anchors (cluster energizations, GPU counts converted at ~1.4 kW/H100e all-in). NOT the same as contracted/announced capacity, which is far larger (e.g. OpenAI 9+ GW, Anthropic 8.5+ GW, Meta 14 GW by 2027). Meta's figure covers its whole AI accelerator fleet incl. ranking/recsys, so it is not directly comparable to model-only estimates. Error bars are wide: +/-30% or more."),
    ("Quarterly revenue", "Booked revenue per calendar quarter, $B. Anthropic Q2 2025/Q2 2026 and Alibaba Q2 2026 are disclosed; OpenAI annual totals (2024: $3.7B, 2025: $13.07B) are disclosed and split across quarters along the run-rate curve; everything else interpolated from disclosed run rates. Alibaba = 'AI-related product revenue' only (excludes broader AI Cloud & Compute). Meta has no separable model revenue (monetizes via ads). Note OpenAI reports net of cloud-partner costs while Anthropic includes gross reseller revenue - not perfectly comparable."),
    ("Revenue run rate (ARR)", "Annualized revenue run rate at quarter end, $B, as disclosed by companies or estimated. This is the metric labs quote publicly (e.g. Anthropic $65B Jul 2026, OpenAI $40B Aug 2026). Alibaba ARR = quarterly AI-related product revenue x4."),
    ("Compute cost", "Estimated quarterly compute cost = operational GW x $8.5B/GW/yr annualized TCO (Epoch AI benchmark: 5-yr server / 14-yr facility depreciation + opex) / 4. This proxies economic cost, not cash capex. Reported anchors for calibration: OpenAI ~$17B/yr burn, xAI ~$1B/month, Anthropic adjusted-operating-income positive in Q2 2026. Leased capacity (OpenAI, Anthropic) likely carries a rental premium above this benchmark; efficient self-builds (xAI, DeepSeek) may run below."),
    ("Build capex per GW", "Industry all-in cost to BUILD 1 GW (facility + power + IT), $B/GW, by dominant GPU generation: Hopper ~$32-35B (2024), Blackwell ~$38-40.5B (2025), Vera Rubin ~$47B (2026). Sources: Epoch AI ($38B GB200-class), Bernstein/Foxconn ($47B Rubin), Turner & Townsend ($45-55B)."),
    ("Revenue per GW", "Annualized run-rate revenue / operational GW, $B per GW per year. Benchmark: a GW must generate ~$31B/yr of inference revenue to clear base-case economics (YK Research). Labs below that line are subsidizing growth with capital; labs near/above it have workable unit economics. Google's low figure partly reflects un-monetized internal usage (Search AI, Workspace)."),
    ("Cost per GW", "Two views: (1) annualized TCO ~$8.5B/GW/yr (same for all labs, Epoch benchmark); (2) build capex per GW time series above. Per-lab true costs vary with chip generation, lease vs own, power strategy, and utilization but are not separately disclosed."),
    ("Token volume", "Trillions of tokens processed per quarter, all surfaces (first-party apps + API + self-hosted where estimable). Google is derived from three disclosed monthly anchors (9.7T -> 480T -> 3,200T/month) via exponential interpolation. OpenAI from disclosed API tokens/min (6B Oct 2025, 15B Q1 2026) x ~1.8 for first-party ChatGPT traffic. Anthropic/xAI/Mistral estimated from revenue and blended price per token. DeepSeek/Qwen/Meta include self-hosted open-weight usage, which no one measures directly - treat as order-of-magnitude."),
    ("Lab selection", "Frontier (closed): OpenAI, Anthropic, Google DeepMind, xAI. Open-weight: Meta (Llama), DeepSeek, Alibaba (Qwen), Mistral AI. Other open-weight labs (Moonshot/Kimi ~$500M ARR, Zhipu/Z.AI ~$480M, MiniMax, Xiaomi MiMo ~17.5% of OpenRouter tokens) are material to token share but excluded as rows for lack of capacity/revenue disclosure; see Sources."),
    ("General caveats", "Compiled Aug 23, 2026 from public reporting, company disclosures, and analyst estimates. Q3 2026 is partial (quarter-to-date run rates annualized/extrapolated). Private-company figures are unaudited and sometimes conflicting across sources (conflicts noted in Sources sheet, e.g. xAI ARR $0.5B vs $1.5B; Anthropic FY2025 ~$4.5B booked vs ~$9-10B gross). All dollar figures nominal USD billions."),
]

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

DARK = "1F3864"
FRONTIER_FILL = PatternFill("solid", fgColor="DDEBF7")
OPEN_FILL = PatternFill("solid", fgColor="E2EFDA")
HEADER_FILL = PatternFill("solid", fgColor=DARK)
SUBTOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
BENCH_FILL = PatternFill("solid", fgColor="FCE4D6")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color=DARK)
SUB_FONT = Font(italic=True, size=10, color="595959")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FRONTIER_LABS = [l for l, c in LABS if c.startswith("Frontier")]
OPEN_LABS = [l for l, c in LABS if c == "Open-weight"]


def sheet_title(ws, title, subtitle, width=13):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 17
    for i in range(3, 3 + N_Q + 2):
        ws.column_dimensions[get_column_letter(i)].width = width


def header_row(ws, row, first_cols):
    for j, h in enumerate(first_cols, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for j, q in enumerate(QUARTERS, start=len(first_cols) + 1):
        c = ws.cell(row=row, column=j, value=q)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER


def metric_sheet(wb, name, subtitle, data, numfmt, disclosed_key=None,
                 totals=True, extra_rows=None):
    """Standard labs x quarters sheet. Returns (ws, first_data_row)."""
    ws = wb.create_sheet(name)
    sheet_title(ws, name, subtitle)
    hr = 4
    header_row(ws, hr, ["Lab", "Category"])
    r = hr + 1
    lab_rows = {}
    for lab, cat in LABS:
        fill = FRONTIER_FILL if cat.startswith("Frontier") else OPEN_FILL
        ws.cell(row=r, column=1, value=lab).font = BOLD
        ws.cell(row=r, column=2, value=cat)
        for j in range(1, 3):
            ws.cell(row=r, column=j).fill = fill
            ws.cell(row=r, column=j).border = BORDER
        series = data[lab]
        for qi, val in enumerate(series):
            c = ws.cell(row=r, column=3 + qi)
            c.border = BORDER
            c.fill = fill
            if val is None:
                c.value = "n/a"
                c.font = Font(italic=True, color="808080")
                c.alignment = Alignment(horizontal="center")
            else:
                c.value = val
                c.number_format = numfmt
                if disclosed_key and qi in DISCLOSED.get(disclosed_key, {}).get(lab, []):
                    c.font = Font(bold=True, color="C00000")
        lab_rows[lab] = r
        r += 1

    if totals:
        for label, group in (("Frontier subtotal", FRONTIER_LABS),
                             ("Open-weight subtotal", OPEN_LABS),
                             ("TOTAL (all 8 labs)", [l for l, _ in LABS])):
            ws.cell(row=r, column=1, value=label).font = BOLD
            for j in range(1, 3 + N_Q):
                ws.cell(row=r, column=j).fill = SUBTOTAL_FILL
                ws.cell(row=r, column=j).border = BORDER
            for qi in range(N_Q):
                col = get_column_letter(3 + qi)
                cells = [f"{col}{lab_rows[l]}" for l in group if data[l][qi] is not None]
                c = ws.cell(row=r, column=3 + qi)
                c.value = f"={'+'.join(cells)}" if cells else None
                c.number_format = numfmt
                c.font = BOLD
            r += 1

    if extra_rows:
        r += 1
        for label, values, fmt in extra_rows:
            ws.cell(row=r, column=1, value=label).font = BOLD
            for j in range(1, 3 + N_Q):
                ws.cell(row=r, column=j).fill = BENCH_FILL
                ws.cell(row=r, column=j).border = BORDER
            for qi, v in enumerate(values):
                c = ws.cell(row=r, column=3 + qi, value=v)
                c.number_format = fmt
            r += 1

    ws.freeze_panes = "C5"
    return ws, hr + 1, lab_rows


def add_line_chart(ws, title, y_title, hr, n_rows, anchor, log=False):
    """Line chart of lab rows over quarters. hr = header row of the table."""
    ch = LineChart()
    ch.title = title
    ch.style = 12
    ch.height = 10
    ch.width = 26
    ch.y_axis.title = y_title
    ch.x_axis.title = "Quarter"
    if log:
        ch.y_axis.scaling.logBase = 10
    for i in range(n_rows):
        row = hr + 1 + i
        vals = Reference(ws, min_col=3, min_row=row, max_col=2 + N_Q, max_row=row)
        s = ChartSeries(vals, title=ws.cell(row=row, column=1).value)
        ch.series.append(s)
    cats = Reference(ws, min_col=3, min_row=hr, max_col=2 + N_Q, max_row=hr)
    ch.set_categories(cats)
    ws.add_chart(ch, anchor)


# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------

wb = Workbook()

# ---- Read Me ----
ws = wb.active
ws.title = "Read Me"
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 130
ws["B2"] = "AI Lab Compute Economics - Quarterly Time Series"
ws["B2"].font = Font(bold=True, size=18, color=DARK)
ws["B3"] = "GW capacity, revenue, cost, and token volume for frontier and open-weight AI labs | Q1 2024 - Q3 2026 | Compiled Aug 23, 2026"
ws["B3"].font = SUB_FONT
rows = [
    ("", ""),
    ("What's inside", ""),
    ("1. Summary (Q2 2026)", "Latest-full-quarter snapshot: GW, ARR, revenue/GW, est. compute cost, cost/GW, tokens - with frontier vs open-weight subtotals."),
    ("2. GW Capacity", "Estimated operational AI compute capacity (owned + leased IT power, GW, end of quarter) per lab, with subtotals and chart."),
    ("3. Revenue (Quarterly)", "Booked revenue per calendar quarter, $B."),
    ("4. Revenue Run Rate", "Annualized run-rate revenue at quarter end, $B (the metric labs quote publicly)."),
    ("5. Compute Cost", "Estimated quarterly compute cost per lab (GW x $8.5B/GW/yr TCO), industry build-capex-per-GW benchmark, and reported cost datapoints."),
    ("6. Rev & Cost per GW", "Derived: annualized revenue per operational GW vs annualized cost per GW, by lab and quarter."),
    ("7. Token Volume", "Estimated tokens processed per quarter (trillions), all surfaces, with global market anchors and chart."),
    ("8. Sources & Methodology", "Every disclosed anchor with source + URL, and metric-by-metric methodology and caveats."),
    ("", ""),
    ("Lab coverage", ""),
    ("Frontier (closed)", "OpenAI, Anthropic, Google DeepMind, xAI"),
    ("Open-weight", "Meta (Llama), DeepSeek, Alibaba (Qwen), Mistral AI"),
    ("", ""),
    ("How to read it", ""),
    ("Red bold cells", "Directly disclosed figures (company statements, filings, earnings). Everything else is estimated/interpolated - see Sources & Methodology."),
    ("'Q3 2026 (part.)'", "Partial quarter: quarter-to-date run rates as of Aug 23, 2026, extrapolated."),
    ("'n/a' (Meta revenue)", "Meta gives Llama away and monetizes via ads - it has no separable model revenue, so revenue-per-GW is not meaningful for it."),
    ("Error bars", "Capacity and token estimates for private companies carry +/-30% or wider uncertainty. Treat as order-of-magnitude, directionally reliable."),
    ("", ""),
    ("Headline takeaways (Q2-Q3 2026)", ""),
    ("Revenue", "Anthropic leads on run rate ($65B Jul 2026) vs OpenAI ($40B Aug 2026); Google AI ~$30B+; everyone else is <$10B. Anthropic turned adjusted-operating-income positive in Q2 2026."),
    ("Capacity", "The 8 labs run ~12 GW operational vs 40+ GW contracted. Meta has the largest fleet (~3.6 GW incl. recsys); Google ~2.8 GW; Anthropic ~1.8 GW; announced pipelines (Stargate 9+ GW, Meta 14 GW) dwarf what is energized."),
    ("Tokens", "Google processes the most tokens (3.2 quadrillion/month May 2026). On neutral routing (OpenRouter), open-weight models flipped to a 60/40 majority by Jun 2026, led by DeepSeek (~23%)."),
    ("Economics", "~$31B/yr revenue per GW is the break-even benchmark; Anthropic (~$26B/GW) and OpenAI (~$22B/GW) are approaching it, Google sits lower on paid revenue per GW (heavy un-monetized internal usage), and open-weight labs monetize a small fraction of their token footprint."),
]
r = 4
for k, v in rows:
    ws.cell(row=r, column=2, value=k).font = Font(bold=True, color=DARK) if v == "" or k else BOLD
    if v:
        ws.cell(row=r, column=3, value=v).alignment = Alignment(wrap_text=True, vertical="top")
    r += 1

# ---- GW Capacity ----
ws_gw, gw_hr, gw_rows = metric_sheet(
    wb, "GW Capacity",
    "Estimated operational AI compute capacity serving each lab (owned + leased IT power, gigawatts, end of quarter). "
    "Red bold = disclosed. Contracted/announced capacity is far larger - see Sources.",
    GW, "0.000", totals=True,
    extra_rows=[("Memo: top-10 models operational GW (jasperbernaers, 2026)",
                 [None]*9 + [7.8, None], "0.0")])
add_line_chart(ws_gw, "Operational GW capacity by lab", "GW", gw_hr - 1, len(LABS), "A20")

# ---- Revenue (Quarterly) ----
ws_rev, rev_hr, rev_rows = metric_sheet(
    wb, "Revenue (Quarterly)",
    "Booked revenue per calendar quarter, US$B. Red bold = disclosed (Anthropic Q2'25/Q2'26; Alibaba Q2'26 AI-related product revenue; xAI Q1'26 reported). "
    "Meta: no separable model revenue. OpenAI reports net of cloud-partner costs; Anthropic gross - not perfectly comparable.",
    REV_Q, "$#,##0.00", disclosed_key="REV_Q", totals=True)

# ---- Revenue Run Rate ----
ws_arr, arr_hr, arr_rows = metric_sheet(
    wb, "Revenue Run Rate",
    "Annualized revenue run rate at quarter end, US$B - the metric labs quote publicly. "
    "Red bold = disclosed (OpenAI $21.5B end-25 / $25B Q1'26 / $40B Aug'26; Anthropic $9B late-25 / $47B May'26 / $65B Jul'26; Mistral $312M Dec'25 / ~$1B May'26).",
    ARR, "$#,##0.00", disclosed_key="ARR", totals=True)
add_line_chart(ws_arr, "Annualized revenue run rate by lab", "US$B / yr", arr_hr - 1, len(LABS), "A20")

# ---- Compute Cost ----
ws_cost = wb.create_sheet("Compute Cost")
sheet_title(
    ws_cost, "Compute Cost",
    "Estimated quarterly compute cost = operational GW x annualized TCO ($8.5B/GW/yr, Epoch AI) / 4. Formula-linked to the GW Capacity sheet. "
    "Bottom: industry build-capex-per-GW benchmark and reported lab cost datapoints.")
ws_cost["A3"] = "TCO assumption ($B per GW per year):"
ws_cost["A3"].font = BOLD
ws_cost["C3"] = TCO_PER_GW_YR
ws_cost["C3"].font = Font(bold=True, color="C00000")
ws_cost["C3"].number_format = "$0.0"
hr = 5
header_row(ws_cost, hr, ["Lab", "Category"])
r = hr + 1
cost_rows = {}
for lab, cat in LABS:
    fill = FRONTIER_FILL if cat.startswith("Frontier") else OPEN_FILL
    ws_cost.cell(row=r, column=1, value=lab).font = BOLD
    ws_cost.cell(row=r, column=2, value=cat)
    for qi in range(N_Q):
        col = get_column_letter(3 + qi)
        c = ws_cost.cell(row=r, column=3 + qi)
        c.value = f"='GW Capacity'!{col}{gw_rows[lab]}*$C$3/4"
        c.number_format = "$#,##0.00"
    for j in range(1, 3 + N_Q):
        ws_cost.cell(row=r, column=j).fill = fill
        ws_cost.cell(row=r, column=j).border = BORDER
    cost_rows[lab] = r
    r += 1
for label, group in (("Frontier subtotal", FRONTIER_LABS),
                     ("Open-weight subtotal", OPEN_LABS),
                     ("TOTAL (all 8 labs)", [l for l, _ in LABS])):
    ws_cost.cell(row=r, column=1, value=label).font = BOLD
    for qi in range(N_Q):
        col = get_column_letter(3 + qi)
        c = ws_cost.cell(row=r, column=3 + qi)
        c.value = "=" + "+".join(f"{col}{cost_rows[l]}" for l in group)
        c.number_format = "$#,##0.00"
        c.font = BOLD
    for j in range(1, 3 + N_Q):
        ws_cost.cell(row=r, column=j).fill = SUBTOTAL_FILL
        ws_cost.cell(row=r, column=j).border = BORDER
    r += 1
r += 1
ws_cost.cell(row=r, column=1, value="Industry build capex per GW ($B/GW, all-in, by GPU generation)").font = BOLD
for qi, v in enumerate(BUILD_COST_PER_GW):
    c = ws_cost.cell(row=r, column=3 + qi, value=v)
    c.number_format = "$#,##0.0"
for j in range(1, 3 + N_Q):
    ws_cost.cell(row=r, column=j).fill = BENCH_FILL
    ws_cost.cell(row=r, column=j).border = BORDER
r += 2
ws_cost.cell(row=r, column=1, value="Reported cost datapoints (calibration)").font = Font(bold=True, size=12, color=DARK)
r += 1
for txt in [
    "OpenAI: 2025 operating loss $20.9B on $13.07B revenue; ~$17B/yr cash burn in 2026; $1.4T+ long-term datacenter commitments.",
    "xAI: ~$1B/month infrastructure + training burn; Q1 2026 operating loss $2.47B on $818M revenue.",
    "Anthropic: positive ADJUSTED operating income in Q2 2026 - compute cost ran below $11.5B quarterly revenue.",
    "Meta: 2026 company-wide capex guidance $130-145B (AI-driven); JVs (BlackRock 80/20 El Paso, Blue Owl Louisiana) keep part off balance sheet.",
    "Alibaba: quarterly capex $9.975B (Jun 2026 quarter, +75% YoY), mostly AI infrastructure.",
    "Benchmarks: $38B upfront capex/GW (Epoch, GB200-class) rising to ~$47B/GW (Vera Rubin, Bernstein/Foxconn); $0.9B/GW/yr opex; ~$31B/GW/yr inference revenue needed to break even (YK Research).",
]:
    ws_cost.cell(row=r, column=1, value="- " + txt).alignment = Alignment(wrap_text=False)
    r += 1
ws_cost.freeze_panes = "C6"

# ---- Rev & Cost per GW ----
ws_pg = wb.create_sheet("Rev & Cost per GW")
sheet_title(
    ws_pg, "Revenue & Cost per GW",
    "Annualized run-rate revenue per operational GW vs annualized cost per GW (US$B per GW per year). Formula-linked. "
    "Break-even benchmark: ~$31B/yr revenue per GW (YK Research, base case). Cost per GW uses the flat $8.5B/GW/yr TCO benchmark.")
hr = 4
ws_pg.cell(row=hr - 1, column=1, value="REVENUE PER GW ($B ARR / operational GW)").font = Font(bold=True, color=DARK)
header_row(ws_pg, hr, ["Lab", "Category"])
r = hr + 1
for lab, cat in LABS:
    fill = FRONTIER_FILL if cat.startswith("Frontier") else OPEN_FILL
    ws_pg.cell(row=r, column=1, value=lab).font = BOLD
    ws_pg.cell(row=r, column=2, value=cat)
    for qi in range(N_Q):
        col = get_column_letter(3 + qi)
        c = ws_pg.cell(row=r, column=3 + qi)
        if ARR[lab][qi] is None:
            c.value = "n/a"
            c.font = Font(italic=True, color="808080")
            c.alignment = Alignment(horizontal="center")
        else:
            c.value = f"='Revenue Run Rate'!{col}{arr_hr + [l for l, _ in LABS].index(lab)}/'GW Capacity'!{col}{gw_rows[lab]}"
            c.number_format = "$#,##0.0"
    for j in range(1, 3 + N_Q):
        ws_pg.cell(row=r, column=j).fill = fill
        ws_pg.cell(row=r, column=j).border = BORDER
    r += 1
ws_pg.cell(row=r, column=1, value="Break-even benchmark (YK Research)").font = BOLD
for qi in range(N_Q):
    c = ws_pg.cell(row=r, column=3 + qi, value=31.0)
    c.number_format = "$#,##0.0"
for j in range(1, 3 + N_Q):
    ws_pg.cell(row=r, column=j).fill = BENCH_FILL
    ws_pg.cell(row=r, column=j).border = BORDER
r += 2
ws_pg.cell(row=r, column=1, value="COST PER GW ($B per GW per year)").font = Font(bold=True, color=DARK)
r += 1
hr2 = r
header_row(ws_pg, hr2, ["Measure", "Scope"])
r += 1
ws_pg.cell(row=r, column=1, value="Annualized TCO per GW (Epoch AI)").font = BOLD
ws_pg.cell(row=r, column=2, value="All labs")
for qi in range(N_Q):
    c = ws_pg.cell(row=r, column=3 + qi, value="='Compute Cost'!$C$3")
    c.number_format = "$#,##0.0"
for j in range(1, 3 + N_Q):
    ws_pg.cell(row=r, column=j).fill = BENCH_FILL
    ws_pg.cell(row=r, column=j).border = BORDER
r += 1
ws_pg.cell(row=r, column=1, value="Build capex per GW (industry)").font = BOLD
ws_pg.cell(row=r, column=2, value="One-time, all-in")
for qi, v in enumerate(BUILD_COST_PER_GW):
    c = ws_pg.cell(row=r, column=3 + qi, value=v)
    c.number_format = "$#,##0.0"
for j in range(1, 3 + N_Q):
    ws_pg.cell(row=r, column=j).fill = BENCH_FILL
    ws_pg.cell(row=r, column=j).border = BORDER
r += 2
ws_pg.cell(row=r, column=1, value="Reading: labs above ~$31B/GW of run-rate revenue clear base-case data-center economics. As of Q2 2026, Anthropic (~$26B/GW) and OpenAI (~$22B/GW) approach it; Google's paid AI revenue per GW is lower because much of its fleet serves un-monetized internal surfaces; open-weight labs (Meta, DeepSeek, Qwen, Mistral) monetize only a small share of the tokens their models generate.").alignment = Alignment(wrap_text=True)
ws_pg.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=13)
ws_pg.freeze_panes = "C5"

# ---- Token Volume ----
ws_tok, tok_hr, tok_rows = metric_sheet(
    wb, "Token Volume",
    "Estimated tokens processed per quarter, TRILLIONS, all surfaces (first-party apps + API + estimable self-hosted). "
    "Google derived from disclosed monthly anchors (9.7T May'24 / 480T May'25 / 3,200T May'26). Others estimated - see Methodology.",
    TOKENS, "#,##0", totals=True,
    extra_rows=[
        ("Memo: global inference, all providers (est.)",
         [30, 55, 140, 380, 900, 2000, 3600, 5800, 9200, 14500, 21000], "#,##0"),
        ("Memo: 8-lab share of global (%)", [None]*11, "0%"),
    ])
# fill share row formulas (8-lab total is at row tok_hr+len(LABS)+2 -> "TOTAL" row)
total_row = tok_hr + len(LABS) + 2
memo_row = total_row + 2
share_row = memo_row + 1
for qi in range(N_Q):
    col = get_column_letter(3 + qi)
    c = ws_tok.cell(row=share_row, column=3 + qi)
    c.value = f"={col}{total_row}/{col}{memo_row}"
    c.number_format = "0%"
add_line_chart(ws_tok, "Tokens processed per quarter by lab (trillions)", "T tokens / quarter", tok_hr - 1, len(LABS), "A22")

# ---- Summary (Q2 2026) ----
ws_sum = wb.create_sheet("Summary (Q2 2026)", 1)
sheet_title(ws_sum, "Summary - Q2 2026 (last full quarter)",
            "Formula-linked snapshot. ARR = end-of-quarter annualized run rate. Break-even revenue/GW benchmark ~$31B/yr (YK Research).", width=16)
q2_col = get_column_letter(3 + LAST_FULL_Q_IDX)
headers = ["Lab", "Category", "Operational GW", "ARR ($B/yr)", "Revenue / GW ($B/yr per GW)",
           "Est. compute cost ($B/yr)", "Cost / GW ($B/yr per GW)", "Tokens Q2'26 (T)", "Notes"]
hr = 4
for j, h in enumerate(headers, start=1):
    c = ws_sum.cell(row=hr, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
ws_sum.column_dimensions["I"].width = 90
r = hr + 1
sum_rows = {}
for lab, cat in LABS:
    fill = FRONTIER_FILL if cat.startswith("Frontier") else OPEN_FILL
    idx = [l for l, _ in LABS].index(lab)
    gw_r = gw_rows[lab]
    arr_r = arr_hr + idx
    tok_r = tok_rows[lab]
    ws_sum.cell(row=r, column=1, value=lab).font = BOLD
    ws_sum.cell(row=r, column=2, value=cat)
    ws_sum.cell(row=r, column=3, value=f"='GW Capacity'!{q2_col}{gw_r}").number_format = "0.000"
    if ARR[lab][LAST_FULL_Q_IDX] is None:
        ws_sum.cell(row=r, column=4, value="n/a").font = Font(italic=True, color="808080")
        ws_sum.cell(row=r, column=5, value="n/a").font = Font(italic=True, color="808080")
    else:
        ws_sum.cell(row=r, column=4, value=f"='Revenue Run Rate'!{q2_col}{arr_r}").number_format = "$#,##0.00"
        ws_sum.cell(row=r, column=5, value=f"=D{r}/C{r}").number_format = "$#,##0.0"
    ws_sum.cell(row=r, column=6, value=f"=C{r}*'Compute Cost'!$C$3").number_format = "$#,##0.00"
    ws_sum.cell(row=r, column=7, value="='Compute Cost'!$C$3").number_format = "$#,##0.0"
    ws_sum.cell(row=r, column=8, value=f"='Token Volume'!{q2_col}{tok_r}").number_format = "#,##0"
    ws_sum.cell(row=r, column=9, value=LAB_NOTES[lab]).alignment = Alignment(wrap_text=True, vertical="top")
    for j in range(1, 10):
        ws_sum.cell(row=r, column=j).fill = fill
        ws_sum.cell(row=r, column=j).border = BORDER
    sum_rows[lab] = r
    r += 1
for label, group in (("Frontier subtotal", FRONTIER_LABS),
                     ("Open-weight subtotal", OPEN_LABS),
                     ("TOTAL (all 8 labs)", [l for l, _ in LABS])):
    ws_sum.cell(row=r, column=1, value=label).font = BOLD
    ws_sum.cell(row=r, column=3, value="=" + "+".join(f"C{sum_rows[l]}" for l in group)).number_format = "0.000"
    arr_cells = [f"D{sum_rows[l]}" for l in group if ARR[l][LAST_FULL_Q_IDX] is not None]
    ws_sum.cell(row=r, column=4, value="=" + "+".join(arr_cells)).number_format = "$#,##0.00"
    ws_sum.cell(row=r, column=5, value=f"=D{r}/C{r}").number_format = "$#,##0.0"
    ws_sum.cell(row=r, column=6, value="=" + "+".join(f"F{sum_rows[l]}" for l in group)).number_format = "$#,##0.00"
    ws_sum.cell(row=r, column=7, value="='Compute Cost'!$C$3").number_format = "$#,##0.0"
    ws_sum.cell(row=r, column=8, value="=" + "+".join(f"H{sum_rows[l]}" for l in group)).number_format = "#,##0"
    for j in range(1, 10):
        ws_sum.cell(row=r, column=j).fill = SUBTOTAL_FILL
        ws_sum.cell(row=r, column=j).border = BORDER
        ws_sum.cell(row=r, column=j).font = BOLD
    r += 1
r += 1
ws_sum.cell(row=r, column=1, value="Current as of Aug 23, 2026: Anthropic ARR $65B (Jul), OpenAI $40B (Aug) - see Revenue Run Rate sheet, Q3 2026 column.").font = SUB_FONT
# bar chart: revenue per GW
ch = BarChart()
ch.type = "col"
ch.title = "Revenue per GW vs break-even (~$31B/GW/yr), Q2 2026"
ch.y_axis.title = "US$B ARR per operational GW"
ch.height = 9
ch.width = 22
labs_with_rev = [l for l, _ in LABS if ARR[l][LAST_FULL_Q_IDX] is not None]
# contiguous rows except Meta; use full range - Meta cell is text 'n/a', excel skips
vals = Reference(ws_sum, min_col=5, min_row=hr + 1, max_col=5, max_row=hr + len(LABS))
cats = Reference(ws_sum, min_col=1, min_row=hr + 1, max_col=1, max_row=hr + len(LABS))
ch.add_data(vals, titles_from_data=False)
ch.set_categories(cats)
ch.legend = None
ws_sum.add_chart(ch, f"A{r + 2}")

# ---- Sources & Methodology ----
ws_src = wb.create_sheet("Sources & Methodology")
ws_src.column_dimensions["A"].width = 18
ws_src.column_dimensions["B"].width = 24
ws_src.column_dimensions["C"].width = 110
ws_src.column_dimensions["D"].width = 12
ws_src.column_dimensions["E"].width = 38
ws_src.column_dimensions["F"].width = 70
ws_src["A1"] = "Sources & Methodology"
ws_src["A1"].font = TITLE_FONT
ws_src["A2"] = "All disclosed anchors underpinning the series, then metric-by-metric methodology. Compiled Aug 23, 2026."
ws_src["A2"].font = SUB_FONT
hr = 4
for j, h in enumerate(["Metric", "Lab / Scope", "Data point", "As of", "Source", "URL"], start=1):
    c = ws_src.cell(row=hr, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.border = BORDER
r = hr + 1
for metric, lab, point, asof, src, url in SOURCES:
    ws_src.cell(row=r, column=1, value=metric).font = BOLD
    ws_src.cell(row=r, column=2, value=lab)
    ws_src.cell(row=r, column=3, value=point).alignment = Alignment(wrap_text=True, vertical="top")
    ws_src.cell(row=r, column=4, value=asof)
    ws_src.cell(row=r, column=5, value=src)
    c = ws_src.cell(row=r, column=6, value=url)
    c.hyperlink = url
    c.font = Font(color="0563C1", underline="single")
    for j in range(1, 7):
        ws_src.cell(row=r, column=j).border = BORDER
    r += 1
r += 2
ws_src.cell(row=r, column=1, value="Methodology").font = Font(bold=True, size=13, color=DARK)
r += 1
for name, desc in METHODOLOGY:
    ws_src.cell(row=r, column=1, value=name).font = BOLD
    ws_src.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    ws_src.cell(row=r, column=3, value=desc).alignment = Alignment(wrap_text=True, vertical="top")
    ws_src.row_dimensions[r].height = max(30, 14 * (len(desc) // 110 + 1))
    r += 1
ws_src.freeze_panes = "A5"

wb.save("ai_lab_compute_economics_quarterly.xlsx")
print("saved ai_lab_compute_economics_quarterly.xlsx")
