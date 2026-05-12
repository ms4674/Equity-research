# AI vs. Traditional Data Center Economics: Rack Density and OpEx Comparison

**Equity Research Supplement | May 2026**

---

## Executive Summary

AI data centers and traditional cloud/enterprise data centers are fundamentally different infrastructure categories. A single AI training rack now consumes 100-200 kW — the equivalent of 15-30 traditional server racks. This density difference cascades into every aspect of design, construction, cooling, and operating costs. AI data centers carry approximately **3-5x higher OpEx per MW** than traditional facilities, driven primarily by GPU hardware refresh cycles, liquid cooling operations, and specialized staffing. However, they generate **5-10x higher revenue per MW**, making the unit economics overwhelmingly favorable despite the cost premium.

---

## 1. Rack Power Density: AI vs. Traditional

### The Density Gap

| Category | kW per Rack | Typical Examples |
|---|---|---|
| **Traditional enterprise** | 5-8 kW | Corporate IT, email servers, ERP systems |
| **Traditional cloud/colocation** | 8-15 kW | AWS general-purpose, web hosting, SaaS |
| **Hyperscale cloud (modern)** | 15-25 kW | Hyperscaler CPU workloads, storage |
| **Industry average (2026)** | 27 kW | Blended across all types |
| **AI inference** | 30-80 kW | Serving trained models, chatbots, image generation |
| **AI training (current gen)** | 80-132 kW | H100/H200 DGX clusters, GB200 NVL72 |
| **AI training (next gen, 2027)** | 250-600 kW | NVIDIA Rubin NVL576 platform |

### GPU System-Level Power by Platform

| Platform | GPUs per Rack/System | System Power | Effective kW/GPU (system-level) | Cooling Required |
|---|---|---|---|---|
| DGX H100 (8 GPUs) | 8 | 10.2 kW | 1.28 kW | Air (marginal) |
| HGX H100 (8 per node, 4 nodes/rack) | 32 | ~40-45 kW/rack | ~1.3 kW | Air or rear-door liquid |
| DGX B200 (8 GPUs) | 8 | 14.3 kW | 1.79 kW | Liquid (required) |
| GB200 NVL72 (72 GPUs) | 72 | 120 kW/rack | 1.67 kW | Direct liquid (mandatory) |
| GB300 NVL72 (Blackwell Ultra) | 72 | ~140-160 kW/rack (est.) | ~2.0-2.2 kW | Direct liquid (mandatory) |
| Rubin NVL576 (2027) | 576 (multi-rack) | ~500-600 kW | ~1.0 kW | Advanced liquid (mandatory) |

### The Multiplier: AI vs. Traditional

| Comparison | AI (GB200 NVL72) | Traditional Cloud | Ratio |
|---|---|---|---|
| Power per rack | 120 kW | 8-12 kW | **10-15x** |
| Heat density (kW/sq ft) | 3-5 kW/sq ft | 0.2-0.4 kW/sq ft | **10-15x** |
| Cooling requirement | 100+ kW rejection/rack | 8-12 kW rejection/rack | **10-15x** |
| Floor space per MW | ~8 racks | ~80-125 racks | **10-15x fewer racks** |
| Racks per MW | 8 | 80-125 | AI uses far fewer, denser racks |

### Key Insight

A 100 MW AI data center contains roughly **800-1,000 racks** (at 100-120 kW each), while a 100 MW traditional cloud data center contains **8,000-12,000 racks** (at 8-12 kW each). The AI facility is physically much smaller but requires radically different power distribution, cooling infrastructure, and structural engineering.

---

## 2. Cooling Architecture Comparison

The single biggest physical design difference between AI and traditional data centers is the cooling system.

### Cooling Technology by Density Tier

| Rack Density | Viable Cooling Technology | PUE Achievable | Status for AI |
|---|---|---|---|
| 5-15 kW/rack | Traditional hot-aisle/cold-aisle air | 1.50-1.80 | Irrelevant — cannot support AI |
| 15-25 kW/rack | Enhanced air + in-row cooling | 1.30-1.50 | Marginal — legacy AI inference only |
| 25-40 kW/rack | Rear-door heat exchangers (RDHx) | 1.25-1.45 | Transitional — older GPU clusters |
| 40-80 kW/rack | Direct-to-chip liquid cooling (DLC) | 1.10-1.20 | Standard for current AI training |
| 80-132 kW/rack | Full liquid cooling (DLC + CDU) | 1.08-1.15 | Required for GB200 NVL72 |
| 132-600 kW/rack | Advanced immersion or DLC | 1.03-1.10 | Required for next-gen (Rubin) |

### Cooling Performance Impact

Liquid cooling delivers measurable performance benefits beyond just enabling higher density:

| Metric | Air Cooled | Liquid Cooled | Improvement |
|---|---|---|---|
| GPU temperature | 55-71°C | 46-54°C | 15-25°C lower |
| Power per node | Baseline | -1 kW/node (-16%) | 16% energy savings |
| Training throughput | Baseline | +1.4% to +29% | Faster model training |
| Stress test throughput | Baseline | +17% | Higher sustained perf |
| PUE | 1.40-1.80 | 1.08-1.20 | 25-40% overhead reduction |

### Cooling CapEx Comparison

| System | Cost per kW (cooling infrastructure) | Cost for 100 MW |
|---|---|---|
| Traditional air (CRAH/CRAC) | $1,500-2,500/kW | $150-250M |
| Rear-door heat exchangers | $2,500-4,000/kW | $250-400M |
| Direct liquid cooling (DLC) | $3,000-5,000/kW | $300-500M |
| Single-phase immersion | $4,000-6,000/kW | $400-600M |
| Two-phase immersion | $5,000-8,000/kW | $500-800M |

**Net impact:** Liquid cooling adds 15-25% to mechanical construction budgets vs. air cooling, but reduces ongoing energy costs by 20-35% and enables 5-15x higher rack density.

---

## 3. OpEx Comparison: AI vs. Traditional Data Centers

### Overall OpEx per MW

| Category | Traditional Cloud/Colo | AI Data Center | AI Premium |
|---|---|---|---|
| **Total annual OpEx per MW** | $0.8-1.2M | $3.0-4.0M | **3-4x** |
| **Total OpEx (excl. GPU refresh)** | $0.8-1.2M | $1.2-1.8M | **1.5-2x** |
| **Total OpEx (incl. GPU amortization)** | N/A | $8-15M | **10-15x** (not comparable) |

*Note: The massive difference when including GPU amortization reflects the capital-intensive nature of AI infrastructure where IT hardware dominates total cost of ownership.*

### OpEx Breakdown: Component-by-Component

#### A. Electricity Cost (Largest Variable)

| Parameter | Traditional Cloud | AI Data Center | Notes |
|---|---|---|---|
| PUE | 1.40-1.60 | 1.10-1.25 | AI achieves better PUE via liquid cooling |
| Effective power draw (% of capacity) | 60-70% utilization | 80-90% utilization | AI runs hotter, near capacity |
| Annual kWh per MW of capacity | 5.3-6.1M kWh | 7.0-7.9M kWh | Higher utilization in AI |
| Electricity rate | $0.06-0.10/kWh | $0.05-0.08/kWh | AI operators seek cheap power |
| **Annual electricity cost per MW** | **$320K-$610K** | **$350K-$630K** | Roughly comparable per MW |

Despite different architectures, electricity cost per MW is similar because AI's higher utilization is offset by better PUE. The real difference is that AI fills each MW with far more expensive hardware, making the revenue-to-power-cost ratio much more favorable.

#### B. Cooling System Operations

| Parameter | Traditional Cloud (Air) | AI Data Center (Liquid) | Difference |
|---|---|---|---|
| Cooling energy (% of IT load) | 30-45% | 8-20% | AI uses 50-75% less cooling energy |
| Coolant/fluid costs | None (air) | $50-150K/MW/year | New cost category |
| CDU maintenance | N/A | $30-80K/MW/year | New cost category |
| Chiller plant maintenance | $40-80K/MW/year | $20-50K/MW/year | Smaller liquid systems |
| **Total cooling OpEx per MW** | **$180-350K** | **$100-280K** | AI slightly cheaper per MW |

Counterintuitively, liquid cooling can reduce total cooling OpEx per MW because it eliminates massive air-handling infrastructure and reduces the cooling energy multiplier from 30-45% to 8-20% of IT load.

#### C. Staffing

| Parameter | Traditional Cloud | AI Data Center | Difference |
|---|---|---|---|
| Staff per MW | 3-5 FTEs | 2-4 FTEs | AI is denser; fewer racks per MW |
| Avg fully-loaded cost/FTE | $120-160K | $180-250K | AI requires specialized engineers |
| **Annual staffing cost per MW** | **$360K-$800K** | **$400K-$1,000K** | AI ~1.2-1.5x due to salary premium |

AI data centers require fewer total staff per MW (fewer racks to maintain) but demand higher-paid specialists (liquid cooling engineers, GPU/networking experts, ML infrastructure engineers).

#### D. Maintenance and Replacement

| Parameter | Traditional Cloud | AI Data Center | Difference |
|---|---|---|---|
| Server refresh cycle | 4-6 years | 2-3 years (GPUs) | AI refreshes 2x faster |
| Hardware failure rate | 1-3% annually | 3-8% annually | GPUs fail more often at extreme load |
| Maintenance contracts per MW | $50-120K/year | $150-400K/year | Higher due to complexity |
| Network equipment maintenance | $30-60K/MW/year | $100-250K/MW/year | NVLink, InfiniBand, high-speed fabric |
| **Total maintenance per MW** | **$80-180K** | **$250-650K** | AI 2-4x higher |

#### E. Software & Licensing

| Parameter | Traditional Cloud | AI Data Center | Difference |
|---|---|---|---|
| Virtualization/orchestration | $20-50K/MW/year | $100-300K/MW/year | Kubernetes, Slurm, custom orchestration |
| Networking software | $10-30K/MW/year | $50-150K/MW/year | InfiniBand management, NCCL |
| Monitoring/observability | $10-20K/MW/year | $30-80K/MW/year | GPU health monitoring, thermal |
| **Total software per MW** | **$40-100K** | **$180-530K** | AI 3-5x higher |

### Summary: OpEx per MW (Annual)

| Component | Traditional Cloud | AI Data Center | AI Multiple |
|---|---|---|---|
| Electricity | $320-610K | $350-630K | ~1.0x |
| Cooling operations | $180-350K | $100-280K | ~0.7x (better) |
| Staffing | $360-800K | $400-1,000K | ~1.3x |
| Maintenance | $80-180K | $250-650K | ~3x |
| Software/licensing | $40-100K | $180-530K | ~4x |
| Insurance & other | $40-80K | $80-200K | ~2x |
| **Total OpEx/MW (facility only)** | **$1.0-2.1M** | **$1.4-3.3M** | **1.5-2x** |
| GPU amortization (if included) | N/A ($2-4M servers) | $7-12M/MW | Dominant cost |
| **Total incl. IT hardware amort.** | **$3-6M** | **$8-15M** | **2.5-3.5x** |

---

## 4. CapEx Comparison: Building the Facility

### Construction Cost per MW (Shell + MEP, excluding IT)

| Facility Type | CapEx per MW | Key Cost Drivers |
|---|---|---|
| Traditional enterprise (Tier III) | $5-7M/MW | Standard air cooling, conventional power |
| Traditional cloud/hyperscale | $7-9M/MW | Higher redundancy, larger scale |
| AI-optimized (liquid cooled) | $10-14M/MW | Liquid cooling plant, 480V power distribution, structural reinforcement |
| AI premium GW-scale campus | $12-16M/MW | On-site generation, advanced cooling, custom electrical |

### What Drives the AI CapEx Premium

| Factor | Cost Adder vs. Traditional | Explanation |
|---|---|---|
| Liquid cooling infrastructure | +$1.5-3M/MW | CDUs, piping, heat rejection, fluid management |
| Electrical upgrades (480V dist.) | +$0.5-1.5M/MW | Higher amperage, busway upgrades for dense loads |
| Structural reinforcement | +$0.3-0.8M/MW | 120 kW racks are significantly heavier |
| On-site power generation (BTM) | +$1-3M/MW | Gas turbines, fuel cells for queue bypass |
| Enhanced fire suppression | +$0.2-0.5M/MW | Different approach required for liquid-cooled environments |
| **Total AI premium** | **+$3.5-9M/MW** | Roughly 50-100% cost increase vs. traditional |

### Total Deployed Capital per MW (Including IT Hardware)

| Facility Type | Shell + MEP | IT Hardware | Total Capital/MW |
|---|---|---|---|
| Traditional enterprise | $5-7M | $3-5M (servers, storage) | $8-12M |
| Traditional cloud | $7-9M | $4-7M (servers, networking) | $11-16M |
| AI data center (H100 era) | $10-12M | $20-27M (GPUs + networking) | $30-39M |
| AI data center (Blackwell era) | $12-14M | $25-35M (GPUs + networking) | $37-49M |

**The IT hardware (GPUs) now represents 60-70% of total deployed capital** in an AI data center, vs. 30-40% in traditional facilities. This inverts the historical relationship where the building/MEP was the dominant capital cost.

---

## 5. Unit Economics Comparison

### Revenue per MW

| Metric | Traditional Cloud/Colo | AI GPU Cloud | Multiple |
|---|---|---|---|
| Annual revenue per MW | $1.5-3.0M | $6-12M | **4-6x** |
| Revenue per rack per month | $1,000-3,000 | $10,000-80,000 | **5-30x** |
| Revenue per kWh consumed | $0.20-0.40 | $0.80-1.60 | **4-5x** |

### Margin Profile

| Metric | Traditional Colo | AI GPU Cloud (Full-stack) |
|---|---|---|
| Gross margin | 50-65% | 55-70% |
| EBITDA margin | 40-55% | 40-55% |
| OpEx as % of revenue | 35-50% | 25-45% |
| Power as % of revenue | 15-25% | 5-8% |
| **Revenue payback on total capital** | 4-7 years | 2-3.5 years |

Despite AI facilities costing 3-4x more per MW to build and equip, the revenue premium (4-6x) more than compensates, yielding faster payback periods and similar or better margins.

### Return on Invested Capital (ROIC)

| Model | Total Capital/MW | Annual Rev/MW | Annual EBITDA/MW | Implied ROIC |
|---|---|---|---|---|
| Traditional colocation | $10-14M | $2.0-2.5M | $0.9-1.4M | 7-12% |
| AI GPU cloud (full-stack) | $35-45M | $8-12M | $3.5-5.5M | 8-14% |
| AI GPU cloud (bare-metal) | $25-30M | $3-5M | $1.5-2.5M | 5-9% |

ROIC is comparable across models because the higher revenue in AI is matched by proportionally higher capital intensity. The key differentiator is **speed of payback** and **absolute dollar return** per MW, which strongly favors AI.

---

## 6. The Density-Cost Flywheel

### Why Higher Density Lowers Some Costs

Higher rack density in AI data centers creates counterintuitive cost advantages:

1. **Less floor space per MW** → lower real estate cost per MW
2. **Fewer racks per MW** → fewer points of failure, less cabling, fewer maintenance touchpoints
3. **Liquid cooling** → lower PUE → less energy waste per useful compute watt
4. **Higher utilization** → more revenue per MW (AI workloads run at 80-90% vs. cloud at 60-70%)

### Why Higher Density Raises Other Costs

1. **Thermal density** → requires expensive liquid cooling infrastructure
2. **Power distribution** → needs 480V+ delivery at extreme amperage per rack
3. **Single point of failure risk** → one rack failure affects more compute; requires more redundancy
4. **Hardware replacement velocity** → 2-3 year GPU refresh vs. 4-6 year server refresh
5. **Specialized expertise** → higher salaries for liquid cooling and GPU infrastructure engineers

---

## 7. Forward Look: Convergence or Divergence?

### The Density Roadmap

| Year | AI Rack Density | Traditional Rack Density | Gap |
|---|---|---|---|
| 2023 | 40-60 kW | 8-12 kW | 5x |
| 2025 | 80-132 kW | 10-15 kW | 8-10x |
| 2027 (projected) | 250-600 kW | 12-20 kW | 15-30x |
| 2029 (projected) | 500-1,000+ kW | 15-25 kW | 25-50x |

The gap is widening, not converging. AI infrastructure is evolving toward a fundamentally different physical form factor: liquid-cooled, ultra-dense, modular systems that bear little resemblance to traditional raised-floor data centers.

### Implications for Data Center REITs and Operators

1. **Existing traditional facilities cannot be trivially converted to AI.** Retrofitting for liquid cooling, 480V power, and 100+ kW racks requires gutting most of the MEP infrastructure. The floor loading alone (traditional: 150-250 lbs/sq ft; AI: 350-500+ lbs/sq ft) may require structural reinforcement.

2. **AI facilities have shorter economic lives.** With 2-3 year GPU refresh cycles and rapidly evolving cooling requirements, an AI data center may become obsolete in 5-7 years vs. 15-20 years for traditional facilities.

3. **Different site selection criteria.** AI data centers prioritize power availability and cooling water access over proximity to population centers. Traditional data centers prioritize latency to end users.

4. **OpEx profile shifts.** As AI grows from 4% to potentially 40-50% of total data center power demand by 2030, the industry cost structure will increasingly reflect AI economics: higher capital intensity, faster refresh, higher revenue per MW, and greater power procurement sophistication.

---

## 8. Summary Comparison Table

| Parameter | Traditional Cloud | AI Data Center | Ratio |
|---|---|---|---|
| **Rack density** | 8-15 kW | 80-132 kW (current) | 8-15x |
| **Cooling method** | Air | Direct liquid | — |
| **PUE** | 1.40-1.60 | 1.08-1.25 | AI is more efficient |
| **Construction CapEx/MW** | $7-9M | $10-14M | 1.4-1.8x |
| **IT Hardware CapEx/MW** | $4-7M | $20-35M | 4-6x |
| **Total deployed capital/MW** | $11-16M | $35-49M | 3-4x |
| **Annual OpEx/MW (facility)** | $1.0-2.1M | $1.4-3.3M | 1.5-2x |
| **Annual OpEx/MW (incl. hardware amort.)** | $3-6M | $8-15M | 2.5-3.5x |
| **Annual revenue/MW** | $1.5-3.0M | $6-12M | 4-6x |
| **Power as % of revenue** | 15-25% | 5-8% | AI monetizes power better |
| **Payback period** | 4-7 years | 2-3.5 years | AI pays back faster |
| **Economic facility life** | 15-20 years | 5-10 years | AI turns over faster |
| **Staff per MW** | 3-5 FTE | 2-4 FTE | AI needs fewer but pricier |
| **Utilization rate** | 60-70% | 80-90% | AI runs hotter |
| **Racks per MW** | 80-125 | 8-12 | AI is 10x denser |

---

*Disclaimer: Figures represent industry averages and ranges based on publicly available data from Turner & Townsend, Uptime Institute, NVIDIA specifications, KPMG benchmarks, and operator disclosures as of May 2026. Actual costs vary significantly by geography, operator, scale, and specific technology deployment.*
