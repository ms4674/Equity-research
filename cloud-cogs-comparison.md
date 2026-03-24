# Breakdown of COGS: Hyperscale Cloud vs Neocloud Vendors

## Executive Summary

This analysis compares the cost of goods sold (COGS) / cost of revenue structures between **hyperscale cloud providers** (AWS, Microsoft Azure, Google Cloud) and **neocloud vendors** (CoreWeave, Lambda, Crusoe). The two categories exhibit fundamentally different cost profiles driven by asset ownership models, capital structures, and go-to-market strategies. Hyperscalers benefit from diversified service portfolios and amortized infrastructure at scale, while neoclouds operate with higher capital intensity, thinner margins, and concentrated GPU-centric workloads.

---

## 1. Defining the Vendor Categories

### Hyperscale Cloud Providers
- **AWS** (Amazon) — $128.7B cloud revenue (FY2025), $45.6B operating income
- **Microsoft Azure** — 34% Azure revenue growth (FY2025), part of Intelligent Cloud segment
- **Google Cloud** — $58.7B revenue (FY2025), 36% YoY growth

### Neocloud Vendors
- **CoreWeave** — $5.1B revenue (FY2025), 168% YoY growth, ~250K GPUs globally
- **Lambda** — Private, GPU cloud focused on AI/ML workloads
- **Crusoe** — ~$276M revenue (2024 est.), leverages stranded natural gas for power

---

## 2. COGS Component Breakdown

### 2.1 Hyperscale Cloud — Cost of Revenue Components

Hyperscalers do not typically disaggregate cost of revenue at the cloud-segment level in SEC filings. Based on 10-K disclosures, industry analysis, and management commentary, the approximate composition is:

| COGS Component | % of Cost of Revenue | Notes |
|---|---|---|
| **Server & hardware depreciation** | 30–40% | Useful life 5–6 years; Amazon recently shortened AI server depreciation from 6 to 5 years |
| **Power & cooling** | 20–30% | 40–60% of data center OpEx; cooling alone consumes 30–40% of total energy |
| **Data center lease / facility costs** | 10–15% | Mix of owned and leased; long-term commitments amortized over 15–25 years |
| **Network & bandwidth** | 10–15% | Transit, peering, inter-AZ data transfer; ~90% of AWS data transfer costs in outbound + cross-AZ |
| **Technical operations staff** | 8–12% | Site reliability, hardware ops, data center technicians |
| **Software licensing & IP** | 3–5% | Third-party licenses, internal platform costs allocated to COGS |

**Key characteristics:**
- Gross margins are **high and diversified**: AWS operating margin ~35% (FY2025); Google Cloud operating income growing 2.4x YoY
- Managed/PaaS services carry 35–42% operating margins vs. IaaS at 22–26%
- Custom silicon (AWS Graviton, Google TPUs, Azure Cobalt) creates deflationary pressure on compute COGS
- Scale advantages in power procurement (long-term PPAs, renewable energy contracts)

### 2.2 Neocloud Vendors — Cost of Revenue Components

Neoclouds operate a fundamentally different model: GPU-dense infrastructure, often in colocation facilities, with revenue concentrated in bare-metal or near-bare-metal GPU compute.

| COGS Component | % of Cost of Revenue | Notes |
|---|---|---|
| **GPU & hardware depreciation** | 40–55% | Dominant cost; an 8×H100 node costs $222K–$383K; rapid obsolescence drives aggressive depreciation |
| **Power** | 20–30% | 25–40% of OpEx; GPU racks draw 50–100+ kW each; $0.05/kWh spread changes margins by $15K–$30K/node/year |
| **Colocation / facility lease** | 15–20% | Most neoclouds colocate rather than own DCs; 20–30% of total OpEx |
| **Network / interconnect** | 5–10% | InfiniBand/RoCE fabric between nodes; less outbound egress vs. hyperscalers |
| **Technical staff** | 5–8% | Lean operations teams; headcount stays relatively flat as clusters scale |
| **Interest expense (effective COGS)** | Significant | CoreWeave: $264M interest in Q1 2025 alone; $14B+ total debt — not in GAAP COGS but a real cash cost |

**Key characteristics:**
- Gross margins are **thin**: industry estimates of 14–16% after labor, power, and depreciation (McKinsey)
- Adjusted EBITDA margins appear healthy (CoreWeave ~60%) but mask massive depreciation/amortization
- GAAP operating margins have compressed: CoreWeave went from 17% (2024) to -1% (2025) as depreciation ramped
- Capital intensity is extreme: CoreWeave spent $10.3B in CapEx (FY2025) against $5.1B revenue

---

## 3. Deep Dive: Energy Costs as a Percentage of COGS

Energy (electricity for compute + cooling) is one of the most consequential COGS line items because it scales directly with capacity and is subject to external price volatility. The table below summarizes energy cost as a share of COGS across vendor types, then breaks it down further.

### 3.0 Summary Table — Energy as % of COGS

| Metric | Hyperscalers | Neoclouds |
|---|---|---|
| **Energy (power + cooling) as % of cost of revenue** | **20–30%** | **20–30%** |
| **Energy as % of total data center OpEx** | 40–60% | 25–40% (power only, excl. cooling passed through colo fees) |
| **Energy as % of total revenue** | ~7–12% | ~10–18% |
| **Cooling as % of total energy spend** | 30–40% | Bundled into colocation fees for most neoclouds |
| **Absolute cost per 8-GPU H100 node (electricity only)** | $2,400–$5,600/mo | $2,400–$5,600/mo |

The raw electricity cost per GPU hour is roughly equivalent for both vendor types — the same H100 draws the same wattage regardless of who owns it. The difference lies in how that cost flows through the P&L and what share of the revenue base it consumes.

### 3.0.1 Hyperscaler Energy Economics

For hyperscalers, energy costs represent approximately **20–30% of cost of revenue** for infrastructure services (IaaS). At the data center facility level, electricity is the single largest operating cost at **40–60% of total facility OpEx**, with cooling systems consuming an additional **30–40% of total energy draw** (reflected in PUE ratios of 1.1–1.3 for best-in-class hyperscale facilities).

As a percentage of **revenue**, energy runs approximately **7–12%** because hyperscalers carry high gross margins (60–65%) — the large revenue denominator dilutes energy's share. This is calculated as:

> Energy % of revenue = Energy % of COGS × (1 – Gross Margin)
> ≈ 25% × 35–40% COGS-to-revenue = ~9–10% of revenue

Key hyperscaler energy advantages:
- **Long-term PPAs** at $0.03–$0.05/kWh vs. spot rates of $0.07–$0.22/kWh
- **Owned substations and dedicated grid connections** that eliminate markup
- **Custom silicon** (Graviton, TPU, Trainium) that delivers 30–50% better performance per watt vs. NVIDIA GPUs for compatible workloads, effectively reducing energy COGS per unit of compute sold
- **AI-optimized cooling** (Google claims 40% reduction in cooling energy through ML-driven controls)
- **PUE of 1.1–1.2** at best facilities, meaning only 10–20% overhead on top of IT power draw

### 3.0.2 Neocloud Energy Economics

For neoclouds, electricity represents approximately **25–40% of operating expenses** and **20–30% of cost of revenue**. The percentage range is wider than hyperscalers because neoclouds vary significantly in their power sourcing:

| Neocloud | Power Strategy | Estimated Electricity Cost |
|---|---|---|
| **CoreWeave** | Colocation (3.1 GW contracted across 40+ sites) | Market rate via colo providers |
| **Lambda** | Colocation, multiple regions | Market rate, varies by region |
| **Crusoe** | Stranded natural gas generation | ~$0.01–$0.03/kWh (~1/13th standard rates) |

As a percentage of **revenue**, energy costs run higher for neoclouds at approximately **10–18%** because their gross margins are far thinner (14–16% fully loaded), so COGS components consume a larger share of each revenue dollar:

> Energy % of revenue = Energy % of COGS × COGS-to-revenue ratio
> ≈ 25% × 65–85% COGS-to-revenue = ~16–21% of revenue (at fully loaded margins)
> ≈ 25% × 40–50% COGS-to-revenue = ~10–13% of revenue (at adjusted margins)

The neocloud energy cost sensitivity is acute: a **$0.05/kWh difference** in electricity rates changes annual margins by **$15,000–$30,000 per 8-GPU node**. For a fleet of 1,000 nodes, that is a $15M–$30M annual margin swing — which can be the difference between profitability and loss.

### 3.0.3 Why Energy % of COGS Is Similar Despite Different Business Models

The paradox is that energy as a percentage of COGS is roughly comparable (20–30%) for both hyperscalers and neoclouds, despite vastly different business models. This is because:

1. **Physics is the same**: A given GPU or server draws a fixed amount of power regardless of who owns it. The electricity input per unit of compute is an engineering constant.
2. **Depreciation scales proportionally**: Both vendor types have hardware depreciation as their largest COGS component (30–55%), and energy scales roughly in proportion to the hardware base.
3. **The difference shows up in revenue, not COGS**: Hyperscalers charge 3–5x more per GPU hour for the same hardware, so the same energy cost consumes a much smaller share of revenue. Neoclouds compete on price, so energy eats a larger share of their thinner margins.

The strategic implication: **energy efficiency and procurement are margin-critical for neoclouds but margin-enhancing for hyperscalers**. A 20% reduction in power costs might add 1–2 percentage points of margin for a hyperscaler but could double a neocloud's gross profit.

---

## 4. Comparative Analysis

### 4.1 Gross Margin Comparison

| Metric | Hyperscalers | Neoclouds |
|---|---|---|
| **Gross margin (estimated)** | 60–65% | 35–50% (adj.); 14–16% (fully loaded) |
| **EBITDA margin** | 35–45% | 55–65% (adjusted, pre-depreciation) |
| **GAAP operating margin** | 25–37% | -1% to 17% |
| **Depreciation as % of revenue** | 15–20% | 35–50%+ |

The divergence between adjusted EBITDA and GAAP operating income for neoclouds is a critical analytical point. Neocloud EBITDA margins (~60%) suggest a healthy business, but this metric excludes the single largest cost driver — hardware depreciation — which is the core "raw material" of the business.

### 4.2 Unit Economics: GPU Hour

| Cost Component | Hyperscaler (H100/hr) | Neocloud (H100/hr) |
|---|---|---|
| **List price** | ~$12.25 (on-demand) | ~$2.49 (on-demand) |
| **Hardware depreciation** | ~$1.50–2.00 | ~$1.50–2.00 |
| **Power** | ~$0.40–0.80 | ~$0.40–0.80 |
| **Facility / overhead** | ~$0.50–1.00 | ~$0.30–0.60 |
| **Network** | ~$0.30–0.50 | ~$0.10–0.20 |
| **Gross profit per GPU-hr** | ~$7.00–9.00 | ~$-0.50–0.50 |

Hyperscalers bundle GPUs with managed services, monitoring, security, and ecosystem lock-in that justify premium pricing. Neoclouds compete on raw price but carry similar underlying hardware costs, resulting in far thinner margins.

### 4.3 Structural COGS Advantages by Vendor Type

| Advantage | Hyperscalers | Neoclouds |
|---|---|---|
| **Power procurement** | Long-term PPAs, renewable contracts, owned substations | Some advantages (Crusoe: stranded gas at ~1/13th standard rates); most rely on colo provider |
| **Hardware procurement** | Custom silicon + NVIDIA volume discounts | NVIDIA volume pricing, but less diversification |
| **Facility costs** | Own many DCs; amortized over decades | Colocate; shorter lease terms, higher per-MW cost |
| **Software/platform** | Massive internal platform reduces per-unit costs | Minimal software layer; lower platform COGS but also lower ASPs |
| **Labor leverage** | Spread across millions of customers and services | Lean teams but fewer customers to amortize against |
| **Deployment speed** | 3–5 years for new DCs | 6–18 months (speed advantage) |

---

## 5. Capital Structure Impact on Effective COGS

While interest expense is not classified as COGS under GAAP, the neocloud business model is so capital-intensive that debt service is effectively part of the unit cost of delivering compute.

### CoreWeave Capital Intensity (FY2025)
| Metric | Value |
|---|---|
| Revenue | $5.1B |
| CapEx | $10.3B |
| Total debt | $14B+ ($9.7B due within 12 months) |
| Interest expense (Q1 2025) | $264M |
| Annualized interest | ~$1.0–1.2B |
| Interest as % of revenue | ~20–24% |

### Hyperscaler Capital Intensity (FY2025)
| Metric | AWS | Google Cloud | Azure (est.) |
|---|---|---|---|
| Revenue | $128.7B | $58.7B | ~$80B+ |
| Parent CapEx | ~$100B | ~$75B | ~$80B |
| Debt cost allocated to cloud | Minimal (funded from ops) | Minimal | Minimal |
| Interest as % of cloud revenue | <2% | <2% | <2% |

Hyperscalers fund infrastructure from operating cash flow and investment-grade debt at low rates. Neoclouds rely heavily on GPU-backed debt facilities at higher rates, making interest a quasi-COGS item that hyperscalers largely avoid.

---

## 6. COGS Trajectory & Outlook

### Headwinds (Both)
- **Accelerating depreciation**: Hardware generations turning over every 1–2 years (H100 → B200 → next-gen) compresses useful life assumptions
- **Power costs rising**: Data center power demand growing faster than supply; 42% of operators cite power as top cost increase area
- **Cooling complexity**: AI-density racks (50–100+ kW) require liquid cooling, adding 20–30% to facility costs

### Hyperscaler-Specific
- Custom silicon (Graviton, TPU, Trainium) offsets GPU COGS by 30–50% for compatible workloads
- Software-defined infrastructure enables higher utilization rates (70–85%)
- Risk of depreciation charges hitting P&L harder as $75–100B/year CapEx cycles flow through (Alphabet CFO warning: "depreciation is going to be hitting the P&L statements a bit harder than in the past")

### Neocloud-Specific
- Utilization is the margin lever: below 60%, most GPU clusters lose money; above 80%, margins expand rapidly
- Long-term contracts (CoreWeave: $14.2B deal with Meta) reduce revenue volatility but lock in pricing
- Hardware refresh risk: ~$200K+ per node replacement every 3–4 years
- Debt maturity walls: CoreWeave has $9.7B due within 12 months as of Q3 2025

---

## 7. Key Takeaways for Investors

1. **Neocloud EBITDA margins are misleading**: The ~60% adjusted EBITDA margin headline masks the fact that depreciation (the core input cost) is excluded. Fully loaded gross margins are 14–16%.

2. **Hyperscaler COGS benefits from diversification**: Revenue from high-margin PaaS/SaaS services cross-subsidizes infrastructure-heavy IaaS, creating a blended margin profile that neoclouds cannot replicate.

3. **Power is the great equalizer — and differentiator**: Both vendor types face power as 20–30% of COGS, but hyperscalers have structural procurement advantages through scale. Crusoe's stranded-gas model is a notable exception among neoclouds.

4. **Interest expense is the hidden neocloud COGS tax**: At 20–24% of revenue for CoreWeave, debt service is effectively the second-largest cost component after depreciation, yet sits below the gross profit line.

5. **Custom silicon is the hyperscaler moat**: AWS Graviton, Google TPU/Trillium, and Azure Cobalt allow hyperscalers to reduce per-unit compute COGS by 30–50% for workloads that can run on proprietary chips, a structural advantage neoclouds cannot match.

6. **Neocloud pricing advantage is narrowing**: As hyperscalers build dedicated GPU capacity and offer reserved/committed pricing, the 60–70% discount that neoclouds offered in 2023–2024 is compressing, while neocloud costs remain relatively fixed.

---

## Sources & Methodology

- Amazon, Microsoft, and Alphabet SEC filings (10-K, 10-Q, earnings releases) — FY2024 and FY2025
- CoreWeave S-1 (March 2025), quarterly earnings releases, and investor presentations
- McKinsey analysis of GPU cloud margins
- Uptime Institute 2025 data center spending survey
- Industry sources: Data Center Dynamics, The Next Platform, Compute Forecast
- Vendor pricing: GPULeaseIndex, ComputePrices (H100/B200 on-demand rates as of early 2026)

*Note: Hyperscalers do not fully disaggregate cloud-segment COGS in public filings. Percentage estimates are derived from management commentary, footnote disclosures, and industry benchmarks. Neocloud financials are based primarily on CoreWeave's public filings as the only publicly traded pure-play neocloud.*

---

*Last updated: March 2026*
