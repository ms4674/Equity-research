# Power Procurement Strategy for Data Centers: Hyperscalers and Neoclouds

**Equity Research Report | May 2026**

---

## Executive Summary

Power has replaced compute as the primary constraint on AI infrastructure expansion. The hyperscalers (Microsoft, Amazon/AWS, Google, Meta) and neoclouds (CoreWeave, Crusoe, Lambda) are engaged in an unprecedented race to secure gigawatt-scale electricity supply for data centers. This report examines the strategic shift from passive energy consumption to active power infrastructure development, the emerging procurement models, cost structures, regulatory dynamics, and the resulting competitive moats that will define the next decade of AI infrastructure investment.

**Key thesis:** Operators that locked in long-term power agreements early (2024-2025) at favorable pricing have created structural cost advantages that will compound as power demand outstrips supply through 2030. The market is bifurcating into power "haves" and "have-nots," with implications for cloud service pricing, AI training costs, and the investability of the entire data center value chain.

---

## Table of Contents

1. [The Power Demand Landscape](#1-the-power-demand-landscape)
2. [Procurement Models: A Taxonomy](#2-procurement-models-a-taxonomy)
3. [Hyperscaler Strategies: Company-by-Company Analysis](#3-hyperscaler-strategies-company-by-company-analysis)
4. [Neocloud Strategies: Differentiated Approaches](#4-neocloud-strategies-differentiated-approaches)
5. [The Nuclear Pivot](#5-the-nuclear-pivot)
6. [Behind-the-Meter and On-Site Generation](#6-behind-the-meter-and-on-site-generation)
7. [Grid Interconnection: The Bottleneck](#7-grid-interconnection-the-bottleneck)
8. [Regulatory Landscape](#8-regulatory-landscape)
9. [Cost Economics and PPA Pricing](#9-cost-economics-and-ppa-pricing)
10. [Competitive Moats and Investment Implications](#10-competitive-moats-and-investment-implications)
11. [Key Risks](#11-key-risks)
12. [Conclusion](#12-conclusion)

---

## 1. The Power Demand Landscape

### Current State

U.S. data centers consumed approximately 177-192 TWh of electricity in 2024, representing roughly 4-5% of total U.S. electricity consumption. This figure is projected to grow dramatically:

| Metric | 2024 | 2026E | 2028E | 2030E |
|---|---|---|---|---|
| U.S. DC Electricity Demand (TWh) | ~183 | ~260 | 325-580 | 380-790 |
| Share of U.S. Electricity | ~4.5% | ~6% | 6.7-12% | 9-17% |
| Peak Power Demand (GW) | ~53 | ~76 | ~90-110 | ~90-140 |

*Sources: EPRI, IEA, Lawrence Berkeley National Laboratory, Grid Strategies LLC*

### Growth Drivers

- **AI training clusters** requiring 100 MW-1 GW+ of continuous power per facility
- **Inference scaling** as AI models are deployed at consumer/enterprise scale
- **GW-scale campuses** becoming the norm: 16 GW-scale data centers with ~30 GW aggregate demand are scheduled to come online in 2026-2027 alone
- **Average facility size step-up** from ~30 MW (historical average) to 100+ MW as standard and 500 MW-1 GW for frontier projects

### The Fundamental Mismatch

Data centers can be built in 18-24 months. High-voltage transmission upgrades require 7-10 years to plan, approve, and construct. New generation sources spend 4-5 years in interconnection queues before coming online. This temporal mismatch is the central tension driving every procurement strategy discussed in this report.

---

## 2. Procurement Models: A Taxonomy

The market has evolved from a single dominant model (grid PPAs) to a diversified toolkit of power procurement strategies. Each carries distinct risk/return, timeline, and scalability profiles.

### 2.1 Traditional Grid Interconnection

- **Timeline:** 40-60 months (national average 53 months; PJM 58 months; CAISO 56 months)
- **Capex:** $2M-$5M per MW
- **Regulatory complexity:** High (ISO queue, utility approval, transmission upgrades)
- **Best for:** Patient capital, established markets with available capacity

### 2.2 Long-Term Power Purchase Agreements (PPAs)

Corporate PPAs with utilities, independent power producers, or renewable developers. This remains the most common procurement tool but is evolving from single-asset deals to multi-gigawatt portfolio frameworks.

- **Renewable PPAs:** Solar $61.7/MWh (Q4 2025 avg), wind $73.7/MWh (Q4 2025 avg); rising 8-16% YoY
- **Nuclear PPAs:** Existing reactors $45-70/MWh; forward-start SMRs $65-100/MWh
- **Typical term:** 10-20 years
- **Key shift:** Hyperscalers now demand hourly load matching (not annual RECs), which structurally favors 24/7 baseload sources (nuclear, geothermal) over intermittent renewables

### 2.3 Behind-the-Meter (BTM) Generation

On-site power generation that bypasses the grid entirely. The fastest-growing category.

- **Timeline:** 12-24 months (natural gas); 2-6 months (fuel cells)
- **Capex:** $3M-$8M per MW
- **Technologies:** Modular natural gas turbines/reciprocating engines, solid-oxide fuel cells (Bloom Energy), battery-gas hybrids
- **Industry estimate:** 25%+ of incremental data center megawatts this decade will be BTM

### 2.4 Co-Location with Existing Generation

Siting data centers adjacent to existing power plants (nuclear, gas) and leveraging the plant's grid interconnection.

- **Timeline:** 18-36 months
- **Capex:** $4M-$7M per MW
- **Examples:** AWS/Talen Susquehanna campus, Google/Intersect Power partnership
- **Key advantage:** Avoids the load interconnection queue; leverages existing transmission capacity

### 2.5 Direct Asset Ownership/Acquisition

The most aggressive model: hyperscalers directly acquiring, investing in, or co-developing power generation assets.

- **Examples:** Google's acquisition of Intersect Power (Dec 2025); Amazon's $700M investment in X-energy
- **Signals a shift** from energy consumer to integrated tech-energy company

### 2.6 Electric Service Agreements / Utility Partnerships

Structured agreements with regulated utilities where data center operators commit to pay 100% of interconnection costs, fund new generation, and guarantee contracted costs regardless of energy use.

- **Example:** Google/OG&E agreement in Oklahoma (Apr 2026); NiSource deals with Alphabet and Amazon in Indiana
- **Ratepayer protection** built in: data center pays full cost of infrastructure, shielding existing customers

### Comparison Matrix

| Pathway | Speed | Cost/MW | Grid Risk | Carbon Intensity | Scalability |
|---|---|---|---|---|---|
| Grid Interconnection | Slow (40-60 mo) | Low ($2-5M) | Full grid dependency | Grid mix | Unlimited (if available) |
| Renewable PPA | Medium (18-36 mo) | Medium ($3-6M) | Partial | Low/Zero | High |
| Nuclear PPA | Slow (existing) / Very slow (SMR) | High ($5-10M) | Low | Zero | High (if available) |
| BTM Natural Gas | Fast (12-24 mo) | Medium ($3-6M) | None | High | Medium |
| BTM Fuel Cells | Very Fast (2-6 mo) | High ($5-8M) | None | Medium-Low | Low-Medium |
| Co-Location | Medium (18-36 mo) | Medium-High ($4-7M) | Low | Depends on source | Limited by existing assets |
| Direct Ownership | Slow (varies) | Highest | None | Controlled | Highest long-term |

---

## 3. Hyperscaler Strategies: Company-by-Company Analysis

### 3.1 Microsoft

**Strategy: Programmatic, market-making renewable procurement + nuclear restart underwriting**

Microsoft has pursued the most aggressive and diversified power procurement strategy among the hyperscalers. The company has transitioned from incremental PPAs to programmatic, supply-chain-level engagements.

**Scale of commitment:**
- $80 billion global data center infrastructure build-out (announced Jan 2025)
- $50 billion specifically for AI and data centers (Feb 2026)
- 40 GW of renewable capacity contracted across 26 countries (matching 100% of 2025 global electricity consumption)

**Key deals:**
- **Brookfield:** 10.5 GW framework agreement (the largest corporate PPA ever)
- **Qcells:** 12 GW solar supply and EPC alliance (domestic U.S. supply chain)
- **Constellation Energy / Crane Clean Energy Center (TMI Unit 1):** 20-year PPA, 835 MW, ~$100-102/MWh estimated pricing, plant restart expected 2027. Microsoft directly underwrote the $1.6B restart with a $16B total contract value
- **Regional diversification:** $19B Canada, $15.2B UAE, EUR 8.6B Portugal, plus major PPAs in Spain, Italy, Norway, Kenya (geothermal)

**Assessment:** Microsoft is the most advanced in transforming from energy buyer to market-maker. The Brookfield and Qcells deals are not procurement transactions; they are supply chain construction. The TMI restart demonstrates willingness to pay ~$100+/MWh for 24/7 firm power, roughly double merchant rates, setting a price signal for the entire nuclear-for-AI market.

### 3.2 Amazon / AWS

**Strategy: Direct nuclear offtake + co-location + SMR investment**

AWS has taken the most direct approach to securing physical power, combining the largest nuclear PPA in corporate history with strategic investments in next-generation reactor technology.

**Key deals:**
- **Talen Energy / Susquehanna:** $18B, 17-year PPA for 1.92 GW from the 2.5 GW Susquehanna nuclear plant. The deal transitioned from a behind-the-meter arrangement to a front-of-meter PPA after FERC regulatory uncertainty around co-location. Full volume delivery by 2032
- **X-energy:** $700M investment for up to 12 Xe-100 SMRs (960 MW, Cascade Advanced Energy Facility). First phase: 4 units (320 MW), targeting 2030s
- **Dominion Energy:** Virginia SMR development (capacity TBD)
- **NiSource:** Long-term energy deal for Indiana data centers (Apr 2026)

**Implied pricing:** The AWS-Talen deal is estimated at ~$68.90/MWh, establishing a benchmark for existing nuclear capacity offtake that is significantly below the Microsoft/Constellation rate, reflecting the advantage of contracting earlier with a willing seller.

**Assessment:** AWS secured the most economically attractive nuclear deal in the market. The Talen PPA at ~$69/MWh vs. Microsoft's Constellation deal at ~$100+/MWh illustrates the compounding advantage of early commitment. AWS is also hedging technology risk by investing directly in SMR development for medium-term supply.

### 3.3 Google / Alphabet

**Strategy: Co-development + direct asset acquisition + utility partnerships**

Google has pursued the most vertically integrated approach, moving from procurement to direct ownership and co-development of power generation assets.

**Key deals:**
- **Intersect Power / TPG Rise Climate:** $20B co-location and co-development partnership (Dec 2024) to build data centers with dedicated clean power generation
- **Intersect Power acquisition** (Dec 2025): Google acquired Intersect Power outright, gaining direct ownership and control of generation assets
- **Kairos Power:** First-ever corporate SMR fleet deal for 500 MW of KP-FHR capacity (first reactor targeted 2030)
- **Xcel Energy:** PPAs for 1,400 MW wind, 200 MW solar, 300 MW long-duration energy storage for Minnesota data center (Feb 2026)
- **OG&E:** Landmark electric service agreement for three Oklahoma data centers. Google pays 100% of grid connection costs, all contracted costs regardless of energy use, and makes solar generation capacity available from two facilities under construction (Apr 2026)

**Assessment:** Google's acquisition of Intersect Power is the most significant strategic signal in this market. It represents a shift from buying power to owning the means of production. This gives Google supply security that no PPA can match, plus optionality to sell excess capacity. The OG&E deal establishes a template for utility partnerships that protects ratepayers while guaranteeing infrastructure investment.

### 3.4 Meta

**Strategy: Largest nuclear commitment + diversified portfolio approach**

Meta has assembled the largest nuclear power commitment of any corporation in history, using a sophisticated multi-pronged approach targeting near-term, medium-term, and long-term supply.

**Key deals:**
- **Vistra:** 20-year PPAs for 2.1+ GW from existing nuclear plants (Perry, Davis-Besse in Ohio; Beaver Valley in Pennsylvania), plus the largest nuclear uprates ever supported by a corporate customer
- **Constellation Energy / Clinton:** 20-year virtual PPA for full 1,121 MW output starting June 2027. Meta claims carbon-free credits while Constellation sells output to MISO grid
- **TerraPower:** Up to 8 Natrium reactor-and-storage plants providing 2.8 GW baseload + 1.2 GW molten salt storage. First 2 units (690 MW) by 2032; 6 more (2.1 GW) by 2035
- **Oklo:** Advanced nuclear campus in Pike County, Ohio, potentially 1.2 GW, online as early as 2030
- **XGS Energy:** 150 MW waterless geothermal partnership (Jun 2025)

**Total nuclear commitment:** Up to 6.6 GW by 2035 across existing fleet extensions, uprates, and new-build advanced reactors.

**Innovation in deal structure:** Meta introduced prepayment mechanisms to de-risk early-stage nuclear development (fuel procurement, site work, interconnection studies). This is hyperscaler-style capex thinking applied to power: cash pulled forward in exchange for clearer line of sight to future firm power.

**Assessment:** Meta's nuclear portfolio is the most comprehensive and technologically diversified. However, it carries the highest execution risk: the TerraPower Natrium fleet depends on unproven commercial-scale reactor technology and HALEU fuel supply chain development. The Vistra deals for existing nuclear provide essential near-term bridging while the advanced reactor pipeline matures.

---

## 4. Neocloud Strategies: Differentiated Approaches

Neoclouds lack the balance sheet scale of hyperscalers for multi-GW commitments, but they are compensating through distinct strategic advantages in power procurement.

### 4.1 CoreWeave

**Strategy: Portfolio procurement + BTM optionality + acquisition of power-rich assets**

- **Scale:** ~590 MW active power (late 2025), ~2.9 GW contracted across 41+ data centers, targeting 850 MW+ active by year-end 2025
- **Revenue backlog:** $55.6B, 96-98% locked into multi-year take-or-pay contracts
- **Core Scientific acquisition:** All-stock deal providing immediate control of large data center footprint and its power rights (~840 MW already leased to CoreWeave), route toward ~1.3 GW owned capacity plus optionality for another GW
- **Multi-region portfolio approach:** Treats power as a blendable commodity across ISOs. In ERCOT: hedges and tolling/shape products. In vertically integrated territories: large-load specials and economic development riders. In PJM: two-stage constructs connecting to grid-cheap energy post-upgrades
- **BTM bridge strategy:** Modular onsite generation (aeroderivative CTs, large recips, SOFC fuel cell skids) in 25-100 MW tranches to bridge deployment gaps while utility works progress
- **Galaxy Digital Helios campus:** Several hundred MW allocation in ERCOT, leveraging existing transmission and generation infrastructure

**Assessment:** CoreWeave's acquisition of Core Scientific is the clearest example of "if you dislike your landlord's power, buy out the landlord." The company is building a contracted-power specialist model, comfortable with structured power deals and large anchored campuses. The key risk is execution: construction delays from shell providers already caused revenue shifts from Q4 2025 into Q1 2026.

### 4.2 Crusoe

**Strategy: Vertically integrated energy-first model**

- **Origin:** Founded as energy-tech company converting stranded flare gas to compute power (Digital Flare Mitigation)
- **2024 performance:** Converted ~10.4 billion cubic feet of flare gas into ~1.3 TWh of electricity (~87% of total electricity needs), avoiding ~1.3 million tCO2e
- **Key facility:** Iceland data center (Apr 2025) powered entirely by renewable hydroelectric and geothermal
- **Crusoe Spark:** Turnkey prefabricated modular AI data center solution for edge deployment (Jun 2025), partnering with Redwood Materials for repurposed EV batteries
- **OpenAI/Stargate partnership:** Selected as primary partner for the Stargate project, directly due to energy-first model enabling faster deployment at lower cost
- **Financing:** $750M credit facility from Brookfield (Jun 2025) for AI factory development
- **Phase 1 plan (2025-2027):** Gas-fired BTM generation (aeroderivative CTs and recips) with firm gas supply from Energy Transfer for Panhandle operations

**Assessment:** Crusoe's moat is the deepest among neoclouds because it is built on vertical integration of energy. Co-locating with wasted, stranded, or curtailed energy sources provides below-market power rates and faster deployment. This is the leading indicator of a market-wide shift where access to cheap, abundant electricity is the most critical competitive advantage. The Stargate selection validates the commercial viability of this model at scale.

### 4.3 Lambda Labs

**Strategy: Power desk mindset + hyperscaler partnership**

- **Positioning:** "The Superintelligence Cloud" with on-demand and reserved access to latest NVIDIA clusters
- **Power approach:** Treats electricity as a traded commodity embedded in scheduling, placement, and contract design rather than marketing a BYOP campus or generation portfolio
- **Series D:** $480M funding round (Feb 2025) led by Andra Capital and SGW with NVIDIA participation
- **Microsoft partnership:** Multibillion-dollar partnership to deploy tens of thousands of NVIDIA GPUs including GB300 NVL72 systems
- **Growth trajectory:** Expanding toward gigawatt-scale AI factory language in customer communications

**Assessment:** Lambda is the least vertically integrated of the leading neoclouds on energy. Its advantage is commercial agility and developer ecosystem loyalty. However, in a power-constrained market, this may become a vulnerability. Lambda will likely need to develop BTM capabilities or partner with power-rich operators to maintain competitiveness as the market tightens.

### Neocloud Competitive Positioning

| Company | Power Moat | Current Scale | Strategic Edge | Key Risk |
|---|---|---|---|---|
| CoreWeave | Medium (portfolio + acquisition) | 2.9 GW contracted | NVIDIA relationship, contract backlog | Execution/construction delays, leverage |
| Crusoe | High (vertical energy integration) | ~1.3 TWh/yr | Below-market power, speed to deploy | Scaling beyond stranded energy sources |
| Lambda | Low (partnership-dependent) | Not disclosed | Developer loyalty, NVIDIA access | Power availability in constrained market |

---

## 5. The Nuclear Pivot

### The Paradigm Shift

As of April 2026, every major hyperscaler has signed at least one nuclear power deal for AI data center capacity. Thirteen announced projects commit over 9.8 GW of nuclear capacity, representing more capital commitment to nuclear energy than any prior decade in U.S. history.

### All Major Nuclear-for-AI Deals

| Buyer | Developer | Reactor Type | Capacity | Deal Value | Timeline | Status |
|---|---|---|---|---|---|---|
| Microsoft | Constellation Energy | Large reactor restart (TMI Unit 1) | 835 MW | $16B (20-yr PPA) | 2027 | Under Construction |
| AWS | Talen Energy | Existing nuclear (Susquehanna) | 1.92 GW | $18B (17-yr PPA) | Full by 2032 | Contracted |
| AWS | X-energy / Energy Northwest | Xe-100 HTGR (SMR) | 960 MW | $700M equity | 2030s | Planned |
| AWS | Dominion Energy | SMR (TBD) | TBD | TBD | TBD | Planned |
| Google | Kairos Power | KP-FHR (SMR) | 500 MW | Undisclosed | 2030 | Planned |
| Meta | TerraPower | Natrium SFR | 4.0 GW (2.8 GW base + 1.2 GW storage) | Undisclosed | 2032-2035 | Planned |
| Meta | Oklo | Aurora SFR | 1.2 GW | Undisclosed | ~2030 | Planned |
| Meta | Vistra | Existing nuclear fleet | 2.1+ GW | Undisclosed (20-yr PPAs) | Near-term | Contracted |
| Meta | Constellation | Existing nuclear (Clinton) | 1.12 GW | Undisclosed (20-yr PPA) | June 2027 | Contracted |

### Nuclear PPA Pricing Benchmarks

| Deal Structure | Price Range ($/MWh) | Commentary |
|---|---|---|
| Existing large reactor (physical PPA) | $45-70 | AWS/Talen at ~$69; most attractive tier |
| Existing reactor (premium, 24/7 matched) | $100-115 | Microsoft/Constellation (TMI restart); includes restart risk premium |
| Virtual nuclear PPA | $55-85 | Financial settlement, no physical delivery (Meta/Clinton model) |
| Forward-start SMR PPA | $65-100 | First-of-a-kind construction risk premium |
| Projected convergence (all types, 2030) | $55-75 | As SMR costs decline through learning curve effects |

### Why Nuclear?

The shift to nuclear is driven by three interlocking factors:

1. **Capacity factor:** Nuclear operates at ~94.7% capacity factor vs. solar ~25% and wind ~35%. For always-on AI workloads, this means 3-4x less capacity procurement per MW of actual continuous supply.
2. **Hourly carbon-free matching:** Both Google and Microsoft have committed to matching 100% of electricity consumption with carbon-free energy on an hourly basis by 2030. Intermittent renewables are structurally inadequate for this requirement.
3. **Baseload economics:** The total cost of "firm renewable" (solar/wind + battery storage for 24/7 operation) increasingly approaches or exceeds the cost of nuclear PPAs, particularly when curtailment, degradation, and replacement cycles are considered.

### Technology Bets

The hyperscalers are pursuing a two-track nuclear strategy:

- **Near-term (2026-2028):** PPAs with existing nuclear operators (Constellation, Talen, Vistra) for proven, operating reactors. Fastest path to 24/7 firm power.
- **Medium-term (2030-2035):** Investment in and offtake from next-generation SMRs (Kairos KP-FHR, X-energy Xe-100, TerraPower Natrium, Oklo Aurora). Higher execution risk but purpose-built for data center co-location.

The nuclear PPA market is projected to grow from $6.2B in 2025 to $47.8B by 2034 (24.5% CAGR).

---

## 6. Behind-the-Meter and On-Site Generation

### The Grid Bypass Revolution

The inability of public grids to support gigawatt-scale demands on AI timelines has catalyzed the most significant shift in data center power strategy: deploying immediate, on-site generation at scale. Developers are planning to build 56 GW of on-site power generation for data centers, representing ~30% of all planned data center capacity.

### Key Deployments

| Operator | Partner | Technology | Capacity | Location | Timeline |
|---|---|---|---|---|---|
| Oracle | VoltaGrid | Modular natural gas (QPac) | 2.3 GW | Texas (multiple sites) | 2025-2027 |
| Vantage Data Centers | VoltaGrid | BTM natural gas | 1+ GW | North American portfolio | 2025-2026 |
| Oracle | Bloom Energy | Solid-oxide fuel cells | Varies | Multiple U.S. sites | 90-day deployment |
| Chevron | (Data center partner TBD) | Natural gas | Up to 4 GW | West Texas | First power 2027 |

### The VoltaGrid/Oracle Model

The Oracle-VoltaGrid partnership is the flagship example of the BTM gas model:

- **QPac platform:** Modular natural gas reciprocating engine system, each node generating up to 20 MW, combinable to 200 MW under minor source air permits
- **Gas supply:** Firm natural gas from Energy Transfer's 140,000-mile pipeline and storage network
- **Key differentiator:** "AI-optimized performance" with ultra-fast transient response to handle variable AI workload power demands without battery storage
- **Timeline advantage:** 12-24 months vs. 40-60 months for grid interconnection

### Fuel Cell Deployment

Bloom Energy fuel cells offer the fastest path to on-site power:
- 2-6 month deployment timeline
- $5M-$8M per MW capex
- Minimal permitting requirements
- Best suited as bridge power or for sites requiring immediate capacity
- Oracle has deployed across several U.S. data centers since mid-2025

### Carbon and Sustainability Trade-offs

The pivot to on-site natural gas generation creates a fundamental tension with hyperscaler sustainability commitments. Companies are pursuing a contradictory dual approach: using fossil fuels for rapid scaling today while investing in clean technologies (nuclear, renewables, geothermal) for the long term. This introduces reputational and regulatory risk, particularly as the Trump administration's March 2026 White House pledge requires Big Tech to pay for new generation needed to power data centers without impacting utility bills.

---

## 7. Grid Interconnection: The Bottleneck

### Queue Congestion by Market

| Market / ISO | Average Queue Time | Key Constraints |
|---|---|---|
| PJM (Mid-Atlantic, Midwest) | 58 months | 10+ GW queued in N. Virginia alone. Dominion cannot serve new large loads without multi-year transmission buildout |
| CAISO (California) | 56 months | CEQA environmental review adds 12-24 months. Local opposition in drought-prone regions |
| ERCOT (Texas) | 36 months | Fastest market, but SB6 mandates demand flexibility or BYOP for loads >75 MW. 226 GW of large load requests as of Dec 2025, up from 63 GW at end of 2024 |
| MISO/SPP | 36-48 months | Growing secondary market. Less congested but infrastructure still limited |

### The Opportunity Cost of Delay

- Lost revenue: $10M-$12M per MW per year
- At 24 months of queue delay: $20M-$24M per MW in foregone cash flow
- National interconnection queue: 2,600+ GW of generation and storage capacity as of mid-2025, up from <500 GW in 2017

### Queue Reform Efforts

- **FERC Order 2023:** Moved from first-come/first-served to cluster study process. Reduced speculative projects by >50% but did not reduce wait times for legitimate projects
- **PJM Expedited Interconnection Track (EIT):** Filed Feb 2026, would process up to 10 interconnection requests/year with ~10-month target for executed agreements. Pending FERC approval by Jul 2026
- **SPP HILLGA:** High-Impact Large Load Generator Activity process for campus-style data center developments with geographic proximity constraints
- **FERC December 2025 order:** Directed PJM to create new rules for co-located load, responding to the trend of hyperscale operators co-locating with existing generation to bypass queues

### Phantom Queue Requests

A growing problem: developers file multiple or duplicative project requests across utility territories to explore options simultaneously. In ERCOT, this has skewed demand forecasts dramatically (226 GW of requests vs. ~85 GW total existing peak demand). Texas SB6 and FERC reforms attempt to address this, but the structural challenge persists.

---

## 8. Regulatory Landscape

### Federal Level

The regulatory framework for data center power is evolving rapidly, moving from outright rejection of novel power arrangements (Nov 2024) to active rulemaking (Apr 2026).

**Key developments:**
- **FERC co-location ruling (Dec 2025):** Directed PJM to create rules for behind-the-meter co-located load, responding to the AWS/Talen Susquehanna precedent
- **DOE Rulemaking Proposal:** Active rulemaking on BTM generation and large load interconnection, expected to stabilize no earlier than late 2026
- **"One Big Beautiful Bill Act" (OBBBA):** Preserved nuclear production tax credits while phasing out incentives for wind and solar, fundamentally altering the relative economics of power procurement
- **IRA safe harbor deadline (Jul 3, 2026):** Critical for renewable developers needing to demonstrate construction commencement for ITC/PTC tax credits. Potential wave of queue withdrawals in Q2-Q3 2026 could trigger restudies

**White House pledge (Mar 2026):** Amazon, Google, Microsoft, and Meta committed to paying for new energy generation needed to power U.S. data center infrastructure, responding to rising energy prices for the general public. This signals that ratepayer protection will remain a central regulatory concern.

### State Level

- **Texas SB6:** Requires loads >75 MW to provide demand flexibility or bring their own power. Establishes mandatory pre-approval for BTM operations requiring ERCOT studies and PUCT approval
- **Virginia:** Dominion Energy faces 10+ GW in queue; state regulators increasingly scrutinizing data center power impact on residential rates
- **Oklahoma:** OG&E/Google agreement establishes template where data center pays 100% of infrastructure costs with rate protection for existing customers

### Regulatory Risk Summary

The regulatory environment is more defined than six months ago but still taking shape. Key uncertainties for investors:

1. **Co-location rules in PJM:** Will FERC permit large-scale BTM arrangements that could divert generation from the wider grid?
2. **Cost allocation for backup capacity:** Who pays for standby grid capacity that BTM data centers rely on but rarely use?
3. **Rate impact:** As data centers consume a growing share of electricity, regulatory backlash from residential/commercial ratepayers could increase costs or restrict growth
4. **Nuclear licensing timelines:** NRC approval processes for SMRs remain untested at the scale contemplated by hyperscaler agreements

---

## 9. Cost Economics and PPA Pricing

### PPA Price Trends (North America)

| Power Source | Q4 2024 ($/MWh) | Q4 2025 ($/MWh) | YoY Change | 2030E ($/MWh) |
|---|---|---|---|---|
| Solar | ~$56.6 | ~$61.7 | +9% | $55-65 |
| Wind | ~$67.6 | ~$73.7 | +9% (up to +16% in ERCOT) | $65-80 |
| Existing Nuclear | ~$45-55 | ~$45-70 | Rising | $55-75 |
| Nuclear (restart/premium) | N/A | ~$100-115 | N/A | $80-100 |
| SMR (forward-start) | N/A | ~$65-100 | N/A | $55-75 |
| Natural Gas BTM (LCOE) | N/A | ~$60-90 | N/A | Market dependent |
| Fuel Cells (BTM) | N/A | ~$100-140 | N/A | $80-120 |

*Sources: LevelTen Energy, Pexapark, MarketIntelo, analyst estimates*

### Price Drivers

1. **Demand-supply tightening:** AI demand growth outpacing new generation additions
2. **Tax credit uncertainty:** OBBBA phases out wind/solar incentives, supporting nuclear but raising renewable PPA prices
3. **Hourly matching premium:** 24/7 carbon-free power commands 50-100% premium over annual matching
4. **Early mover advantage:** AWS's ~$69/MWh nuclear PPA (2025) vs. subsequent deals at $100+/MWh demonstrates that pricing is a function of timing, not just technology
5. **Interconnection cost pass-through:** Network upgrade costs increasingly allocated to data center operators, adding $2-5M/MW to total project cost

### The Structural Cost Advantage of Early Commitment

A hyperscaler that secured a 15- or 20-year PPA at 2024-2025 pricing locks in a cost structure that competitors face at tomorrow's prices. Over the lifetime of an AI data center, that difference compounds into a structural cost advantage reflected directly in cloud service pricing.

The inverse is also true: operators entering the market in 2026-2027 face higher prices, longer lead times, and fewer available sites. This bifurcation will widen before it narrows, as the pipeline of new generation coming online through 2029 is insufficient to meet aggregate demand from AI infrastructure, renewable energy transition, EV charging, and industrial electrification.

---

## 10. Competitive Moats and Investment Implications

### The Power Hierarchy

The power procurement strategies executed over the past three years have created a clear competitive hierarchy:

**Tier 1 — Integrated Power-Compute Operators:**
- Microsoft, AWS, Google, Meta
- Multi-GW committed capacity across multiple generation technologies and geographies
- Nuclear PPAs providing 24/7 firm power at scale
- Direct investment in or ownership of generation assets
- Balance sheet capacity to underwrite $10B+ power infrastructure projects

**Tier 2 — Differentiated Neoclouds:**
- CoreWeave (portfolio approach + power-rich acquisitions)
- Crusoe (vertical energy integration)
- Power access is a genuine competitive moat, enabling below-market rates or faster deployment

**Tier 3 — Partnership-Dependent Operators:**
- Lambda, Nebius, smaller neoclouds
- Rely on leased capacity from third-party shell providers or hyperscaler partnerships
- Power availability is a vulnerability, not an advantage

**Tier 4 — Traditional Data Center REITs/Operators:**
- Equinix, Digital Realty, CyrusOne, etc.
- Primarily retail colocation; less exposed to AI power arms race
- Growing opportunity as hyperscalers seek third-party capacity for non-training workloads

### Investment Implications

#### Power Suppliers & Developers (Beneficiaries)

| Company | Relevance | Key Catalyst |
|---|---|---|
| **Constellation Energy** (CEG) | Largest U.S. nuclear fleet; PPAs with Microsoft and Meta. Revenue from hyperscaler deals est. $1.5B/yr by 2028 at ~$100+/MWh | TMI restart timeline (2027); additional fleet repricing at premium rates |
| **Talen Energy** (TLN) | $18B AWS PPA; Susquehanna plant as anchor for AI campus | Full delivery ramp by 2032; potential SMR site development |
| **Vistra** (VST) | 2.1+ GW nuclear PPAs with Meta (Perry, Davis-Besse, Beaver Valley) | Nuclear uprate execution; long-term PPA repricing |
| **Brookfield Renewable** (BEP) | 10.5 GW framework with Microsoft; $750M Crusoe credit facility | Renewable project pipeline execution; hyperscaler relationship deepening |
| **VoltaGrid** (private) | 2.3 GW Oracle + 1 GW Vantage BTM deployments | IPO potential; BTM market growth |
| **Bloom Energy** (BE) | Fuel cell deployments at Oracle and other DC operators | Volume ramp; 2-6 month deployment as "bridge to grid" value proposition |

#### GPU/AI Infrastructure (Dependent on Power)

| Company | Power Position | Risk/Opportunity |
|---|---|---|
| **CoreWeave** (CRWV) | 2.9 GW contracted; Core Scientific acquisition | Execution risk on energization timelines; high leverage |
| **Crusoe** (private) | Vertically integrated energy | Stargate execution; scaling beyond stranded energy |
| **Oracle** (ORCL) | 3+ GW BTM generation (VoltaGrid + Bloom + others) | Stargate execution; ERCOT regulatory compliance |

#### Nuclear Technology (Long-Duration Bets)

| Company | Technology | Hyperscaler Backers | Key Risk |
|---|---|---|---|
| **Oklo** (OKLO) | Aurora SFR | Meta (1.2 GW) | NRC licensing; first-of-a-kind construction |
| **Kairos Power** (private) | KP-FHR | Google (500 MW) | Scale-up from demo to commercial |
| **X-energy** (private) | Xe-100 HTGR | Amazon ($700M) | NRC licensing; HALEU fuel supply |
| **TerraPower** (private) | Natrium SFR | Meta (4 GW) | Longest timeline (2032-2035); HALEU dependency |

### Key Lessons for Mid-Tier Operators

1. **Early commitment at sustainable scale delivers the same structural logic** as hyperscaler GW-scale commitment. A 100 MW long-term PPA today is locking in advantages over competitors who secure equivalent capacity later.
2. **Power desk capabilities are now as critical as GPU procurement.** Traders, schedulers, and risk managers who can price time and arbitrage BTM vs. ISO markets are essential hires.
3. **BTM options are available at smaller scales than most assume.** Fuel cells, small-scale gas generation, and battery storage can be deployed at 10-50 MW scales. The operational complexity per MW is higher, but queue-bypass advantages apply regardless of scale.

---

## 11. Key Risks

### Demand Uncertainty
- EPRI's 2030 scenarios range from 380-790 TWh (2x spread). If AI adoption slows, capex shifts, or efficiency breakthroughs reduce power intensity, the massive power commitments could become stranded liabilities
- Phantom queue requests inflate pipeline estimates; actual realization rates may be significantly lower

### Execution and Construction Risk
- CoreWeave's Q4 2025 revenue push-out due to shell provider delays illustrates that power infrastructure construction is the critical path
- SMR technology is unproven at commercial scale. None of the announced hyperscaler SMR deals will deliver power before 2030, and many target 2032-2035
- HALEU fuel supply chain for advanced reactors (TerraPower Natrium, Oklo Aurora) is nascent; DOE committed $2.7B (Jan 2026) but domestic enrichment capacity is years away

### Regulatory Risk
- FERC/DOE rulemaking on BTM generation and co-located load still evolving through late 2026
- Cost allocation for grid backup capacity used by BTM data centers is unresolved
- State-level pushback (Virginia, Texas) as residential/commercial ratepayers face potential rate increases
- Nuclear relicensing and restart timelines (TMI targeted 2027-2028) carry NRC approval uncertainty

### Carbon and Sustainability Risk
- Aggressive BTM natural gas deployment contradicts corporate carbon-neutral pledges
- Rising gas consumption by data centers: IEA projects U.S. gas-fired electricity for data centers rising from ~100 TWh (2025) to well above 200 TWh (2030)
- Reputational risk from "greenwashing" if nuclear/renewable investments are insufficient to offset near-term gas expansion

### Competitive and Market Risk
- Power pricing is rising: 9-16% YoY increases in renewable PPAs; nuclear premiums at 50-100% above merchant rates
- Limited pool of uncontracted existing nuclear capacity creates a competitive race; subsequent deals priced higher than early movers
- Private equity capital (KKR, Brookfield, TPG) flooding into power-for-AI, potentially inflating asset prices

---

## 12. Conclusion

The data center power procurement landscape has undergone a structural transformation. What was a routine operating expense (buying electricity from the grid) is now a strategic capability that separates winners from losers in the AI infrastructure race. The key conclusions:

1. **Power is the new moat.** The hyperscalers that secured GW-scale commitments in 2024-2025 at favorable pricing have created structural cost advantages that will compound through the decade. Late entrants face higher prices, longer queues, and fewer sites.

2. **The procurement playbook has diversified.** No single strategy dominates. The market leaders are pursuing a portfolio approach: renewable PPAs for volume and carbon goals, nuclear PPAs for 24/7 baseload, BTM gas for speed, and direct ownership for control. The optimal mix depends on timeline pressure, capital availability, and sustainability commitments.

3. **Hyperscalers are becoming energy companies.** Google's acquisition of Intersect Power, Meta's prepayment mechanisms for nuclear development, and Microsoft's supply-chain-level renewable deals signal that Big Tech views power generation as a core competency, not an outsourced input.

4. **Neoclouds compete on different terms.** Crusoe's vertical energy integration and CoreWeave's portfolio approach demonstrate that power moats can be built without hyperscaler balance sheets. Lambda and others that remain power-dependent face a structural disadvantage as markets tighten.

5. **Nuclear is the next frontier.** The $47.8B nuclear PPA market projected by 2034 represents the most significant corporate investment in nuclear energy in U.S. history. For investors, the key question is not whether nuclear-for-AI will happen, but which operators (Constellation, Vistra, Talen) and which technology developers (Kairos, X-energy, TerraPower, Oklo) will capture the economics.

6. **Regulatory uncertainty remains the wild card.** FERC co-location rules, DOE BTM rulemaking, state-level rate protection, and NRC licensing timelines will determine how quickly the market can evolve. The companies with the deepest regulatory expertise and most diversified geographic footprints will navigate this best.

**For equity investors:** The power procurement strategies described in this report have created a durable, multi-year investment theme across energy suppliers, nuclear operators, BTM generation providers, and the AI infrastructure operators themselves. The key to evaluating any data center or cloud investment is now answering: *How much power is secured, at what price, for how long, and how firm is the delivery timeline?*

---

*Disclaimer: This report is for informational and educational purposes only. It does not constitute investment advice. All data sourced from publicly available filings, press releases, industry reports, and analyst estimates as of May 2026.*
