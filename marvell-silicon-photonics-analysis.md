# Marvell Technology: Top Customers, Accelerator Stack Position, and Silicon Photonics Exposure

## Executive Summary

Marvell Technology ($MRVL) has emerged as a critical infrastructure enabler for AI data centers, generating $8.2 billion in FY2026 revenue (+42% YoY) with 74% of revenue from data center customers. The company occupies a unique position in the accelerator stack—not as a GPU or accelerator maker, but as the connective tissue that links custom compute, memory, networking, and optics across the AI infrastructure. Its silicon photonics exposure is expanding rapidly through organic products (1.6T optical DSPs, light engines) and the $3.25 billion acquisition of Celestial AI, positioning it at the center of the copper-to-optical transition reshaping both training and inference architectures.

---

## 1. Top Customers and Revenue Concentration

### Revenue Breakdown by Segment (FY2026)

| Segment | Revenue | % of Total | YoY Growth |
|---------|---------|-----------|------------|
| Data Center | $6.1B | 74.4% | ~47% |
| Communications & Other | $2.1B | 25.6% | ~26% |
| **Total** | **$8.2B** | **100%** | **42%** |

### Data Center Sub-Segment Revenue

| Sub-Segment | FY2026 Revenue | FY2027E | Commentary |
|-------------|---------------|---------|------------|
| Custom Silicon (XPUs) | ~$1.5B | ~$1.8B (+20%) | Amazon Trainium/Inferentia, Microsoft Maia |
| Interconnect (Electro-Optics) | ~$1.5B+ | Growing 50%+ | AEC, retimers, optical DSPs; doubling in FY27 |
| Switching (Teralynx) | ~$300M+ | ~$600M+ | 51.2T Ethernet switch ramp |
| Other Data Center | ~$2.8B | Growing | Storage controllers, DPUs, networking ICs |

### Key Customer Relationships

**Amazon Web Services** — Marvell's largest and most strategic customer. The two companies have a 15-year relationship, formalized in a five-year strategic collaboration agreement (December 2024) covering next-generation custom AI ASICs (Trainium and Inferentia families). Amazon's custom chip business has exceeded a $10 billion run rate with triple-digit growth. Marvell provides high-value IP blocks (SerDes, interconnect, packaging) at 50%+ gross margins, distinguishing it from lower-margin ASIC service houses like Alchip.

**Microsoft** — Customer for the Maia custom AI accelerator program. Microsoft's Maia 200 (3nm, launched January 2026) is purpose-built for inference. However, reports indicate Microsoft may be evaluating Broadcom as an alternative supplier, introducing a key risk to Marvell's second-largest custom silicon relationship.

**Google** — Marvell supplies custom silicon components and networking products used in Google's data center infrastructure. Google's TPU program (now on v7 Ironwood) uses Broadcom as its primary ASIC partner, but Marvell captures networking and interconnect content.

**Meta** — Primarily a Broadcom custom silicon customer (MTIA program), but Marvell competes for networking and optical interconnect content across Meta's massive AI buildout.

**Other Hyperscalers and Cloud Providers** — Marvell has secured design wins with 20+ customers, reflecting a diversification strategy beyond the top 3-4 hyperscalers. This includes enterprise cloud, sovereign cloud, and GPU cloud (neocloud) customers.

### Customer Concentration Risk

Marvell faces meaningful concentration risk. While specific 10%+ customer disclosures from the FY2026 10-K are not publicly broken out by name in press materials, Amazon is widely understood to represent a disproportionate share of custom silicon revenue. The Microsoft/Broadcom risk adds further uncertainty. Broadcom commands roughly 80% of the custom AI ASIC market (via Google, Meta, and others) versus Marvell's approximately 20% share—making customer retention existential for Marvell's custom silicon growth narrative.

---

## 2. Where Marvell Fits in the Accelerator Stack

Marvell does not compete at the compute layer (GPUs/TPUs/XPUs) but supplies nearly every other layer of the AI data center stack. Its position is best understood as **infrastructure silicon for AI**, spanning four connectivity tiers:

### Marvell's Position Across the AI Stack

```
┌─────────────────────────────────────────────────────────────────┐
│                    AI DATA CENTER STACK                         │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  COMPUTE LAYER (GPUs / Custom XPUs / TPUs)                     │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  NVIDIA GPUs │ Google TPUs │ AWS Trainium │ MSFT Maia   │   │
│  │                                                         │   │
│  │  Marvell Role: CUSTOM ASIC DESIGN PARTNER               │   │
│  │  → IP blocks: SerDes, interconnect, Arm subsystems      │   │
│  │  → Advanced packaging (UHDI, multi-chip)                │   │
│  │  → 3nm/5nm process on TSMC                              │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  INTRA-SERVER CONNECTIVITY (Scale-Up)                          │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  PCIe / CXL / NVLink / Infinity Fabric                  │   │
│  │                                                         │   │
│  │  Marvell Products:                                      │   │
│  │  → Alaska P PCIe Gen5/Gen6 Retimers                     │   │
│  │  → Structera S CXL 3.0 Switches (260-lane)             │   │
│  │  → Structera A Near-Memory Accelerators                 │   │
│  │  → Structera X Memory Expansion Controllers             │   │
│  │  → 112G/224G XSR SerDes IP                              │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  INTER-SERVER / RACK NETWORKING (Scale-Out)                    │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  Ethernet / InfiniBand Fabrics                          │   │
│  │                                                         │   │
│  │  Marvell Products:                                      │   │
│  │  → Teralynx 10: 51.2T Ethernet Switch (volume prod.)   │   │
│  │  → Next-gen 102.4T switch in development                │   │
│  │  → Active Electrical Cables (AEC)                       │   │
│  │  → PAM4 linear-drive solutions                          │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  OPTICAL INTERCONNECT (Scale-Out + DCI)                        │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  800G → 1.6T → 3.2T Optical Connectivity               │   │
│  │                                                         │   │
│  │  Marvell Products:                                      │   │
│  │  → Ara 1.6T Optical DSP Platform (3nm, mass prod.)     │   │
│  │  → 1.6T Silicon Photonics Light Engine                  │   │
│  │  → Colorz 800-ZR/ZR+ DCI Modules                       │   │
│  │  → Aquila Coherent-Lite DSP (O-band)                    │   │
│  │  → Celestial AI Photonic Fabric (16 Tbps/chiplet)       │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  STORAGE & DATA PROCESSING                                     │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  → SSD/HDD Controllers                                  │   │
│  │  → Data Processing Units (DPUs)                         │   │
│  │  → Security/Encryption ASICs                            │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### Competitive Landscape by Layer

| Stack Layer | Marvell Products | Primary Competitors | Marvell's Position |
|-------------|-----------------|---------------------|-------------------|
| Custom ASIC Design | SerDes IP, packaging, Arm subsystems | Broadcom (~80% share), Alchip | ~20% share; IP-rich, high-margin model |
| Intra-Server (PCIe/CXL) | Alaska P retimers, Structera CXL switches | Broadcom, Astera Labs, Microchip | Growing; CXL memory pooling is early |
| Ethernet Switching | Teralynx 10 (51.2T) | Broadcom Tomahawk (102.4T), Cisco | Challenger; ~1 generation behind Broadcom |
| Optical DSP | Ara 1.6T (3nm), Aquila, Petra | Broadcom, Credo, Coherent | Leader in 1.6T; first 200G/lane PAM4 DSP |
| Silicon Photonics | Light engines, Photonic Fabric (Celestial AI) | Intel, Broadcom, Cisco, Lumentum | Differentiated via Celestial AI acquisition |
| DCI Optics | Colorz 800-ZR/ZR+ | Ciena, Infinera, Cisco | Niche; growing with hyperscaler DCI |
| Storage | SSD/HDD controllers | Broadcom, Microchip | Mature; modest growth |

### Key Takeaway

Marvell's differentiation is its **breadth across the stack**—it can supply custom accelerator IP, the retimers and CXL switches connecting them, the Ethernet fabric switching packets between racks, the optical DSPs driving the transceivers, and increasingly the silicon photonic engines themselves. No other company outside of Broadcom covers this many layers, and Marvell's optical portfolio (post-Celestial AI) is arguably more forward-looking.

---

## 3. Marvell's Silicon Photonics Exposure

### Current Products and Revenue

Marvell's silicon photonics exposure spans three product categories:

#### A. Optical DSPs (Currently in Mass Production)

These are the digital signal processing chips inside optical transceivers—Marvell's largest and most mature optical revenue stream.

| Product | Speed | Process | Status | Use Case |
|---------|-------|---------|--------|----------|
| Nova | 1.6T (8x200G) | 5nm | Shipping since 2023 | Pluggable 1.6T optics |
| Ara | 1.6T (8x200G) | 3nm | Mass production 2025+ | Hyperscaler 1.6T deployments |
| Ara T | 1.6T | 3nm | Announced Mar 2026 | Transmit-retimed optics |
| Ara X | 1.6T | 3nm | Announced Mar 2026 | Advanced link reliability |
| Petra | 8x100G → 4x200G | 3nm | Announced Mar 2026 | Gearbox for mixed-speed networks |
| Aquila M | Coherent-lite | 3nm | Announced Mar 2026 | O-band DCI with MACsec |

Optical DSPs and electro-optic interconnect products are part of the Interconnect sub-segment, which is expected to grow 50%+ in FY2027 and has been cited as doubling within the year. This implies a trajectory from roughly $1.5 billion toward $3 billion+ in annual revenue.

#### B. Silicon Photonics Light Engines (Ramping)

Marvell's 1.6T Silicon Photonics Light Engine—honored in the 2026 Lightwave Innovation Reviews—delivers 8 lanes of 200G PAM4 optical connectivity in a single package. It consolidates hundreds of discrete components into a ready-to-use module compatible with both linear-drive pluggable optics (LPO) and on-board optical integration.

This is a direct play into the silicon photonics manufacturing chain, moving Marvell beyond DSPs into the photonic integrated circuit (PIC) layer.

#### C. Celestial AI Photonic Fabric (FY2028+ Revenue)

The $3.25 billion acquisition of Celestial AI (completed February 2026) is Marvell's most consequential silicon photonics bet. Key parameters:

| Metric | Detail |
|--------|--------|
| Technology | Photonic Fabric: optical interconnect embedded directly in processor packages |
| Bandwidth | Up to 16 Tbps per chiplet (~10x current optical ports) |
| Architecture | Replaces copper for chip-to-chip, chip-to-memory, and rack-to-rack links |
| Revenue Ramp | Initial revenue H2 FY2028; $500M ARR by Q4 FY2028; $1B ARR by Q4 FY2029 |
| OpEx Impact | ~$50M annual operating expense addition |
| Deal Structure | $1B cash + ~27.2M Marvell shares (~$2.25B) |

The Photonic Fabric platform enables multi-rack AI supercomputers where hundreds of processors share direct memory access across optical links—a fundamental architectural shift from single-rack GPU servers connected by copper.

### Quantifying Total Silicon Photonics Exposure

| Category | FY2026E Revenue | FY2028E Revenue | Notes |
|----------|----------------|----------------|-------|
| Optical DSPs & Electro-Optics | ~$1.5B | ~$3B+ | Part of interconnect sub-segment |
| Silicon Photonics Light Engines | Early ramp | Scaling | Integrated into transceiver modules |
| Celestial AI Photonic Fabric | $0 | ~$500M ARR | Revenue begins H2 FY2028 |
| **Total Optical/SiPh Exposure** | **~$1.5B+** | **~$3.5B+** | **~18% of FY2026 → ~23%+ of FY2028 revenue** |

As a percentage of revenue, Marvell's silicon photonics and optical interconnect exposure is set to grow from roughly 18% of total revenue in FY2026 to over 23% by FY2028—and potentially 25-30%+ by FY2029 as Celestial AI ramps to $1 billion.

---

## 4. Silicon Photonics: Scale-Up vs. Scale-Out Networking

### Definitions

- **Scale-up networking**: GPU-to-GPU, chip-to-chip, and chip-to-memory interconnects *within* a node or supernode (e.g., NVLink, PCIe/CXL, proprietary GPU fabrics). These are short-reach, ultra-high-bandwidth links—historically copper.
- **Scale-out networking**: Rack-to-rack, cluster-to-cluster, and data-center-to-data-center connectivity via Ethernet or InfiniBand switches and optical transceivers. These are medium-to-long-reach links—historically optical.

### Where Is Silicon Photonics Exposure Higher Today?

**Scale-out dominates current silicon photonics revenue.** The vast majority of optical transceivers, DSPs, and silicon photonics light engines sold today serve scale-out Ethernet and InfiniBand backend fabrics connecting racks of GPUs. This is Marvell's primary optical revenue stream (Ara 1.6T DSPs, light engines, Colorz DCI modules). The scale-out optical transceiver market for AI data centers reached ~$16.5 billion in 2025 and is projected to hit ~$26 billion in 2026 (60% growth), driven by the 800G-to-1.6T speed transition.

Scale-up has been almost entirely copper-based through 2025. NVIDIA's NVLink, AMD's Infinity Fabric, and custom hyperscaler XPU interconnects have relied on copper traces, passive copper cables, and active electrical cables (AEC) for GPU-to-GPU links within racks.

### Where Is the Bigger Opportunity Emerging?

**Scale-up is the larger long-term TAM for silicon photonics and is growing faster, but from a near-zero base.**

The reasons are structural:

| Dimension | Scale-Up | Scale-Out |
|-----------|----------|-----------|
| **Bandwidth per link** | Up to 10x higher than scale-out (multi-Tbps per GPU) | 800G–1.6T per port today |
| **Current optical penetration** | Near zero (copper dominates) | Near 100% (already optical beyond 2m) |
| **Net new TAM creation** | Massive—entire copper domain converting to optical | Incremental—speed upgrades (800G→1.6T→3.2T) on existing optical links |
| **Links per GPU** | 8-18 NVLink/scale-up connections per GPU | 1-2 scale-out network connections per GPU |
| **Total bandwidth demand** | Growing faster (training clusters 100K→1M GPUs, each needing multi-Tbps scale-up) | Growing (more racks, faster ports) |
| **Technology required** | Co-packaged optics (CPO), photonic fabric, on-package optics | Pluggable transceivers (LPO), silicon photonics light engines |
| **Revenue timeline** | 2027-2028 volume ramp | Now and growing |
| **Key suppliers** | NVIDIA + Lumentum/Coherent/Ayar Labs (CPO), Marvell/Celestial AI (Photonic Fabric) | Marvell (DSPs), Broadcom, Credo, Coherent, Lumentum |

### The Scale-Up Inflection

Several catalysts are accelerating the scale-up optical transition:

1. **NVIDIA's $4B optical investment (March 2026)**: $2B each to Lumentum and Coherent, plus Ayar Labs Series E participation, all targeting co-packaged optics for scale-up GPU interconnects. This is a clear signal that NVLink's copper era has a finite horizon.

2. **NVIDIA Spectrum-X Photonics**: First fully integrated 512-lane 200G CPO switch, delivering 3.5x energy savings and 10x network resilience versus pluggable optics. Available 2026, targeting scale-up AI factory networks.

3. **Copper physical limits**: At 200G/lane speeds (required for 1.6T and beyond), passive copper cannot reliably span beyond a single rack. Every next-generation AI cluster being designed today must plan for optical scale-up.

4. **CPO power economics**: Pluggable optics consume ~15 picojoules/bit; CPO reduces this to ~5 pJ/bit with a path to <1 pJ/bit. At megawatt-scale clusters, this difference is existential for power budgets.

### Marvell's Exposure: Scale-Up vs. Scale-Out

| Product | Scale-Up or Scale-Out | Status | Revenue Contribution |
|---------|----------------------|--------|---------------------|
| Ara 1.6T Optical DSP | **Scale-out** (pluggable transceivers for Ethernet/IB fabrics) | Mass production | Large — core of current interconnect revenue |
| 1.6T SiPh Light Engine | **Both** (pluggable and on-board optics) | Ramping | Growing |
| AEC (Active Electrical Cables) | **Scale-up** (short-reach copper-based, pre-optical) | Shipping | Moderate — bridge product before CPO |
| Alaska P PCIe Retimers | **Scale-up** (copper-based, intra-server) | Shipping | Moderate |
| Structera CXL Switches | **Scale-up** (memory pooling, electrical today) | Early | Small |
| Celestial AI Photonic Fabric | **Scale-up** (chip-to-chip, chip-to-memory optical) | Pre-revenue | $0 now; $500M ARR by FY2028, $1B by FY2029 |
| Colorz DCI | **Scale-out** (inter-data-center) | Shipping | Small |
| Teralynx Ethernet Switch | **Scale-out** (spine/leaf switching) | Volume production | ~$300M FY2026 |

**Today**: Marvell's silicon photonics exposure is ~90% scale-out (optical DSPs and light engines for pluggable transceivers in Ethernet/IB fabrics).

**By FY2029**: The mix shifts significantly toward scale-up, driven by Celestial AI Photonic Fabric ($1B ARR target) and on-board optical integration. Scale-up could represent 25-35% of total optical/SiPh revenue, up from near-zero today.

### Market Size Comparison

```
SCALE-OUT OPTICAL MARKET (PLUGGABLE TRANSCEIVERS + DSPs)
├── 2025: ~$16.5B (mature, already optical)
├── 2026: ~$26B (800G→1.6T upgrade cycle)
├── 2028: ~$35B+ (1.6T→3.2T, more racks)
└── Growth: ~20-25% CAGR

SCALE-UP OPTICAL MARKET (CPO + PHOTONIC FABRIC)
├── 2025: <$0.5B (near-zero, copper dominates)
├── 2026: ~$1-2B (early CPO deployments)
├── 2028: ~$5-8B (CPO volume ramp, photonic fabric)
├── 2030: ~$10-15B+ (standard for all new clusters)
└── Growth: ~80-100%+ CAGR from near-zero base

COMBINED AI OPTICAL INTERCONNECT TAM
├── 2025: ~$18B
├── 2028: ~$40-43B
└── 2030: ~$50B+
```

### Answer: Which Has Higher Exposure?

**Scale-out is bigger today and remains the larger absolute market through at least 2028.** Marvell's current silicon photonics revenue is almost entirely scale-out (optical DSPs for pluggable transceivers). This market is large ($26B in 2026), growing steadily (20-25% CAGR), and Marvell has a leadership position with its Ara DSP platform.

**Scale-up is the more transformative opportunity and will grow faster.** The copper-to-optical transition in scale-up creates entirely new TAM (not just speed upgrades). NVIDIA's $4B optical investment, the Spectrum-X Photonics platform, and Marvell's own Celestial AI acquisition all target this transition. Scale-up bandwidth demand per link is up to 10x higher than scale-out, and every GPU in a cluster needs 8-18 scale-up connections versus 1-2 scale-out connections.

**Marvell's strategic bet (Celestial AI) is explicitly a scale-up play.** The Photonic Fabric embeds optics directly into processor packages for chip-to-chip and chip-to-memory links—this is scale-up by definition. The $3.25B acquisition price and $1B ARR target by FY2029 signal Marvell's view that scale-up optical is the higher-growth vector, even though scale-out is the larger revenue base today.

---

## 5. The Role of Silicon Photonics in Training vs. Inference

### Why Silicon Photonics Matters for AI Workloads

The fundamental bottleneck in AI infrastructure has shifted from compute to interconnect. Whether training a frontier model across 100,000+ GPUs or serving inference at scale, the limiting factor is how fast data moves between chips, nodes, memory, and racks—not raw FLOPS.

Copper-based electrical interconnects face three compounding constraints at AI scale:
1. **Bandwidth ceiling**: Copper degrades at distances beyond ~2 meters at high data rates
2. **Power consumption**: Electrical signaling consumes ~2x the energy of optical at equivalent bandwidth
3. **Distance limitations**: Training clusters span hundreds of racks; copper cannot reach

Silicon photonics solves all three by integrating lasers, modulators, and photodetectors onto silicon chips using standard CMOS fabrication processes, enabling optical links with higher bandwidth, lower power, and distance-independent signal integrity.

### Training Workloads

Training frontier AI models is the primary driver of silicon photonics adoption today. Key characteristics:

| Requirement | Detail | Silicon Photonics Benefit |
|-------------|--------|--------------------------|
| **Scale** | 100,000+ GPUs per training cluster | Optical enables multi-rack/multi-building clusters impossible with copper |
| **Bandwidth** | Hundreds of petabits/sec aggregate | 1.6T–3.2T links per port; photonic fabric enables 16 Tbps/chiplet |
| **Latency** | Synchronous all-reduce operations | Co-packaged optics reduce latency by eliminating electrical conversion |
| **Traffic Pattern** | Predictable collective communication | Optical circuit switching optimized for fixed all-to-all patterns |
| **Power** | MW-scale clusters constrained by cooling | CPO reduces power 30-40% vs. pluggable optics |
| **Reliability** | Jobs run weeks/months | Optical links show 10x higher resilience, 5x longer uptime |

Training clusters use silicon photonics primarily for:
- **Scale-out fabric** (rack-to-rack Ethernet/InfiniBand via 800G/1.6T optical transceivers)
- **Scale-up interconnect** (emerging: chip-to-chip optical links replacing NVLink copper)
- **Backend network switching** (optical DSPs inside every transceiver in the fabric)

SiP-ML research demonstrates 1.3–9.1x training time improvements when silicon photonic networks replace electrical alternatives, by exploiting the predictability of ML collective communication patterns.

### Inference Workloads

Inference has evolved from a simple, single-GPU task into a distributed, cluster-scale problem—and this shift is rapidly increasing silicon photonics demand for inference:

| Trend | Impact on Optical Interconnect |
|-------|-------------------------------|
| **Reasoning models (o1, o3, R1)** | Test-time compute scaling → multi-GPU inference → cluster-scale bandwidth needs |
| **Agentic workloads** | Long-running, multi-step inference chains → sustained bandwidth demand |
| **Disaggregated inference** | Prefill (compute-bound) and decode (bandwidth-bound) split across different hardware pools connected by high-bandwidth fabric |
| **MoE architectures** | Expert routing across GPUs requires low-latency, high-bandwidth interconnect |
| **KV-cache sharing** | Shared KV-cache across inference instances requires memory-fabric bandwidth |
| **Scale of serving** | Inference may consume 10-50x the compute of training over a model's lifetime |

#### Disaggregated Inference Architecture (Emerging, 2026+)

The most significant inference trend for silicon photonics is disaggregated inference, where prefill and decode phases are physically separated:

```
┌──────────────────────────────────────────────────────────────┐
│              DISAGGREGATED INFERENCE CLUSTER                  │
│                                                              │
│  ┌──────────────┐    OPTICAL     ┌──────────────────────┐   │
│  │  PREFILL POOL │◄─────────────►│  DECODE POOL          │   │
│  │  (Compute-    │   FABRIC      │  (Bandwidth-          │   │
│  │   bound)      │  (SiPh)       │   bound)              │   │
│  │              │               │                        │   │
│  │  High-FLOPS  │               │  High-bandwidth        │   │
│  │  GPUs        │               │  memory access          │   │
│  │              │               │  (HBM or pooled mem)    │   │
│  └──────────────┘               └──────────────────────┘   │
│         │                                │                   │
│         └────────────┐  ┌────────────────┘                   │
│                      ▼  ▼                                    │
│              ┌───────────────┐                               │
│              │ SHARED MEMORY │                               │
│              │ POOL (Optical │                               │
│              │ Fabric, CXL)  │                               │
│              └───────────────┘                               │
└──────────────────────────────────────────────────────────────┘
```

Silicon photonic memory pooling (e.g., Celestial AI's Photonic Fabric Appliance) delivers 32 TB shared memory at HBM3E bandwidth with 115 Tbps all-to-all switching, achieving up to 3.66x throughput and 1.40x latency improvements for 405B parameter LLM inference versus traditional NVLink configurations.

### Training vs. Inference: Silicon Photonics Comparison

| Dimension | Training | Inference | Winner for SiPh Growth |
|-----------|----------|-----------|----------------------|
| **Current adoption** | High—already deployed at scale in training clusters | Emerging—growing rapidly with reasoning models | Training (today) |
| **Growth rate** | Scaling with cluster size (100K → 1M+ GPUs) | Exploding with disaggregation, reasoning, agents | Inference (faster growth) |
| **Traffic pattern** | Predictable, collective (all-reduce) | Variable, bursty, request-driven | Training (easier to optimize) |
| **Bandwidth per GPU** | Multi-Tbps required | Growing toward training-class bandwidth | Converging |
| **Latency sensitivity** | Moderate (batch jobs) | High (user-facing, real-time) | Inference (higher bar) |
| **Distance requirements** | Multi-rack to multi-building | Rack-scale to multi-rack | Training (longer reach needed) |
| **Memory fabric need** | Moderate (HBM per GPU sufficient) | Critical (KV-cache sharing, disaggregated memory) | Inference (new TAM for SiPh) |
| **Cluster size trend** | Plateauing at 100K-500K GPUs | Growing from single-GPU to 1,000s | Inference (more new deployments) |
| **Revenue timeline** | Now | Now + accelerating FY2027-2029 | Both |

### Key Insight

Training built the initial market for silicon photonics in AI data centers, but inference is becoming the larger long-term opportunity. The shift to reasoning models, agentic workflows, and disaggregated architectures transforms inference from a single-GPU problem into a distributed-systems problem with bandwidth requirements approaching those of training. Marvell's Celestial AI Photonic Fabric—with its memory-pooling and chip-to-chip optical capabilities—is specifically designed for this disaggregated inference architecture, not just training scale-out.

---

## 6. Financial Outlook and Valuation Context

| Metric | FY2026A | FY2027E | FY2028E |
|--------|---------|---------|---------|
| Revenue | $8.2B | ~$11B | ~$15B |
| Data Center Revenue | $6.1B | ~$8.5B+ | ~$12B+ |
| Custom Silicon | ~$1.5B | ~$1.8B | ~$2.5B+ |
| Interconnect / Optics | ~$1.5B | ~$3B+ | ~$3.5B+ |
| Switching | ~$300M | ~$600M | ~$1B+ |
| Celestial AI Revenue | $0 | $0 | ~$500M ARR |
| Non-GAAP Gross Margin | 59.0% | ~60%+ | ~61%+ |
| Non-GAAP EPS | $2.84 | ~$4.50+ | ~$6.50+ |

### Revenue Growth Bridge (FY2026 → FY2028)

```
FY2026: $8.2B
  + Custom Silicon growth (~20% CAGR)      → +$1.0B
  + Interconnect/Optics (50%+ growth)       → +$2.0B
  + Switching (doubling)                    → +$0.7B
  + Celestial AI ramp                       → +$0.5B
  + Other Data Center + Comms               → +$2.6B
FY2028: ~$15B
```

---

## 7. Risks

1. **Customer concentration**: Amazon and Microsoft represent an outsized share of custom silicon revenue. Loss of either would be material.
2. **Microsoft/Broadcom switching risk**: Reports of Microsoft evaluating Broadcom for Maia could reduce Marvell's custom ASIC TAM by $500M+.
3. **Broadcom dominance**: Broadcom holds ~80% custom ASIC share and leads in switching (102.4T Tomahawk 6). Marvell is the clear #2 in a winner-take-most market.
4. **Celestial AI execution**: The $3.25B acquisition must deliver on a $1B ARR target by FY2029—execution risk is non-trivial for a pre-revenue technology.
5. **Gross margin pressure**: Custom silicon carries lower margins than Marvell's networking/optics products; mix shift toward custom could compress margins.
6. **Switching gap**: Marvell's 51.2T Teralynx is roughly one generation behind Broadcom's 102.4T Tomahawk 6, limiting share gains in the highest-bandwidth tier.

---

## 8. Summary

Marvell is the second-most-important infrastructure silicon company for AI data centers after Broadcom. Its strength lies in the combination of custom ASIC design partnerships (Amazon, Microsoft), optical DSP leadership (1.6T Ara platform), and a forward-looking silicon photonics strategy (Celestial AI Photonic Fabric). The company's silicon photonics exposure—currently ~18% of revenue and growing toward 25-30%—positions it at the epicenter of the copper-to-optical transition that is essential for scaling both training and inference workloads. Scale-out optical (pluggable transceivers, DSPs) is the larger revenue base today (~$26B market in 2026), but scale-up optical (co-packaged optics, photonic fabric) is the faster-growing and more transformative opportunity, converting an entirely copper-based domain to optical over the next 3-5 years. Marvell's $3.25B Celestial AI acquisition is a direct bet on scale-up optical, while its Ara DSP platform captures the scale-out cycle. Across both vectors, training built the initial market but disaggregated inference architectures represent the larger long-term TAM.

---

*Sources: Marvell FY2026 10-K, Q4 FY2026 earnings call and press release, Marvell product announcements (OFC 2025/2026), Celestial AI acquisition filings, industry research (IDTechEx, Lightwave, Nature npj Nanophotonics), Lambda AI research, SiP-ML (MIT CSAIL).*
