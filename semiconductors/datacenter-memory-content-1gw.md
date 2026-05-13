# NAND, DRAM & HBM Content in a 1 GW AI Datacenter

## Training vs. Inference Workloads

---

## Executive Summary

A 1 GW AI datacenter represents the frontier of hyperscale infrastructure, with facilities of this magnitude under construction or announced by Meta, Microsoft, xAI, and others as of mid-2026. This analysis estimates the NAND flash, DRAM, and HBM semiconductor content required to equip a 1 GW facility optimized for (a) training and (b) inference workloads, using NVIDIA GB300 NVL72 as the reference platform.

**Key findings:**

| Metric | Training | Inference | Delta |
|---|---|---|---|
| GPU count | ~295,000 | ~323,000 | +10% |
| HBM3e | ~85 PB | ~93 PB | +9% |
| DRAM (LPDDR5X/DDR5) | ~75 PB | ~80 PB | +7% |
| NAND flash | ~5.4 EB | ~6.1 EB | +13% |
| Total memory semiconductor $ | ~$1.7B | ~$1.9B | +12% |

Inference datacenters carry **higher total memory content per facility** because they allocate less power to networking (no all-reduce fabric), freeing capacity for more GPU racks and dedicated context-memory storage (NVIDIA CMX). Training facilities consume more NAND in shared storage tiers (checkpointing, training datasets) while inference facilities concentrate NAND in per-GPU context-memory tiers. HBM content per GPU is identical across workloads since it is fixed in silicon.

The transition to NVIDIA Vera Rubin (HBM4, shipping 2026) increases memory semiconductor content per facility by an estimated ~35%, to ~$2.6B for a 1 GW inference build.

---

## Reference Platform: NVIDIA GB300 NVL72

All estimates use the GB300 NVL72 rack as the primary building block, representing the state-of-the-art rack-scale GPU platform shipping in volume in 2025-2026.

| Specification | GB300 NVL72 |
|---|---|
| GPUs per rack | 72 (NVIDIA B300) |
| CPUs per rack | 36 (NVIDIA Grace) |
| Rack TDP | 132 kW |
| **HBM3e per GPU** | **288 GB** |
| **HBM3e per rack** | **20.7 TB** |
| **LPDDR5X system memory** | **17 TB** |
| **E1.S NVMe drive bays** | **144** |
| Local NAND per rack (at 4 TB/drive) | 576 TB |
| Cooling | Direct-to-chip liquid |
| Rack size | 48U |

Sources: NVIDIA, Supermicro, Lenovo product datasheets (May 2026).

---

## Methodology & Key Assumptions

### Facility Power Envelope

| Parameter | Training | Inference | Rationale |
|---|---|---|---|
| Total facility power | 1,000 MW | 1,000 MW | Assumed equal for comparison |
| PUE | 1.15 | 1.18 | Training fully liquid-cooled; inference partially mixed |
| IT load | 870 MW | 847 MW | = Facility power / PUE |

### IT Power Allocation

Training and inference workloads impose fundamentally different infrastructure requirements, leading to different power allocation profiles.

| Subsystem | Training | Inference | Notes |
|---|---|---|---|
| GPU compute racks | 62% (540 MW) | 70% (593 MW) | Inference needs less interconnect, freeing power for GPUs |
| Network fabric | 15% (130 MW) | 8% (68 MW) | Training requires all-reduce / all-to-all across thousands of GPUs |
| Context memory storage (CMX) | -- | 8% (68 MW) | Inference-specific: NVIDIA CMX / ICMS tier |
| Shared storage tier | 13% (113 MW) | 5% (42 MW) | Training needs far more shared storage (data, checkpoints) |
| Management / other | 10% (87 MW) | 9% (76 MW) | Control plane, monitoring, preprocessing |

### Pricing Assumptions (mid-2026)

| Component | $/GB | $/TB |
|---|---|---|
| HBM3e | $10.00 | $10,240 |
| LPDDR5X / DDR5 | $3.50 | $3,584 |
| NAND (die-level content) | $0.11 | $113 |
| HBM4 (Vera Rubin) | $14.00 | $14,336 |

---

## 1 GW Training Datacenter

### GPU Infrastructure

| | Value |
|---|---|
| GPU rack power budget | 540 MW |
| Rack platform | GB300 NVL72 @ 132 kW |
| Number of GPU racks | 4,091 |
| Total GPUs | 294,545 |

### Memory Content: GPU Racks

| Component | Per Rack | Total | $ Value |
|---|---|---|---|
| HBM3e | 20.7 TB | **84.7 PB** | $867M |
| LPDDR5X (system) | 17 TB | **69.5 PB** | $249M |
| Local NAND (E1.S NVMe) | 576 TB | **2,356 PB (2.4 EB)** | $266M |

### Shared Storage Tier

Training clusters require extensive shared storage for:

- **Training datasets**: Curated web-scale corpora (Common Crawl, proprietary data) — typically 50-100+ PB per major training run
- **Checkpoints**: Full model + optimizer state snapshots at regular intervals. For a 1T-parameter model in mixed precision, each checkpoint is ~8-12 TB. Keeping 100+ checkpoints for fault recovery across multiple concurrent experiments can require 50+ PB
- **Model artifacts**: Intermediate outputs, evaluation results, logs

At this GPU scale (~295K GPUs), the facility likely supports dozens of concurrent training runs of varying size, sharing common datasets.

| Component | Estimate | Rationale |
|---|---|---|
| Shared NAND | **~3.0 EB** | ~10 TB per GPU; driven by checkpoint frequency and dataset duplication for locality |
| Storage server DRAM | **~4 PB** | Metadata servers, caching layers for parallel file systems (WEKA, VAST, DDN) |

### Network Fabric

Training at this scale demands a massive InfiniBand or NVLink-network spine-leaf fabric supporting all-to-all communication patterns. The 130 MW budget implies thousands of high-radix switches (NVIDIA Quantum/Spectrum).

| Component | Estimate | Notes |
|---|---|---|
| Switch DRAM | ~0.5 PB | ~3,000-5,000 switches, 64-128 GB each |
| Switch NAND | Negligible | Firmware/boot only |

### CPU / Management Servers

Data preprocessing, orchestration, monitoring, and job scheduling servers.

| Component | Estimate |
|---|---|
| Server DRAM (DDR5) | ~1 PB |
| Server NAND | ~50 PB |

### Training Datacenter: Total Memory Content

| Memory Type | Quantity | $ Value | % of Total $ |
|---|---|---|---|
| **HBM3e** | **85 PB** | **$867M** | **50%** |
| **DRAM (LPDDR5X + DDR5)** | **75 PB** | **$269M** | **15%** |
| **NAND** | **5.4 EB** | **$610M** | **35%** |
| **Total** | | **$1,746M** | **100%** |

### Per-GPU Memory Intensity (Training)

| | Per GPU |
|---|---|
| HBM3e | 288 GB |
| DRAM | 255 GB |
| NAND | 18.3 TB |
| Memory semiconductor $ | $5,927 |

---

## 1 GW Inference Datacenter

### GPU Infrastructure

| | Value |
|---|---|
| GPU rack power budget | 593 MW |
| Rack platform | GB300 NVL72 @ 132 kW |
| Number of GPU racks | 4,492 |
| Total GPUs | 323,424 |

The inference facility fits ~10% more GPUs because the reduced networking requirement (no all-to-all fabric) frees power for additional GPU racks.

### Memory Content: GPU Racks

| Component | Per Rack | Total | $ Value |
|---|---|---|---|
| HBM3e | 20.7 TB | **93.0 PB** | $952M |
| LPDDR5X (system) | 17 TB | **76.4 PB** | $274M |
| Local NAND (E1.S NVMe) | 576 TB | **2,587 PB (2.6 EB)** | $292M |

### Context Memory Storage (NVIDIA CMX / ICMS)

This is the defining architectural difference for inference. NVIDIA's Context Memory Storage platform (CMX, formerly ICMS) creates a new storage tier between GPU HBM and shared network storage, using BlueField-4 DPUs managing NVMe SSDs. The system stores KV cache for long-context, multi-turn, and agentic AI workloads, addressing the "HBM black hole" problem where context data consumes GPU memory that should be used for compute.

CMX delivers 5x higher tokens-per-second and 5x better power efficiency versus re-computing context from shared storage.

| Parameter | Estimate | Rationale |
|---|---|---|
| CMX NAND per GPU | 10 TB | Conservative for GB300 era; Vera Rubin targets 16 TB/GPU |
| Total CMX NAND | **3,234 PB (3.2 EB)** | 323,424 GPUs x 10 TB |
| CMX DRAM (BF4 DPU memory) | ~2 PB | BlueField-4 DPU memory for metadata and indexing |

### Shared Storage Tier

Inference requires far less shared storage than training — no checkpointing or training data. Storage is primarily for:

- Model weight repositories (serving dozens to hundreds of distinct models)
- Request/response logging and audit trails
- Input/output data buffers

| Component | Estimate | Rationale |
|---|---|---|
| Shared NAND | **~0.3 EB** | ~1 TB per GPU; models are read-mostly and replicated across racks |
| Storage server DRAM | ~1 PB | Less metadata overhead than training file systems |

### Network Fabric

Inference networking is simpler — primarily load balancing and request routing rather than all-to-all GPU communication.

| Component | Estimate |
|---|---|
| Switch DRAM | ~0.3 PB |

### CPU / Management Servers

| Component | Estimate |
|---|---|
| Server DRAM (DDR5) | ~1 PB |
| Server NAND | ~30 PB |

### Inference Datacenter: Total Memory Content

| Memory Type | Quantity | $ Value | % of Total $ |
|---|---|---|---|
| **HBM3e** | **93 PB** | **$952M** | **50%** |
| **DRAM (LPDDR5X + DDR5)** | **80 PB** | **$286M** | **15%** |
| **NAND** | **6.1 EB** | **$689M** | **36%** |
| **Total** | | **$1,927M** | **100%** |

### Per-GPU Memory Intensity (Inference)

| | Per GPU |
|---|---|
| HBM3e | 288 GB |
| DRAM | 247 GB |
| NAND | 18.9 TB |
| Memory semiconductor $ | $5,959 |

---

## Side-by-Side Comparison

### Absolute Content

| | Training 1 GW | Inference 1 GW | Inference Premium |
|---|---|---|---|
| GPUs | 294,545 | 323,424 | +10% |
| **HBM3e** | **85 PB** | **93 PB** | **+9%** |
| **DRAM** | **75 PB** | **80 PB** | **+7%** |
| **NAND** | **5.4 EB** | **6.1 EB** | **+13%** |
| **Memory $ total** | **$1.75B** | **$1.93B** | **+10%** |

### Memory Content per MW (Total Facility Power)

| | Training | Inference | Unit |
|---|---|---|---|
| HBM3e / MW | 85 | 93 | TB/MW |
| DRAM / MW | 75 | 80 | TB/MW |
| NAND / MW | 5.4 | 6.1 | PB/MW |
| Memory $ / MW | $1.75M | $1.93M | $/MW |

### NAND Composition: Where the Flash Goes

| NAND Tier | Training | Inference |
|---|---|---|
| Local (GPU rack E1.S) | 2.4 EB (44%) | 2.6 EB (43%) |
| Context Memory (CMX) | -- | 3.2 EB (52%) |
| Shared Storage | 3.0 EB (56%) | 0.3 EB (5%) |
| **Total** | **5.4 EB** | **6.1 EB** |

The NAND composition tells a sharply different story despite similar totals:
- **Training** concentrates NAND in shared storage tiers (checkpoint and dataset repositories)
- **Inference** concentrates NAND in per-GPU context-memory tiers (CMX/ICMS for KV cache)

This distinction matters for NAND suppliers: training storage favors high-capacity, throughput-oriented enterprise SSDs (61.44 TB+ form factors), while CMX favors lower-latency, write-endurance-optimized drives in E1.S form factor.

---

## Forward Look: NVIDIA Vera Rubin (2026-2027)

The Vera Rubin platform (R100 GPU + Vera CPU) shifts memory content materially upward, particularly for inference.

### Vera Rubin NVL72 Estimated Specifications

| Specification | GB300 NVL72 | Vera Rubin NVL72 (est.) | Change |
|---|---|---|---|
| HBM generation | HBM3e | HBM4 | New gen |
| HBM per GPU | 288 GB | 384 GB | +33% |
| HBM per rack | 20.7 TB | 27.6 TB | +33% |
| System memory per rack | 17 TB (LPDDR5X) | ~20 TB (LPDDR6, est.) | +18% |
| CMX NAND per GPU | 10 TB (est.) | 16 TB (confirmed) | +60% |
| Rack TDP | 132 kW | ~150 kW (est.) | +14% |

Sources: NVIDIA product announcements (GTC 2026, CES 2026), Citi estimates on ICMS NAND demand, Tom's Hardware / WCCFTech reporting on R100 specifications.

### 1 GW Inference on Vera Rubin

| Component | GB300 NVL72 | Vera Rubin NVL72 | Change |
|---|---|---|---|
| GPU racks (at 70% IT allocation) | 4,492 | 3,953 | -12% |
| Total GPUs | 323,424 | 284,616 | -12% |
| **HBM** | **93 PB (HBM3e)** | **109 PB (HBM4)** | **+17%** |
| **DRAM** | **80 PB** | **82 PB** | **+3%** |
| **NAND** | **6.1 EB** | **7.5 EB** | **+23%** |

### Dollar Content Comparison (1 GW Inference)

| | GB300 | Vera Rubin | Change |
|---|---|---|---|
| HBM $ | $952M | $1,563M | +64% |
| DRAM $ | $286M | $294M | +3% |
| NAND $ | $689M | $848M | +23% |
| **Total memory $** | **$1,927M** | **$2,704M** | **+40%** |

The HBM4 price premium (~$14/GB vs. $10/GB for HBM3e) combined with the larger capacity per GPU drives the most significant dollar uplift. NAND growth is driven by the 60% increase in CMX storage per GPU (16 TB vs. 10 TB).

---

## Supplier Implications

### HBM (50% of memory $ content)

HBM remains the highest-dollar-content memory component per facility. At ~$0.9-1.6B per 1 GW datacenter, HBM is a concentrated revenue pool split among three suppliers:

| Supplier | Estimated HBM Share (2026) | Notes |
|---|---|---|
| SK Hynix | ~50-55% | Technology leader, primary NVIDIA supplier for HBM3e/HBM4 |
| Samsung | ~30-35% | Ramping HBM4 mass production (early 2026) |
| Micron | ~15-20% | Competitive on HBM3e, HBM4 expected late 2026 |

### DRAM (15% of memory $ content)

System DRAM (LPDDR5X for Grace-based systems, DDR5 for x86 storage/management servers) contributes a smaller but steady share. The LPDDR5X in NVL72 racks is a differentiated market segment with premium pricing.

### NAND (35% of memory $ content)

NAND is the fastest-growing memory category in AI datacenters, driven by CMX/ICMS adoption. Citi estimates ICMS alone could consume 2.8% of global NAND supply in 2026 and 9.3% in 2027. Enterprise SSD demand is forecast to grow 58% in 2026 (Goldman Sachs).

| Supplier | Position | Notes |
|---|---|---|
| Samsung | Market leader | Gen6 PCIe SSDs up to 256 TB for AI servers |
| Kioxia / WDC | Strong in enterprise | Competitive in high-capacity E1.S form factors |
| SK Hynix / Solidigm | Growing share | Solidigm (ex-Intel NAND) strong in datacenter |
| Micron | Competitive | Broad enterprise SSD portfolio |

---

## Key Risks & Sensitivities

| Factor | Impact on Estimates |
|---|---|
| **PUE variance** (1.10-1.25) | +/- 5-8% on all memory content; lower PUE = more GPUs per MW |
| **GPU rack power allocation** (60-75%) | +/- 10-15% on GPU count and associated HBM/DRAM |
| **CMX adoption rate** | Binary risk for inference NAND; without CMX, inference NAND drops ~50% |
| **Model quantization** | Reduces effective HBM utilization but does not change physical HBM per GPU |
| **Mix of GPU platforms** | Real facilities may mix generations (H100/H200/B200/B300), altering per-GPU memory content |
| **NAND pricing** | Enterprise SSD pricing is volatile; Goldman Sachs forecasts significant shortage in 2026-2027 |
| **HBM pricing** | Premium pricing may compress as Samsung/Micron ramp HBM4 capacity |
| **Non-NVIDIA platforms** | AMD MI350X and custom ASICs (Google TPU, Amazon Trainium) have different memory architectures |
| **Shared storage ratio** | Training shared storage is workload-dependent; could range from 5-20 TB/GPU |

---

## Appendix: Data Sources

- NVIDIA GB300 NVL72 product specifications (Supermicro, Lenovo, HPE datasheets, May 2026)
- NVIDIA DGX B300 User Guide (docs.nvidia.com)
- NVIDIA CMX / Inference Context Memory Storage Platform specifications (nvidia.com, CES 2026)
- NVIDIA Vera Rubin platform announcements (GTC 2026)
- Samsung HBM4 mass production announcement (early 2026)
- Goldman Sachs: DRAM/NAND supply-demand analysis (2026)
- Morgan Stanley: Memory supercycle and HBM demand estimates (2026)
- Citi: ICMS NAND demand projections (2026-2027)
- Meta Prometheus cluster specifications (1,020 MW, 500K GPUs)
- Blocks & Files, SDX Central, WCCFTech, Tom's Hardware industry reporting

---

*Analysis prepared May 2026. All estimates are forward-looking and subject to revision as platform specifications, pricing, and deployment patterns evolve.*
