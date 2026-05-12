# Key Drivers of PUE for AI Data Centers

**Equity Research Supplement | May 2026**

---

## Executive Summary

PUE (Power Usage Effectiveness) measures how efficiently a data center converts total facility power into useful IT compute. A PUE of 1.0 means every watt entering the building reaches IT equipment — theoretically perfect but physically unachievable. The industry average remains stuck at ~1.58 (37% wasted on overhead), while best-in-class AI data centers achieve 1.04-1.10. For a 1 GW AI campus, the difference between PUE 1.60 and PUE 1.10 equates to **~310 MW of saved overhead power** — enough to run an additional 200,000+ GPUs. Understanding what drives PUE is therefore a direct lever on both cost and compute capacity.

---

## 1. PUE Fundamentals

### The Formula

```
PUE = Total Facility Power / IT Equipment Power
```

| PUE | Overhead % | Meaning |
|---|---|---|
| 1.00 | 0% | Theoretical perfection (impossible) |
| 1.04 | 3.8% | Best-in-class hyperscaler (Google Lancaster, OH) |
| 1.10 | 9.1% | Leading AI data center with liquid cooling |
| 1.20 | 16.7% | Good modern hyperscale facility |
| 1.40 | 28.6% | Industry median (2026) |
| 1.58 | 36.7% | Global weighted average |
| 1.80 | 44.4% | Legacy air-cooled enterprise facility |
| 2.00 | 50.0% | Poorly designed or aged facility |

### Where PUE Overhead Goes (Traditional Air-Cooled Facility, PUE ~1.6)

| Component | Share of Overhead Power | Share of Total Facility Power |
|---|---|---|
| **Cooling (HVAC/CRAC/chillers)** | 60-70% | 22-26% |
| **Power distribution losses (UPS, PDU, transformers)** | 20-25% | 7-9% |
| **Fans and air movement** | 8-12% | 3-4% |
| **Lighting, security, BMS, other** | 3-5% | 1-2% |
| **Total overhead** | 100% | ~37% |

**Cooling dominates.** It accounts for 60-70% of all non-IT power consumption. This is why cooling architecture choice is the single most impactful PUE decision.

---

## 2. The Five Key PUE Drivers for AI Data Centers

### Driver 1: Cooling Architecture (Largest Impact)

**Contribution to PUE: 0.03-0.60 of total overhead**

This is the dominant variable. The choice between air, liquid, and immersion cooling creates the widest performance spread.

| Cooling Technology | Achievable PUE | Cooling Energy as % of IT Load | Viable Density |
|---|---|---|---|
| Traditional CRAC/CRAH (air) | 1.50-1.80 | 30-45% | Up to 15-20 kW/rack |
| Hot-aisle containment (air) | 1.30-1.50 | 20-30% | Up to 25-30 kW/rack |
| Rear-door heat exchangers (RDHx) | 1.20-1.35 | 15-25% | Up to 40-50 kW/rack |
| Direct-to-chip liquid (DLC) | 1.08-1.15 | 5-12% | Up to 150 kW/rack |
| Single-phase immersion | 1.03-1.08 | 2-6% | Up to 200+ kW/rack |
| Two-phase immersion | 1.02-1.07 | 1-5% | Up to 250+ kW/rack |

**Why liquid cooling is transformational for PUE:**

1. **Eliminates fan energy.** In air-cooled facilities, fans consume 8-12% of overhead. Liquid cooling eliminates most or all internal air movement for the IT load.
2. **Higher thermal conductivity.** Water carries ~3,500x more heat per unit volume than air. Less energy is needed to move the same heat load.
3. **Warmer operating temperatures.** DLC typically operates at 35-45°C supply water — warm enough to reject heat to ambient via dry coolers for most of the year, avoiding energy-intensive compressor-based refrigeration.
4. **Eliminates raised floor/plenum.** Air-cooled facilities waste significant fan energy pushing air through raised floors. Liquid piping is more direct.

**For AI data centers (80-132+ kW/rack), liquid cooling is not optional — it is physically mandatory.** Air cannot remove 120 kW from a single rack without extreme (and inefficient) airflow volumes. This is why AI data centers inherently achieve better PUE than traditional facilities: the density forces a cooling choice that happens to be more efficient.

#### Sub-Factors Within Cooling:

**a) Heat rejection method (outdoor side):**
| Method | Energy Consumption | Water Use | Climate Dependency |
|---|---|---|---|
| Evaporative cooling towers | Low | High (1.8L/kWh) | Moderate |
| Dry coolers (air-blast radiators) | Medium | Zero | High (hot climates penalized) |
| Adiabatic hybrid coolers | Low-Medium | Low | Low |
| Seawater/river/lake cooling | Very Low | Zero (pass-through) | Site-specific |
| Ground-source/geothermal | Very Low | Zero | Site-specific |

**b) Free cooling hours (economizer utilization):**
The number of annual hours where ambient conditions allow "free" cooling without mechanical refrigeration. This is strongly climate-dependent:

| Climate Zone | Free Cooling Hours/Year | Impact on PUE |
|---|---|---|
| Nordic (Finland, Iceland, Norway) | 8,000+ (>90%) | PUE reduction of 0.15-0.30 |
| Pacific NW / UK / Ireland | 6,500-7,500 (~80%) | PUE reduction of 0.10-0.20 |
| Midwest US (Iowa, Ohio) | 5,000-6,500 (~65%) | PUE reduction of 0.08-0.15 |
| Mid-Atlantic (Virginia) | 4,000-5,500 (~55%) | PUE reduction of 0.05-0.10 |
| Hot-dry (Phoenix, Texas) | 2,000-3,500 (~30%) | PUE reduction of 0.02-0.05 |
| Hot-humid (Singapore, Houston) | 500-2,000 (<20%) | Minimal free cooling benefit |

Google's fleet-wide PUE of 1.09 is achieved partly through strategic site selection in climates offering high free-cooling hours (Finland seawater, Belgium canal water, Iowa/Ohio cold winters).

**c) Supply temperature setpoints:**
Running cooling systems at higher temperatures (warm-water cooling at 35-45°C vs. chilled water at 7-12°C) dramatically expands free-cooling hours. ASHRAE guidelines have progressively raised recommended inlet temperatures from 18-27°C (A1 class, 2004) to 15-40°C (A3/A4 class), enabling warmer operation.

---

### Driver 2: Power Distribution Efficiency (Second Largest Impact)

**Contribution to PUE: 0.03-0.15 of total overhead**

Every power conversion step between the utility feed and the GPU chip introduces losses. Minimizing conversion stages is the key principle.

| Component | Typical Efficiency | Loss per Conversion | Approaches to Minimize |
|---|---|---|---|
| Medium-voltage transformer (MV→LV) | 98-99% | 1-2% | Right-sizing, high-efficiency cores |
| UPS (double-conversion) | 92-97% | 3-8% | Eco-mode, lithium-ion, modular |
| UPS (transformerless) | 96-99% | 1-4% | Eliminates internal transformer |
| PDU (step-down transformer) | 97-99% | 1-3% | Direct 400V distribution |
| Busway distribution | 99-99.5% | 0.5-1% | Short runs, right-sized copper |
| Server PSU (AC→DC) | 92-96% | 4-8% | Titanium-rated PSUs, 48V DC |

**Total power distribution path efficiency:**
- Traditional (480V→UPS→PDU→208V→server): 82-89% (11-18% losses)
- Optimized (400V direct→transformerless UPS→server): 92-96% (4-8% losses)
- HVDC (utility→direct DC→server): 94-97% (3-6% losses)

**Key strategies for AI data centers:**

1. **Eliminate conversion steps.** Google famously eliminates central UPS systems, using server-level battery backup instead. This removes 3-8% of distribution losses.
2. **400V direct distribution.** Keeping power at 400V AC to the rack eliminates the 480V→208V step-down transformer loss (~4%). This is becoming standard for AI facilities.
3. **Higher voltage busbars.** 480V or 600V+ distribution to reduce copper losses (I²R) at high amperages typical of 120 kW racks.
4. **UPS eco-mode operation.** Running UPS in bypass (eco-mode) during normal conditions improves efficiency from 92-94% to 98-99%, at the cost of slightly reduced power protection. Acceptable for facilities with redundant utility feeds.
5. **48V DC architecture.** Pioneered by Google's Open Rack designs. Eliminates multiple AC-DC conversions. Reduces total distribution losses by 50%.
6. **Right-sizing.** Oversized UPS and transformers operating at low load have poor efficiency. Modular UPS allows scaling with actual load, keeping each module at optimal 60-80% loading.

**AI-specific consideration:** GPU power factor. NVIDIA GPU servers have a power factor of ~0.9 (not unity). This creates reactive power that flows through the distribution system, causing additional losses and requiring appropriately rated equipment. Power factor correction at the rack or row level can recover 2-5% of distribution losses.

---

### Driver 3: IT Equipment Efficiency (Indirect PUE Impact)

**Contribution to PUE: Indirect — determines how much useful compute is generated per watt consumed**

While IT equipment power is the denominator of PUE (and thus doesn't directly appear in the ratio), the efficiency of IT equipment determines the heat load that cooling must handle and influences facility design.

| Factor | Impact on PUE | Mechanism |
|---|---|---|
| GPU utilization rate | Indirect | Higher utilization = more consistent heat load = easier to optimize cooling |
| Power management (idle states) | Moderate | Idle GPUs still draw 30-40% of peak power but produce zero useful compute |
| Workload scheduling | Indirect | Consolidating workloads reduces active nodes, reducing cooling load |
| Server PSU efficiency | Direct (0.01-0.03 PUE) | Titanium vs. Gold PSU saves 3-5% at the server level |
| Fan elimination (liquid-cooled) | Direct (0.02-0.05 PUE) | Internal server fans in air-cooled nodes consume 5-15% of node power |

**NVIDIA Blackwell's PUE advantage:** The GB200 NVL72 system is designed exclusively for liquid cooling, with no internal fans. This eliminates 5-15% of node-level power that would otherwise appear as waste heat, directly improving effective PUE. The system also operates at higher inlet temperatures (35-45°C), enabling more economizer hours for the facility.

---

### Driver 4: Climate and Site Selection (Fixed at Design Time)

**Contribution to PUE: 0.05-0.30 depending on geography**

Once a site is selected, the climate imposes a ceiling on PUE performance that no amount of engineering can overcome without energy-intensive mechanical cooling.

| Site Factor | PUE Impact | Examples |
|---|---|---|
| Ambient temperature (annual avg) | 0.05-0.25 | Iceland (2°C avg) vs. Singapore (27°C avg) |
| Humidity | 0.02-0.10 | Dry climates allow more evaporative cooling |
| Altitude | 0.01-0.03 | Higher altitude = thinner air = less effective air cooling |
| Water availability | 0.05-0.15 | Abundant water enables evaporative towers |
| Renewable heat sinks | 0.03-0.10 | Seawater, lake water, river water for direct cooling |
| Diurnal temperature range | 0.02-0.05 | Large day/night swings enable nighttime free cooling |

**Why Iceland, Nordics, and Pacific Northwest are favored for AI:**
- Average temperatures of 0-12°C enable year-round free cooling
- Abundant renewable power (hydro, geothermal, wind)
- Cool water bodies for direct heat rejection
- Achievable PUE: 1.03-1.08

**Why Northern Virginia is challenging despite being the largest DC market:**
- Hot, humid summers (30-35°C with high dew points)
- Only ~55% free-cooling hours annually
- Requires significant mechanical cooling May-September
- Achievable PUE: 1.15-1.30 (with DLC); 1.35-1.60 (air-cooled)

**Why Texas (ERCOT) presents a PUE trade-off:**
- Power is cheap and available (fast interconnection)
- But hot climate (35-40°C summers) degrades cooling efficiency
- ERCOT operators typically accept PUE 1.15-1.25 (with DLC) as cost of faster power access
- The economics work: cheap power ($0.04-0.06/kWh) offsets slightly worse PUE vs. Nordic sites

---

### Driver 5: Operational Optimization and Controls (Continuous Improvement)

**Contribution to PUE: 0.02-0.10 ongoing improvement**

After design and construction, software-driven optimization provides the final PUE improvements.

| Optimization | PUE Improvement | Mechanism |
|---|---|---|
| ML-based cooling control | 0.03-0.10 | Predicts heat load; adjusts pumps, fans, tower sequencing in real-time |
| Variable speed drives (VFDs) | 0.02-0.05 | Reduces pump/fan energy at partial load (cube law: 50% speed = 12.5% power) |
| Hot/cold aisle containment | 0.05-0.15 | Eliminates mixing of supply and return air |
| Setpoint optimization | 0.02-0.05 | Raising supply temp from 18°C to 27°C reduces chiller load |
| Airflow management (blanking panels, grommets) | 0.01-0.03 | Prevents bypass airflow in air-cooled environments |
| Predictive maintenance | 0.01-0.02 | Keeps equipment at peak efficiency; prevents degradation drift |
| Real-time PUE monitoring | 0.01-0.03 | Enables rapid identification and correction of inefficiencies |

**Google's DeepMind cooling AI** reduced cooling energy by 40% vs. manual operation by predicting heat loads 30 minutes ahead and adjusting cooling plant operation pre-emptively. This was one of the earliest production applications of reinforcement learning at Google.

**Key principle: the cube law.** Pump and fan power scales with the cube of flow rate. Reducing cooling flow by 20% reduces pumping energy by 49%. Variable-speed drives (VFDs) on all pumps, fans, and compressors are essential for AI data centers with dynamic workloads.

---

## 3. PUE Breakdown: AI Data Center at 1.10 vs. Traditional at 1.58

### Component-by-Component Comparison (Per MW of IT Load)

| PUE Component | Traditional Air-Cooled (PUE 1.58) | AI Liquid-Cooled (PUE 1.10) | How AI Achieves Savings |
|---|---|---|---|
| **Cooling (chillers/towers)** | 0.25-0.30 | 0.02-0.04 | DLC with warm water enables dry coolers; no chiller needed 80%+ of year |
| **Cooling (fans/air handling)** | 0.08-0.12 | 0.00-0.01 | Liquid eliminates CRAH units and internal server fans |
| **Cooling (pumps)** | 0.02-0.04 | 0.02-0.03 | Pumps still needed for liquid loops, but at lower energy |
| **UPS losses** | 0.04-0.06 | 0.01-0.02 | Eco-mode UPS or distributed server-level batteries |
| **Transformer/PDU losses** | 0.03-0.05 | 0.01-0.02 | 400V direct distribution; fewer conversion stages |
| **Lighting, security, BMS** | 0.01-0.02 | 0.01-0.01 | Similar (small) in both; LEDs standard |
| **Total PUE overhead** | **0.43-0.59** | **0.07-0.13** | |
| **PUE** | **1.43-1.59** | **1.07-1.13** | |

### What This Means in Absolute Power (1 GW Facility)

| Scenario | IT Power Delivered | Overhead Power Consumed | GPUs Powered (B200 @ 1kW) |
|---|---|---|---|
| PUE 1.60 (legacy air) | 625 MW | 375 MW | ~625,000 |
| PUE 1.40 (good air) | 714 MW | 286 MW | ~714,000 |
| PUE 1.20 (good DLC) | 833 MW | 167 MW | ~833,000 |
| PUE 1.10 (excellent DLC) | 909 MW | 91 MW | ~909,000 |
| PUE 1.05 (frontier) | 952 MW | 48 MW | ~952,000 |

**Moving from PUE 1.40 to 1.10 in a 1 GW facility frees up 195 MW** — equivalent to ~195,000 additional B200 GPUs. At $10-12M revenue per MW, this is **$1.9-2.3 billion in additional annual revenue** from the same utility feed.

---

## 4. PUE Sensitivity Analysis

### What Moves PUE the Most?

| Change | PUE Impact | Difficulty | Cost |
|---|---|---|---|
| Air cooling → Direct liquid cooling | **-0.20 to -0.40** | High (retrofit) / Medium (new build) | +$1.5-3M/MW CapEx |
| Add free cooling economizer | **-0.10 to -0.25** | Medium | +$0.3-0.8M/MW |
| UPS double-conversion → eco-mode | **-0.03 to -0.06** | Low | Minimal (firmware/config) |
| Eliminate PDU transformer (400V direct) | **-0.02 to -0.04** | Medium | Requires 400V-rated IT |
| ML-optimized cooling controls | **-0.03 to -0.10** | Medium | Software + sensors |
| Site selection (Virginia → Iowa) | **-0.05 to -0.15** | Fixed at design | Land cost differential |
| Site selection (Virginia → Iceland) | **-0.10 to -0.25** | Fixed at design | Latency/connectivity tradeoff |
| Raise inlet temp setpoint 18°C → 27°C | **-0.03 to -0.08** | Low | None (server tolerance) |
| VFDs on all motors | **-0.02 to -0.05** | Low-Medium | +$50-150K/MW |
| Server-level batteries (eliminate UPS) | **-0.03 to -0.06** | High | Custom server design |

### Diminishing Returns Curve

PUE optimization follows a diminishing returns curve. The first improvements are cheap and impactful; the last basis points are expensive and complex:

| PUE Target | Difficulty Level | Primary Levers |
|---|---|---|
| 1.80 → 1.50 | Easy | Basic containment, economizers, right-sizing |
| 1.50 → 1.30 | Moderate | Full containment, VFDs, warm-aisle, efficient UPS |
| 1.30 → 1.15 | Hard | Liquid cooling, 400V distribution, eco-mode UPS |
| 1.15 → 1.08 | Very Hard | Full DLC, ML controls, optimal climate site, server-level batteries |
| 1.08 → 1.04 | Extreme | Immersion cooling, Nordic/cold-water site, HVDC, custom server design |

---

## 5. PUE Economics: The Financial Case

### Cost of PUE Improvement vs. Benefit

For a 100 MW AI data center at $0.07/kWh:

| From PUE | To PUE | MW Saved | Annual Energy Savings | Incremental CapEx | Simple Payback |
|---|---|---|---|---|---|
| 1.40 → 1.20 | 1.20 | 11.9 MW | $7.3M/year | $15-25M (liquid cooling) | 2-3 years |
| 1.20 → 1.10 | 1.10 | 7.6 MW | $4.6M/year | $8-15M (optimization) | 2-3 years |
| 1.10 → 1.05 | 1.05 | 4.3 MW | $2.7M/year | $10-20M (extreme measures) | 4-7 years |

**But the real value isn't just energy savings — it's the additional compute capacity:**

| PUE Improvement | Additional MW Available for GPUs | Additional Revenue (@ $10M/MW) |
|---|---|---|
| 1.40 → 1.20 | +11.9 MW | +$119M/year |
| 1.20 → 1.10 | +7.6 MW | +$76M/year |
| 1.10 → 1.05 | +4.3 MW | +$43M/year |

When the revenue from additional compute is factored in (not just energy savings), PUE improvements have sub-1-year payback periods for AI data centers. This is why hyperscalers invest aggressively in PUE despite already being far ahead of industry averages.

---

## 6. Best-in-Class: How Hyperscalers Achieve PUE 1.04-1.10

### Google (PUE 1.09 fleet-wide; 1.04 best facility)

Key techniques:
- **Server-level batteries** instead of central UPS (eliminates 3-6% distribution losses)
- **Custom server design** with 48V DC distribution
- **DeepMind ML cooling control** (40% cooling energy reduction)
- **Strategic site selection** (cold climates with water access: Finland, Belgium, Iowa)
- **Seawater/canal water cooling** at Nordic/European sites
- **Hot-aisle temperatures of 95°F+** allowed (expanding free-cooling hours)
- **No raised floors** — overhead cable/pipe distribution

### Meta (PUE 1.08 at Prineville; ~1.10 fleet average)

Key techniques:
- **Open Compute Project** hardware (custom, efficient servers)
- **Evaporative cooling** leveraging dry Western US climates (Prineville, OR)
- **Penthouse air-handling units** instead of traditional CRAH
- **480V distribution directly to racks**
- **Warm-water cooling loops** for newer AI facilities

### Microsoft (Approaching 1.05 for new AI facilities)

Key techniques:
- **Direct-to-chip liquid cooling** deployment across Azure fleet (started Jul 2025)
- **Zero-water cooling designs** for next-gen datacenters (Dec 2024 announcement)
- **Two-phase immersion exploration** for highest-density AI workloads
- **Custom power distribution** with minimal conversion stages

---

## 7. Emerging PUE Metrics and Limitations

### Beyond PUE: New Metrics for AI Data Centers

PUE has limitations for AI facilities. Emerging alternatives:

| Metric | Formula | What It Captures |
|---|---|---|
| **PUE** | Total Power / IT Power | Basic facility efficiency |
| **PCE** (Power Compute Effectiveness) | Total Power / Compute Output | Revenue-generating efficiency |
| **WUE** (Water Usage Effectiveness) | Water Use (L) / IT Power (kWh) | Water sustainability |
| **CUE** (Carbon Usage Effectiveness) | CO₂ Emissions / IT Power | Carbon intensity |
| **TUE** (Total-resource Usage Effectiveness) | Combined energy + water + carbon | Holistic sustainability |
| **Tokens per Watt-hour** | Inference tokens / Total Wh | AI-specific output efficiency |

**PCE is increasingly relevant** because two facilities can have identical PUE but vastly different useful output. An AI facility at PUE 1.15 generating 60,000 tokens/sec/GPU delivers more value than the same facility running idle GPUs at PUE 1.10.

### PUE Limitations

1. **PUE doesn't measure IT efficiency.** A facility running idle servers at PUE 1.10 wastes more total energy than a fully utilized facility at PUE 1.30.
2. **Partial PUE (pPUE)** measures individual systems within a facility, useful for identifying improvement targets.
3. **Annualized vs. instantaneous.** PUE varies seasonally (worse in summer). Annual average can mask periods of poor performance.
4. **Doesn't account for location.** A PUE 1.10 in Iceland (cheap hydro power) is different economically from PUE 1.10 in Singapore (expensive imported LNG).

---

## 8. Summary: PUE Driver Hierarchy for AI Data Centers

| Rank | Driver | PUE Impact | Controllability | When Determined |
|---|---|---|---|---|
| **1** | Cooling architecture (air vs. liquid vs. immersion) | 0.20-0.60 | High | Design phase |
| **2** | Climate/site selection | 0.05-0.30 | Fixed | Site selection |
| **3** | Power distribution design (voltage, conversion stages) | 0.03-0.15 | High | Design phase |
| **4** | Operational optimization (ML controls, VFDs, setpoints) | 0.02-0.10 | High | Ongoing |
| **5** | IT equipment design (server fans, PSU, power factor) | 0.01-0.05 | Medium | Hardware procurement |

**For AI data centers specifically, the hierarchy simplifies to:**

1. **Liquid cooling is mandatory** (due to 80-132+ kW/rack density), which inherently drives PUE below 1.20
2. **Site climate determines whether you reach 1.05 or 1.15** — the spread between Nordic and hot-climate facilities
3. **Power distribution architecture** provides the next 0.03-0.08 improvement via 400V direct, eco-mode UPS, or server-level batteries
4. **ML-based optimization** continuously squeezes the last 0.02-0.05

The most important insight: **AI data centers achieve better PUE than traditional facilities not primarily because operators invest more in efficiency, but because the physics of GPU density forces adoption of liquid cooling — which happens to be inherently more efficient than air.** The density constraint and the efficiency outcome are structurally linked.

---

*Sources: Google Environmental Report 2025, Uptime Institute Global Survey 2025, Schneider Electric WP110, NVIDIA GB200 NVL72 specifications, Open Compute Project 2026 guidelines, Build.inc analysis, Introl deployment guides, industry disclosures. May 2026.*
