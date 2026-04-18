# Agentic CX Vendor Contrast: Sierra vs. Decagon vs. Intercom Fin vs. Zendesk

*As of April 2026. Figures are sourced from public reporting (Sacra, The Information / TBPN, Sierra/Intercom/Zendesk press releases, OpenAI DevDay 2025 trillion‑token slide, Hackernoon, Letter Two, SiliconANGLE, Sacra, etc.). Where numbers are not publicly disclosed, the table flags "n/d" (not disclosed) rather than guessing.*

---

## 1. One‑page contrast

| Dimension | **Sierra** | **Decagon** | **Intercom – Fin** | **Zendesk – Resolution Platform / AI Agents** |
|---|---|---|---|---|
| Founded / agent launched | Feb 2024 (public debut) | 2023 | Mar 2023 (Fin v1) | Pre‑existing platform; "AI Agents" SKU launched 2024, Resolution Platform 2025 |
| Pricing model | Outcome‑based (per resolved outcome / completed work) | Per‑conversation **and** per‑resolution (~$1.50/resolution mid‑market) | Per‑resolution ($0.99) on top of per‑seat Intercom | Outcome‑based per AI resolution + per‑seat Suite |
| Model strategy | "Constellation of models" – 15+ frontier + open‑weight + fine‑tuned proprietary, multi‑provider router | Multi‑model, primarily OpenAI (GPT‑3.5 / 4 / 4o); >1T OpenAI tokens consumed | Anthropic Claude 3.5 Sonnet (Fin 2, switched from OpenAI in Oct 2024) + RAG stack | OpenAI GPT‑4o → GPT‑5 + MCP; adds Forethought self‑improving agents (acq. Mar 2026) |
| Resolution rate (avg) | Not disclosed (outcome‑priced) | ~70–80%+ deflection on enterprise deployments | 66–67% (top customers ~70%) | Up to ~80% automation for top customers; ~60% reported avg |
| Headline customers | OpenAI, WeightWatchers, SiriusXM, Sonos, Discord, Rivian, SoFi, ADT, Cigna, Vans, Deliveroo, Bissell | Notion, Duolingo, Rippling, Bilt, Eventbrite, Substack, Avis, Mercado Libre, Deutsche Telekom | Anthropic, Vanta, Clay, Whoop, Miro, Sharesies (~7–8k Fin customers) | ~20,000 AI customers across the Zendesk base |
| Latest valuation | $10B (Series D, Sep 2025) | $4.5B (Jan 2026 round) | Private (Intercom standalone) | $9.6B (PE‑owned by H&F / Permira since 2022) |

---

## 2. ARR time series for the *agentic* offering

All figures are the **agentic / AI‑agent product line only** (not whole‑company), unless noted. Time stamps are when the figure was first publicly reported.

### Sierra (entire company is agentic)
| Date | ARR | Source |
|---|---|---|
| Oct 2024 (≈8 mo post‑launch) | **~$20M** | Read the Signal / Sacra |
| End 2024 | **~$26M** | Sacra |
| ~Q3 2025 | **~$100M** (hit in 7 quarters) | Sierra blog / TBPN, Apr 2026 |
| End 2025 | **~$130–150M** (first $50M quarter) | Sierra "Year Two in Review" |
| Jan 2026 | **~$150M** | Sacra |

Implied trajectory: ~7.5× in 15 months (Oct '24 → Jan '26).

### Decagon (entire company is agentic)
| Date | ARR | Source |
|---|---|---|
| End 2024 | **~$10M** | Sacra |
| Apr 2025 | **~$17M** | Sacra |
| Oct 2025 | **~$35M** (annualized; Q3 ARR +3× YoY) | Sacra / SiliconANGLE |
| Jan 2026 | n/d at round, but $250M Series D at $4.5B | SiliconANGLE |

Implied trajectory: ~3.5× in 12 months; ~283% YoY.

### Intercom – Fin
| Date | Fin ARR | Source |
|---|---|---|
| Mar 2024 (Fin's first year) | **~$1M** ARR exiting | Mostly Metrics |
| Mar 2025 (Fin's second year) | **~$12M** ARR | Mostly Metrics / Lenny's Vault |
| Late 2025 | **~$100M** (approaching) | Complete AI Training, citing Intercom Berlin announcement Oct 2025 |
| Early 2026 | Fin projected to be ~50% of Intercom's ~$400M total ARR (~$200M run‑rate) | Mostly Metrics |

Note: Intercom company‑wide ARR re‑accelerated from $250M (2023) → $343M (2024) → ~$400M (early 2026) precisely because of Fin and outcome pricing.

### Zendesk – AI / Resolution Platform
| Date | AI ARR | Source |
|---|---|---|
| 2023 | **~$0** | Letter Two |
| 2024 | n/d (early ramp) | – |
| End 2025 | **~$200M** AI ARR, ~20,000 AI customers | Letter Two / Zendesk PR |
| 2026 (projected) | **~$400–500M** AI ARR (~150% YoY) | Letter Two / Permira note |

Reference point: Zendesk total revenue was **~$1.93B in 2024** (per Latka), so AI is going from 0% → ~10% of total in two years; Forethought acquisition (Mar 2026) is the bolt‑on accelerator.

### Side‑by‑side ARR snapshot

| Period | Sierra | Decagon | Intercom Fin | Zendesk AI |
|---|---|---|---|---|
| End 2023 | n/a (pre‑launch) | <$5M | ~$1M (just launched) | $0 |
| End 2024 | ~$26M | ~$10M | ~$12M (exit) | low single‑digit $M est. |
| End 2025 | ~$130–150M | ~$35M | ~$100M | ~$200M |
| 2026 run‑rate (early/projected) | $150M+ ($50M Q) | n/d (>$50M implied at $4.5B) | ~$200M | $400–500M |

---

## 3. Total tokens / interaction volume

There is no apples‑to‑apples public token disclosure across all four. Here is the closest contrast available:

| Vendor | Public token / volume signal | Inference |
|---|---|---|
| **Sierra** | Powers **2B+ monthly customer interactions** (per FourWeekMBA / Sierra). Multi‑provider router across 15+ models including a 70B‑parameter, 128k‑context proprietary "Agent LLM." Not on the OpenAI 1T‑tokens list (uses many providers + own models). | If avg interaction ≈ 4–8k tokens incl. tool use, implies **~10–15T+ tokens / month** across all providers — likely the largest in this set on a per‑interaction basis. |
| **Decagon** | Explicitly named in OpenAI DevDay 2025 "1T+ tokens" list. Handles "millions of customer conversations" for Notion, Duolingo, Substack et al. Mostly OpenAI today (GPT‑3.5 / 4 / 4o). | **>1T cumulative OpenAI tokens** disclosed, plus additional Anthropic / open‑model usage not reported. |
| **Intercom Fin** | ~**2M resolved customer queries / week** (~8–10M / month); 40M+ cumulative resolutions; primarily on Anthropic Claude 3.5 Sonnet since Oct 2024 (was OpenAI before). Not on OpenAI 1T list (consistent with the migration to Anthropic). | At ~5–10k tokens / resolution, **~50–100B tokens / month on Anthropic**, ~1–2T annualized. |
| **Zendesk AI** | ~**5B issues resolved annually** across the platform (only a fraction agentic today, growing fast). On OpenAI 1T list (GPT‑4o powering Agent Copilot, Resolution Platform now uses GPT‑5 + MCP). | **>1T cumulative OpenAI tokens** disclosed. Largest raw ticket footprint by far, but lower autonomous‑resolution mix than Sierra/Decagon/Fin today. |

### Token‑volume ranking (best estimate, monthly, all providers)
1. **Sierra** – highest per‑interaction token intensity (constellation routing, long context, voice growing).
2. **Zendesk AI** – highest absolute resolution count (5B/yr) but each automated resolution is shorter, RAG‑heavy.
3. **Intercom Fin** – ~8–10M monthly autonomous resolutions, Claude‑heavy.
4. **Decagon** – millions of conversations / month across ~150+ enterprise customers; >1T cumulative tokens on OpenAI alone.

---

## 4. Strategic contrast

- **Pure‑play vs. embedded.** Sierra and Decagon are **standalone agent platforms** with outcome‑priced revenue that *is* the AI ARR. Intercom Fin and Zendesk AI are **agent layers on incumbent CX suites**, where AI ARR is reported as a fast‑growing slice of a much larger seat‑based ARR base.

- **ARR per customer.** Decagon's median ACV is ~$400k and Sierra's land deals appear well into 7 figures (40% of Fortune 50); Fin lands at $0.99/resolution on top of ~$200/seat Intercom; Zendesk AI piggy‑backs on a ~$2B installed base with ~20,000 customers attached.

- **Model strategy.** Sierra is the most aggressive on **owning the stack** (proprietary 70B Agent LLM, multi‑provider router); Decagon and (now) Zendesk are deeply tied to **OpenAI** (both confirmed >1T‑token customers); Fin made a **defensive switch to Anthropic** in late 2024.

- **Growth ranking by ARR, end 2025**:  
  Zendesk AI ($200M) > Sierra (~$130–150M) > Intercom Fin (~$100M) > Decagon (~$35M).  
  But on **growth rate**, Sierra (>5×) and Decagon (~3.5×) are running ahead of Intercom Fin (~8×, off a smaller base) and Zendesk (which is growing AI ARR ~150% off a much larger ramp).

- **Risk asymmetry.** Sierra and Decagon's outcome pricing ties revenue directly to volume of work resolved, so any acceleration in voice + multi‑step workflows compounds ARR; Intercom Fin and Zendesk are exposed to seat‑erosion as autonomous agents reduce the number of human seats they have historically priced against.

---

## 5. Caveats

- ARR figures for private companies are reported numbers from journalism (Sacra, TBPN, The Information, mostlymetrics, Letter Two, SiliconANGLE) and have not been audited.
- Token figures for Sierra, Decagon and Fin are partly inferred from interaction counts; only Decagon and Zendesk are explicitly disclosed by OpenAI as 1T+ token customers.
- "AI ARR" definitions vary by vendor — Sierra and Decagon report only AI revenue; Intercom and Zendesk separate AI ARR from total subscription ARR.
