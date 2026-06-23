# Aggregated US Government Spending → Beneficiary Companies (Quantum & PQC/Cybersecurity)

**Companion data file to:** `research/2026-06-23-trump-quantum-cyber-eo-impact.md`
**Compiled:** June 23, 2026
**Scope:** Disclosed US federal funding (grants, equity-stake incentives, contracts, LOIs, IDIQ vehicles) flowing to quantum computing and post-quantum cryptography (PQC) / cybersecurity companies, with public-ticker mapping.

> **Disclaimer.** Informational only; not investment advice. Figures are nominal headline values from public sources and frequently differ from *obligated* (actually funded) dollars — see the IonQ note in §4. "Up to" amounts are ceilings/LOIs subject to definitive agreements and appropriations. Always reconcile against company filings (10-K/10-Q), USASpending.gov, and FPDS before relying on any number.

---

## 1. Headline programs driving the spend

| Program / vehicle | Sponsor | Size | Date | Mechanism | Notes |
|---|---|---|---|---|---|
| **CHIPS & Science quantum LOIs** | Dept. of Commerce / NIST | **$2.013B** across 9 cos. | May 21, 2026 | Grants **+ minority, non-controlling equity stake** in each recipient | Largest single federal intervention in quantum to date; extends the Intel-style equity model (Aug 2025) to quantum |
| **EO 14411** "Ushering In the Next Frontier of Quantum Innovation" | White House | Directive (no $ attached) | Jun 22, 2026 | QC-ADDS program; advance market commitments (AMCs); 2028 science-grade computer + sensor targets | Demand-pull; DOE specs due +90d, partnership models +180d |
| **EO 14409** "Securing the Nation Against Advanced Cryptographic Attacks" | White House | Directive (no $ attached) | Jun 22, 2026 | PQC deadlines (key est. 2030 / signatures 2031) + **FAR contractor mandate** | Converts PQC into a market-wide procurement requirement |
| **OMB/ONCD government-wide PQC migration** | OMB / ONCD | **~$7.1B** (2025–2035, 2024 $) | Est. 2024, ongoing | Annual agency inventories + budget justifications | Excludes classified/National Security Systems (DoD/IC estimate separately) |
| **DARPA Quantum Benchmarking Initiative (QBI), Stage B** | DARPA (DoW) | **$1M–$5M per award**, 11 cos. | Nov 6, 2025 | Staged R&D validation toward "utility-scale by 2033" | Stage A → B down-select; not a procurement competition |
| **CISA cybersecurity (FY26)** | DHS / CISA | **~$1.367B** for FCEB protection | FY2026 budget | Appropriated cyber defense for federal civilian networks | Broader than PQC; funds CDM, operations, services |

---

## 2. CHIPS & Science Act — quantum LOIs ($2.013B, May 21, 2026)

Every recipient grants Commerce a **minority, non-controlling equity stake** as a condition of funding.

| Company | Public ticker | Planned funding | Modality / purpose | Investability |
|---|---|---|---|---|
| **IBM** | IBM (NYSE) | **$1,000M** | Domestic quantum **foundry** | Public, mega-cap (immaterial to EPS) |
| **GlobalFoundries** | GFS (Nasdaq) | **$375M** | Secure multi-modality quantum **foundry**; ~1% govt equity | Public |
| **D-Wave Quantum** | **QBTS** (NYSE) | **$100M** | Annealing + gate-model superconducting; would issue $100M common stock to Commerce | Public pure-play |
| **Rigetti Computing** | **RGTI** (Nasdaq) | **up to $100M** | Next-gen superconducting (readout electronics, cryostats) | Public pure-play |
| **Infleqtion** | **INFQ** (Nasdaq) | **$100M** | Neutral-atom systems integration | Public (SPAC, Feb 2026) |
| **Quantinuum** | (Honeywell **HON**; **IPO filed**) | **$100M** | Fault-tolerant trapped-ion (QCCD) | Majority Honeywell; IPO pending |
| **PsiQuantum** | Private | **$100M** | Photonic, fault-tolerant | Not directly investable |
| **Atom Computing** | Private | **$100M** | Neutral-atom (tens of thousands of qubits) | Not directly investable |
| **Diraq** | Private | **up to $38M** | Silicon-spin (CMOS) qubits | Not directly investable |
| **Total** | | **$2.013B** | | 5 of 9 recipients have public exposure |

**Equity-day stock reaction (May 21, 2026):** INFQ +31.4%, D-Wave (QBTS) +33%, Rigetti (RGTI) +30.6%; non-recipients also rallied: IonQ (IONQ) +12.3%, Quantum Computing Inc. (QUBT) +19%.

---

## 3. DARPA QBI — Stage B cohort (Nov 6, 2025; $1M–$5M/award)

11 companies advanced; public-ticker mapping:

| Company | Ticker | In CHIPS LOI? |
|---|---|---|
| IBM | IBM | Yes ($1B foundry) |
| IonQ | IONQ | No |
| Quantinuum | HON (IPO filed) | Yes ($100M) |
| Atom Computing | Private | Yes ($100M) |
| Diraq | Private | Yes ($38M) |
| QuEra Computing | Private | No |
| Nord Quantique | Private (Canada) | No |
| Photonic Inc. | Private (Canada) | No |
| Quantum Motion | Private (UK) | No |
| Silicon Quantum Computing | Private (Australia) | No |
| Xanadu | Private (Canada) | No |

- **Final-phase US2QC (≈ QBI Stage C):** Microsoft (MSFT), PsiQuantum (private).
- **Stage A but did not advance:** Rigetti (RGTI), HPE (HPE), Atlantic Quantum (private), Oxford Ionics — *Oxford Ionics acquired by IonQ Sept 2025 for ~$1.075B.*

---

## 4. Company-level federal contract detail (public names)

### IonQ (IONQ)
| Award / vehicle | Sponsor | Headline value | Obligated / note |
|---|---|---|---|
| Quantum networking contract | AFRL | **$54.5M ceiling** (4-yr, Sep 2024) | **Only ~$11.99M obligated**; earmark left **unfunded in FY2026** (zeroed in FY2025 budget) |
| Pentagon (AFRL et al.) FY22–24 | AFRL | ~$51M total obligated | Earmark-driven |
| SHIELD IDIQ (missile defense) | MDA | $151B program **ceiling** (IDIQ pool) | IonQ added as eligible vendor 2026; not a direct award |
| Oxford Ionics acquisition | — | ~$1.075B (Sep 2025) | M&A, not federal |
| SkyWater (SKYT) acquisition | — | ~$1.8B | Onshore quantum foundry |

> **Caveat (important):** IonQ's FY2025 revenue was **$130M** (>60% commercial); 2026 guidance **$225–245M**. Headline government "contract" figures (e.g., the $54.5M AFRL number) materially overstate funded dollars — a short-seller analysis flagged the obligated AFRL amount at ~$12M and noted FY26 de-funding. Treat all "ceiling" numbers with skepticism.

### Rigetti (RGTI)
| Award | Sponsor | Value | Date |
|---|---|---|---|
| Superconducting quantum networking (w/ QphoX) | AFRL | **$5.8M** (3-yr) | Sep 2025 |
| ABAA chip-fab consortium (incl. LLNL) | AFOSR | **$5.48M** | Apr 2025 |
| CHIPS LOI | Commerce/NIST | **up to $100M** | May 2026 |

### D-Wave Quantum (QBTS)
| Item | Counterparty | Value / note |
|---|---|---|
| CHIPS LOI | Commerce/NIST | **$100M** (issues $100M common stock to Commerce) |
| Los Alamos National Lab | DOE/LANL | Ongoing research collaboration |
| Public-sector distribution | Carahsoft | Reseller channel into federal/SLED |
| Recent bookings | — | Q bookings **$33.4M (+1,994% y/y)** |

### PQC / Cybersecurity names
| Company | Ticker | Federal linkage | Disclosed value |
|---|---|---|---|
| **SandboxAQ** | Private | 5-yr DoW CIO agreement (AQtive Guard, ACDI) Dec 2025; DISA QRC PKI prototype; USAF quantum-nav SBIR (2022) | DoW value **undisclosed**; SBIR **$1.2M** |
| **Cloudflare** | NET | First SASE with full-platform PQC (hybrid ML-KEM in TLS/IPsec); FedRAMP footprint | Productized; no single award |
| **Palo Alto Networks** | PANW | 7+ PQC ciphersuites in NGFW; large federal install base | Flows via FAR mandate / refresh |
| **IBM** | IBM | Co-author of NIST PQC (ML-KEM/ML-DSA), Quantum Safe consulting; + $1B foundry LOI | See §2 |
| **CrowdStrike / Zscaler / Fortinet / Cisco / SentinelOne** | CRWD / ZS / FTNT / CSCO / S | Federal platform vendors that must ship PQC to retain eligibility | Indirect / retention |
| **SEALSQ** | LAES | PQC semiconductor / secure-element pure-play | High-beta small cap |
| **DigiCert / Entrust / Thales / QuSecure / Fortanix / SafeLogic** | Private | Purest PKI / HSM / crypto-discovery beneficiaries | **Not directly investable** |

---

## 5. PQC migration market sizing (demand side)

| Segment | Estimate | Source basis |
|---|---|---|
| Government-wide migration (FCEB, excl. NSS) | **~$7.1B** (2025–2035, 2024 $) | OMB/ONCD annual agency cost rollup |
| Small agency | $5–20M each | CISA/agency budget justifications |
| Medium agency | $50–200M each | " |
| Large agency (DoD, VA, Treasury) | $500M–$2B+ each | " |
| Classified / NSS (DoD + IC) | Separate, undisclosed | NSM-10 / CNSA 2.0 track (faster) |

**Procurement gating already live:** CISA published an initial list of **PQC-capable product categories effective Jan 23, 2026** — agencies must procure PQC-capable products in those categories; EO 14409's FAR rule (proposed within 180 days of Jun 22, 2026) extends this to **all covered contractors by end-2030.**

---

## 6. Key takeaways for investors

1. **The biggest, most concrete dollars are upstream and partly private.** The $2.013B CHIPS package is the clearest "spending → company" line, but **4 of 9 recipients are private** (Atom, Diraq, PsiQuantum) plus Quantinuum (HON/IPO). Public pure-play capture: **QBTS, RGTI, INFQ** ($100M each, RGTI "up to").
2. **Headline ≠ obligated.** IonQ's experience (a $54.5M ceiling that funded ~$12M and was later zeroed) is the cautionary template — discount "ceiling," "IDIQ," and "up to" figures heavily.
3. **PQC is a bigger, more durable pool (~$7.1B+ over a decade) but leaks to private vendors.** Public investors capture it mainly **indirectly** via platform vendors (NET, PANW, IBM, CRWD, ZS, FTNT, CSCO) and via likely **M&A of private PQC tooling**.
4. **Watch the conversion catalysts:** definitive CHIPS agreements (from LOIs), the FAR proposed rule (~Dec 2026), DOE QC-ADDS specs/partnership models, the Quantinuum IPO, and annual agency PQC budget submissions.

---

### Sources
- NIST / Dept. of Commerce CHIPS quantum LOI release (May 21, 2026); FedScoop; Manufacturing Dive; TechTimes
- DARPA QBI Stage B selection page (Nov 6, 2025); ExecutiveGov; postquantum.com; entangledfuture.com QBI tracker
- IonQ Q4/FY2025 results (Feb 25, 2026); IonQ AFRL release (Sep 25, 2024); Wolfpack Research (obligation analysis)
- Rigetti AFRL ($5.8M) & AFOSR ($5.48M) releases; D-Wave CHIPS/LANL/Carahsoft disclosures
- SandboxAQ DoW CIO release (Dec 10, 2025); USAF SBIR; OMB/ONCD $7.1B PQC estimate (Federal News Network; 2024 ONCD report); CISA FY26 Congressional Justification; White House EO 14409 / 14411 + fact sheets
