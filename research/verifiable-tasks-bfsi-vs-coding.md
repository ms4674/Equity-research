# Verifiable Tasks: BFSI AI Agents vs General Coding Agents

**Companion to:** [`data/bfsi_vs_coding_agent_verifiability.xlsx`](../data/bfsi_vs_coding_agent_verifiability.xlsx)
**Generator:** [`scripts/build_verifiability_workbook.py`](../scripts/build_verifiability_workbook.py)
**Date:** April 2026
**Status:** Educational / discussion only — **not investment advice**

---

## 1. Why "verifiable tasks" matter

Most of the recent step‑change in agent capability — OpenAI's o‑series, Claude 3.5/3.7 Sonnet on agentic coding, Cursor / Claude Code on real codebases — has come from training and evaluating against **verifiable tasks**: tasks where an automated oracle can produce a clean pass/fail signal at low cost. Compilers, type‑checkers, and unit‑test suites are the canonical example. They are what makes coding such an unusually fertile RL substrate.

This note compares the verifiability of agent tasks in **banking, payments, and finance (BFSI)** against the verifiability of **general coding agents**, using a shared rubric and a side‑by‑side scoring of 17 tasks in each domain.

The aggregated workbook contains:

| Sheet | Contents |
|---|---|
| `README` | Methodology, rubric, headline finding, disclaimer |
| `Verifiability_Rubric` | The 6‑dimension scoring framework |
| `BFSI_Tasks` | 17 finance‑domain tasks scored 1–5 on each dimension |
| `Coding_Tasks` | 17 general coding‑agent tasks scored 1–5 on each dimension |
| `SideBySide` | Aggregate comparison + a combined ranking across both domains |
| `Benchmarks` | Public benchmarks/oracles per domain (SWE‑bench, HumanEval, FinanceBench, FinQA, etc.) |
| `Sources` | Bibliography with URLs |

---

## 2. Verifiability rubric

Each task is scored 1 (low) to 5 (high) on six dimensions, with the following weights:

| Dimension | Weight | What it measures |
|---|---|---|
| GroundTruthAvailability | 0.25 | Does a single, agreed reference answer exist (or can it be cheaply produced)? |
| AutomatedOracleStrength | 0.25 | Can the check be run by a fast, deterministic program with no human in the loop? |
| OutputDeterminism | 0.15 | Is the correct output well‑defined and stable, or is there a wide manifold of acceptable answers? |
| MeasurementLatency | 0.10 | How quickly can a verdict be produced after the agent acts? |
| RegulatoryAuditability | 0.10 | Can the verification artifact serve as evidence to a regulator, auditor, or model‑risk committee? |
| RewardSignalCleanliness | 0.15 | Suitability as a training reward (sparse vs dense, low‑noise, low reward‑hacking surface) |

Composite = weighted sum (max 5.0).

---

## 3. Headline numbers

From running the scorer on 17 tasks per domain (full lists in the spreadsheet):

| Metric | BFSI agents | Coding agents | Delta (BFSI – Coding) |
|---|---|---|---|
| Number of tasks scored | 17 | 17 | — |
| Avg GroundTruthAvailability | 3.59 | 4.06 | −0.47 |
| Avg AutomatedOracleStrength | 3.35 | 4.12 | −0.77 |
| Avg OutputDeterminism | 3.41 | 3.47 | −0.06 |
| Avg MeasurementLatency | 3.35 | 4.29 | −0.94 |
| Avg RegulatoryAuditability | 3.76 | 2.47 | **+1.29** |
| Avg RewardSignalCleanliness | 3.24 | 4.06 | −0.82 |
| **Avg CompositeScore** | **3.44** | **3.85** | **−0.41** |
| Min CompositeScore | 1.10 | 1.00 | +0.10 |
| Max CompositeScore | 5.00 | 4.80 | +0.20 |

Two things jump out:

1. **General coding agents have a measurably higher average verifiability** (~0.4 points on a 1–5 scale, ~10%). This is mostly driven by `AutomatedOracleStrength`, `MeasurementLatency`, and `RewardSignalCleanliness` — exactly the dimensions where compilers and unit tests dominate.
2. **BFSI tasks beat coding tasks on `RegulatoryAuditability` by ~1.3 points.** When a BFSI task *is* verifiable, the verification artifact (a recomputed P&L, a re‑derived LCR, a passing ISO 20022 schema check) is also a regulator‑grade piece of evidence. Coding artifacts (tests passing) are useful internally but rarely audit‑grade.

---

## 4. Top‑5 most verifiable tasks per domain

### 4.1 BFSI agent tasks

| Rank | Task | Composite |
|---|---|---|
| 1 | Trade reconciliation (front‑to‑back match against custodian / clearing file) | 5.00 |
| 2 | Bank statement / GL reconciliation (entries to confirmations) | 5.00 |
| 3 | Payments rail formatting & validation (ISO 20022, SWIFT MT/MX, NACHA, SEPA) | 4.90 |
| 4 | Recompute regulatory ratios (LCR, NSFR, RWA, leverage ratio) | 4.50 |
| 5 | KYC document extraction (passport, utility bill, corporate registry) | 4.25 |

These are the **"finance unit tests"** of BFSI: schema‑driven, row‑level, deterministic, and regulator‑auditable. They are the natural first targets for verifier‑driven training and for production gating.

### 4.2 General coding agent tasks

| Rank | Task | Composite |
|---|---|---|
| 1 | Compile / typecheck a codebase | 4.80 |
| 2 | Run unit tests | 4.80 |
| 3 | Pass a coding‑interview problem (HumanEval / MBPP / LeetCode) | 4.70 |
| 4 | Implement an HTTP/gRPC API to a spec | 4.65 |
| 5 | SQL query authoring against a known schema (Spider / BIRD) | 4.65 |

These are the canonical RL‑friendly tasks behind SWE‑bench, HumanEval, MBPP, LiveCodeBench, Spider, and BIRD.

### 4.3 Bottom of each list

The least‑verifiable tasks in each domain:

- **BFSI:** Equity research / earnings note synthesis (1.25); SAR narrative drafting (~2.0); wealth advisor research synthesis (~2.2). All narrative‑heavy, with no automatable oracle and noisy human ratings.
- **Coding:** System design / architecture proposal (1.00); developer documentation (~2.0); frontend Figma‑to‑UI (~2.7). Same pattern — open‑ended, judgment‑laden, no reference answer.

The *lowest* score in the entire workbook is on the coding side (system design = 1.00), and the *highest* scores are on the BFSI side (reconciliations = 5.00). The interesting story is therefore not "coding > finance" but **"finance verifiability is bimodal"**.

---

## 5. Why BFSI verifiability is bimodal

Two opposing forces shape the BFSI distribution:

- **Pulling up:** finance has unusually strong machine‑readable structure when the task is operational. Double‑entry accounting, ISO 20022, FIX, SWIFT MT/MX, NACHA, SEPA, FpML, and book‑of‑record systems all give an agent a *schema* and a *recomputable ground truth*. BCBS 239 effectively requires banks to maintain the data lineage that makes recomputation possible. This is why reconciliation, payments formatting, and ratio recomputation score as high as anything in coding.
- **Pulling down:** finance also has unusually strong **judgment‑laden, narrative tasks** (SAR drafts, credit memos, equity research, advisor notes) where the "correct" output is undefined. These score below almost any coding task because there is no compiler, no tests, and no execution result.

By contrast, coding's distribution is much more clustered in the high‑verifiability range, because almost any code task can fall back on **execution + tests** as a partial oracle. Even the lowest‑verifiability coding tasks (system design, docs) sit alongside compiler‑gated artifacts in the same workflow, which lifts the practical signal available.

---

## 6. The `RegulatoryAuditability` flip

The single largest dimension‑level difference in favor of BFSI is `RegulatoryAuditability` (+1.29). This matters more than its 0.10 weight suggests, because:

1. Under **SR 11‑7** (Fed/OCC model‑risk guidance) and the **EU AI Act** high‑risk regime, banks need a *defensible verification record* for any model influencing decisions.
2. A reconciliation diff, an ISO 20022 schema validation, a re‑derived LCR, or a `terraform plan` with policy‑as‑code attached all qualify as audit artifacts. A passing pytest run typically does not, on its own, satisfy a model‑risk reviewer of an external decision's correctness.
3. Therefore in BFSI, **the same verifier doubles as the model‑risk audit artifact**. This compounds the value of building strong automated oracles for finance tasks: they pay back twice (RL signal + audit evidence).

This is one reason the `RegulatoryAuditability` weight is non‑zero in the rubric even though it is not strictly part of "can a model learn from this signal"; it is where the BFSI economics of verifier construction are largest.

---

## 7. Practical implications

For builders of BFSI agents:

1. **Industrialize the high‑verifiability quadrant first.** Reconciliation, payments formatting, KYC extraction, regulatory‑ratio recomputation, and chargeback‑evidence generation should be the first BFSI workflows wrapped in agent‑callable tools with strict oracle gating. They are where both production safety and RL signal are richest.
2. **Treat narrative finance tasks as RAG + grounding problems, not RL problems.** SAR drafting, advisor notes, and equity research will not get a clean reward function any time soon. The right verification target is **citation grounding** ("every claim is supported by a cited source") and **structured‑field completeness**, not output quality.
3. **Build BFSI evals that look like SWE‑bench, not like GLUE.** Public BFSI benchmarks (FinanceBench, FinQA, ConvFinQA, PIXIU/FinBen) are still mostly QA‑style. The next generation should look like *task suites with oracles* — apply this patch, recompute the LCR, file this dispute, post this journal entry — graded by deterministic checks. The closest existing analog is WebArena/OSWorld‑style goal predicates.
4. **Use BFSI verifiers as audit artifacts.** Every oracle you build should produce machine‑readable output that an auditor or MRM reviewer can replay. This is a cheap thing to design in upfront and an expensive thing to retrofit.

For builders of general coding agents working with BFSI customers:

1. **Coding agents already operate in the most‑verifiable regime in this workbook.** That is a *capability* moat, not just a benchmark moat: it explains why agentic coding has scaled so fast.
2. **The translation into BFSI is to wrap finance domain tools as code‑like APIs** (a `recompute_lcr(...)` tool, a `validate_iso20022(...)` tool) so that the agent can use the same verifier‑in‑the‑loop reasoning patterns it already uses for `pytest`.

---

## 8. Caveats

- Scores are author‑assigned on a 1–5 scale. They are intended as a *consistent rubric* across tasks, not as objective truth. The script makes it trivial to override scores or weights and re‑emit the workbook.
- "Verifiability" here is about whether a *post‑hoc* check exists. It is not the same as task **importance**, **value**, or **maturity** — those were addressed in the prior note ([`where-ai-agents-are-most-effective-in-bfsi.md`](where-ai-agents-are-most-effective-in-bfsi.md)).
- The benchmark list is illustrative, not exhaustive. Notably absent: closed industry benchmarks held by individual banks and vendors, and the rapidly growing set of agentic finance evals being built by Bloomberg, S&P, Patronus, and others.

---

## 9. Reproducing the workbook

```bash
pip install openpyxl
python3 scripts/build_verifiability_workbook.py
# -> writes data/bfsi_vs_coding_agent_verifiability.xlsx
```

Edit `RUBRIC` (weights), `BFSI_TASKS`, or `CODING_TASKS` in the script to change the analysis; the composite scores and ranking recompute on the next run.
