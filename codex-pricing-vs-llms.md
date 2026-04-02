# OpenAI Codex: Usage, Pricing & Comparison vs Frontier LLM Offerings

**Date:** April 2026

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [OpenAI Codex Overview](#openai-codex-overview)
3. [API Token Pricing Comparison](#api-token-pricing-comparison)
4. [Subscription & Seat-Based Plans](#subscription--seat-based-plans)
5. [Coding Agent Products](#coding-agent-products)
6. [GitHub Commit Activity & Market Adoption](#github-commit-activity--market-adoption)
7. [Benchmark Performance](#benchmark-performance)
8. [Cost-Efficiency Analysis](#cost-efficiency-analysis)
9. [Key Takeaways](#key-takeaways)

---

## Executive Summary

OpenAI Codex has evolved from a standalone code-completion model into a full-fledged coding agent integrated across ChatGPT subscriptions and the OpenAI API. This report compares Codex pricing and capabilities against competing frontier LLM offerings from Anthropic (Claude), Google (Gemini), GitHub (Copilot), and Cursor, covering both API token economics and subscription-based access models.

**Key findings:**
- On a per-token basis, Codex models (GPT-5.x-Codex) are **price-competitive with Gemini** and **significantly cheaper than Claude Opus** for output-heavy coding workloads.
- At the subscription level, ChatGPT Pro ($200/mo) competes directly with Claude Max ($100-200/mo) and Cursor Ultra ($200/mo) for power-user unlimited access.
- Benchmark performance across SWE-Bench Verified has converged (~80% for all frontier models), making **cost, context window, and tooling integration** the primary differentiators.
- **Claude Code dominates public GitHub commit share** at 5.78% (~20.8M cumulative commits), dwarfing Codex, Cursor, and Copilot in detected autonomous commits — though this gap is inflated by attribution methodology.
- **Google Antigravity is the newest entrant** (Nov 2025, free preview), powered by Gemini 3 Pro with a 1-2M context window, but has minimal detectable GitHub commit footprint so far.

---

## OpenAI Codex Overview

### What Is Codex?

OpenAI Codex is an AI coding agent available through ChatGPT and the OpenAI API. Originally launched as a specialized code-generation model, it has evolved into a comprehensive development tool powered by the GPT-5.x model family. Core capabilities include:

| Capability | Description |
|---|---|
| **Code Generation** | Produces code from natural language, adapting to project structure and conventions |
| **Code Understanding** | Reads and explains complex/legacy codebases |
| **Code Review** | Identifies bugs, logic errors, and unhandled edge cases |
| **Debugging** | Traces failures, diagnoses root causes, suggests fixes |
| **Task Automation** | Automates refactoring, testing, migrations, and DevOps workflows |

### March 2026: Plugin System

OpenAI introduced a plugin system for Codex in March 2026, adding:
- **Skills** — packaged workflows for automating narrow tasks
- **MCP integrations** — connections to version control, project management, cloud infrastructure, and enterprise databases
- **Pre-packaged connectors** — 12+ ready-made integrations including GitHub repos and Google Drive
- **Enterprise governance** — controls for data access, compliance, and AI behavior policies

### Access Channels

| Channel | Models Available | Billing Model |
|---|---|---|
| ChatGPT Plus/Pro/Team/Enterprise | GPT-5.4, GPT-5.3-Codex | Credits / subscription |
| OpenAI API (pay-as-you-go) | GPT-5.3-Codex, GPT-5.2-Codex, GPT-5.1-Codex, GPT-5.1-Codex-Mini | Per-token |

---

## API Token Pricing Comparison

### Per-Million-Token Rates (April 2026)

#### OpenAI Codex Models

| Model | Input (per 1M tokens) | Output (per 1M tokens) | Context Window |
|---|---|---|---|
| GPT-5.3-Codex | $1.75 | $14.00 | 256K |
| GPT-5.2-Codex | $1.75 | $14.00 | 256K |
| GPT-5-Codex | $1.25 | $10.00 | 256K |
| GPT-5.1-Codex | $1.25 | $10.00 | 256K |
| GPT-5.1-Codex-Mini | $0.25 | $2.00 | 128K |
| Codex Mini | $0.75 | $3.00 | 128K |

#### Anthropic Claude Models

| Model | Input (per 1M tokens) | Output (per 1M tokens) | Context Window |
|---|---|---|---|
| Claude Opus 4.6 | $5.00 | $25.00 | 200K |
| Claude Sonnet 4.6 | $3.00 | $15.00 | 200K-1M |
| Claude Haiku 4.5 | $1.00 | $5.00 | 200K |
| Claude Haiku 3.5 | $0.80 | $4.00 | 200K |

#### Google Gemini Models

| Model | Input (per 1M tokens) | Output (per 1M tokens) | Context Window |
|---|---|---|---|
| Gemini 3.1 Pro | $2.00 | $12.00 | 1M |
| Gemini 2.5 Pro | $1.25 | $10.00 | 1M |
| Gemini 2.5 Flash | $0.15 | $0.60 | 1M |

### Normalized Price Comparison (Flagship Models)

The following chart compares the flagship code-capable model from each provider:

```
Output Cost per 1M Tokens (USD)
═══════════════════════════════════════════════════════

Claude Opus 4.6      ████████████████████████████████████████████████  $25.00
Claude Sonnet 4.6    ██████████████████████████████                    $15.00
GPT-5.3-Codex        ████████████████████████████                      $14.00
Gemini 3.1 Pro       ████████████████████████                          $12.00
GPT-5-Codex          ████████████████████                              $10.00
Gemini 2.5 Pro       ████████████████████                              $10.00
Claude Haiku 4.5     ██████████                                         $5.00
Codex Mini           ██████                                             $3.00
GPT-5.1-Codex-Mini   ████                                               $2.00
Gemini 2.5 Flash     █                                                  $0.60
```

### Input-to-Output Cost Ratio

All providers charge significantly more for output tokens than input tokens, reflecting the higher compute cost of generation:

| Model | Output : Input Ratio |
|---|---|
| Claude Opus 4.6 | 5:1 |
| Claude Sonnet 4.6 | 5:1 |
| GPT-5.3-Codex | 8:1 |
| GPT-5-Codex | 8:1 |
| Gemini 2.5 Pro | 8:1 |
| Gemini 3.1 Pro | 6:1 |
| GPT-5.1-Codex-Mini | 8:1 |
| Gemini 2.5 Flash | 4:1 |

Codex and Gemini models carry a higher output:input ratio (8:1) compared to Claude (5:1), meaning output-heavy workloads (code generation) disproportionately favor Claude on a per-token basis, while input-heavy workloads (code review, analysis) favor Codex and Gemini.

---

## Subscription & Seat-Based Plans

### Individual Developer Plans

| Plan | Monthly Cost | Key Coding Features |
|---|---|---|
| **ChatGPT Plus** | $20 | GPT-5.4 Codex, 80-100 msgs/3hr, basic agentic features |
| **ChatGPT Pro** | $200 | Unlimited GPT-5.4, full autonomous coding, 256K context |
| **Claude Pro** | $20 | Claude Code access, rate-limited, Sonnet 4 default |
| **Claude Max** | $100-200 | 5-20x Pro limits, Opus 4 access at $200 tier |
| **Cursor Pro** | $20 | Unlimited tab completions, $20 usage credits |
| **Cursor Pro+** | $60 | 3x model usage on all models |
| **Cursor Ultra** | $200 | 20x model usage, priority feature access |
| **Copilot Pro** | $10 | Unlimited completions, unlimited chat, 300 premium reqs |
| **Copilot Pro+** | $39 | 1,500 premium reqs, all models including Opus 4.6 |

### Team & Enterprise Plans

| Plan | Per-User/Month | Highlights |
|---|---|---|
| **ChatGPT Team** | $25-30 | Plus features + collaboration, data not used for training |
| **ChatGPT Enterprise** | ~$50-60 (custom) | Unlimited, SSO, SCIM, compliance |
| **Cursor Teams** | $40 | Shared chats, RBAC, SSO |
| **Cursor Enterprise** | Custom | Pooled usage, audit logs, SCIM |
| **Copilot Business** | $19 | 300 premium reqs, centralized mgmt |
| **Copilot Enterprise** | $39 | 1,000 premium reqs, full enterprise features |

### Value-for-Money Positioning

At the **$20/mo entry tier**, all four products (ChatGPT Plus, Claude Pro, Cursor Pro, Copilot Pro) provide basic coding AI access, but Copilot Pro at $10/mo is the most affordable entry point with unlimited completions.

At the **power-user tier** ($100-200/mo), ChatGPT Pro, Claude Max, and Cursor Ultra compete for developers who need unrestricted high-volume usage. ChatGPT Pro offers unlimited GPT-5.4 access for $200/mo; Claude Max offers comparable usage at $100-200/mo depending on whether Opus 4 access is needed.

---

## Coding Agent Products

### Feature Comparison

| Feature | OpenAI Codex | Claude Code | GitHub Copilot | Cursor |
|---|---|---|---|---|
| **Inline completion** | Via ChatGPT | No (terminal-based) | Yes (IDE-native) | Yes (IDE-native) |
| **Multi-file editing** | Yes | Yes | Yes (Copilot Edits) | Yes (Agent mode) |
| **Autonomous agent mode** | Yes | Yes | Yes (Agent mode) | Yes (Agent mode) |
| **Terminal/CLI agent** | Via plugins | Yes (native) | Limited | Yes |
| **Plugin/MCP support** | Yes (March 2026) | Yes | Limited | Yes |
| **Git integration** | Via plugins | Native | Native | Native |
| **IDE integration** | ChatGPT web/desktop | Terminal + IDE extensions | VS Code, JetBrains, etc. | Custom IDE (VS Code fork) |
| **Codebase indexing** | Yes | Yes | Yes (repo-level) | Yes |
| **Context window** | 256K tokens | 200K-1M tokens | Varies by model | Varies by model |
| **Enterprise governance** | Yes | Yes | Yes | Yes |

### Differentiation

- **Codex** excels in autonomous multi-step workflows and terminal/DevOps tasks; new plugin system provides the broadest external integration story.
- **Claude Code** provides the most natural terminal-native coding experience; Sonnet 4.6's 1M context window enables full-repository reasoning at moderate cost.
- **GitHub Copilot** offers the deepest IDE integration and lowest entry cost ($10/mo); strongest value for completion-heavy workflows.
- **Cursor** provides the most polished AI-native IDE experience with multi-model support; best for developers wanting model flexibility.

---

## GitHub Commit Activity & Market Adoption

A critical measure of real-world traction — distinct from benchmarks — is how much code each agent actually ships. Public GitHub commit tracking provides the most transparent (if imperfect) signal of adoption.

### Detected Public GitHub Commit Share (March 2026)

Sources: [Powerset GitHub Coding Agent Tracker](https://github.com/powerset-co/github-coding-agent-tracker), [claudescode.dev](https://www.claudescode.dev), [botcommits.dev](https://botcommits.dev), CoreMention.

| Agent | Share of Public Commits (10-day avg) | Estimated Daily Commits | Total Commits (cumulative) |
|---|---|---|---|
| **Claude Code** | **5.78%** (~9.7% by late March) | 326,000-377,000 | **~20.8 million** (since Feb 2025 launch) |
| **GitHub Copilot** | 0.24% | ~13,500* | Not disclosed |
| **OpenAI Codex** | <0.24%** | Low (tracked via `chatgpt-codex-connector[bot]`) | Not disclosed |
| **Cursor** | <0.24%** | Low (tracked via `cursoragent@cursor.com`) | Not disclosed |
| **Google Jules** | 0.05% | ~2,800* | Not disclosed |
| **Google Antigravity** | Not detected*** | Unknown | Not tracked |

*Estimated from share × ~5.6M daily public GitHub commits. **Specific percentages not broken out in public tracker data but listed below Claude Code and Copilot. ***Antigravity does not appear to leave detectable commit signatures in public tracking data.

```
Share of Public GitHub Commits (March 2026)
═══════════════════════════════════════════════════════

Claude Code       ██████████████████████████████████████████████████████████  5.78%
GitHub Copilot    ██                                                          0.24%
Google Jules      ▌                                                           0.05%
Codex             ▏                                                          <0.24%
Cursor            ▏                                                          <0.24%
Antigravity       (not detected)                                              —
```

### Claude Code: The Dominant Commit Generator

Claude Code's commit footprint is an order of magnitude larger than any competitor:

| Metric | Value |
|---|---|
| Total commits since launch (Feb 2025) | 20,807,124 |
| Active repositories | 1,087,408 |
| Lines of code added | 50.4 billion |
| Lines of code deleted | 19.8 billion |
| New repos per week (first Claude Code commit) | ~114,785 |
| Week-over-week growth | +8% |
| Doubling time | 61 days |
| Peak daily commits (late March 2026) | 377,162 |

**Top languages by Claude Code commits:**

| Language | Share | Total Commits |
|---|---|---|
| TypeScript | 34.8% | 7.25M |
| Python | 18.9% | 3.92M |
| JavaScript | 10.2% | 2.13M |

**Quality caveat:** ~90% of Claude Code commits land in repositories with fewer than 2 GitHub stars, indicating heavy personal/experimental use rather than high-impact production code.

### Why the Disparity Is Misleading

The massive gap between Claude Code and other agents in public commit data reflects **detection methodology**, not necessarily true market share:

| Factor | Impact |
|---|---|
| **Commit signature attribution** | Claude Code leaves a `Co-authored-by: Claude` trailer by default; Copilot, Cursor, and Antigravity often leave no detectable signature |
| **Inline completions vs. agent commits** | Copilot and Cursor primarily operate via inline completions that blend into human commits — invisible to trackers |
| **Public vs. private repos** | Only ~18.5% of GitHub activity is public; enterprise-heavy tools (Copilot, Codex) skew toward private repos |
| **Agent mode maturity** | Codex's autonomous agent mode launched later than Claude Code; Antigravity launched Nov 2025 |

**Bottom line:** Claude Code genuinely leads in autonomous commit generation, but the 5.78% vs. <0.24% gap overstates the true difference in code influence because Copilot, Cursor, and Codex generate enormous volumes of code that is committed under human authorship.

### Google Antigravity: Early but Growing

Google Antigravity (launched November 2025) is the newest entrant. Key adoption signals:

| Metric | Value |
|---|---|
| Pricing | Free (public preview) |
| Powered by | Gemini 3 Pro (1-2M context window) |
| GitHub ecosystem | 73+ public repos, 28,000+ stars across projects |
| Attention growth | +244.4% (from near-zero baseline) |
| SWE-Bench Verified | 76.2% |
| Task completion (first attempt) | 72% |
| Built on | Google's $2.4B Windsurf acquisition |

Antigravity does not appear in public commit trackers, likely because: (1) it operates primarily as a browser-based cloud IDE, (2) commits may not carry detectable agent signatures, and (3) the product is still in early public preview with rate-limited access.

### Projection: AI-Authored Commits by End of 2026

SemiAnalysis analyst Dylan Patel projects Claude Code alone will exceed **20% of all daily GitHub commits by end of 2026**, driven by 8% week-over-week growth.

Combined with Copilot, Cursor, Codex, and Antigravity, **total AI-assisted code is estimated to already represent 42% of new code on GitHub** as of early 2026, though only a fraction carries detectable attribution.

### Developer Satisfaction Rankings

Complementing commit volume, developer satisfaction surveys show:

| Agent | Developer Satisfaction |
|---|---|
| Claude Code | 46% |
| Cursor | 19% |
| GitHub Copilot | 9% |

### PR Merge Rates (Quality Signal)

Pull request acceptance rates across agents (from PR Arena):

| Agent | PR Merge Rate |
|---|---|
| Cursor Agents | 96.1% |
| GitHub Copilot | 95.6% |
| Google Jules | 93.6% |
| OpenAI Codex | 87.0% |
| Devin | 61.2% |

Cursor and Copilot lead in PR merge rates, suggesting higher-quality code output in agentic PR workflows, though Codex's lower rate may reflect more ambitious autonomous tasks.

---

## Benchmark Performance

### SWE-Bench Verified (Real-World Bug Fixing)

SWE-Bench Verified is the industry-standard benchmark measuring an AI's ability to resolve real GitHub issues end-to-end.

| Model | Score | Notes |
|---|---|---|
| GPT-5.3 Codex | ~80-85% | Leading on certain task subsets |
| Claude Opus 4.6 | 80.8% | First model to break 80% (early 2026) |
| Gemini 3.1 Pro | 75-80.6% | Strong given price advantage |

### Terminal-Bench 2.0 (CLI/DevOps)

| Model | Score |
|---|---|
| GPT-5.3 Codex | 77-82% |
| Gemini 3.1 Pro | 68.5% |
| Claude Opus 4.6 | 65.4% |

Codex dominates terminal and DevOps workflows by a significant margin (12-16 points over competitors).

### LiveCodeBench (Competitive Programming)

| Model | Score |
|---|---|
| Gemini 3.1 Pro | 75.6% |
| DeepSeek V3.2 | 74.1% |
| Claude Opus 4.6 | ~62% |

Gemini leads competitive programming tasks, while Claude lags significantly.

---

## Cost-Efficiency Analysis

### Scenario: Typical Daily Coding Session

Assumptions: 1-hour focused coding session generating ~50K output tokens and consuming ~200K input tokens (context + code).

| Provider / Model | Input Cost | Output Cost | **Total** |
|---|---|---|---|
| GPT-5.3-Codex | $0.35 | $0.70 | **$1.05** |
| GPT-5-Codex | $0.25 | $0.50 | **$0.75** |
| GPT-5.1-Codex-Mini | $0.05 | $0.10 | **$0.15** |
| Claude Sonnet 4.6 | $0.60 | $0.75 | **$1.35** |
| Claude Opus 4.6 | $1.00 | $1.25 | **$2.25** |
| Gemini 2.5 Pro | $0.25 | $0.50 | **$0.75** |
| Gemini 3.1 Pro | $0.40 | $0.60 | **$1.00** |
| Gemini 2.5 Flash | $0.03 | $0.03 | **$0.06** |

### Monthly Cost Projection (20 Working Days)

| Provider / Model | Monthly API Cost | Best Subscription Alternative |
|---|---|---|
| GPT-5.3-Codex | ~$21 | ChatGPT Plus ($20) — breakeven |
| Claude Sonnet 4.6 | ~$27 | Claude Pro ($20) — subscription wins |
| Claude Opus 4.6 | ~$45 | Claude Max ($100) — API wins until heavy use |
| Gemini 2.5 Pro | ~$15 | No subscription equivalent (free tier available) |
| GPT-5.1-Codex-Mini | ~$3 | API is far cheaper than any subscription |

### Cost per SWE-Bench Point

Normalizing cost against SWE-Bench Verified performance (proxy for coding quality):

| Model | Output $/1M | SWE-Bench % | $/1M per SWE-Bench Point |
|---|---|---|---|
| GPT-5.3-Codex | $14.00 | ~82% | $0.17 |
| Claude Opus 4.6 | $25.00 | 80.8% | $0.31 |
| Claude Sonnet 4.6 | $15.00 | ~75%* | $0.20 |
| Gemini 2.5 Pro | $10.00 | ~76% | $0.13 |
| Gemini 3.1 Pro | $12.00 | ~78% | $0.15 |
| GPT-5.1-Codex-Mini | $2.00 | ~65%* | $0.03 |

*Estimated scores for non-flagship variants.

**Gemini 2.5 Pro offers the best cost-per-performance ratio** among flagship models, while **GPT-5.1-Codex-Mini is unmatched for budget-conscious workloads** where peak accuracy is not required.

---

## Key Takeaways

### 1. Codex Is Price-Competitive at the API Level
GPT-5.3-Codex ($1.75/$14.00) is priced between Gemini 3.1 Pro ($2.00/$12.00) and Claude Sonnet 4.6 ($3.00/$15.00). For output-heavy code generation, Codex is 44% cheaper than Claude Opus and roughly on par with Gemini.

### 2. Subscription Models Favor ChatGPT Plus for Light Users
At $20/mo, ChatGPT Plus with Codex access is competitive with Claude Pro and Cursor Pro. GitHub Copilot Pro at $10/mo remains the cheapest entry point for developers who primarily need inline completions.

### 3. Power Users Face a Three-Way Choice
At the $200/mo tier, developers choose between ChatGPT Pro (unlimited GPT-5.4), Claude Max $200 (Opus 4 access with 20x limits), and Cursor Ultra (multi-model flexibility with 20x usage). The right choice depends on workflow preference: web-based (Codex), terminal-native (Claude Code), or IDE-native (Cursor).

### 4. Benchmark Convergence Makes Cost the Differentiator
With SWE-Bench Verified scores clustering around 80% for all frontier models, raw coding ability is no longer a meaningful differentiator. Cost-efficiency, context window size (Gemini's 1M tokens is 4x Codex's 256K), tooling integration, and workflow fit now drive the decision.

### 5. Codex's Plugin System Is a Strategic Moat
The March 2026 plugin system gives Codex the broadest integration surface across external tools and enterprise systems, positioning it as the most extensible coding agent for complex enterprise workflows.

### 6. Budget-Tier Models Are Disruptively Cheap
GPT-5.1-Codex-Mini ($0.25/$2.00) and Gemini 2.5 Flash ($0.15/$0.60) enable coding assistance at 10-20x lower cost than flagship models, making AI-assisted coding economically viable for high-volume, lower-stakes tasks like boilerplate generation, test writing, and documentation.

### 7. Claude Code Leads Autonomous Commit Volume by an Order of Magnitude
With 20.8M cumulative commits and 5.78% of public GitHub activity, Claude Code generates more detected autonomous commits than Codex, Cursor, Copilot, and Antigravity combined. However, this reflects Claude Code's default commit attribution (`Co-authored-by: Claude`) more than a true 20:1 usage gap — Copilot and Cursor generate substantial code that is committed under human authorship without agent signatures.

### 8. Google Antigravity Is a Wild Card
Free, powered by Gemini 3 Pro with a 1-2M context window, and backed by Google's $2.4B Windsurf acquisition, Antigravity has compelling positioning but minimal market traction so far. Its browser-based cloud IDE approach and lack of commit attribution make it invisible in public tracking data. The +244.4% attention growth (from a near-zero base) suggests early interest but not yet meaningful adoption.

### 9. 42% of GitHub Code Is Already AI-Assisted
Combining all agents (detected and undetected), approximately 42% of new code on GitHub involves AI assistance as of early 2026. Only 48% of AI-generated code receives human review before merge, raising quality and security questions — AI-assisted commits leak secrets at 2x the rate of human-only commits.

---

*Sources: OpenAI, Anthropic, Google, GitHub, and Cursor official pricing pages; Powerset GitHub Coding Agent Tracker; claudescode.dev; botcommits.dev; CoreMention; SemiAnalysis; PR Arena leaderboard (accessed April 2026). Benchmark data from SWE-Bench Verified, Terminal-Bench 2.0, and LiveCodeBench leaderboards.*
