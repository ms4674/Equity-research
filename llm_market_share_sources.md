# LLM Market Share: Non-OpenRouter Sources Showing Similar Charts/Data

This document catalogs company slides, reports, and datasets that show LLM market share breakdowns similar to the OpenRouter enterprise chart, from sources **other than OpenRouter**.

---

## 1. Menlo Ventures — "State of Generative AI in the Enterprise"

**The closest match to the original OpenRouter chart.** Menlo Ventures publishes annual reports with stacked bar charts showing enterprise LLM API market share by provider across 2023, 2024, and 2025. Their data is based on surveys of 150+ enterprise technical leaders.

### Market Share Data (Enterprise LLM API Usage)

| Provider   | 2023 | 2024 | Mid-2025 | End-2025 |
|------------|------|------|----------|----------|
| OpenAI     | 50%  | 34%  | 25%      | 27%      |
| Anthropic  | 12%  | 24%  | 32%      | 40%      |
| Google     | 7%   | 16%  | 20%      | 21%      |
| Meta       | 16%  | 12%  | 9%       | 8%       |
| Mistral    | 6%   | 5%   | —        | —        |
| Cohere     | 3%   | 3%   | —        | —        |
| DeepSeek   | —    | —    | 1%       | 1%       |
| Other      | 6%   | 6%   | 13%      | 3%       |

### Coding Market Share (Mid-2025)

| Provider   | Share |
|------------|-------|
| Anthropic  | 42-54%|
| OpenAI     | 21%   |
| Google     | 11-16%|
| Meta       | 8-10% |
| Other      | 6-11% |

### Links
- **2025 Mid-Year Report**: https://menlovc.com/2025-mid-year-llm-market-update/
- **2025 Annual Report**: https://menlovc.com/2025-the-state-of-generative-ai-in-the-enterprise/
- **2024 Annual Report**: https://menlovc.com/2024-the-state-of-generative-ai-/
- **2023 Annual Report**: https://menlovc.com/2023-the-state-of-generative-ai-in-the-enterprise-report/
- **Mid-Year PDF**: https://menlovc.com/wp-content/uploads/2025/11/menlo_ventures_mid-year_llm_report-2025.pdf
- **Annual PDF**: https://menlovc.com/wp-content/uploads/2025/12/menlo_ventures_enterprise_ai_report-2025.pdf

---

## 2. a16z (Andreessen Horowitz) — "Enterprise AI Arms Race" CIO Survey

a16z surveyed 100 Global 2000 CIOs (VP+ level, companies >$500M revenue). This focuses on **enterprise production adoption** and **wallet share**, not token volume.

### Enterprise Production Adoption Rates

| Provider   | In Production | Including Testing |
|------------|--------------|-------------------|
| OpenAI     | 78%          | ~85%              |
| Anthropic  | 44%          | 63%               |
| Google     | Not specified| —                 |

### Enterprise Wallet Share
- **OpenAI**: ~56% (but declining)
- **Anthropic**: Growing fastest (25% share increase since May 2025)
- **Google Gemini**: Steadily gaining

### Key Findings
- 81% of enterprises use 3+ model families in testing/production
- Anthropic leads coding; OpenAI leads chatbots and knowledge management
- Microsoft dominates enterprise apps (M365 Copilot, GitHub Copilot)
- Enterprises expect ~65% AI spend growth in 2026

### Links
- **Report**: https://a16z.com/leaders-gainers-and-unexpected-winners-in-the-enterprise-ai-arms-race/
- **CIO Survey**: https://a16z.com/ai-enterprise-2025

---

## 3. Ramp — AI Index (Corporate Card Transaction Data)

Ramp tracks AI adoption via corporate card and bill pay data from 50,000+ U.S. businesses. Different methodology from API usage — measures **business purchasing** of AI tools.

### Business AI Adoption Rates (March 2026)

| Provider   | Adoption Rate | Trend         |
|------------|--------------|---------------|
| OpenAI     | 34.4%        | -1.5% MoM     |
| Anthropic  | 24.4%        | +4.9% MoM     |
| Google     | 4.5%         | Understated*  |
| xAI (Grok) | 1.5%        | Growing        |
| DeepSeek   | <1%          | —              |

*Google AI spending often goes through existing Cloud contracts, not corporate cards.

### First-Time Enterprise AI Buyers
- **Anthropic**: ~73% of first-time spending (March 2026)
- **OpenAI**: ~27% of first-time spending

### Links
- **AI Index Dashboard**: https://ramp.com/data/ai-index
- **March 2026 Report**: https://ramp.com/velocity/ai-index-march-2026
- **February 2026 Report**: https://ramp.com/velocity/ai-index-february-2026

---

## 4. Langfuse — "State of LLMs on the Application Layer"

Langfuse is an open-source LLM observability platform tracking 20,000+ organizations with billions of traces. Measures **application-layer LLM usage** (not consumer chatbot visits).

### Provider Market Share (September 2025)

| Provider   | Sep 2025 | Oct 2024 | Change     |
|------------|----------|----------|------------|
| OpenAI     | 55.3%    | 82.7%    | -27.4 pts  |
| Google     | 13.1%    | 0.5%     | +12.6 pts  |
| Anthropic  | 7.3%     | 6.8%     | +0.5 pts   |
| Other      | 24.3%    | 10.0%    | +14.3 pts  |

### Top 5 Models (September 2025)
1. GPT-4o mini (OpenAI): 14.2%
2. GPT-4o (OpenAI): 10.9%
3. GPT-4.1 (OpenAI): 9.8%
4. gemini-2.5-flash (Google): 9.4%
5. GPT-4.1 mini (OpenAI): 8.3%

### Links
- **Report**: https://langfuse.com/blog/2025-10-13-state-of-llms-september-2025

---

## 5. Similarweb — AI Chatbot Consumer Traffic

Measures **consumer web traffic** to AI chatbot platforms. Very different from enterprise API usage.

### Market Share by Web Traffic (January 2026)

| Platform     | Share   | 12-Month Change     |
|--------------|---------|---------------------|
| ChatGPT      | 68%     | Down from 87.2%     |
| Google Gemini| 18.2%   | Up from 5.4%        |
| DeepSeek     | 4%      | New entrant         |
| Grok (xAI)   | 2.9%    | New entrant         |
| Perplexity   | 2%      | Growing             |
| Claude       | 2%      | Stable              |
| Copilot      | 1.2%    | Declining           |

### Links
- **Analysis**: https://vertu.com/lifestyle/ai-chatbot-market-share-2026-chatgpt-drops-to-68-as-google-gemini-surges-to-18-2/
- **Traffic Tool**: https://www.similarweb.com/ai-traffic/

---

## 6. Statista — U.S. LLM Provider Market Share

Statista tracks estimated dollars spent based on proportion of production API usage. **Paywalled** — specific percentages require subscription.

### Links
- **Chart**: https://www.statista.com/statistics/1659536/us-llm-provider-market-share/

---

## 7. Gartner — AI Vendor Assessment

Gartner ranks AI vendors across segments: cloud hyperscalers, foundation models, and packaged AI solutions.

### Key Rankings
- **Cloud/Infrastructure**: Microsoft, AWS, Google dominate
- **Foundation Models**: OpenAI, Anthropic, Cohere lead
- **Packaged AI Solutions**: ServiceNow, Salesforce lead
- **Silicon**: NVIDIA dominant

### Links
- **Assessment**: https://www.aicerts.ai/news/market-assessment-gartners-top-ai-vendor-frontrunners-ranked/

---

## 8. McKinsey — "State of AI" Annual Survey

McKinsey's global survey of organizations on AI adoption. Focuses on **adoption breadth** rather than provider market share.

### Key Stats (2025)
- 88% of organizations use AI in at least one function (up from 78% in 2024)
- ~2/3 have not begun scaling AI across the enterprise
- 62% experimenting with AI agents
- Programming rose from 11% to 50%+ of token usage

### Links
- **Report PDF**: https://www.mckinsey.com/~/media/mckinsey/business%20functions/quantumblack/our%20insights/the%20state%20of%20ai/november%202025/the-state-of-ai-2025-agents-innovation_cmyk-v1.pdf

---

## 9. Yipit — Enterprise AI Vendor Adoption Panel

Yipit tracks a proprietary panel of ~1,000 mid-market and enterprise companies. Referenced by a16z to corroborate their findings.

### Adoption Rates
- **OpenAI**: ~85%
- **Anthropic**: ~55% and rising

### Links
- Referenced in a16z report: https://a16z.com/leaders-gainers-and-unexpected-winners-in-the-enterprise-ai-arms-race/

---

## 10. GAI Insights — Corporate Buyers Guide to Enterprise Intelligence

GAI Insights researches enterprise demand for LLMs with detailed market assessments.

### Enterprise Preference (H1 2025)
| Provider   | Enterprise Preference |
|------------|----------------------|
| OpenAI     | 84%                  |
| Google     | 80%                  |
| Anthropic  | 67%                  |
| Meta       | 42%                  |

### Links
- **Report**: https://gaiinsights.com/hubfs/Q3%202025%20Corporate%20Buyers%20Guide%20to%20Enterprise%20Intelligence%20Applications%20(EIA)%20-%20Report%20Excerpt.pdf

---

## Summary: Comparison Across Sources

The table below compares market share estimates across all major sources. Differences reflect different methodologies (API usage vs. spending vs. adoption vs. traffic).

| Source                  | Methodology              | OpenAI | Anthropic | Google | Meta | Notes                              |
|-------------------------|--------------------------|--------|-----------|--------|------|------------------------------------|
| **OpenRouter**          | Token volume (all users) | 8%     | 22%       | 22%    | 4%   | Developer/price-sensitive bias     |
| **Menlo Ventures**      | Enterprise API usage     | 25-27% | 32-40%    | 20-21% | 8-9% | 150 enterprise leaders surveyed    |
| **a16z CIO Survey**     | Enterprise wallet share  | ~56%   | Growing   | Growing| —    | Global 2000 CIOs only             |
| **Ramp**                | Corporate card spending  | 34.4%  | 24.4%     | 4.5%   | —    | Misses bundled cloud contracts     |
| **Langfuse**            | App-layer LLM traces     | 55.3%  | 7.3%      | 13.1%  | —    | 20K+ orgs on platform             |
| **Similarweb**          | Consumer web traffic     | 68%    | 2%        | 18.2%  | —    | Consumer chatbot only             |
| **GAI Insights**        | Enterprise preference    | 84%    | 67%       | 80%    | 42%  | Preference, not actual usage       |
| **Yipit**               | Enterprise panel         | ~85%   | ~55%      | —      | —    | Adoption, not share               |

### Key Takeaway
The **most comparable** source to the original OpenRouter stacked bar chart is **Menlo Ventures**, which produces nearly identical visualizations (horizontal stacked bar charts showing enterprise LLM market share by provider, across 2023-2025). Their data tells a similar story of market fragmentation but with Anthropic — not Google — as the biggest gainer.
