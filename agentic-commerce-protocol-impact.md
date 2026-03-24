# Agentic Commerce Protocol and Its Impact on Performance Ads and Visual Search

## Executive Summary

The emergence of agentic commerce protocols -- most notably the **Universal Commerce Protocol (UCP)** co-developed by Google and Shopify, and the **Agentic Commerce Protocol (ACP)** co-developed by OpenAI and Stripe -- represents the most significant structural shift in digital commerce since the rise of mobile shopping. These open standards enable AI agents to discover, evaluate, and complete purchase transactions autonomously through machine-readable APIs, fundamentally bypassing the visual, browser-based shopping interfaces that underpin the ~$970 billion global digital advertising market.

This report examines how agentic commerce protocols are reshaping performance advertising economics, disrupting retail media networks, transforming visual search from a human-centric tool to a hybrid human-agent modality, and forcing the dominant ad platforms -- Google, Meta, and Amazon -- to reinvent their monetization architectures.

---

## 1. The Protocol Landscape: UCP vs. ACP

### 1.1 Universal Commerce Protocol (UCP)

UCP is an open-source standard designed as the foundational "HTTP of agentic commerce." Co-developed by Google with partners including Shopify and Walmart, UCP enables AI agents to understand merchant capabilities -- inventory availability, cart management, checkout flows, and order management -- and execute end-to-end transactions across heterogeneous systems.

**Key characteristics:**
- Open-source, multi-platform standard
- Endorsed by 20+ major retailers
- Powers direct purchasing from retailers like Etsy and Wayfair within Google AI Mode and Gemini
- Shopify, Target, and Walmart integrations announced for 2026

**Adoption metrics (as of March 2026):**
- Nearly 100% of Shopify stores are technically UCP-ready, with 2,646 verified live UCP manifests
- UCP domain tracking grew from ~450 domains in January 2026 to over 3,000 by mid-March 2026
- 285 non-Shopify platforms appeared in UCP benchmarks for the first time in March 2026, representing 24% of 1,183 benchmarked stores
- Verification rates hold at 95%

**Readiness gap:** Despite near-universal technical availability on Shopify, only 9.8% of 305,991 audited Shopify stores score as "AI-Ready" across five readiness signals, with just 0.5% achieving full readiness. This gap between protocol availability and merchant preparedness is a critical bottleneck.

### 1.2 Agentic Commerce Protocol (ACP)

ACP, co-developed by OpenAI and Stripe, specializes in conversational, chat-to-buy transactions. It provides standardized checkout coordination and secure payment credential sharing between AI platforms and merchants.

**Key characteristics:**
- Open-source, payment-processor agnostic
- Designed for integration with ChatGPT and other conversational AI surfaces
- Powered ChatGPT's "Instant Checkout" feature (launched September 2025)
- Supports over one million Shopify merchants and Etsy sellers

**Strategic pivot (March 2026):** OpenAI is shifting away from direct Instant Checkout within ChatGPT product listings toward merchant-specific ChatGPT apps. Retailers including Target, DoorDash, and Instacart are launching branded apps within the ChatGPT ecosystem, with ACP serving as connective infrastructure rather than a direct checkout surface. This pivot reflects merchants' concerns over inventory management complexity, sales tax compliance, and proprietary data control.

### 1.3 Walmart's Cautionary Signal

Walmart withdrew from OpenAI's ChatGPT Instant Checkout in March 2026 after internal data showed conversion rates approximately 3x lower than on Walmart's own site. The company is instead embedding its AI assistant "Sparky" into ChatGPT while keeping transactions on its own platform -- a hybrid model that preserves merchant control while leveraging agent distribution.

---

## 2. Impact on Performance Advertising

### 2.1 The Core Disruption: Machines Don't See Ads

The fundamental disruption is architectural. AI agents interact with commerce through structured API calls and JSON data feeds, not through visual web pages. They do not "see" banner ads, display placements, or visually rich product listings. When an agent queries for products, it processes machine-readable data -- product attributes, pricing, availability, reviews -- not rendered HTML with embedded advertising.

**Quantitative impact:** If 30% of e-commerce traffic becomes agent-driven (a plausible medium-term scenario), banner ad inventory effectively drops by 30%, eroding the impression-based economics that sustain display advertising.

### 2.2 From Click-Based to Transaction-Based Metrics

Traditional performance advertising runs on a **Search -> Click -> Landing Page -> Checkout** flow, monetized through CPC (cost-per-click) and CPM (cost-per-mille impressions). Agentic commerce compresses this into a single conversational or API-mediated flow where there is no discrete "click" or "page view."

**Key metric shifts:**
| Traditional Model | Agentic Model |
|---|---|
| CPC (cost per click) | Cost per agent query inclusion |
| CPM (impressions) | Cost per structured data ranking |
| CTR (click-through rate) | Agent selection rate |
| Landing page conversion | In-conversation checkout rate |
| ROAS (return on ad spend) | Return on protocol spend (ROPS) |

Downstream commerce events -- cart additions, checkout completions, purchases -- become the primary performance signals, displacing upstream engagement metrics.

### 2.3 The Shift from "Sponsored Listings" to "Sponsored Context"

Performance advertising is not disappearing but is being fundamentally restructured. In the agentic paradigm:

- **Banner ads** become "privileged positions in structured data feeds"
- **Sponsored product ads** become "ranking boosts within API payloads"
- **Display retargeting** becomes "agent memory context injection"

When an AI agent queries for "best running shoes under $150," the monetization opportunity lies in which products appear in the structured JSON response and how they are ranked -- not in visual banner placements around the results.

Google is already operationalizing this shift through:
- **AI Mode ads:** Retailer information embedded alongside AI-generated product recommendations, clearly marked as sponsored
- **Direct Offers:** Personalized discounts surfaced within AI Mode conversations based on conversational context, using a pay-per-click model
- **Business Agents:** Brand-specific conversational AI that shoppers can interact with directly on Google Search

### 2.4 Semantic Governance: The Trust Constraint

Unlike human shoppers who tolerate irrelevant ads (ad blindness), AI agents actively penalize poor relevance. An agent that receives irrelevant sponsored items may downgrade or blacklist the entire API endpoint, marking the merchant as "untrustworthy." This imposes a new constraint:

**Relevance guardrails** must prevent low-relevance sponsored items from damaging trust scores, even when brands bid aggressively. Retailers must balance monetization through sponsored placements against the existential risk of losing agent traffic entirely. This is a qualitatively different dynamic from the traditional "user experience vs. ad load" trade-off.

### 2.5 Market Context

Despite these structural shifts, the overall digital advertising market continues to grow:
- **Global digital ad spend:** projected to reach $972.5 billion in 2026, growing 13.2% annually
- **US digital ad spend:** expected to reach $413.24 billion in 2026 (+14.2% YoY)
- **US retail media spend:** projected at $69 billion in 2026, up from $62 billion in 2025
- There are 277 retail media networks globally as of late 2025

The near-term impact is therefore one of margin compression and model migration rather than absolute revenue decline. The advertising ecosystem is growing fast enough to absorb protocol-driven disruption in the short term, but the structural economics are shifting beneath the surface.

---

## 3. Impact on Retail Media Networks

### 3.1 The Retail Media Threat Model

Retail media networks (RMNs) -- Amazon Ads, Walmart Connect, Target Roundel, Instacart Ads, and 277+ others -- built competitive moats on three advantages:

1. **First-party purchase data** for targeting and measurement
2. **Closed-loop attribution** from ad impression to transaction
3. **Visual real estate** on high-intent shopping surfaces

Agentic commerce erodes all three:
- **First-party data:** Agent providers (Google, OpenAI, Anthropic) accumulate deeper cross-retailer customer insights, diluting any single retailer's data advantage
- **Closed-loop attribution:** When transactions complete within agent interfaces, the attribution chain is controlled by the agent platform, not the retailer
- **Visual real estate:** Agent-mediated queries bypass the retailer's visual interface entirely

### 3.2 The New Monetization Architecture

RMNs are pivoting from visual placements to structured data monetization:

- **Sponsored product ads** -> **Sponsored data attributes** (boosted visibility in API responses)
- **Display banners** -> **Contextual relevance boosts** (higher ranking weight in agent queries)
- **On-site search ads** -> **Protocol-layer sponsorship** (preferred routing in UCP/ACP resolution)

Google's **Retail-MCP** (Model Context Protocol) standardizes how agents access retail systems, creating a new governance layer where monetization must co-exist with agent trust. This is analogous to how Google's search quality team balances ad revenue against organic search quality -- but with the stakes dramatically higher, since agents can switch retailers instantly.

---

## 4. Impact on Visual Search

### 4.1 Visual Search Market Scale

Visual search has grown into a major commerce channel:
- **Google Lens:** ~20 billion monthly visual searches, of which ~4 billion are shopping-related
- **Pinterest Lens:** 600 million monthly visual searches, 500 million MAUs globally
- **Amazon StyleSnap:** 70% YoY increase in visual product searches, 40% higher AOV
- **Market size:** Visual search technology market reached $15 billion in 2026, projected to hit $27.8 billion by 2032 (20-25% CAGR)
- **Revenue impact:** Visual search expected to generate $14.7 billion in direct US e-commerce revenue in 2026

### 4.2 The Dual-Track Future of Visual Search

Agentic commerce creates a bifurcation in how visual search evolves:

**Track 1: Human-Facing Visual Search (Growing)**
Visual search remains deeply human-centric in its primary use case: a consumer points their phone camera at a product, and AI identifies it, finds purchase options, and surfaces reviews. This modality is inherently visual and grows with smartphone penetration and camera AI capabilities.

Google is expanding this channel with:
- Shopping product details directly in Google Lens results (price, deals, reviews, where to buy)
- Circle to Search for shopping items seen on social media or in video content
- Combined text + image searches for refined discovery (e.g., "this chair but in blue velvet")

Retailers implementing AI-powered visual discovery report 30% increases in conversion rates and 25% lifts in customer engagement versus traditional text search.

**Track 2: Agent-Mediated Visual Understanding (Emerging)**
In the agentic paradigm, visual search becomes a **data extraction layer** rather than a shopping interface. An AI agent might use visual recognition to:
- Parse product images into structured attribute data (color, material, style, brand)
- Cross-reference visual matches against protocol-accessible inventory feeds
- Extract product identifiers from physical environments for automated repurchase

In this mode, the "visual search" happens upstream of the commerce protocol -- it's a perception step, not a shopping step. The commerce transaction flows through UCP/ACP regardless of whether the initial product discovery was visual, conversational, or text-based.

### 4.3 Implications for Visual Search Advertising

Visual search advertising (e.g., Google Lens Shopping Ads, Pinterest Shopping Pins) faces the same structural pressure as traditional display:

- **If the agent mediates:** Visual search ads become irrelevant because the agent processes the visual input into structured data and routes the transaction through protocol-layer monetization
- **If the human shops:** Visual search ads remain highly effective because the consumer is engaging with a visual, browsable interface

The net effect is that visual search advertising becomes tied to the share of visual search queries that remain human-directed versus agent-mediated. In the near term (2026-2027), the overwhelming majority of visual searches are still human-initiated and human-completed, protecting this channel. Over the medium term (2028-2030), as agents become the default commerce interface, the visual search advertising surface area will compress.

---

## 5. Platform-Level Strategic Responses

### 5.1 Google (Alphabet)

**Capital commitment:** $175-185 billion in AI infrastructure spending for 2026.

**Strategy:** Google is pursuing a dual approach -- leading the protocol standard (UCP) while simultaneously building monetization into AI-first shopping surfaces (AI Mode, Gemini). This mirrors the company's historical playbook of controlling the platform layer while monetizing the application layer.

**Key moves:**
- UCP as the open standard for agent-to-merchant communication
- AI Mode ads embedded in Gemini shopping conversations
- Direct Offers for personalized discounts in conversational contexts
- Business Agents enabling brand-specific chat on Google Search
- Google Lens integration with structured shopping data

**Risk:** AI Mode queries are 23x longer than traditional search queries, indicating a fundamental behavioral shift. If agents increasingly handle shopping intent, Google's keyword-auction revenue model faces structural dilution.

**Opportunity:** By owning the protocol layer (UCP) and the agent layer (Gemini), Google can capture value at both ends -- charging merchants for protocol access and brands for preferential positioning within agent responses.

### 5.2 Meta

**Capital commitment:** $115-135 billion in AI infrastructure spending.

**Strategy:** Meta is positioning itself as the "compute rail" for autonomous shopping agents, leveraging its 3.58 billion daily active users for personal context data. Meta's advantage is unique behavioral and preference data at scale that competitors cannot replicate.

**Key moves:**
- Building AI infrastructure to support agentic commerce at scale
- Q4 2025 ad revenue of $59.9 billion funding the AI transition
- Exploring agent-mediated commerce within Instagram and WhatsApp surfaces

**Risk:** Meta's entire revenue model ($164 billion in 2024 ad revenue) depends on visual, feed-based advertising. Agentic commerce fundamentally threatens this by replacing browsing with autonomous agent transactions.

**Opportunity:** Meta's social graph and preference data could make its agents the most personalized shopping assistants, creating a new monetization model based on transaction facilitation rather than impression delivery.

### 5.3 Amazon

**Strategy:** Amazon is building a closed agent ecosystem to protect its card-based payment dominance and keep transactions internal rather than flowing through external agent networks.

**Key moves:**
- Rufus AI shopping assistant embedded throughout the Amazon experience
- Alexa integration for voice-initiated agentic commerce
- Resistance to open protocol standards that would enable cross-platform agent shopping

**Risk:** If open protocols (UCP/ACP) become dominant, Amazon's walled-garden advantage erodes as agents can compare and transact across all retailers transparently.

**Opportunity:** Amazon's fulfillment infrastructure and Prime membership create switching costs that agents cannot easily disintermediate.

---

## 6. Revenue Model Migration: From Ads to Pay-Per-Use

### 6.1 The Emerging Revenue Architecture

Agentic commerce is catalyzing a shift from advertising-funded commerce to transaction-funded commerce:

| Revenue Model | Current State | Agentic State |
|---|---|---|
| **Brand monetization** | Ad spend (CPC/CPM) | Protocol access fees + sponsored data ranking |
| **Platform monetization** | Ad auction revenue | Transaction facilitation fees (5-10% of value) |
| **Retailer monetization** | Retail media ad sales | Agent API access fees + sponsored context |
| **Attribution** | Multi-touch ad attribution | Protocol-layer transaction attribution |

AI platforms are positioned to capture 5-10% of transaction value through facilitation fees, creating a direct revenue transfer away from the traditional ad ecosystem.

### 6.2 Early Adopter Performance

Early adopters of protocol-optimized commerce report:
- **9x conversion increases** by optimizing for the protocol layer rather than traditional SEO
- **Higher AOV** through agent-curated product recommendations
- **Lower CAC** due to reduced intermediary friction

These results, while early-stage, suggest that protocol-optimized merchants can achieve superior unit economics compared to ad-dependent customer acquisition.

---

## 7. Key Risks and Uncertainties

### 7.1 Adoption Friction
- Only 9.8% of Shopify stores are truly "AI-Ready" despite near-universal UCP technical availability
- Walmart's 3x lower conversion rate on ChatGPT versus its own site suggests agents have not yet earned consumer trust for high-consideration purchases
- Payment infrastructure limitations (e.g., stablecoin processing restrictions) constrain certain transaction models

### 7.2 Regulatory Uncertainty
- Agent-mediated sponsored placements raise novel questions about advertising disclosure
- Cross-border agent transactions introduce jurisdictional complexity for consumer protection
- Data portability requirements may apply to agent-accumulated purchase histories

### 7.3 Consumer Behavior Lag
- Consumers must trust AI agents with purchasing authority, a behavior change that historically takes years
- High-consideration categories (electronics, furniture, apparel) may resist agent mediation longer than commodity categories
- Privacy concerns about agents accessing payment credentials and purchase history may slow adoption

### 7.4 Protocol Fragmentation
- UCP (Google/Shopify) and ACP (OpenAI/Stripe) are not yet fully interoperable
- Proprietary agent ecosystems (Amazon, Walmart's Sparky) may resist open protocol adoption
- Competing standards could balkanize the agent commerce landscape, slowing network effects

---

## 8. Investment Implications

### 8.1 Structural Winners
- **Protocol infrastructure providers:** Companies building the middleware between agents and merchants (Shopify, Stripe)
- **Structured data platforms:** Companies helping merchants optimize product data for agent consumption (Syndigo, Salsify)
- **Agent platform operators:** Companies controlling the agent interface (Google, OpenAI, Anthropic)

### 8.2 Structural Risks
- **Pure-play display advertising:** Companies whose revenue depends on visual impression delivery
- **Traditional SEO/SEM agencies:** The keyword-auction optimization model faces obsolescence
- **Retailers without protocol readiness:** Merchants that fail to implement UCP/ACP will become invisible to agent-mediated commerce

### 8.3 Transitional Dynamics
- **Near-term (2026-2027):** Protocol adoption expands but human-directed shopping remains dominant; advertising revenue growth continues to mask structural shifts
- **Medium-term (2028-2030):** Agent-mediated transactions reach meaningful scale (Gartner forecasts 60%+ of brands using agentic AI for one-to-one marketing by 2028); ad model migration accelerates
- **Long-term (2030+):** The dominant commerce interface is agent-mediated; advertising economics permanently restructure around protocol-layer monetization

---

## 9. Conclusion

Agentic commerce protocols represent a platform-level shift comparable to the transition from desktop to mobile commerce. The impact on performance advertising is structural: the visual, impression-based, click-driven model that generated $972 billion in global digital ad spend in 2026 is being challenged by machine-readable, protocol-mediated, transaction-driven commerce.

Visual search occupies a unique position in this transition -- it remains inherently human-centric in its primary use case (camera-based product discovery) while simultaneously being absorbed as a perception layer in agent workflows. The $15 billion visual search market will bifurcate between a growing human-facing channel and an emerging agent-facing data extraction modality.

The winners will be those who control the protocol layer (Google via UCP, OpenAI via ACP), the data infrastructure (Shopify, Stripe), and the agent interfaces (Gemini, ChatGPT). The losers will be those whose business models depend on humans visually browsing web pages -- an activity that agentic commerce is designed to eliminate.

---

*Sources: Google Blog, OpenAI Blog, Stripe Documentation, UCP Checker, Particular Audience, Simon-Kucher, Everest Group, PYMNTS, Coresight Research, Global Growth Insights, Stormy AI, StoreInspect, various industry analyses. Data current as of March 2026.*
