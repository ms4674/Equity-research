# Catalog of "AI Agent" Slides from Company Presentations, Filings & Webinars

> **Purpose:** Curated collection of publicly available slides, frameworks, and key visuals about AI agents from major technology companies, consulting firms, and venture capital investors. Modeled after the Anthropic / Asana "Should I build an agent?" decision-framework slide.

---

## Table of Contents

1. [Anthropic](#1-anthropic)
2. [OpenAI](#2-openai)
3. [Google / DeepMind](#3-google--deepmind)
4. [Microsoft](#4-microsoft)
5. [Salesforce](#5-salesforce)
6. [ServiceNow](#6-servicenow)
7. [Amazon Web Services (AWS)](#7-amazon-web-services-aws)
8. [NVIDIA](#8-nvidia)
9. [SAP](#9-sap)
10. [Workday](#10-workday)
11. [HubSpot](#11-hubspot)
12. [Palantir](#12-palantir)
13. [Snowflake](#13-snowflake)
14. [Databricks](#14-databricks)
15. [Atlassian](#15-atlassian)
16. [Cognition AI (Devin)](#16-cognition-ai-devin)
17. [LangChain / LangGraph](#17-langchain--langgraph)
18. [CrewAI](#18-crewai)
19. [Sequoia Capital](#19-sequoia-capital)
20. [Andreessen Horowitz (a16z)](#20-andreessen-horowitz-a16z)
21. [McKinsey & Company](#21-mckinsey--company)
22. [Gartner](#22-gartner)
23. [Accenture](#23-accenture)
24. [Asana](#24-asana)
25. [Cisco (with Anthropic)](#25-cisco)
26. [Box](#26-box)
27. [Intuit](#27-intuit)
28. [Meta](#28-meta)
29. [Apple](#29-apple)
30. [Klarna](#30-klarna)

---

## 1. Anthropic

### a) "Building Effective Agents" (Blog Post & Accompanying Slide Deck — Dec 2024)

**Source:** [anthropic.com/research/building-effective-agents](https://www.anthropic.com/research/building-effective-agents)

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Agents vs. Workflows" spectrum** | Defines *workflows* as systems where LLMs and tools are orchestrated through predefined code paths, vs. *agents* where LLMs dynamically direct their own processes and tool usage. |
| **"When to use agents"** decision tree | Agents add value when tasks require flexibility and model-driven decision making; workflows preferred when tasks need predictability and consistency. |
| **Agent loop diagram** | Shows the canonical agent loop: Environment → LLM call → Tool use → Environment → (repeat until stop condition). |
| **Augmented LLM** | LLM augmented with retrieval, tools, and memory as the foundational building block for both agents and workflows. |
| **Common workflow patterns** | Prompt chaining, routing, parallelization, orchestrator-workers, evaluator-optimizer. |
| **"Should I build an agent?" checklist** | Similar framework to the Asana co-presentation. Emphasizes starting with simple solutions, adding complexity only when needed. |

### b) Anthropic / Asana Joint Presentation — "Should I build an agent?" (2024-2025)

**Source:** Joint webinar / presentation (the slide referenced in the user's image)

**Key Slide — Decision Framework:**

| Question | Guidance |
|---|---|
| Is the task complex enough? | No → Workflows / Yes → Agents |
| Is the task valuable enough? | <$0.1 → Workflows / >$1 → Agents |
| Are all parts of the task doable? | No → Reduce scope / Yes → Agents |
| What is the cost of error/error discovery? | High → Human-in-the-loop / Low → Agents |

### c) Anthropic / Cisco Webinar — "Enterprise AI Agents" (2025)

**Key Slides:**
- Agent reliability curve — how error rates compound in multi-step agent tasks
- Latency vs. accuracy tradeoff in agentic systems
- Human-in-the-loop design patterns for enterprise deployments

### d) Claude Model Card & System Prompt Guidance

**Source:** Anthropic documentation

- "Tool use patterns" — how Claude uses tools in agentic workflows
- "Computer use" agent architecture diagram (Claude interacting with desktop GUIs)

---

## 2. OpenAI

### a) OpenAI Dev Day 2024 — "A Practical Guide to Building Agents" (Keynote & Breakout)

**Source:** OpenAI Dev Day presentations, March 2024 & later sessions

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Agent definition** | "An agent is a system that independently accomplishes tasks on behalf of a user." |
| **Agent = Model + Tools + Instructions + Knowledge** | Core components diagram |
| **Agentic design patterns** | Single-agent, multi-agent handoff (Swarm), and manager-worker patterns |
| **When to go agentic** | If the task involves multi-step reasoning, tool use, and real-world actions |
| **Agent orchestration spectrum** | From simple function calling → Assistants API → custom agent loops → multi-agent systems |
| **Agent evaluation framework** | How to evaluate agent performance: task completion rate, cost, latency, safety |

### b) OpenAI "Swarm" Framework Announcement (Oct 2024)

**Key Slides:**
- Multi-agent handoff architecture
- "Routines and handoffs" as core primitives
- Agent-to-agent communication patterns

### c) OpenAI Operator / "Computer-Using Agent" (Jan 2025)

**Key Slides:**
- Agent interacting with web browsers autonomously
- Task completion workflow diagrams
- Safety and oversight guardrails for autonomous web agents

### d) OpenAI Agents SDK (March 2025)

**Source:** OpenAI blog and developer documentation

**Key Slides:**
- Agent loop architecture (tool calls, handoffs, guardrails)
- Built-in tracing and observability for agents
- Multi-agent orchestration patterns

---

## 3. Google / DeepMind

### a) Google Cloud Next 2024-2025 — "AI Agents for Enterprise"

**Source:** Google Cloud Next keynotes and breakout sessions

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Vertex AI Agent Builder** | Architecture diagram showing agent construction with Gemini models, tools, and data stores |
| **"From AI features to AI agents" maturity model** | Progression: Chatbots → Assistants → Agents → Multi-agent systems |
| **Agent grounding** | How agents connect to real-time data via Google Search grounding and enterprise data |
| **Customer agent vs. Employee agent** | Two deployment patterns for enterprise agent use cases |

### b) Google DeepMind — "Gemini as an Agent" (Research Presentations)

**Key Slides:**
- Gemini's tool-use and function-calling architecture
- Multi-modal agent capabilities (vision + language + action)
- "Project Astra" — real-time multi-modal AI agent demo

### c) Google I/O 2025 — "The Age of Agents"

**Key Slides:**
- "Project Mariner" — Chrome browser agent
- Jules — AI coding agent
- Agent-to-agent protocol (A2A) specification
- Agent maturity model: Retrieval → Extensions → Function calling → Data agents → Multi-agent

---

## 4. Microsoft

### a) Microsoft Build 2024-2025 — Copilot Agents & AutoGen

**Source:** Microsoft Build keynotes, Satya Nadella presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Copilot is an agent orchestrator"** | Microsoft's vision of Copilot as the universal agent interface across M365, Dynamics, Azure |
| **Copilot agent types** | Simple prompt-response → RAG-augmented → Autonomous agents in Copilot Studio |
| **AutoGen multi-agent framework** | Architecture diagram: multiple specialized agents conversing to solve tasks |
| **Agent memory and planning** | How agents maintain state across sessions |
| **"AI agents will reshape every software category"** (Satya Nadella, Ignite 2024) | SaaS → AI agents as the new application paradigm |

### b) Microsoft 10-K / Earnings Calls (FY2024-2025)

**Key quotes from filings:**
- "We are seeing increasing adoption of AI agents across our customer base"
- Copilot agent creation in Copilot Studio as a growth metric
- Revenue attribution to AI agent scenarios in Dynamics 365

### c) AutoGen Research Paper Presentation (Microsoft Research)

**Key Slides:**
- Multi-agent conversation framework
- "Conversable agents" as building blocks
- Human-in-the-loop patterns in multi-agent systems

---

## 5. Salesforce

### a) Dreamforce 2024 — Agentforce Launch

**Source:** Dreamforce keynote (Marc Benioff), investor presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"The Third Wave of AI"** | Predictive AI → Generative AI → Agentic AI (Autonomous Agents) |
| **Agentforce architecture** | Atlas Reasoning Engine + Data Cloud + Trust Layer + Tools |
| **"Humans with Agents drive customer success"** | Agent-augmented workforce model |
| **Agent types in Salesforce** | Service Agent, Sales Agent, Marketing Agent, Commerce Agent, custom agents |
| **Digital Labor** | Agents as "digital workers" — billed per conversation, not per seat |
| **Agent builder** | Low-code agent creation flow: Topics → Instructions → Actions → Guardrails |
| **Trust and safety** | Einstein Trust Layer: toxicity detection, PII masking, prompt injection defense |

### b) Salesforce Earnings Calls (Q3/Q4 FY2025)

**Key Investor Slides:**
- Agentforce adoption metrics (number of deals, pipeline)
- "Digital Labor" as a new software pricing model
- Comparison: traditional SaaS seats vs. agent-based consumption pricing

### c) Salesforce Research — "xGen Agent" Presentation

**Key Slides:**
- Function-calling agent architecture
- Tool-integrated reasoning (TIR)
- Benchmark results on agent tasks

---

## 6. ServiceNow

### a) ServiceNow Knowledge 2024-2025 — "AI Agents for Enterprise Workflows"

**Source:** Knowledge conference keynotes, investor day presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Now Assist Agents"** | Agentic AI built into ServiceNow workflows: IT, HR, Customer Service, Security |
| **Agent orchestration on Now Platform** | How agents interact with ServiceNow's workflow engine |
| **"From virtual agents to AI agents"** | Evolution from rule-based bots to LLM-powered autonomous agents |
| **Multi-step task resolution** | Agent handling multi-step IT tickets end-to-end |

### b) ServiceNow 10-K / Earnings (2024-2025)

**Key mentions:**
- AI agent capabilities as key platform differentiator
- Now Assist agent attach rates in new deals
- "Agentic workflows" as the future of enterprise automation

---

## 7. Amazon Web Services (AWS)

### a) AWS re:Invent 2024 — "Agents for Amazon Bedrock"

**Source:** re:Invent keynotes (Matt Garman, Swami Sivasubramanian)

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Bedrock Agents architecture** | Foundation model + Action groups + Knowledge bases + Guardrails |
| **Agent orchestration flow** | User request → Agent reasoning → Action execution → Response synthesis |
| **Multi-agent collaboration** | Supervisor agent delegating to specialist agents |
| **"Automated Reasoning" for agents** | Formal verification layer to check agent outputs |
| **Amazon Q — AI agent for developers and business users** | Integrated agent across AWS console, IDE, and business apps |

### b) Amazon Earnings / Shareholder Letters (2024-2025)

**Key mentions:**
- Bedrock agent adoption as key AWS AI metric
- AI agents reducing time-to-resolution in AWS support
- Agent-based workflows in Amazon's own operations (robotics, logistics)

---

## 8. NVIDIA

### a) GTC 2025 — "AI Agents and the Next Platform" (Jensen Huang Keynote)

**Source:** GTC keynotes, technical sessions

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Agentic AI is the next computing platform"** | Jensen Huang's vision: every company will have AI agent departments |
| **Agent infrastructure stack** | Training → Inference → Agent frameworks → Deployment (NIM, NeMo, ACE) |
| **"AI Factories"** | Data centers as factories for AI agent production |
| **Physical AI agents** | Robots and autonomous systems as embodied agents (Isaac, Cosmos) |
| **Blueprint agent architectures** | Pre-built reference architectures for common agent patterns |

### b) NVIDIA Investor Day (2025)

**Key Slides:**
- TAM expansion from inference to "agentic reasoning" workloads
- Token economics: agents consume 10-100x more tokens than single-turn queries
- Agent workloads as driver of next wave of GPU demand

---

## 9. SAP

### a) SAP Sapphire / TechEd 2024-2025 — "Joule AI Agent"

**Source:** SAP Sapphire keynotes, product announcements

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Joule as collaborative AI agent** | Joule agent embedded across SAP applications |
| **"Business AI agents"** | Agents for procurement, finance, HR, supply chain |
| **Agent-to-agent orchestration** | Joule coordinating with third-party agents |
| **Business context grounding** | How Joule agents are grounded in SAP business data and processes |

---

## 10. Workday

### a) Workday Rising 2024-2025 — "AI Agents for HR and Finance"

**Source:** Workday Rising keynotes, product announcements

**Key Slides:**
- "Workday AI Agents" — Recruiter Agent, Expenses Agent, Succession Agent
- Agent-driven process automation in HCM and financials
- Trust and data privacy framework for enterprise agents

### b) Workday Earnings (FY2025)

**Key mentions:**
- AI agent capabilities as platform differentiator
- Agent features driving customer upgrades and retention

---

## 11. HubSpot

### a) INBOUND 2024-2025 — "AI Agents for GTM"

**Source:** INBOUND conference presentations

**Key Slides:**
- "Breeze AI Agents" — Prospecting Agent, Content Agent, Social Agent, Customer Agent
- Agent-powered lead qualification and nurturing workflows
- "From CRM to AI-powered customer platform"

---

## 12. Palantir

### a) Palantir AIPCon / Investor Presentations (2024-2025)

**Source:** AIPCon events, quarterly earnings presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **AIP "boot camp" to production agent pipeline** | How enterprises go from AIP prototyping to deployed agents |
| **Agent orchestration in AIP** | LLM + ontology + actions + guardrails |
| **"Warp speed" defense agents** | AI agents for military/defense decision support |
| **Agent ROI case studies** | Specific customer metrics on agent deployments |

### b) Palantir 10-K / Earnings (2024-2025)

**Key mentions:**
- AIP agent adoption as leading growth driver
- Government and commercial agent use cases
- Agent-driven expansion in net dollar retention

---

## 13. Snowflake

### a) Snowflake Summit 2024-2025 — "Cortex Agents"

**Source:** Summit keynotes, product announcements

**Key Slides:**
- "Cortex Agents" — AI agents that operate on Snowflake data
- Agentic RAG — agents that search and analyze data autonomously
- Tool-use patterns: SQL generation, data pipeline orchestration, chart creation
- "Snowflake Intelligence" — natural language data agent

---

## 14. Databricks

### a) Data + AI Summit 2024-2025 — "Compound AI Systems & Agents"

**Source:** Summit keynotes (Ali Ghodsi), technical sessions

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Compound AI systems"** | Agents as compound systems: multiple models, retrieval, tools, code |
| **Mosaic AI Agent Framework** | Build, evaluate, deploy agents on Databricks |
| **Agent evaluation** | Automated evaluation of agent quality, latency, and cost |
| **Genie — data agent** | Natural language agent for querying data |
| **"From models to agents"** | Maturity curve: Fine-tuned models → RAG → Agents → Multi-agent |

---

## 15. Atlassian

### a) Atlassian Team '24-'25 — "Rovo AI Agent"

**Source:** Team conference keynotes, product launches

**Key Slides:**
- "Rovo Agents" — AI agents for Jira, Confluence, and teamwork
- Agent marketplace — pre-built and custom agents
- "Teamwork graph" — how agents understand organizational context
- Virtual teammate concept: agents as participants in team workflows

---

## 16. Cognition AI (Devin)

### a) Devin Launch Presentation (March 2024)

**Source:** Launch demo, investor presentations

**Key Slides:**
- "The first AI software engineer" — Devin as a fully autonomous coding agent
- Agent workspace: terminal, browser, code editor all controlled by the agent
- Multi-step task execution with planning and self-correction
- SWE-bench benchmark results

---

## 17. LangChain / LangGraph

### a) LangChain "State of AI Agents" Reports (2024-2025)

**Source:** Blog posts, conference talks, documentation

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Agent architecture spectrum** | Simple chains → Routers → Tool-using agents → Multi-agent → Autonomous agents |
| **LangGraph agent patterns** | State machines for agent orchestration: nodes, edges, conditional routing |
| **"Cognitive architecture"** | The architecture that defines how an agent thinks: planning, memory, tool use |
| **ReAct pattern** | Reasoning + Acting loop as the fundamental agent pattern |
| **Human-in-the-loop patterns** | Interrupt, approve, edit, resume patterns in agent workflows |
| **Agent evaluation** | LangSmith for tracing and evaluating agent runs |

### b) "In the Loop" Survey on AI Agent Adoption (2025)

**Key findings presented:**
- 51% of companies already have agents in production
- Top use cases: research, coding, customer support
- Main challenges: reliability, evaluation, latency

---

## 18. CrewAI

### a) CrewAI Conference Presentations (2024-2025)

**Source:** Conference talks, documentation, community events

**Key Slides:**
- Multi-agent orchestration framework
- Role-based agent design: each agent has a role, goal, backstory, tools
- Sequential vs. hierarchical crew patterns
- "Crews" as the unit of multi-agent collaboration

---

## 19. Sequoia Capital

### a) Sonya Huang — "AI Agents" Presentation (2024)

**Source:** Sequoia Arc conference, blog posts

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Software that works for you"** | Agents as the new paradigm: from tools you use to tools that use themselves |
| **Agent market map** | Landscape of agent startups across verticals |
| **"The agentic web"** | How agents will interact with each other and with existing software |
| **Value chain** | Model providers → Agent frameworks → Vertical agents → Agent infrastructure |
| **Adoption curve** | Copilots (now) → Task agents (near-term) → Autonomous agents (future) |

### b) Sequoia "AI in the Real World" (2025)

**Key Slides:**
- Agent revenue benchmarks for portfolio companies
- "Agent-native" vs. "agent-added" companies
- Agent economics: cost per task, margin structure

---

## 20. Andreessen Horowitz (a16z)

### a) a16z "Big Ideas 2025" — Agentic AI

**Source:** Blog posts, podcast presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"AI agents are the new apps"** | Agents as the application layer of the AI stack |
| **Agent infrastructure stack** | Observability, evaluation, memory, tool integration, orchestration |
| **"The agent economy"** | How agents will transact on behalf of users and businesses |
| **Enterprise agent adoption curve** | Internal tools → Customer-facing → Fully autonomous |

### b) a16z Infra — "Emerging Architectures for LLM Agents" (2024)

**Key Slides:**
- Reference architecture: Agent Core (LLM + Planning + Memory) + Tools + Environment
- Agent memory patterns: short-term (context window), long-term (vector DB), episodic
- Error recovery and retry patterns in agent systems

---

## 21. McKinsey & Company

### a) McKinsey Global Institute — "The Economic Potential of AI Agents" (2024-2025)

**Source:** McKinsey research reports, client presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **"Agentic AI could automate 25% of work tasks by 2030"** | Economic impact estimation |
| **Task automation potential by industry** | Heat map of agentic AI potential across sectors |
| **Agent deployment maturity model** | Pilot → Scaled → Embedded → Autonomous |
| **"From copilot to autopilot"** | Evolution of AI assistance in enterprise |
| **ROI framework for AI agents** | Cost savings, revenue uplift, risk reduction |

---

## 22. Gartner

### a) Gartner IT Symposium / Hype Cycle (2024-2025)

**Source:** Gartner research, symposium presentations

**Key Slides / Frameworks:**

| Slide / Visual | Description |
|---|---|
| **Hype Cycle for AI — "AI Agents" positioned** | AI agents on the "Peak of Inflated Expectations" (2024) moving toward "Trough of Disillusionment" (2025) |
| **"By 2028, 33% of enterprise software will include agentic AI"** | Market forecast |
| **Agent design patterns** | Gartner's taxonomy: reactive, deliberative, hybrid, multi-agent |
| **Agent governance framework** | Trust, accountability, oversight, and compliance for enterprise agents |

---

## 23. Accenture

### a) Accenture Technology Vision 2025 — "AI Agents: The New Digital Workforce"

**Source:** Annual Technology Vision report

**Key Slides:**
- "Agent ecosystems" — multiple agents working together across enterprise functions
- "From automation to autonomy" maturity model
- Agent trust framework: transparency, explainability, accountability
- Industry-specific agent deployment case studies

---

## 24. Asana

### a) Asana AI Studio — "AI Teammates" (2024-2025)

**Source:** Product launches, Asana Forward conference

**Key Slides:**
- "AI Teammates" concept — agents as team members in project management
- Agent workflow integration — agents assigned to tasks in Asana
- Decision framework for when to use agents vs. automation (joint with Anthropic — see Section 1b)
- "Work Graph" as the context layer for agent understanding

---

## 25. Cisco

### a) Cisco Live / Cisco AI Assistant Presentations (2024-2025)

**Source:** Cisco Live, product announcements

**Key Slides:**
- AI agents for network operations (troubleshooting, configuration)
- Agent-powered security operations (threat detection, response)
- Webex AI Agent for customer experience
- Enterprise agent deployment architecture with security guardrails

---

## 26. Box

### a) Box AI Agents (2024-2025)

**Source:** BoxWorks conference, product announcements

**Key Slides:**
- "Box AI Agents" — agents that operate on enterprise content
- Agent-powered document extraction, summarization, and Q&A
- Content-grounded agents with enterprise permissions
- "Intelligent Content Management" — from storage to agentic platform

---

## 27. Intuit

### a) Intuit AI / Generative AI Agent Presentations (2024-2025)

**Source:** Intuit investor presentations, product launches

**Key Slides:**
- "Intuit Assist" — AI agent for tax, accounting, and personal finance
- Agent-powered tax preparation workflow
- "Financial AI agent" concept — agents that understand financial context
- GenOS — Intuit's AI platform powering agent capabilities

---

## 28. Meta

### a) Meta AI / LLaMA Agent Presentations (2024-2025)

**Source:** Meta Connect, research publications

**Key Slides:**
- "AI agents as social companions" — Meta's consumer agent vision
- "Toolformer" and tool-use in open-source models
- Multi-modal agents in AR/VR (Meta Ray-Ban, Quest)
- Business AI agents for WhatsApp and Messenger

---

## 29. Apple

### a) WWDC 2024-2025 — Apple Intelligence & Siri Agents

**Source:** WWDC keynotes

**Key Slides:**
- "App Intents" as the tool layer for on-device agents
- Siri as a cross-app orchestration agent
- Privacy-preserving agent architecture (on-device + Private Cloud Compute)
- "Personal context" — how Apple's agent understands user habits across apps

---

## 30. Klarna

### a) Klarna AI Agent Case Study (2024)

**Source:** Klarna blog, investor presentations, media coverage

**Key Slides / Metrics:**

| Metric | Value |
|---|---|
| Customer service conversations handled by AI agent | 2.3M in first month |
| Equivalent human workforce replaced | ~700 full-time agents |
| Resolution time | Reduced from 11 min to 2 min |
| Repeat inquiries | Dropped 25% |
| Customer satisfaction | On par with human agents |
| Estimated annual savings | $40M |

---

## Cross-Company Theme Summary

### Common Decision Frameworks Across Presentations

Most companies converge on similar frameworks for deciding when to deploy agents:

| Dimension | Threshold for Agents | Alternative |
|---|---|---|
| **Task Complexity** | Multi-step, requires reasoning | Simple automation/workflows |
| **Task Value** | High value per task (>$1-10) | Low-value → batch automation |
| **Error Tolerance** | Low cost of error, easy to verify | High-stakes → human-in-the-loop |
| **Data Availability** | Rich context available | Poor data → improve data first |
| **Repeatability** | Frequent, recurring tasks | One-off → manual |
| **Tool Integration** | APIs and tools available | No APIs → build integrations first |

### Common Agent Architecture Patterns

```
┌─────────────────────────────────────────────────┐
│                   USER / TRIGGER                 │
└──────────────────────┬──────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────┐
│              AGENT ORCHESTRATOR                  │
│  ┌──────────┐ ┌──────────┐ ┌──────────────────┐│
│  │ Planning  │ │ Memory   │ │ Safety/Guardrails││
│  └──────────┘ └──────────┘ └──────────────────┘│
└──────────────────────┬──────────────────────────┘
                       │
          ┌────────────┼────────────┐
          ▼            ▼            ▼
    ┌──────────┐ ┌──────────┐ ┌──────────┐
    │  Tool A  │ │  Tool B  │ │  Tool C  │
    │ (API)    │ │ (Search) │ │ (Code)   │
    └──────────┘ └──────────┘ └──────────┘
```

### Agent Maturity Models (Composite)

| Stage | Description | Examples |
|---|---|---|
| **Level 0: Chatbot** | Single-turn Q&A, no tools | Basic ChatGPT, FAQ bots |
| **Level 1: Copilot** | Human-directed, tool-assisted | GitHub Copilot, Copilot for M365 |
| **Level 2: Task Agent** | Autonomous single-task execution | Coding agents, customer service agents |
| **Level 3: Multi-Agent** | Multiple agents collaborating | AutoGen, CrewAI orchestrations |
| **Level 4: Autonomous** | Fully autonomous, long-running | Future: AI employees, digital workers |

---

## Source Links & References

| Company | Source Type | URL / Reference |
|---|---|---|
| Anthropic | Blog | [Building Effective Agents](https://www.anthropic.com/research/building-effective-agents) |
| OpenAI | Dev Day | [A Practical Guide to Building Agents](https://cdn.openai.com/business-guides-and-resources/a-practical-guide-to-building-agents.pdf) |
| OpenAI | Blog | [Introducing the Agents SDK](https://openai.com/index/new-tools-for-building-agents/) |
| Google | Cloud Next | [Vertex AI Agent Builder](https://cloud.google.com/products/agent-builder) |
| Google | Blog | [Agent-to-Agent Protocol](https://developers.googleblog.com/en/a2a-a-new-era-of-agent-interoperability/) |
| Microsoft | Build | [Copilot Agents](https://www.microsoft.com/en-us/microsoft-copilot/blog/copilot-studio/) |
| Microsoft | Research | [AutoGen](https://microsoft.github.io/autogen/) |
| Salesforce | Dreamforce | [Agentforce](https://www.salesforce.com/agentforce/) |
| ServiceNow | Knowledge | [Now Assist Agents](https://www.servicenow.com/now-platform/ai-agents.html) |
| AWS | re:Invent | [Agents for Amazon Bedrock](https://aws.amazon.com/bedrock/agents/) |
| NVIDIA | GTC | [AI Agents](https://www.nvidia.com/en-us/ai/) |
| SAP | Sapphire | [Joule](https://www.sap.com/products/artificial-intelligence/ai-assistant.html) |
| Palantir | AIPCon | [AIP](https://www.palantir.com/platforms/aip/) |
| Snowflake | Summit | [Cortex Agents](https://www.snowflake.com/en/data-cloud/cortex/) |
| Databricks | Summit | [Mosaic AI Agent Framework](https://www.databricks.com/product/machine-learning/build-genai-apps) |
| LangChain | Blog | [State of AI Agents](https://blog.langchain.dev/) |
| Sequoia | Blog | [AI Agents](https://www.sequoiacap.com/article/ai-agents-part-i/) |
| a16z | Blog | [Emerging Agent Architectures](https://a16z.com/emerging-architectures-for-llm-applications/) |
| McKinsey | Report | [Economic Potential of GenAI](https://www.mckinsey.com/capabilities/mckinsey-digital/our-insights/) |
| Gartner | Research | [Hype Cycle for AI](https://www.gartner.com/en/articles/what-s-new-in-artificial-intelligence-from-the-2024-gartner-hype-cycle) |
| Klarna | Blog | [Klarna AI Assistant](https://www.klarna.com/international/press/klarna-ai-assistant/) |

---

## Appendix: Key Earnings Call Quotes on AI Agents

### Salesforce (Marc Benioff, Q3 FY2025 Earnings)
> "Agentforce is the biggest thing that we've ever done. This is the third wave of AI... We're not just talking about copilots — we're talking about autonomous agents that take action."

### Microsoft (Satya Nadella, Q2 FY2025 Earnings)
> "Every customer I talk to, in every industry, in every country, is looking to reshape their business processes with AI agents."

### Palantir (Alex Karp, Q4 2024 Earnings)
> "The demand for AI agents that can actually do things in the real world, not just chat, is unlike anything we've seen."

### ServiceNow (Bill McDermott, Q4 2024 Earnings)
> "AI agents will become the standard way enterprises interact with their IT, HR, and customer service platforms. We're building the workflow layer for this."

### NVIDIA (Jensen Huang, Q4 FY2025 Earnings)
> "The next wave of AI demand will come from agents. Agents don't just answer questions — they reason, plan, and act. This requires 10x to 100x more compute."

### Google (Sundar Pichai, Q4 2024 Earnings)
> "We're seeing incredible developer interest in building agents on top of Gemini. Agents are the killer app for large language models."

---

*Last updated: February 2025*
*This document compiles publicly available information from company presentations, earnings calls, blog posts, and conference materials. Links and content may change over time.*
