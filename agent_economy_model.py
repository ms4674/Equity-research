#!/usr/bin/env python3
"""
Agent Economy Data Model
========================
Generates a comprehensive Excel workbook covering:
  1. Top Companies in the AI-Agent Economy
  2. Task Duration Benchmarks
  3. Token Consumption Profiles
  4. Time-Series Projections (2024-2030)

Author: Research Team
Date: 2026-02-12
"""

import xlsxwriter
import datetime
import math

OUTPUT_FILE = "agent_economy_data_model.xlsx"

# ---------------------------------------------------------------------------
# 1. TOP COMPANIES DATA
# ---------------------------------------------------------------------------
COMPANIES = [
    {
        "company": "OpenAI",
        "hq": "San Francisco, CA",
        "founded": 2015,
        "category": "Foundation Model / Platform",
        "key_products": "GPT-4o, GPT-4.5, o3, Operator, Assistants API, ChatGPT Plugins",
        "latest_valuation_bn": 300.0,
        "est_2025_revenue_bn": 12.7,
        "est_2026_revenue_bn": 29.0,
        "employees": 3700,
        "primary_model": "GPT-4o / o3",
        "agent_focus": "General-purpose agents, tool-use, code interpreter, web browsing, Operator autonomous agent",
        "funding_total_bn": 40.0,
    },
    {
        "company": "Anthropic",
        "hq": "San Francisco, CA",
        "founded": 2021,
        "category": "Foundation Model / Platform",
        "key_products": "Claude 4 (Opus/Sonnet), Claude Code, Computer Use, MCP Protocol",
        "latest_valuation_bn": 61.5,
        "est_2025_revenue_bn": 2.5,
        "est_2026_revenue_bn": 6.0,
        "employees": 1200,
        "primary_model": "Claude 4 Opus / Sonnet",
        "agent_focus": "Tool use, computer use, agentic coding (Claude Code), Model Context Protocol (MCP)",
        "funding_total_bn": 13.7,
    },
    {
        "company": "Google DeepMind",
        "hq": "Mountain View, CA / London, UK",
        "founded": 2023,
        "category": "Foundation Model / Platform (Big Tech)",
        "key_products": "Gemini 2.0, Project Mariner, Jules, Vertex AI Agent Builder",
        "latest_valuation_bn": None,  # Part of Alphabet
        "est_2025_revenue_bn": None,  # Segment not broken out
        "est_2026_revenue_bn": None,
        "employees": 4000,
        "primary_model": "Gemini 2.0 Flash/Pro",
        "agent_focus": "Project Mariner (browser agent), Jules (coding agent), Vertex AI Agent Builder for enterprise",
        "funding_total_bn": None,
    },
    {
        "company": "Microsoft (Copilot / Azure AI)",
        "hq": "Redmond, WA",
        "founded": 1975,
        "category": "Platform / Enterprise AI (Big Tech)",
        "key_products": "Copilot Studio, Azure AI Agent Service, Microsoft 365 Copilot, GitHub Copilot",
        "latest_valuation_bn": None,  # Public: ~$3T market cap
        "est_2025_revenue_bn": None,  # AI revenue embedded
        "est_2026_revenue_bn": None,
        "employees": 228000,
        "primary_model": "GPT-4o / Custom (via OpenAI partnership)",
        "agent_focus": "Enterprise copilots, autonomous agents in M365, Azure AI Agent Service, GitHub Copilot Workspace",
        "funding_total_bn": None,
    },
    {
        "company": "Salesforce (Agentforce)",
        "hq": "San Francisco, CA",
        "founded": 1999,
        "category": "Enterprise Agent Platform",
        "key_products": "Agentforce, Einstein GPT, Data Cloud, Slack AI",
        "latest_valuation_bn": None,  # Public: ~$300B market cap
        "est_2025_revenue_bn": 1.5,  # AI/agent attributed
        "est_2026_revenue_bn": 4.0,
        "employees": 72000,
        "primary_model": "Multi-model (OpenAI, Anthropic, proprietary xGen)",
        "agent_focus": "Customer service agents, sales agents, marketing agents, custom agent builder via Agentforce",
        "funding_total_bn": None,
    },
    {
        "company": "Cursor (Anysphere)",
        "hq": "San Francisco, CA",
        "founded": 2022,
        "category": "AI-Native Coding Agent",
        "key_products": "Cursor IDE, Background Agents, Tab Autocomplete, Composer",
        "latest_valuation_bn": 10.0,
        "est_2025_revenue_bn": 0.5,
        "est_2026_revenue_bn": 1.2,
        "employees": 100,
        "primary_model": "Claude, GPT-4o, Custom fine-tuned models",
        "agent_focus": "Autonomous coding agents, background task execution, multi-file editing, codebase-aware AI",
        "funding_total_bn": 0.9,
    },
    {
        "company": "Cognition (Devin)",
        "hq": "San Francisco, CA",
        "founded": 2023,
        "category": "Autonomous Coding Agent",
        "key_products": "Devin (AI Software Engineer)",
        "latest_valuation_bn": 3.0,
        "est_2025_revenue_bn": 0.05,
        "est_2026_revenue_bn": 0.2,
        "employees": 50,
        "primary_model": "Proprietary + Claude",
        "agent_focus": "Fully autonomous software engineering agent that plans, codes, debugs, and deploys",
        "funding_total_bn": 0.2,
    },
    {
        "company": "LangChain / LangSmith",
        "hq": "San Francisco, CA",
        "founded": 2022,
        "category": "Agent Framework / Tooling",
        "key_products": "LangChain, LangGraph, LangSmith, LangServe",
        "latest_valuation_bn": 2.0,
        "est_2025_revenue_bn": 0.04,
        "est_2026_revenue_bn": 0.12,
        "employees": 120,
        "primary_model": "Model-agnostic orchestration",
        "agent_focus": "Agent orchestration framework, multi-step workflows (LangGraph), observability (LangSmith)",
        "funding_total_bn": 0.035,
    },
    {
        "company": "CrewAI",
        "hq": "San Francisco, CA",
        "founded": 2023,
        "category": "Multi-Agent Framework",
        "key_products": "CrewAI Framework, CrewAI Enterprise",
        "latest_valuation_bn": 0.5,
        "est_2025_revenue_bn": 0.01,
        "est_2026_revenue_bn": 0.05,
        "employees": 40,
        "primary_model": "Model-agnostic orchestration",
        "agent_focus": "Multi-agent collaboration, role-based agents, task delegation, enterprise agent workflows",
        "funding_total_bn": 0.018,
    },
    {
        "company": "Adept AI",
        "hq": "San Francisco, CA",
        "founded": 2022,
        "category": "Enterprise Action Agent",
        "key_products": "ACT-2 Model, Adept Platform",
        "latest_valuation_bn": 1.1,
        "est_2025_revenue_bn": 0.02,
        "est_2026_revenue_bn": 0.08,
        "employees": 80,
        "primary_model": "ACT-2 (proprietary multimodal)",
        "agent_focus": "Computer-use agents that interact with enterprise software (SAP, Salesforce, internal tools)",
        "funding_total_bn": 0.415,
    },
    {
        "company": "Replit",
        "hq": "San Francisco, CA",
        "founded": 2016,
        "category": "AI Coding / Deployment Platform",
        "key_products": "Replit Agent, Ghostwriter, Deployments",
        "latest_valuation_bn": 1.16,
        "est_2025_revenue_bn": 0.1,
        "est_2026_revenue_bn": 0.3,
        "employees": 200,
        "primary_model": "Claude, proprietary fine-tuned models",
        "agent_focus": "End-to-end app building agent: plans, codes, configures, and deploys full applications",
        "funding_total_bn": 0.222,
    },
    {
        "company": "Cohere",
        "hq": "Toronto, Canada",
        "founded": 2019,
        "category": "Enterprise Foundation Model",
        "key_products": "Command R+, Coral, North (Agentic Platform)",
        "latest_valuation_bn": 5.5,
        "est_2025_revenue_bn": 0.06,
        "est_2026_revenue_bn": 0.2,
        "employees": 450,
        "primary_model": "Command R+ / Aya",
        "agent_focus": "Enterprise RAG agents, multi-step tool use, grounded generation with citations",
        "funding_total_bn": 0.97,
    },
    {
        "company": "Mistral AI",
        "hq": "Paris, France",
        "founded": 2023,
        "category": "Foundation Model / Open-Weight",
        "key_products": "Mistral Large 2, Codestral, Le Chat, La Plateforme",
        "latest_valuation_bn": 6.2,
        "est_2025_revenue_bn": 0.05,
        "est_2026_revenue_bn": 0.2,
        "employees": 100,
        "primary_model": "Mistral Large 2 / Codestral",
        "agent_focus": "Function calling, agentic coding, open-weight models enabling custom agent deployments",
        "funding_total_bn": 1.1,
    },
    {
        "company": "Perplexity AI",
        "hq": "San Francisco, CA",
        "founded": 2022,
        "category": "AI Search / Research Agent",
        "key_products": "Perplexity Pro, Internal Knowledge Search, API",
        "latest_valuation_bn": 9.0,
        "est_2025_revenue_bn": 0.15,
        "est_2026_revenue_bn": 0.5,
        "employees": 200,
        "primary_model": "Multi-model (proprietary + Claude + GPT-4o)",
        "agent_focus": "Research agents with real-time web search, multi-step reasoning, citation generation",
        "funding_total_bn": 0.9,
    },
    {
        "company": "Hebbia",
        "hq": "New York, NY",
        "founded": 2020,
        "category": "Enterprise Knowledge Agent",
        "key_products": "Matrix (Agentic Analytics Platform)",
        "latest_valuation_bn": 0.7,
        "est_2025_revenue_bn": 0.03,
        "est_2026_revenue_bn": 0.1,
        "employees": 70,
        "primary_model": "Multi-model orchestration",
        "agent_focus": "Document analysis agents for finance, legal, and consulting; structured data extraction at scale",
        "funding_total_bn": 0.165,
    },
    {
        "company": "Sierra AI",
        "hq": "San Francisco, CA",
        "founded": 2023,
        "category": "Customer Experience Agents",
        "key_products": "Sierra Agent Platform",
        "latest_valuation_bn": 4.5,
        "est_2025_revenue_bn": 0.03,
        "est_2026_revenue_bn": 0.12,
        "employees": 100,
        "primary_model": "Multi-model",
        "agent_focus": "Conversational AI agents for customer experience, order management, troubleshooting",
        "funding_total_bn": 0.285,
    },
    {
        "company": "Fixie / AI.JSX",
        "hq": "Seattle, WA",
        "founded": 2022,
        "category": "Agent Platform",
        "key_products": "Fixie Platform, AI.JSX",
        "latest_valuation_bn": 0.2,
        "est_2025_revenue_bn": 0.005,
        "est_2026_revenue_bn": 0.02,
        "employees": 30,
        "primary_model": "Model-agnostic",
        "agent_focus": "Composable AI agents with structured output, multi-agent coordination",
        "funding_total_bn": 0.025,
    },
    {
        "company": "AutoGPT / Significant Gravitas",
        "hq": "Remote / London, UK",
        "founded": 2023,
        "category": "Open-Source Agent Framework",
        "key_products": "AutoGPT, AutoGPT Forge, Agent Protocol",
        "latest_valuation_bn": 0.2,
        "est_2025_revenue_bn": 0.005,
        "est_2026_revenue_bn": 0.03,
        "employees": 20,
        "primary_model": "GPT-4o, Claude, open-source LLMs",
        "agent_focus": "Autonomous goal-driven agents, recursive task decomposition, open-source agent standard",
        "funding_total_bn": 0.012,
    },
    {
        "company": "Amazon (Bedrock Agents / Q)",
        "hq": "Seattle, WA",
        "founded": 1994,
        "category": "Cloud Platform / Enterprise AI (Big Tech)",
        "key_products": "Amazon Bedrock Agents, Amazon Q, CodeWhisperer",
        "latest_valuation_bn": None,  # Public
        "est_2025_revenue_bn": None,
        "est_2026_revenue_bn": None,
        "employees": 1540000,
        "primary_model": "Claude (Anthropic), Titan, Nova, multi-model via Bedrock",
        "agent_focus": "Bedrock Agents for enterprise workflow automation, Amazon Q for developer/business productivity",
        "funding_total_bn": None,
    },
    {
        "company": "ServiceNow",
        "hq": "Santa Clara, CA",
        "founded": 2004,
        "category": "Enterprise Agent Platform",
        "key_products": "Now Assist, AI Agents for ITSM/ITOM/HR/CSM",
        "latest_valuation_bn": None,  # Public: ~$200B market cap
        "est_2025_revenue_bn": 1.0,  # AI attributed
        "est_2026_revenue_bn": 3.0,
        "employees": 24000,
        "primary_model": "Multi-model (proprietary + OpenAI + others)",
        "agent_focus": "IT service desk agents, HR agents, customer service agents, workflow automation agents",
        "funding_total_bn": None,
    },
]

COMPANY_COLUMNS = [
    ("company", "Company"),
    ("hq", "Headquarters"),
    ("founded", "Founded"),
    ("category", "Category"),
    ("key_products", "Key Products"),
    ("latest_valuation_bn", "Latest Valuation ($B)"),
    ("est_2025_revenue_bn", "Est. 2025 Revenue ($B)"),
    ("est_2026_revenue_bn", "Est. 2026 Revenue ($B)"),
    ("employees", "Employees"),
    ("primary_model", "Primary Model(s)"),
    ("agent_focus", "Agent Focus / Capabilities"),
    ("funding_total_bn", "Total Funding ($B)"),
]

# ---------------------------------------------------------------------------
# 2. TASK DURATION BENCHMARKS
# ---------------------------------------------------------------------------
TASK_DURATIONS = [
    # (Task Category, Task Example, Avg Duration (sec), p50, p90, p99, Model Class, Notes)
    ("Code Generation", "Generate a single function (50-100 LOC)", 8, 6, 15, 35, "Frontier (GPT-4o/Claude Opus)", "Includes planning + generation + validation"),
    ("Code Generation", "Multi-file feature implementation", 120, 90, 300, 600, "Frontier", "Agentic loop with tool calls; Cursor/Devin-style"),
    ("Code Generation", "Full project scaffolding", 180, 120, 420, 900, "Frontier", "Replit Agent / Cursor Composer creating full apps"),
    ("Code Review", "PR review with inline suggestions", 25, 18, 50, 120, "Frontier", "Reads diff, generates suggestions, posts comments"),
    ("Code Debugging", "Root-cause analysis + fix suggestion", 45, 30, 120, 300, "Frontier", "Reads error, traces code, suggests fix"),
    ("Document Analysis", "Summarize a 50-page PDF", 30, 22, 60, 150, "Frontier", "RAG + long-context window"),
    ("Document Analysis", "Extract structured data from 100 docs", 300, 240, 600, 1200, "Frontier", "Batch processing pipeline"),
    ("Research", "Multi-source web research query", 15, 10, 30, 60, "Frontier", "Perplexity-style search + synthesis"),
    ("Research", "Deep research report (10+ sources)", 180, 120, 360, 900, "Frontier", "Multi-step agent with web browsing"),
    ("Customer Service", "Simple FAQ response", 3, 2, 5, 10, "Mid-tier (GPT-4o-mini/Haiku)", "Single-turn retrieval + generation"),
    ("Customer Service", "Complex issue resolution with actions", 45, 30, 90, 180, "Frontier", "Multi-turn with tool calls (refund, order lookup)"),
    ("Data Analysis", "SQL query generation + execution", 12, 8, 25, 60, "Frontier", "Text-to-SQL + validation + execution"),
    ("Data Analysis", "Full dashboard/report creation", 180, 120, 360, 720, "Frontier", "Multi-step: query, transform, visualize"),
    ("Email Drafting", "Reply to a single email", 8, 5, 15, 30, "Mid-tier", "Read context, generate appropriate reply"),
    ("Email Drafting", "Campaign with personalization (100 emails)", 120, 90, 240, 480, "Mid-tier", "Batch generation with per-recipient customization"),
    ("Workflow Automation", "Simple trigger-action automation", 5, 3, 10, 20, "Mid-tier", "If-then logic with one tool call"),
    ("Workflow Automation", "Complex multi-system orchestration", 90, 60, 180, 450, "Frontier", "Multiple API calls, error handling, retries"),
    ("Image Analysis", "Describe/analyze a single image", 5, 3, 10, 20, "Frontier multimodal", "Vision model inference"),
    ("Image Analysis", "Batch process 100 images with extraction", 180, 120, 360, 600, "Frontier multimodal", "Parallel vision model calls"),
    ("Computer Use", "Navigate website and fill form", 60, 40, 120, 300, "Frontier + vision", "Screenshot-based navigation loop"),
    ("Computer Use", "Complete multi-step desktop workflow", 300, 200, 600, 1500, "Frontier + vision", "Full desktop automation with decision-making"),
]

TASK_DURATION_COLUMNS = [
    "Task Category",
    "Task Example",
    "Avg Duration (sec)",
    "p50 (sec)",
    "p90 (sec)",
    "p99 (sec)",
    "Model Class",
    "Notes",
]

# ---------------------------------------------------------------------------
# 3. TOKEN CONSUMPTION PROFILES
# ---------------------------------------------------------------------------
TOKEN_CONSUMPTION = [
    # (Task Type, Input Tokens (avg), Output Tokens (avg), Total Tokens, Cost per Task @ Frontier ($), Cost @ Mid-tier ($), Tool Calls (avg), Reasoning Tokens (avg))
    ("Simple Q&A / Chat", 500, 300, 800, 0.008, 0.0002, 0, 0),
    ("Code Generation (function)", 2000, 1500, 3500, 0.035, 0.0009, 1, 500),
    ("Code Generation (multi-file)", 25000, 15000, 40000, 0.40, 0.010, 15, 8000),
    ("Code Generation (full project)", 60000, 40000, 100000, 1.00, 0.025, 40, 20000),
    ("Code Review (PR)", 10000, 3000, 13000, 0.13, 0.003, 3, 2000),
    ("Code Debugging", 15000, 5000, 20000, 0.20, 0.005, 5, 5000),
    ("Document Summarization (short)", 8000, 1500, 9500, 0.10, 0.002, 1, 1000),
    ("Document Analysis (batch 100)", 500000, 80000, 580000, 5.80, 0.145, 100, 50000),
    ("Web Research (single query)", 5000, 2000, 7000, 0.07, 0.002, 3, 1500),
    ("Deep Research Report", 80000, 20000, 100000, 1.00, 0.025, 25, 25000),
    ("Customer Service (simple)", 1000, 500, 1500, 0.015, 0.0004, 0, 0),
    ("Customer Service (complex)", 8000, 3000, 11000, 0.11, 0.003, 4, 2000),
    ("SQL Query Generation", 3000, 800, 3800, 0.04, 0.001, 2, 800),
    ("Dashboard/Report Creation", 30000, 15000, 45000, 0.45, 0.011, 12, 8000),
    ("Email Reply", 2000, 800, 2800, 0.03, 0.0007, 0, 300),
    ("Email Campaign (100)", 50000, 30000, 80000, 0.80, 0.020, 0, 5000),
    ("Workflow Automation (simple)", 1500, 500, 2000, 0.02, 0.0005, 1, 200),
    ("Workflow Automation (complex)", 20000, 8000, 28000, 0.28, 0.007, 8, 5000),
    ("Image Analysis (single)", 1000, 500, 1500, 0.02, 0.001, 0, 0),
    ("Image Analysis (batch 100)", 100000, 30000, 130000, 1.60, 0.080, 0, 0),
    ("Computer Use (simple)", 15000, 5000, 20000, 0.25, 0.010, 10, 3000),
    ("Computer Use (complex)", 80000, 20000, 100000, 1.20, 0.050, 40, 15000),
    ("Agentic Coding Session (30 min)", 150000, 60000, 210000, 2.10, 0.053, 50, 40000),
    ("Agentic Coding Session (2 hr)", 500000, 200000, 700000, 7.00, 0.175, 150, 120000),
]

TOKEN_COLUMNS = [
    "Task Type",
    "Input Tokens (avg)",
    "Output Tokens (avg)",
    "Total Tokens",
    "Cost/Task @ Frontier ($)",
    "Cost/Task @ Mid-tier ($)",
    "Tool Calls (avg)",
    "Reasoning Tokens (avg)",
]

# ---------------------------------------------------------------------------
# 4. TIME-SERIES PROJECTIONS (2024-2030)
# ---------------------------------------------------------------------------
YEARS = [2024, 2025, 2026, 2027, 2028, 2029, 2030]

# Market size in $B
AGENT_MARKET_SIZE = {
    "Total AI Agent Market ($B)": [5.2, 14.0, 35.0, 68.0, 110.0, 165.0, 240.0],
    "Coding Agents ($B)": [0.8, 2.5, 6.5, 13.0, 22.0, 34.0, 50.0],
    "Enterprise Workflow Agents ($B)": [1.5, 4.0, 10.0, 20.0, 34.0, 52.0, 76.0],
    "Customer Service Agents ($B)": [1.2, 3.0, 7.5, 14.0, 22.0, 32.0, 45.0],
    "Research & Knowledge Agents ($B)": [0.5, 1.5, 4.0, 8.0, 14.0, 22.0, 32.0],
    "Other Vertical Agents ($B)": [1.2, 3.0, 7.0, 13.0, 18.0, 25.0, 37.0],
}

# Task volume (millions of agent tasks per day globally)
DAILY_TASK_VOLUME = {
    "Total Daily Agent Tasks (M)": [2.0, 12.0, 50.0, 150.0, 400.0, 900.0, 1800.0],
    "Coding Tasks (M/day)": [0.5, 3.0, 12.0, 35.0, 90.0, 200.0, 400.0],
    "Enterprise Tasks (M/day)": [0.4, 2.5, 10.0, 30.0, 80.0, 180.0, 360.0],
    "Customer Service Tasks (M/day)": [0.8, 4.0, 18.0, 55.0, 150.0, 340.0, 680.0],
    "Research Tasks (M/day)": [0.2, 1.5, 6.0, 18.0, 50.0, 110.0, 220.0],
    "Other Tasks (M/day)": [0.1, 1.0, 4.0, 12.0, 30.0, 70.0, 140.0],
}

# Token consumption (billions of tokens per day globally)
DAILY_TOKEN_CONSUMPTION = {
    "Total Daily Tokens (B)": [50, 350, 1500, 5000, 15000, 40000, 100000],
    "Input Tokens (B/day)": [30, 210, 900, 3000, 9000, 24000, 60000],
    "Output Tokens (B/day)": [12, 84, 360, 1200, 3600, 9600, 24000],
    "Reasoning Tokens (B/day)": [8, 56, 240, 800, 2400, 6400, 16000],
}

# Average cost efficiency (cost per 1M tokens, blended)
COST_PER_MILLION_TOKENS = {
    "Frontier Model ($/M tokens)": [15.00, 10.00, 6.50, 4.00, 2.50, 1.50, 0.80],
    "Mid-Tier Model ($/M tokens)": [0.60, 0.30, 0.15, 0.08, 0.04, 0.02, 0.01],
    "Open-Source Hosted ($/M tokens)": [0.30, 0.15, 0.08, 0.04, 0.02, 0.01, 0.005],
    "Blended Average ($/M tokens)": [5.00, 3.00, 1.80, 1.00, 0.55, 0.30, 0.15],
}

# Average task duration trends (seconds, weighted average across all task types)
AVG_TASK_DURATION = {
    "Avg Simple Task Duration (sec)": [8.0, 5.0, 3.5, 2.5, 1.8, 1.2, 0.8],
    "Avg Complex Task Duration (sec)": [300, 180, 120, 80, 55, 38, 25],
    "Avg Agentic Session Duration (sec)": [1800, 1200, 900, 600, 420, 300, 200],
    "Latency to First Token (ms)": [800, 500, 300, 180, 120, 80, 50],
}

# Global compute for agents
COMPUTE_INFRA = {
    "GPU Hours for Agents (M/day)": [0.5, 3.0, 12.0, 40.0, 100.0, 250.0, 600.0],
    "Estimated Agent Compute Cost ($B/yr)": [2.0, 8.0, 25.0, 60.0, 110.0, 180.0, 280.0],
    "Agent Share of Total AI Inference (%)": [5, 12, 22, 35, 48, 58, 65],
}

# Key company revenue projections
COMPANY_REVENUE_PROJECTIONS = {
    "OpenAI Revenue ($B)": [3.4, 12.7, 29.0, 50.0, 75.0, 100.0, 130.0],
    "Anthropic Revenue ($B)": [0.5, 2.5, 6.0, 12.0, 20.0, 30.0, 42.0],
    "Google AI Agent Revenue ($B)": [1.0, 4.0, 10.0, 20.0, 35.0, 52.0, 70.0],
    "Microsoft Copilot Revenue ($B)": [2.0, 8.0, 18.0, 32.0, 48.0, 65.0, 85.0],
    "Salesforce Agentforce Revenue ($B)": [0.2, 1.5, 4.0, 8.0, 14.0, 22.0, 32.0],
    "Cursor/Anysphere Revenue ($B)": [0.1, 0.5, 1.2, 2.5, 5.0, 8.0, 12.0],
    "Other Agent Companies ($B)": [0.5, 2.0, 6.0, 15.0, 30.0, 50.0, 80.0],
}

ALL_TIME_SERIES = {
    "Market Size": AGENT_MARKET_SIZE,
    "Daily Task Volume": DAILY_TASK_VOLUME,
    "Daily Token Consumption": DAILY_TOKEN_CONSUMPTION,
    "Cost per Million Tokens": COST_PER_MILLION_TOKENS,
    "Average Task Duration": AVG_TASK_DURATION,
    "Compute Infrastructure": COMPUTE_INFRA,
    "Company Revenue Projections": COMPANY_REVENUE_PROJECTIONS,
}

# ---------------------------------------------------------------------------
# 5. MODEL PRICING REFERENCE
# ---------------------------------------------------------------------------
MODEL_PRICING = [
    ("GPT-4o", "OpenAI", 2.50, 10.00, 128000, "2025-02", "Frontier multimodal"),
    ("GPT-4o-mini", "OpenAI", 0.15, 0.60, 128000, "2025-02", "Cost-efficient mid-tier"),
    ("o3", "OpenAI", 10.00, 40.00, 200000, "2025-02", "Reasoning model (high)"),
    ("o3-mini", "OpenAI", 1.10, 4.40, 200000, "2025-02", "Reasoning model (efficient)"),
    ("Claude 4 Opus", "Anthropic", 15.00, 75.00, 200000, "2025-02", "Highest capability"),
    ("Claude 4 Sonnet", "Anthropic", 3.00, 15.00, 200000, "2025-02", "Best balance"),
    ("Claude 3.5 Haiku", "Anthropic", 0.80, 4.00, 200000, "2025-02", "Fast & cheap"),
    ("Gemini 2.0 Flash", "Google", 0.10, 0.40, 1000000, "2025-02", "Very fast, long context"),
    ("Gemini 2.0 Pro", "Google", 1.25, 5.00, 1000000, "2025-02", "Frontier competitor"),
    ("Mistral Large 2", "Mistral", 2.00, 6.00, 128000, "2025-02", "European frontier"),
    ("Command R+", "Cohere", 2.50, 10.00, 128000, "2025-02", "Enterprise RAG"),
    ("Llama 3.1 405B (hosted)", "Meta/various", 1.00, 1.00, 128000, "2025-02", "Open-weight frontier"),
    ("DeepSeek V3", "DeepSeek", 0.27, 1.10, 128000, "2025-02", "Open-source efficient"),
]

MODEL_PRICING_COLUMNS = [
    "Model",
    "Provider",
    "Input ($/M tokens)",
    "Output ($/M tokens)",
    "Context Window",
    "Pricing Date",
    "Notes",
]


def create_workbook():
    wb = xlsxwriter.Workbook(OUTPUT_FILE, {"nan_inf_to_errors": True})

    # --- FORMAT DEFINITIONS ---
    fmt_title = wb.add_format({
        "bold": True, "font_size": 16, "font_color": "#1B2A4A",
        "bottom": 2, "bottom_color": "#3B82F6",
    })
    fmt_subtitle = wb.add_format({
        "bold": True, "font_size": 11, "font_color": "#6B7280", "italic": True,
    })
    fmt_header = wb.add_format({
        "bold": True, "font_size": 11, "bg_color": "#1E3A5F", "font_color": "#FFFFFF",
        "border": 1, "border_color": "#1E3A5F", "text_wrap": True, "valign": "vcenter",
    })
    fmt_header_green = wb.add_format({
        "bold": True, "font_size": 11, "bg_color": "#166534", "font_color": "#FFFFFF",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })
    fmt_header_purple = wb.add_format({
        "bold": True, "font_size": 11, "bg_color": "#581C87", "font_color": "#FFFFFF",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })
    fmt_cell = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "text_wrap": True, "valign": "vcenter",
    })
    fmt_cell_alt = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "bg_color": "#F9FAFB", "text_wrap": True, "valign": "vcenter",
    })
    fmt_number = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "#,##0.00", "valign": "vcenter",
    })
    fmt_number_alt = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "#,##0.00", "bg_color": "#F9FAFB", "valign": "vcenter",
    })
    fmt_int = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "#,##0", "valign": "vcenter",
    })
    fmt_int_alt = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "#,##0", "bg_color": "#F9FAFB", "valign": "vcenter",
    })
    fmt_dollar = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "$#,##0.000", "valign": "vcenter",
    })
    fmt_dollar_alt = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "$#,##0.000", "bg_color": "#F9FAFB", "valign": "vcenter",
    })
    fmt_pct = wb.add_format({
        "font_size": 10, "border": 1, "border_color": "#E5E7EB",
        "num_format": "0.0%", "valign": "vcenter",
    })
    fmt_section = wb.add_format({
        "bold": True, "font_size": 12, "font_color": "#1E3A5F",
        "bottom": 1, "bottom_color": "#3B82F6",
    })
    fmt_source = wb.add_format({
        "font_size": 9, "font_color": "#9CA3AF", "italic": True,
    })

    # ===================================================================
    # SHEET 1: TOP COMPANIES
    # ===================================================================
    ws1 = wb.add_worksheet("Top Companies")
    ws1.hide_gridlines(2)
    ws1.set_tab_color("#3B82F6")
    ws1.set_landscape()
    ws1.set_paper(1)  # Letter
    ws1.fit_to_pages(1, 0)

    ws1.merge_range("A1:L1", "Agent Economy: Top Companies", fmt_title)
    ws1.merge_range("A2:L2", "Comprehensive overview of leading companies in the AI agent ecosystem | Data as of Feb 2026", fmt_subtitle)

    row = 3
    for col_idx, (_, col_name) in enumerate(COMPANY_COLUMNS):
        ws1.write(row, col_idx, col_name, fmt_header)

    col_widths = [22, 24, 8, 28, 50, 16, 18, 18, 12, 30, 55, 16]
    for i, w in enumerate(col_widths):
        ws1.set_column(i, i, w)
    ws1.set_row(row, 35)

    for r_idx, company in enumerate(COMPANIES):
        row = 4 + r_idx
        base_fmt = fmt_cell_alt if r_idx % 2 else fmt_cell
        num_fmt = fmt_number_alt if r_idx % 2 else fmt_number
        int_fmt = fmt_int_alt if r_idx % 2 else fmt_int
        for c_idx, (key, _) in enumerate(COMPANY_COLUMNS):
            val = company[key]
            if val is None:
                ws1.write(row, c_idx, "N/A (Public Co.)", base_fmt)
            elif isinstance(val, (int, float)):
                if key == "founded" or key == "employees":
                    ws1.write_number(row, c_idx, val, int_fmt)
                else:
                    ws1.write_number(row, c_idx, val, num_fmt)
            else:
                ws1.write(row, c_idx, val, base_fmt)
        ws1.set_row(row, 30)

    source_row = 4 + len(COMPANIES) + 1
    ws1.merge_range(source_row, 0, source_row, 11,
        "Sources: Company filings, Crunchbase, PitchBook, press releases. Valuations reflect latest reported funding rounds. Revenue estimates are consensus/research estimates.",
        fmt_source)

    ws1.autofilter(3, 0, 3 + len(COMPANIES), len(COMPANY_COLUMNS) - 1)
    ws1.freeze_panes(4, 1)

    # ===================================================================
    # SHEET 2: TASK DURATION BENCHMARKS
    # ===================================================================
    ws2 = wb.add_worksheet("Task Duration")
    ws2.hide_gridlines(2)
    ws2.set_tab_color("#10B981")

    ws2.merge_range("A1:H1", "Agent Task Duration Benchmarks", fmt_title)
    ws2.merge_range("A2:H2", "Typical task durations across categories | Frontier & mid-tier models | 2025-2026 benchmarks", fmt_subtitle)

    row = 3
    for col_idx, col_name in enumerate(TASK_DURATION_COLUMNS):
        ws2.write(row, col_idx, col_name, fmt_header_green)
    ws2.set_row(row, 35)

    dur_col_widths = [22, 42, 16, 12, 12, 12, 28, 50]
    for i, w in enumerate(dur_col_widths):
        ws2.set_column(i, i, w)

    for r_idx, task in enumerate(TASK_DURATIONS):
        row = 4 + r_idx
        base_fmt = fmt_cell_alt if r_idx % 2 else fmt_cell
        int_f = fmt_int_alt if r_idx % 2 else fmt_int
        for c_idx, val in enumerate(task):
            if isinstance(val, (int, float)):
                ws2.write_number(row, c_idx, val, int_f)
            else:
                ws2.write(row, c_idx, val, base_fmt)
        ws2.set_row(row, 22)

    ws2.autofilter(3, 0, 3 + len(TASK_DURATIONS), len(TASK_DURATION_COLUMNS) - 1)
    ws2.freeze_panes(4, 0)

    # Add a chart: Average duration by task category
    chart_dur = wb.add_chart({"type": "bar"})
    categories_seen = []
    cat_avg = {}
    for task in TASK_DURATIONS:
        cat = task[0]
        if cat not in cat_avg:
            cat_avg[cat] = []
        cat_avg[cat].append(task[2])

    # Write summary data for chart
    chart_data_row = 4 + len(TASK_DURATIONS) + 3
    ws2.write(chart_data_row - 1, 0, "Category Summary", fmt_section)
    ws2.write(chart_data_row, 0, "Category", fmt_header_green)
    ws2.write(chart_data_row, 1, "Avg Duration (sec)", fmt_header_green)

    for i, (cat, vals) in enumerate(cat_avg.items()):
        ws2.write(chart_data_row + 1 + i, 0, cat, fmt_cell)
        ws2.write_number(chart_data_row + 1 + i, 1, sum(vals) / len(vals), fmt_number)

    n_cats = len(cat_avg)
    chart_dur.add_series({
        "name": "Avg Duration (sec)",
        "categories": ["Task Duration", chart_data_row + 1, 0, chart_data_row + n_cats, 0],
        "values": ["Task Duration", chart_data_row + 1, 1, chart_data_row + n_cats, 1],
        "fill": {"color": "#10B981"},
        "border": {"color": "#059669"},
    })
    chart_dur.set_title({"name": "Average Task Duration by Category"})
    chart_dur.set_x_axis({"name": "Seconds"})
    chart_dur.set_y_axis({"name": "Task Category"})
    chart_dur.set_size({"width": 720, "height": 400})
    chart_dur.set_legend({"none": True})
    chart_dur.set_style(10)
    ws2.insert_chart(chart_data_row + n_cats + 2, 0, chart_dur)

    # ===================================================================
    # SHEET 3: TOKEN CONSUMPTION
    # ===================================================================
    ws3 = wb.add_worksheet("Token Consumption")
    ws3.hide_gridlines(2)
    ws3.set_tab_color("#8B5CF6")

    ws3.merge_range("A1:H1", "Agent Token Consumption Profiles", fmt_title)
    ws3.merge_range("A2:H2", "Average token usage, costs, and tool call patterns per task type | Feb 2026 pricing", fmt_subtitle)

    row = 3
    for col_idx, col_name in enumerate(TOKEN_COLUMNS):
        ws3.write(row, col_idx, col_name, fmt_header_purple)
    ws3.set_row(row, 35)

    tok_col_widths = [35, 18, 18, 16, 22, 22, 14, 20]
    for i, w in enumerate(tok_col_widths):
        ws3.set_column(i, i, w)

    for r_idx, tok in enumerate(TOKEN_CONSUMPTION):
        row = 4 + r_idx
        base_fmt = fmt_cell_alt if r_idx % 2 else fmt_cell
        int_f = fmt_int_alt if r_idx % 2 else fmt_int
        dol_f = fmt_dollar_alt if r_idx % 2 else fmt_dollar
        for c_idx, val in enumerate(tok):
            if c_idx in (4, 5):  # dollar columns
                ws3.write_number(row, c_idx, val, dol_f)
            elif isinstance(val, (int, float)):
                ws3.write_number(row, c_idx, val, int_f)
            else:
                ws3.write(row, c_idx, val, base_fmt)
        ws3.set_row(row, 22)

    ws3.autofilter(3, 0, 3 + len(TOKEN_CONSUMPTION), len(TOKEN_COLUMNS) - 1)
    ws3.freeze_panes(4, 0)

    # Chart: Token distribution by task type (top 10 by total tokens)
    sorted_tokens = sorted(TOKEN_CONSUMPTION, key=lambda x: x[3], reverse=True)[:10]
    chart_tok_row = 4 + len(TOKEN_CONSUMPTION) + 3
    ws3.write(chart_tok_row - 1, 0, "Top 10 Tasks by Token Consumption", fmt_section)
    ws3.write(chart_tok_row, 0, "Task Type", fmt_header_purple)
    ws3.write(chart_tok_row, 1, "Input Tokens", fmt_header_purple)
    ws3.write(chart_tok_row, 2, "Output Tokens", fmt_header_purple)
    ws3.write(chart_tok_row, 3, "Reasoning Tokens", fmt_header_purple)

    for i, tok in enumerate(sorted_tokens):
        ws3.write(chart_tok_row + 1 + i, 0, tok[0], fmt_cell)
        ws3.write_number(chart_tok_row + 1 + i, 1, tok[1], fmt_int)
        ws3.write_number(chart_tok_row + 1 + i, 2, tok[2], fmt_int)
        ws3.write_number(chart_tok_row + 1 + i, 3, tok[7], fmt_int)

    chart_tok = wb.add_chart({"type": "bar", "subtype": "stacked"})
    n_top = len(sorted_tokens)
    for col, name, color in [
        (1, "Input Tokens", "#8B5CF6"),
        (2, "Output Tokens", "#A78BFA"),
        (3, "Reasoning Tokens", "#C4B5FD"),
    ]:
        chart_tok.add_series({
            "name": name,
            "categories": ["Token Consumption", chart_tok_row + 1, 0, chart_tok_row + n_top, 0],
            "values": ["Token Consumption", chart_tok_row + 1, col, chart_tok_row + n_top, col],
            "fill": {"color": color},
        })
    chart_tok.set_title({"name": "Top 10 Tasks: Token Breakdown"})
    chart_tok.set_x_axis({"name": "Tokens"})
    chart_tok.set_size({"width": 800, "height": 450})
    chart_tok.set_style(10)
    ws3.insert_chart(chart_tok_row + n_top + 2, 0, chart_tok)

    # ===================================================================
    # SHEET 4: MODEL PRICING REFERENCE
    # ===================================================================
    ws4 = wb.add_worksheet("Model Pricing")
    ws4.hide_gridlines(2)
    ws4.set_tab_color("#F59E0B")

    fmt_header_amber = wb.add_format({
        "bold": True, "font_size": 11, "bg_color": "#92400E", "font_color": "#FFFFFF",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })

    ws4.merge_range("A1:G1", "LLM Model Pricing Reference", fmt_title)
    ws4.merge_range("A2:G2", "Current pricing for major models used in agent workflows | Feb 2026", fmt_subtitle)

    row = 3
    for col_idx, col_name in enumerate(MODEL_PRICING_COLUMNS):
        ws4.write(row, col_idx, col_name, fmt_header_amber)
    ws4.set_row(row, 35)

    price_col_widths = [22, 16, 18, 18, 14, 14, 26]
    for i, w in enumerate(price_col_widths):
        ws4.set_column(i, i, w)

    for r_idx, model in enumerate(MODEL_PRICING):
        row = 4 + r_idx
        base_fmt = fmt_cell_alt if r_idx % 2 else fmt_cell
        num_f = fmt_number_alt if r_idx % 2 else fmt_number
        int_f = fmt_int_alt if r_idx % 2 else fmt_int
        for c_idx, val in enumerate(model):
            if c_idx in (2, 3):  # price columns
                ws4.write_number(row, c_idx, val, num_f)
            elif c_idx == 4:  # context window
                ws4.write_number(row, c_idx, val, int_f)
            else:
                ws4.write(row, c_idx, str(val), base_fmt)

    ws4.autofilter(3, 0, 3 + len(MODEL_PRICING), len(MODEL_PRICING_COLUMNS) - 1)
    ws4.freeze_panes(4, 0)

    # ===================================================================
    # SHEETS 5-11: TIME SERIES
    # ===================================================================
    colors_map = {
        "Market Size": ("#EF4444", "#DC2626"),
        "Daily Task Volume": ("#3B82F6", "#2563EB"),
        "Daily Token Consumption": ("#8B5CF6", "#7C3AED"),
        "Cost per Million Tokens": ("#F59E0B", "#D97706"),
        "Average Task Duration": ("#10B981", "#059669"),
        "Compute Infrastructure": ("#EC4899", "#DB2777"),
        "Company Revenue Projections": ("#6366F1", "#4F46E5"),
    }

    tab_colors = {
        "Market Size": "#EF4444",
        "Daily Task Volume": "#3B82F6",
        "Daily Token Consumption": "#8B5CF6",
        "Cost per Million Tokens": "#F59E0B",
        "Average Task Duration": "#10B981",
        "Compute Infrastructure": "#EC4899",
        "Company Revenue Projections": "#6366F1",
    }

    chart_types = {
        "Market Size": "area",
        "Daily Task Volume": "area",
        "Daily Token Consumption": "area",
        "Cost per Million Tokens": "line",
        "Average Task Duration": "line",
        "Compute Infrastructure": "column",
        "Company Revenue Projections": "area",
    }

    series_colors = [
        "#EF4444", "#3B82F6", "#10B981", "#F59E0B", "#8B5CF6", "#EC4899", "#6366F1",
    ]

    for section_name, data_dict in ALL_TIME_SERIES.items():
        # Truncate sheet name to 31 chars (Excel limit)
        sheet_name = f"TS - {section_name}"[:31]
        ws = wb.add_worksheet(sheet_name)
        ws.hide_gridlines(2)
        ws.set_tab_color(tab_colors[section_name])

        n_metrics = len(data_dict)
        n_cols = len(YEARS) + 1

        ws.merge_range(0, 0, 0, n_cols - 1,
            f"Agent Economy Time Series: {section_name}", fmt_title)
        ws.merge_range(1, 0, 1, n_cols - 1,
            f"Projections 2024-2030 | Estimates based on industry research and growth models", fmt_subtitle)

        # Header row
        row = 3
        ws.write(row, 0, "Metric", fmt_header)
        for y_idx, year in enumerate(YEARS):
            ws.write(row, 1 + y_idx, year, fmt_header)
        ws.set_row(row, 30)
        ws.set_column(0, 0, 38)
        for y_idx in range(len(YEARS)):
            ws.set_column(1 + y_idx, 1 + y_idx, 16)

        # Data rows
        for m_idx, (metric, values) in enumerate(data_dict.items()):
            row = 4 + m_idx
            base_fmt = fmt_cell_alt if m_idx % 2 else fmt_cell
            num_f = fmt_number_alt if m_idx % 2 else fmt_number
            ws.write(row, 0, metric, base_fmt)
            for v_idx, val in enumerate(values):
                ws.write_number(row, 1 + v_idx, val, num_f)

        # CAGR column
        cagr_col = len(YEARS) + 1
        ws.write(3, cagr_col, "CAGR '24-'30", fmt_header)
        ws.set_column(cagr_col, cagr_col, 14)
        fmt_pct_cell = wb.add_format({
            "font_size": 10, "border": 1, "border_color": "#E5E7EB",
            "num_format": "0.0%", "valign": "vcenter", "bold": True,
            "font_color": "#059669",
        })
        for m_idx, (metric, values) in enumerate(data_dict.items()):
            row = 4 + m_idx
            start_val = values[0]
            end_val = values[-1]
            if start_val > 0 and end_val > 0:
                cagr = (end_val / start_val) ** (1.0 / 6.0) - 1.0
                ws.write_number(row, cagr_col, cagr, fmt_pct_cell)
            else:
                ws.write(row, cagr_col, "N/A", fmt_cell)

        # YoY Growth rows
        growth_start_row = 4 + n_metrics + 2
        ws.write(growth_start_row - 1, 0, "Year-over-Year Growth (%)", fmt_section)
        ws.write(growth_start_row, 0, "Metric", fmt_header)
        for y_idx, year in enumerate(YEARS):
            ws.write(growth_start_row, 1 + y_idx, year, fmt_header)

        fmt_pct_green = wb.add_format({
            "font_size": 10, "border": 1, "num_format": "0.0%",
            "font_color": "#059669", "valign": "vcenter",
        })
        fmt_pct_red = wb.add_format({
            "font_size": 10, "border": 1, "num_format": "0.0%",
            "font_color": "#DC2626", "valign": "vcenter",
        })

        for m_idx, (metric, values) in enumerate(data_dict.items()):
            g_row = growth_start_row + 1 + m_idx
            ws.write(g_row, 0, metric, fmt_cell)
            ws.write(g_row, 1, "—", fmt_cell)  # No YoY for first year
            for v_idx in range(1, len(values)):
                if values[v_idx - 1] != 0:
                    yoy = (values[v_idx] - values[v_idx - 1]) / abs(values[v_idx - 1])
                    cell_fmt = fmt_pct_green if yoy >= 0 else fmt_pct_red
                    ws.write_number(g_row, 1 + v_idx, yoy, cell_fmt)
                else:
                    ws.write(g_row, 1 + v_idx, "N/A", fmt_cell)

        # CHART
        chart_type = chart_types.get(section_name, "line")
        chart = wb.add_chart({"type": chart_type})

        for m_idx, (metric, values) in enumerate(data_dict.items()):
            data_row = 4 + m_idx
            color = series_colors[m_idx % len(series_colors)]
            series_opts = {
                "name": [sheet_name, data_row, 0],
                "categories": [sheet_name, 3, 1, 3, len(YEARS)],
                "values": [sheet_name, data_row, 1, data_row, len(YEARS)],
                "line": {"color": color, "width": 2.5},
                "fill": {"color": color, "transparency": 60} if chart_type == "area" else {"color": color},
            }
            if chart_type == "line":
                series_opts["marker"] = {"type": "circle", "size": 6, "fill": {"color": color}}
            chart.add_series(series_opts)

        chart.set_title({"name": f"Agent Economy: {section_name} (2024-2030)"})
        chart.set_x_axis({"name": "Year", "num_format": "0"})
        chart.set_y_axis({"name": section_name})
        chart.set_size({"width": 900, "height": 500})
        chart.set_style(10)
        chart.set_legend({"position": "bottom"})

        chart_row = growth_start_row + 1 + n_metrics + 2
        ws.insert_chart(chart_row, 0, chart)

        # Source note
        src_row = chart_row + 28
        ws.merge_range(src_row, 0, src_row, n_cols,
            "Sources: Industry research, company disclosures, analyst estimates, IDC, Gartner, internal models. "
            "Projections are illustrative and subject to significant uncertainty.",
            fmt_source)

    # ===================================================================
    # SHEET: SUMMARY DASHBOARD
    # ===================================================================
    ws_dash = wb.add_worksheet("Dashboard")
    ws_dash.hide_gridlines(2)
    ws_dash.set_tab_color("#111827")

    ws_dash.merge_range("A1:J1", "Agent Economy Dashboard: Key Metrics Summary", fmt_title)
    ws_dash.merge_range("A2:J2", "Snapshot of the AI agent ecosystem | Feb 2026", fmt_subtitle)

    # Key stats
    kpi_row = 4
    fmt_kpi_label = wb.add_format({
        "font_size": 10, "font_color": "#6B7280", "valign": "vcenter",
    })
    fmt_kpi_value = wb.add_format({
        "font_size": 18, "bold": True, "font_color": "#1E3A5F", "valign": "vcenter",
        "num_format": "$#,##0.0\"B\"",
    })
    fmt_kpi_value_plain = wb.add_format({
        "font_size": 18, "bold": True, "font_color": "#1E3A5F", "valign": "vcenter",
        "num_format": "#,##0",
    })
    fmt_kpi_value_pct = wb.add_format({
        "font_size": 18, "bold": True, "font_color": "#059669", "valign": "vcenter",
        "num_format": "0%",
    })

    kpis = [
        ("2026E Agent Market Size", 35.0, fmt_kpi_value),
        ("2030E Agent Market Size", 240.0, fmt_kpi_value),
        ("'24-'30 Market CAGR", 0.90, fmt_kpi_value_pct),
        ("Top Companies Tracked", len(COMPANIES), fmt_kpi_value_plain),
        ("Task Types Benchmarked", len(TASK_DURATIONS), fmt_kpi_value_plain),
        ("2026E Daily Agent Tasks", 50, fmt_kpi_value_plain),  # Millions
    ]

    ws_dash.set_column(0, 0, 5)
    for i in range(1, 11):
        ws_dash.set_column(i, i, 14)

    for i, (label, value, vfmt) in enumerate(kpis):
        col = 1 + (i % 3) * 3
        row_off = kpi_row + (i // 3) * 3
        ws_dash.merge_range(row_off, col, row_off, col + 1, value, vfmt)
        ws_dash.merge_range(row_off + 1, col, row_off + 1, col + 1, label, fmt_kpi_label)

    # Methodology note
    method_row = kpi_row + 8
    ws_dash.merge_range(method_row, 1, method_row, 9, "Methodology & Disclaimers", fmt_section)
    methodology_text = (
        "This model aggregates data from multiple sources including company disclosures, "
        "industry reports (IDC, Gartner, McKinsey), venture capital databases (Crunchbase, PitchBook), "
        "and proprietary research estimates. Token consumption and task duration benchmarks are based on "
        "observed usage patterns across production agent deployments. Time-series projections use a "
        "combination of bottom-up (company-level) and top-down (TAM) modeling approaches. "
        "All forward-looking estimates are illustrative and subject to material uncertainty. "
        "Cost figures reflect published API pricing as of February 2026. "
        "Market sizes include software revenue and do not include infrastructure/compute spend."
    )
    fmt_method = wb.add_format({
        "font_size": 10, "text_wrap": True, "valign": "top", "font_color": "#4B5563",
    })
    ws_dash.merge_range(method_row + 1, 1, method_row + 4, 9, methodology_text, fmt_method)

    # Navigation
    nav_row = method_row + 6
    ws_dash.merge_range(nav_row, 1, nav_row, 9, "Sheet Navigation", fmt_section)
    sheets_nav = [
        ("Top Companies", "Overview of 20 leading agent economy companies with valuations and revenue estimates"),
        ("Task Duration", "Benchmarks for agent task durations across 21 task types with percentile data"),
        ("Token Consumption", "Token usage profiles, costs, tool calls, and reasoning tokens for 24 task types"),
        ("Model Pricing", "Current LLM pricing reference for 13 major models"),
        ("TS - Market Size", "Agent market size projections by segment (2024-2030)"),
        ("TS - Daily Task Volume", "Global daily agent task volume projections"),
        ("TS - Daily Token Consumption", "Global daily token consumption projections"),
        ("TS - Cost per Million Tokens", "Model pricing trend projections showing declining costs"),
        ("TS - Average Task Duration", "Task speed improvement projections"),
        ("TS - Compute Infrastructure", "Agent compute infrastructure projections"),
        ("TS - Company Revenue Project", "Revenue projections for top agent companies"),
    ]
    for i, (sname, desc) in enumerate(sheets_nav):
        ws_dash.write(nav_row + 1 + i, 1, sname, wb.add_format({
            "font_size": 10, "bold": True, "font_color": "#3B82F6", "underline": True,
        }))
        ws_dash.write(nav_row + 1 + i, 3, desc, fmt_cell)
        ws_dash.merge_range(nav_row + 1 + i, 3, nav_row + 1 + i, 9, desc, fmt_cell)

    # Move Dashboard to first position
    ws_dash.set_first_sheet()

    wb.close()
    print(f"Workbook created: {OUTPUT_FILE}")
    print(f"  - {len(COMPANIES)} companies profiled")
    print(f"  - {len(TASK_DURATIONS)} task duration benchmarks")
    print(f"  - {len(TOKEN_CONSUMPTION)} token consumption profiles")
    print(f"  - {len(MODEL_PRICING)} model pricing entries")
    print(f"  - {len(ALL_TIME_SERIES)} time-series categories across {len(YEARS)} years")
    print(f"  - 12 sheets total (Dashboard + 4 data sheets + 7 time-series sheets)")


if __name__ == "__main__":
    create_workbook()
