import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Styling constants ──
header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
subheader_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
subheader_font = Font(name="Calibri", bold=True, size=11)
category_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
category_font = Font(name="Calibri", bold=True, size=11)
normal_font = Font(name="Calibri", size=10)
wrap_alignment = Alignment(wrap_text=True, vertical="top")
center_alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
openclaw_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border


def style_row(ws, row, cols, fill=None, font=None):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font or normal_font
        cell.alignment = wrap_alignment
        cell.border = thin_border
        if fill:
            cell.fill = fill


# ═══════════════════════════════════════════════════════════════
# SHEET 1 — Platform Overview & Architecture Comparison
# ═══════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Architecture Comparison"

headers = [
    "Dimension",
    "OpenClaw",
    "Devin (Cognition)",
    "Cursor",
    "GitHub Copilot",
    "Claude Code",
    "Windsurf (Codeium)",
    "Factory AI Droids",
    "Replit Agent",
    "Amazon Q Developer",
    "SWE-agent (OSS)",
]

rows = [
    headers,
    [
        "Category",
        "General-purpose AI assistant",
        "Autonomous SW engineer",
        "AI-native IDE",
        "IDE-integrated copilot",
        "Terminal-native agent",
        "Agentic IDE",
        "Autonomous coding agent",
        "Autonomous app builder",
        "AWS-native coding agent",
        "Open-source coding agent",
    ],
    [
        "Open Source",
        "Yes – MIT License\n322K+ GitHub stars",
        "No – Proprietary",
        "No – Proprietary\n(Composer model details shared)",
        "No – Proprietary",
        "No – Proprietary",
        "No – Proprietary",
        "No – Proprietary",
        "No – Proprietary",
        "No – Proprietary",
        "Yes – MIT License",
    ],
    [
        "Primary Language / Stack",
        "TypeScript (87.8%)\nNode.js runtime",
        "Compound AI system\nMultiple specialized models",
        "Electron (VS Code fork)\nCustom MoE model (Composer)",
        "Cloud-hosted\nMultiple LLM backends",
        "Cloud-hosted\nClaude model family",
        "Electron (VS Code fork)\nSWE-1.5 proprietary model",
        "Cloud-hosted\nMulti-model sampling",
        "Cloud-hosted\nClaude Sonnet 4 + GPT-4o",
        "Cloud-hosted\nClaude models",
        "Python\nYAML-config driven",
    ],
    [
        "Core Architecture",
        "Long-running Node.js daemon (Gateway).\nMessage routing, session persistence,\ncron jobs, tool execution.\nI/O bound for cloud LLMs;\nCPU-bound for tool/browser execution.",
        "Compound AI system:\n- Planner (high-reasoning model)\n- Critic (code review)\n- Coder (specialized code model)\n- Browser Agent (web scraping)\nSeparates planning from execution.",
        "ReAct loop with context injection.\nMoE model (Composer) trained via RL.\nSub-agent spawning for parallel tasks.\nSandboxed file/network access.",
        "Cloud-hosted inference.\nGitHub Issues → PR generation.\nPremium request allocation model.\nDeep GitHub ecosystem integration.",
        "Terminal-native autonomous agent.\nCloud inference only.\nPrompt caching (84% hit rate).\nAPI or subscription billing.",
        "Cascade agent with hybrid indexing\n(AST + semantic embeddings).\nPersistent Memories (vector store).\nFlow-awareness tracking.\n21 MCP connectors.",
        "HyperCode codebase representation.\nByteRank retrieval.\nCustom Droids (subagents).\nMulti-model task-dependent sampling.",
        "Full-stack autonomous builder.\nParallel task execution.\nDesign-to-code (Figma import).\nOne-click deployment.",
        "IDE + CLI + workspace agent.\nAWS API integration.\nJava upgrade automation.\nArchitecture diagram generation.",
        "YAML-config agent framework.\nSWE-ReX remote execution.\nDocker-based sandboxing.\n100+ parallel agents.",
    ],
    [
        "Deployment Model",
        "Self-hosted (VPS, bare metal,\nKubernetes, Docker).\nFull data sovereignty.",
        "SaaS (multi-tenant) or\nCustomer-dedicated SaaS\n(single-tenant VPC).\nAWS PrivateLink / IPSec.",
        "Desktop app (local IDE).\nCloud agents for background tasks.",
        "Cloud SaaS.\nGitHub-native integration.",
        "Cloud SaaS.\nNo self-hosting option.",
        "Desktop app (local IDE).\nCloud inference.",
        "Cloud SaaS.\nPlatform-managed or BYOM compute.",
        "Cloud SaaS.\nReplit-hosted infrastructure.",
        "Cloud SaaS.\nAWS-native deployment.",
        "Self-hosted.\nLocal Docker, AWS, Modal,\nFargate execution.",
    ],
    [
        "Model Support",
        "Any model: OpenAI, Anthropic,\nGoogle, local (Ollama, vLLM,\nTensorRT-LLM). Model-agnostic.",
        "Proprietary compound models.\nNo user model choice.",
        "Custom Composer MoE +\nClaude, GPT-4, Gemini access.",
        "Multiple LLM backends\n(model selection per request).",
        "Claude family only\n(Opus 4.6, Sonnet 4.6, Haiku 4.5).",
        "Proprietary SWE-1.5 +\nthird-party model access.",
        "Multi-model with task-dependent\nselection and multipliers.",
        "Claude Sonnet 4 + GPT-4o.\nNo user model choice.",
        "Claude models.\nNo user model choice.",
        "Any model: GPT-4o, Claude,\nlocal models. Fully configurable.",
    ],
    [
        "Agent Paradigm",
        "Persistent daemon.\nMulti-channel messaging.\nCron-scheduled tasks.\nBrowser automation.",
        "Fully autonomous loop.\nOwn environment (Devbox).\nPlanner-Critic-Coder pipeline.",
        "Interactive agent in IDE.\nSub-agent spawning.\nBackground cloud agents.",
        "Task-centric (Issue → PR).\nInline suggestions.\nChat-based assistance.",
        "Terminal-based autonomous agent.\nReAct loop.\nFile/shell access.",
        "Cascade planner.\nMulti-file reasoning.\nFlow-aware context tracking.",
        "Autonomous task execution.\nCustom subagent delegation.\nCI/CD integration.",
        "Autonomous build loop.\nIterative debugging.\nDeploy-in-loop.",
        "Agentic requests.\nCode transformation.\nAWS service integration.",
        "Autonomous fix loop.\nGitHub issue resolution.\nSecurity vulnerability detection.",
    ],
    [
        "Extensibility / Plugins",
        "3,200+ MCP skills on ClawHub.\nHot-reloading plugins.\nCryptographic MCP signing.",
        "Limited – Slack, IDE extensions.\nNo plugin marketplace.",
        "MCP support.\nCustom rules/commands.\nExtensions via VS Code ecosystem.",
        "GitHub Actions integration.\nLimited extension model.",
        "MCP support.\nBash tool access.",
        "21 first-party MCP connectors\n(Figma, Slack, Stripe, etc.).\n40+ IDE compatibility bridges.",
        "Custom Droids (subagents).\nLinter/analyzer integration.",
        "Figma import.\nDatabase integrations.\nDeployment pipelines.",
        "AWS service integrations.\nIDE plugins (VS Code, JetBrains,\nVisual Studio, Eclipse).",
        "YAML-based configuration.\nCustom tool definitions.\nSWE-ReX execution backends.",
    ],
    [
        "SWE-bench Verified Score",
        "N/A (not a coding agent)",
        "~55-60% (estimated)",
        "Not publicly disclosed",
        "Not publicly disclosed",
        "79.6-82%\n(highest as of early 2026)",
        "Not publicly disclosed",
        "Not publicly disclosed",
        "Not publicly disclosed",
        "Not publicly disclosed",
        "~65% (mini-swe-agent)",
    ],
]

for r_idx, row_data in enumerate(rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws1.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws1, r_idx, len(headers))
    else:
        fill = None
        font = normal_font
        if row_data[0] in ("Category", "Open Source", "Core Architecture", "SWE-bench Verified Score"):
            fill = subheader_fill
            font = subheader_font
        style_row(ws1, r_idx, len(headers), fill=fill, font=font)
        # Highlight OpenClaw column
        ws1.cell(row=r_idx, column=2).fill = openclaw_fill

ws1.column_dimensions["A"].width = 26
for c in range(2, len(headers) + 1):
    ws1.column_dimensions[get_column_letter(c)].width = 32

# ═══════════════════════════════════════════════════════════════
# SHEET 2 — Compute Requirements
# ═══════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Compute Requirements")

compute_headers = [
    "Specification",
    "OpenClaw\n(Self-Hosted)",
    "Devin\n(Enterprise VPC)",
    "Cursor\n(Cloud Agent Infra)",
    "Factory AI\n(Platform-Managed)",
    "SWE-agent\n(Self-Hosted)",
]

compute_rows = [
    compute_headers,
    # Category: Minimum Hardware
    ["— MINIMUM HARDWARE —", "", "", "", "", ""],
    [
        "CPU (Minimum)",
        "1-2 vCPU",
        "70+ vCPUs (AWS VPC)\ni3.metal instances",
        "N/A (cloud-hosted;\nuser runs desktop IDE)",
        "4 CPU\n(platform-managed)",
        "1-2 vCPU\n(Docker container)",
    ],
    [
        "RAM (Minimum)",
        "2 GB\n(hard floor; crashes below)",
        "Not publicly specified\n(metal instances imply 32+ GB)",
        "N/A (cloud-hosted)",
        "8 GB + 6 GB swap\n(platform-managed)",
        "2-4 GB\n(depends on tooling)",
    ],
    [
        "Storage (Minimum)",
        "20 GB",
        "Persistent Docker volumes\n(S3 backing)",
        "N/A (cloud-hosted)",
        "Not specified",
        "10-20 GB",
    ],
    # Category: Recommended / Production
    ["— RECOMMENDED / PRODUCTION —", "", "", "", "", ""],
    [
        "CPU (Recommended)",
        "2-4+ vCPU",
        "Auto-scaling EC2\n(enterprise-grade)",
        "Thousands of GPUs\n(Composer training)\nLarge Linux VMs (agent execution)",
        "BYOM: any machine\nwith network access",
        "2-4 vCPU per agent\n(100+ parallel agents supported)",
    ],
    [
        "RAM (Recommended)",
        "4-8 GB (standard)\n8-16 GB (browser automation)\n16+ GB (multi-agent / local LLM)",
        "Not specified\n(sandboxed Devbox per session)",
        "Not specified publicly",
        "BYOM: user-defined",
        "4-8 GB per agent",
    ],
    [
        "Storage (Recommended)",
        "40-80+ GB",
        "Persistent storage option\n(Large Performant tier)",
        "Not specified publicly",
        "BYOM: user-defined",
        "20-50 GB",
    ],
    # Category: GPU
    ["— GPU REQUIREMENTS —", "", "", "", "", ""],
    [
        "GPU Required?",
        "NO for cloud LLM usage.\nYES only for local LLM inference.",
        "Not disclosed.\nLikely GPU-accelerated\non Cognition's infrastructure.",
        "YES – thousands of NVIDIA GPUs\nfor Composer model training.\nNot needed for end-user IDE.",
        "Not required for users.\nCloud-hosted inference.",
        "NO for cloud LLM.\nOptional for local model.",
    ],
    [
        "GPU Specs (if local LLM)",
        "RTX 3060 (12GB) → 8B models\nRTX 3090 (24GB) → 34B models\nRTX 4090 (24GB) → 34-70B\nA100 (40-80GB) → 70B+\nH100 (80GB) → 70-405B\nRule: params × 1GB VRAM (Q4)",
        "N/A (proprietary infra)",
        "MXFP8 MoE kernels\nwith expert parallelism.\nPyTorch + Ray training infra.\nThousands of NVIDIA GPUs.",
        "N/A (cloud inference)",
        "Optional: depends on\nchosen local model.",
    ],
    # Category: Resource consumption
    ["— RUNTIME RESOURCE CONSUMPTION —", "", "", "", "", ""],
    [
        "Idle CPU Load",
        "1-3% (background router)",
        "Minimal (session-based)",
        "Low (desktop IDE idle)",
        "Auto-pauses when idle",
        "Minimal (Docker idle)",
    ],
    [
        "Active CPU Load",
        "Single agent + cloud LLM: minimal\nBrowser automation (3 nodes): 75-90%\nLocal 8B model: 15-25%\nLocal 14B model: 45-60%",
        "Not disclosed",
        "Not disclosed\n(250 tokens/sec inference)",
        "Not disclosed",
        "Not disclosed",
    ],
    [
        "Memory Footprint",
        "Node.js Gateway: 300-500 MB idle.\nBrowser automation: 8 GB hard req.\nLocal 8B model: 2.8 GB peak.\nLocal 14B model: 5.4 GB peak.",
        "Not disclosed",
        "Not disclosed",
        "Not disclosed",
        "Not disclosed",
    ],
    [
        "Scaling Approach",
        "Single-instance by default.\nKubernetes (ClawPod) for 1-100+ agents\nwith per-agent isolation, resource limits,\nnetwork policies, persistent volumes.",
        "Auto-scaling EC2 instances.\nSandboxed Devbox per session.\nEnterprise SaaS or dedicated VPC.",
        "Single large Linux VM\nwith shared state files.\nThousands of agents per VM.",
        "Platform-managed compute.\nAuto-pause on idle.\nBYOM for custom infra.",
        "SWE-ReX framework.\nDocker containers per agent.\n100+ simultaneous agents\nvia AWS/Modal/Fargate.",
    ],
]

for r_idx, row_data in enumerate(compute_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws2.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws2, r_idx, len(compute_headers))
    elif row_data[0].startswith("—"):
        style_row(ws2, r_idx, len(compute_headers), fill=category_fill, font=category_font)
    else:
        style_row(ws2, r_idx, len(compute_headers))
        ws2.cell(row=r_idx, column=2).fill = openclaw_fill

ws2.column_dimensions["A"].width = 28
for c in range(2, len(compute_headers) + 1):
    ws2.column_dimensions[get_column_letter(c)].width = 36

# ═══════════════════════════════════════════════════════════════
# SHEET 3 — Pricing Comparison
# ═══════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Pricing Comparison")

pricing_headers = [
    "Pricing Dimension",
    "OpenClaw",
    "Devin (Cognition)",
    "Cursor",
    "GitHub Copilot",
    "Claude Code",
    "Windsurf",
    "Factory AI",
    "Replit Agent",
    "Amazon Q Developer",
    "SWE-agent",
]

pricing_rows = [
    pricing_headers,
    ["— SUBSCRIPTION / LICENSE —", "", "", "", "", "", "", "", "", "", ""],
    [
        "Software License Cost",
        "FREE (MIT open source)",
        "$500/month (unlimited seats)\nEnterprise: custom",
        "Free / $20 / $60 / $200 per month\nTeams: $40/user/mo\nEnterprise: custom",
        "Pro: $10/mo\nPro+: $39/mo\nBusiness: $19/user/mo\nEnterprise: $39/user/mo",
        "Pro: $20/mo\nMax 5x: $100/mo\nMax 20x: $200/mo\nTeams: $150/user/mo",
        "Free / $15/mo\nTeams: $30/user/mo\nEnterprise: custom",
        "Pro: $20/mo\nMax: $200/mo\nEnterprise: custom",
        "Effort-based pricing.\nCredits per plan tier.\nTurbo mode: 6× Power cost.",
        "Free: $0\nPro: $19/user/mo",
        "FREE (MIT open source)",
    ],
    ["— INFRASTRUCTURE / COMPUTE COSTS —", "", "", "", "", "", "", "", "", "", ""],
    [
        "Hosting / Compute Cost",
        "Self-hosted VPS:\nBudget: $4-6/mo\nStandard: $5-12/mo\nPerformance: $16-24/mo\nFree: Oracle Cloud Always Free",
        "Included in subscription.\nEnterprise VPC: customer\npays AWS infrastructure.",
        "Included in subscription.\nCloud agents included.",
        "Included in subscription.",
        "Included in subscription\n(or API pay-as-you-go).",
        "Included in subscription.",
        "Included (platform-managed)\nor BYOM (user-paid infra).",
        "Included in subscription.",
        "Included in subscription.",
        "Self-hosted: $4-24/mo\n(VPS) or free (local Docker).",
    ],
    [
        "LLM / API Costs\n(variable)",
        "$5-200+/month\ndepends on model & usage.\nClaude Opus heavy: ~$420/mo.\nDeepSeek/Gemini Flash: $5-10/mo.",
        "Included in subscription.",
        "Included within credit pool.\nOverages billed per usage.",
        "Included.\nPremium requests: $0.04 each\n(overage).",
        "API: $3-25/M tokens\n(varies by model).\nSubscription: included.",
        "Included within credits.\nAdd-on: $10/250 credits.",
        "Included within tokens.\nOverage: $2.70/M tokens.",
        "Included within credits.\nEffort-based pricing.",
        "Included.\nCode transformation:\n$0.003/line (overage).",
        "User pays own API costs.\nModel-agnostic.",
    ],
    ["— TOTAL COST OF OWNERSHIP (TCO) —", "", "", "", "", "", "", "", "", "", ""],
    [
        "Light Usage\n(individual, occasional)",
        "$10-22/month\n(VPS + cheap model APIs)",
        "$500/month",
        "$0-20/month",
        "$10-19/month",
        "$20/month (Pro)",
        "$0-15/month",
        "$20/month",
        "~$5-25/month\n(effort-based)",
        "$0 (Free tier:\n50 agentic requests)",
        "$0-10/month\n(local + API costs)",
    ],
    [
        "Moderate Usage\n(daily professional use)",
        "$27-72/month\n(VPS + moderate API)",
        "$500/month",
        "$60/month (Pro+)",
        "$39/month (Pro+)",
        "$100-200/month (Max)",
        "$15-30/month",
        "$200/month (Max)",
        "~$50-150/month",
        "$19/month (Pro)",
        "$10-50/month",
    ],
    [
        "Heavy / Enterprise\n(multi-agent, team)",
        "$62-224+/month per instance.\nPer-person: $300-500/mo/agent.\n100 people: $30-50K/mo.",
        "$500/mo + enterprise pricing.\nEnterprise: Goldman Sachs,\nSantander, Nubank.",
        "$200/mo (Ultra) or\n$40/user/mo (Teams).\nEnterprise: custom.",
        "$39/user/mo (Enterprise).\nVolume discounts available.",
        "$150/user/mo (Teams).\nEnterprise: custom.",
        "$30/user/mo (Teams).\nEnterprise: custom.",
        "Enterprise: custom pricing.\nUnlimited features.",
        "Pro plan + high usage.\nEnterprise: custom.",
        "$19/user/mo.\nVolume discounts.",
        "$20-100+/month\n(infra + API costs).",
    ],
    ["— HIDDEN / ADDITIONAL COSTS —", "", "", "", "", "", "", "", "", "", ""],
    [
        "Hidden Costs",
        "System prompt token burn: $5-30/mo.\nHeartbeat checks: $0-90/mo.\nMaintenance time: 2-5 hrs/mo.\nModel switching overhead.",
        "AWS VPC infra costs (enterprise).\nData transfer charges.\nOnboarding/integration time.",
        "Overage charges beyond\ncredit pool.\nBugbot add-on: $40/user/mo.",
        "Premium request overages\n($0.04/request).\nModel upgrade costs.",
        "Token-based billing can\nspike unexpectedly.\nNo caching = 4x cost.",
        "Credit overage ($10/250).\nModel-specific multipliers.",
        "Token overage: $2.70/M.\nCached vs. uncached pricing.",
        "Turbo mode: up to 6×\nPower mode cost.\nComplex task surcharges.",
        "Code transformation\noverage: $0.003/line.",
        "API costs are unbounded.\nInfra maintenance time.",
    ],
]

for r_idx, row_data in enumerate(pricing_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws3.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws3, r_idx, len(pricing_headers))
    elif row_data[0].startswith("—"):
        style_row(ws3, r_idx, len(pricing_headers), fill=category_fill, font=category_font)
    else:
        style_row(ws3, r_idx, len(pricing_headers))
        ws3.cell(row=r_idx, column=2).fill = openclaw_fill

ws3.column_dimensions["A"].width = 28
for c in range(2, len(pricing_headers) + 1):
    ws3.column_dimensions[get_column_letter(c)].width = 30

# ═══════════════════════════════════════════════════════════════
# SHEET 4 — Key Differentiators & Strategic Notes
# ═══════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Strategic Notes")

notes_headers = [
    "Platform",
    "Key Differentiator",
    "Target User",
    "Competitive Moat",
    "Risk / Weakness",
    "Data Sovereignty",
    "Growth / Traction",
]

notes_rows = [
    notes_headers,
    [
        "OpenClaw",
        "Open-source, self-hosted, model-agnostic\npersonal AI assistant.\n3,200+ MCP plugins.\nFull data sovereignty.\nLowest TCO at scale.",
        "Power users, privacy-conscious,\nSMBs wanting automation\nwithout vendor lock-in.",
        "322K+ GitHub stars.\nFastest-growing GH project ever.\nMassive plugin ecosystem.\nModel-agnostic flexibility.",
        "NOT a coding agent (general assistant).\nSelf-hosting maintenance burden.\nNo enterprise support/SLAs.\nAPI costs can spike with expensive models.",
        "FULL – data never leaves\nuser's infrastructure.",
        "322K stars, 62K forks,\n360 contributors.\n300-400K users.\nSurpassed React stars in 60 days.",
    ],
    [
        "Devin (Cognition)",
        "Fully autonomous SW engineer.\nCompound AI architecture\n(Planner/Critic/Coder/Browser).\n10M+ token context (Enterprise).",
        "Engineering teams wanting\nasynchronous task execution.\nEnterprise customers\n(Goldman Sachs, Santander).",
        "Proprietary compound AI system.\nEnterprise relationships.\nDeep autonomy capabilities.\nAcquired Windsurf/Codeium.",
        "High cost ($500/mo baseline).\nNo open-source alternative.\nBlack-box architecture.\nVendor lock-in.",
        "SaaS (multi-tenant) or\ncustomer-dedicated VPC.\nAWS PrivateLink option.",
        "Major enterprise customers.\nAcquired Windsurf.\nLed by ex-IOI champions.",
    ],
    [
        "Cursor",
        "AI-native IDE with custom\nComposer MoE model.\n4× faster inference.\nSub-agent spawning.",
        "Senior developers wanting\nflow-state amplification.\nTeams needing shared\nAI-powered workflows.",
        "Custom-trained Composer model.\n250 tokens/sec inference.\nRL-trained on real dev environments.\nLargest AI IDE user base.",
        "Not fully autonomous\n(IDE-dependent).\nCredit-based pricing can\nbe unpredictable.",
        "Local IDE + cloud inference.\nCode stays local;\nprompts sent to cloud.",
        "Dominant AI IDE.\nRapidly growing Teams/Enterprise.",
    ],
    [
        "GitHub Copilot",
        "Deepest GitHub integration.\nIssue → PR workflow.\nLargest enterprise installed base.",
        "Enterprise teams already\non GitHub ecosystem.\nCompliance-focused orgs.",
        "GitHub ecosystem lock-in.\nMicrosoft/OpenAI backing.\nLargest installed base.\nIP indemnity.",
        "Lower autonomy than Devin.\nPremium request limits.\nLess flexible model choice.",
        "Cloud SaaS.\nMicrosoft enterprise\ncompliance certifications.",
        "Largest installed base among\nAI coding tools.\nEnterprise standard.",
    ],
    [
        "Claude Code",
        "Highest SWE-bench score\n(79.6-82%).\nTerminal-native.\n84% prompt cache hit rate.",
        "Expert developers wanting\ntop code quality.\nTerminal-first workflows.",
        "Highest benchmark scores.\n46% developer satisfaction\n(vs. 9% Copilot, 19% Cursor).\nAnthropic model quality.",
        "No self-hosting.\nClaude-only (no model choice).\nToken costs can spike.\nNo IDE integration.",
        "Cloud only.\nAnthropic handles all data.",
        "Most-loved AI coding tool\nin 2026 (46% satisfaction).",
    ],
    [
        "Windsurf",
        "SWE-1.5 at 950 tokens/sec\n(13× Sonnet, 6× Haiku).\n21 MCP connectors.\nPersistent Memories.",
        "Developers wanting fast,\ncontext-aware IDE assistance.\n40+ IDE support.",
        "Fastest inference speed.\nCognition (Devin) backing.\nBroad IDE compatibility.",
        "Acquired by Cognition –\nfuture direction uncertain.\nProprietary model dependency.",
        "Local IDE + cloud inference.\nPersistent memory vectors\nstored locally.",
        "800K+ active users.\n$82M ARR by mid-2025.",
    ],
    [
        "Factory AI",
        "HyperCode codebase representation.\nByteRank retrieval.\nCustom Droids (subagents).",
        "Teams wanting specialized\nautonomous agents per task.\nCI/CD integration focus.",
        "Multi-model task routing.\nCustom subagent architecture.\nToken-based pricing model.",
        "Smaller ecosystem.\nToken overage can be costly.\nLess brand recognition.",
        "Cloud SaaS or BYOM.\nBYOM keeps data local.",
        "Growing but smaller\nthan major competitors.",
    ],
    [
        "Replit Agent",
        "Full-stack app builder\nfrom natural language.\nOne-click deployment.\nDesign-to-code.",
        "Non-technical users,\nprototypers, indie hackers.\nRapid MVP builders.",
        "End-to-end build + deploy.\nLowest barrier to entry.\nNo local setup required.",
        "Limited for complex codebases.\nEffort-based pricing unpredictable.\nNot for enterprise SW engineering.",
        "Cloud SaaS.\nAll code on Replit servers.",
        "Popular with beginners\nand prototypers.",
    ],
    [
        "Amazon Q Developer",
        "Deep AWS service integration.\nJava upgrade automation.\nIP indemnity.",
        "AWS-native teams.\nJava enterprise shops.\nCompliance-focused orgs.",
        "AWS ecosystem integration.\nAmazon backing.\nFree tier available.\nIP indemnity.",
        "AWS-centric (limited outside AWS).\nSmaller agentic capability\nvs. Devin/Claude Code.",
        "Cloud SaaS.\nAWS enterprise compliance.",
        "Growing within AWS\ncustomer base.",
    ],
    [
        "SWE-agent (OSS)",
        "Research-grade open-source.\n100-line mini-swe-agent.\n65%+ SWE-bench.\nMassively parallel execution.",
        "Researchers, OSS contributors.\nTeams wanting full control\nover agent behavior.",
        "Fully open-source.\nResearch pedigree (Princeton/Stanford).\nSimplest architecture.",
        "Maintenance-only mode.\nNo commercial support.\nRequires technical expertise.",
        "FULL – self-hosted.\nAll data stays local.",
        "Influential research project.\nShifting to mini-swe-agent.",
    ],
]

for r_idx, row_data in enumerate(notes_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws4.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws4, r_idx, len(notes_headers))
    else:
        fill = openclaw_fill if row_data[0] == "OpenClaw" else None
        style_row(ws4, r_idx, len(notes_headers), fill=fill)

ws4.column_dimensions["A"].width = 22
for c in range(2, len(notes_headers) + 1):
    ws4.column_dimensions[get_column_letter(c)].width = 36

# ═══════════════════════════════════════════════════════════════
# SHEET 5 — OpenClaw Deep Dive
# ═══════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("OpenClaw Deep Dive")

deep_headers = ["Aspect", "Details"]

deep_rows = [
    deep_headers,
    ["— IDENTITY —", ""],
    ["What It Is", "Open-source, self-hosted personal AI assistant designed to take actions (not just answer questions). Connects to messaging apps (WhatsApp, Telegram, Slack, Discord, Signal, iMessage) and executes real tasks."],
    ["GitHub Stats", "322,031 stars | 61,928 forks | 360 contributors | TypeScript 87.8% | MIT License | 68 releases (latest v2026.3.13-1)"],
    ["Growth Milestone", "Fastest-growing project in GitHub history. Surpassed React's star count in ~60 days (React took 13 years). 300,000-400,000 users worldwide by March 2026."],
    ["— ARCHITECTURE —", ""],
    ["Runtime", "Node.js long-running daemon process (Gateway). Handles message routing, session persistence, cron jobs, and tool execution."],
    ["I/O Profile", "I/O bound when using cloud LLM APIs. Becomes CPU-intensive during tool execution, browser automation, and local LLM inference."],
    ["MCP Integration", "3,200+ MCP skills on ClawHub marketplace. Hot-reloading plugins. Cryptographic message signing (ECDSA P-256) for MCP transport security."],
    ["Multi-Agent Patterns", "Per-Function (most cost-effective), Agent Teams (coordinated specialists), Hierarchical (chief-of-staff orchestrator). Single-instance by default; Kubernetes (ClawPod) for 1-100+ agents."],
    ["— MINIMUM COMPUTE —", ""],
    ["Testing / Light", "1-2 vCPU | 2 GB RAM | 20 GB storage"],
    ["Daily Use (Recommended)", "2 vCPU | 4 GB RAM | 40 GB storage"],
    ["Production / Browser Automation", "4+ vCPU | 8-16 GB RAM | 80+ GB storage"],
    ["— RESOURCE CONSUMPTION —", ""],
    ["Gateway Idle", "Node.js Gateway idles at 300-500 MB RAM. CPU: 1-3%."],
    ["Browser Automation", "8 GB RAM hard requirement. 3 browser nodes: 75-90% CPU."],
    ["Local LLM (8B model)", "15-25% CPU | 2.8 GB peak RAM"],
    ["Local LLM (14B model)", "45-60% CPU | 5.4 GB peak RAM"],
    ["— GPU REQUIREMENTS —", ""],
    ["Cloud LLM Mode", "NO GPU required. OpenClaw acts as HTTP client to cloud APIs."],
    ["Local LLM Mode", "GPU required. Rule of thumb: model params × 1GB VRAM (Q4 quantization)."],
    ["GPU Tiers", "RTX 3060 (12GB) → 8B models | RTX 3090 (24GB) → 34B models | RTX 4090 (24GB) → 34-70B | A100 (40-80GB) → 70B+ | H100 (80GB) → 70-405B"],
    ["Local LLM Stacks", "Ollama, vLLM, TensorRT-LLM for GPU acceleration. LM Studio for CPU-only (limited)."],
    ["— COST BREAKDOWN —", ""],
    ["Budget Setup", "$5-10/month total. Cheaper models (DeepSeek V3, Gemini Flash). Hetzner/Contabo VPS."],
    ["Standard Setup", "$15-30/month. Claude Sonnet 4 or GPT-4o. 4 GB RAM VPS."],
    ["Heavy Use", "$62-224+/month. Browser automation + expensive models."],
    ["Enterprise (100 people)", "$30,000-50,000/month (per-person model with Claude Opus)."],
    ["Hidden Costs", "System prompt token burn: $5-30/mo | Heartbeat checks: $0-90/mo | Maintenance: 2-5 hrs/mo"],
    ["Free Option", "Oracle Cloud Always Free tier (24 GB RAM)."],
    ["— KEY DISTINCTION —", ""],
    ["Not a Coding Agent", "OpenClaw is a general-purpose personal assistant, NOT a coding agent like Devin or Claude Code. It automates tasks like email, calendar, smart home, and workflows across messaging platforms."],
]

for r_idx, row_data in enumerate(deep_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws5.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws5, r_idx, len(deep_headers))
    elif row_data[0].startswith("—"):
        style_row(ws5, r_idx, len(deep_headers), fill=category_fill, font=category_font)
    else:
        style_row(ws5, r_idx, len(deep_headers))

ws5.column_dimensions["A"].width = 30
ws5.column_dimensions["B"].width = 100

# ── Save ──
output_path = "/workspace/OpenClaw_Architecture_Comparison.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
