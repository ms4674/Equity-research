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
    ["— NVIDIA GTC 2026 VALIDATION —", ""],
    ["NVIDIA Partnership", "Jensen Huang (GTC 2026): \"Every single company in the world today has to have an OpenClaw strategy.\" Positioned as \"the operating system for personal AI.\""],
    ["NemoClaw Stack", "NVIDIA's enterprise wrapper: OpenClaw + Nemotron models + OpenShell sandboxed runtime. Apache 2.0 license. Single-command install. Production-ready in <1 hour."],
    ["Nemotron 3 Super 120B", "Top open model on PinchBench (85.6%), only 1.3% behind Claude Sonnet 4.6 (86.9%). 12B active params (MoE). 1M token context. 42,855 tok/s on DGX Spark."],
    ["Hardware Ecosystem", "DGX Spark (128 GB, 1 PFLOP, desktop form) → DGX Station (748 GB, enterprise) → RTX PCs (Nemotron Nano 4B). Full upgrade path for local OpenClaw inference."],
    ["Enterprise Partners", "Adobe, Salesforce, SAP, ServiceNow, Siemens, CrowdStrike, Atlassian, Palantir — all building with NVIDIA Agent Toolkit / NemoClaw."],
    ["Adoption (GTC Data)", "250K+ GitHub stars in 60 days. 2.2M weekly npm downloads. 65% of users in enterprise sectors. Surpassed Linux adoption speed (3 weeks vs. 30 years)."],
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

# ═══════════════════════════════════════════════════════════════
# SHEET 6 — OpenClaw vs Claude Cowork Head-to-Head
# ═══════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("OpenClaw vs Claude Cowork")

cowork_fill = PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid")
verdict_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
verdict_font = Font(name="Calibri", bold=True, size=10)

cw_headers = ["Attribute", "OpenClaw", "Claude Cowork", "Verdict / Edge"]

cw_rows = [
    cw_headers,
    # ── Identity & Positioning ──
    ["— IDENTITY & POSITIONING —", "", "", ""],
    [
        "Category",
        "Open-source, self-hosted personal AI assistant.\nGeneral-purpose task automation across\nmessaging channels.",
        "Anthropic's desktop AI knowledge-work agent.\nBuilt into the Claude app for\nmacOS (Windows planned).",
        "Both are general-purpose AI assistants (NOT coding agents).\nOpenClaw: messaging-first. Cowork: desktop-first.",
    ],
    [
        "Launch / Maturity",
        "Open-source since 2025.\n68 releases as of March 2026.\nMature community-driven project.",
        "Launched Jan 12, 2026 as research preview.\nBuilt in ~1.5 weeks (by Claude Code itself).\nStill early-stage.",
        "OpenClaw: More mature with larger ecosystem.\nCowork: Newer but backed by Anthropic.",
    ],
    [
        "Open Source",
        "YES – MIT License.\n322K+ GitHub stars, 62K forks, 360 contributors.",
        "NO – Proprietary.\nIntegrated into Claude app.\nAgent Skills spec is open.",
        "OpenClaw wins for transparency & community.\nCowork wins for polish & managed experience.",
    ],
    [
        "Primary Use Cases",
        "Multi-channel messaging automation.\nEmail/calendar management.\nSmart home control.\nWorkflow automation via cron.\nBrowser automation.\n24/7 persistent daemon.",
        "File management & organization.\nDocument processing (XLSX, PPTX, DOCX, PDF).\nResearch synthesis & report generation.\nData extraction (images → spreadsheets).\nBrowser automation via Chrome extension.",
        "OpenClaw: Best for persistent, always-on automation across channels.\nCowork: Best for desktop knowledge work & document tasks.",
    ],
    # ── Architecture ──
    ["— ARCHITECTURE —", "", "", ""],
    [
        "Core Runtime",
        "Node.js long-running daemon (Gateway).\nMessage routing, session persistence,\ncron jobs, tool execution.\nRuns 24/7 on user's infrastructure.",
        "Full Ubuntu 22.04 Linux VM\nrunning locally on macOS via\nApple Virtualization.framework.\nClaude Code CLI inside bubblewrap sandbox.",
        "OpenClaw: Lightweight Node.js process.\nCowork: Heavier (full Linux VM on macOS).\nOpenClaw is more resource-efficient at idle.",
    ],
    [
        "Execution Model",
        "Gateway-first, channel-native architecture.\nPersistent daemon running 24/7.\nCron-scheduled autonomous tasks.\nHeartbeat monitoring (15-60 min intervals).",
        "Session-based desktop agent.\nRuns within Claude app.\nTerminates if computer sleeps.\nNo persistent background execution.",
        "OpenClaw: Always-on, autonomous.\nCowork: Session-dependent, interactive.\nOpenClaw wins for unattended automation.",
    ],
    [
        "Sandboxing / Isolation",
        "Process-level isolation.\nDocker containers for multi-agent (ClawPod).\nKubernetes network policies for production.\nFull system access by default.",
        "VM-level isolation (Apple Virtualization.framework).\nBubblewrap sandbox + seccomp filtering.\nFolder-specific read/write/create permissions.\nStronger default isolation.",
        "Cowork: Stronger sandbox by default (VM + seccomp).\nOpenClaw: Full system access (power but risk).\nCowork wins on security posture.",
    ],
    [
        "Parallel Execution",
        "Multi-agent via Kubernetes (ClawPod).\n1-100+ agents with per-agent isolation.\nRequires Kubernetes infrastructure.",
        "Sub-agent coordination.\nMultiple concurrent Claude instances.\nParallel independent subtasks.\nNative to the app, no infra needed.",
        "Cowork: Easier parallel execution (built-in).\nOpenClaw: More scalable (Kubernetes-based).\nTrade-off: simplicity vs. scale.",
    ],
    # ── Models & Intelligence ──
    ["— MODELS & INTELLIGENCE —", "", "", ""],
    [
        "Model Support",
        "Model-agnostic: OpenAI, Anthropic, Google,\nDeepSeek, local models (Ollama, vLLM,\nTensorRT-LLM). Any model via API.",
        "Claude family ONLY:\nClaude Opus 4.5/4.6, Sonnet 4.6, Haiku 4.5.\n1M token extended context (Max plans).\nNo third-party model support.",
        "OpenClaw: Maximum flexibility (any model).\nCowork: Locked to Claude (but highest quality).\nOpenClaw wins on choice; Cowork wins on depth.",
    ],
    [
        "Context Window",
        "Depends on chosen model.\nTypically 128K-200K tokens.\nNo built-in extended context.",
        "Up to 1M tokens (Max plans).\nExtended thinking mode (Max only).\nProgressive disclosure for large files.",
        "Cowork wins: 1M token context + extended thinking\nis significantly larger than typical model defaults.",
    ],
    [
        "Intelligence Quality",
        "Varies by model selected.\nCan use best-in-class models (Claude Opus)\nor budget models (DeepSeek, Gemini Flash).\nUser controls quality-cost trade-off.",
        "Always Claude (Anthropic's top models).\nConsistently high quality.\nNo ability to trade down for cost savings.",
        "Cowork: Consistently top-tier quality.\nOpenClaw: Flexible quality (user's choice).\nDifferent philosophies, not directly comparable.",
    ],
    # ── Compute & Hardware ──
    ["— COMPUTE & HARDWARE —", "", "", ""],
    [
        "Minimum CPU",
        "1-2 vCPU",
        "Any modern processor (1-2 cores).\nRecommended: 2+ cores.",
        "Comparable minimum requirements.\nBoth lightweight for cloud LLM usage.",
    ],
    [
        "Minimum RAM",
        "2 GB (hard floor; crashes below).",
        "2 GB minimum; 4-8 GB recommended.\nVM bundle alone consumes significant memory.",
        "Similar floor, but Cowork's VM overhead\nmakes 4-8 GB more realistic in practice.",
    ],
    [
        "Storage",
        "20 GB minimum.\n40-80+ GB recommended.",
        "1 GB installation + 10 GB VM bundle.\n~11 GB minimum; 20+ GB recommended.\nVM image regenerates after sessions.",
        "OpenClaw needs more disk for persistent data.\nCowork's VM bundle is temporary but large (10 GB).",
    ],
    [
        "GPU Required?",
        "NO for cloud LLM.\nYES for local LLM inference.\n(RTX 3060 → A100/H100 depending on model)",
        "NO – cloud inference only.\nNo local model support.\nGPU never needed.",
        "OpenClaw: Optional GPU for local models.\nCowork: No GPU ever needed.\nCowork simpler; OpenClaw more capable with GPU.",
    ],
    [
        "Platform Support",
        "Linux, macOS, Windows.\nDocker, Kubernetes, VPS, bare metal.\nAny platform with Node.js.",
        "macOS ONLY (currently).\nWindows support planned.\nRequires Apple Virtualization.framework.",
        "OpenClaw: Runs anywhere.\nCowork: macOS only (major limitation).\nOpenClaw wins decisively on platform support.",
    ],
    [
        "Idle Resource Consumption",
        "Node.js Gateway: 300-500 MB RAM, 1-3% CPU.\nLightweight persistent process.",
        "VM process: reported 24-55% CPU idle on some systems.\nMemory leaks reported over time.\n1.9 GB RAM during VM startup.",
        "OpenClaw: Much lighter at idle (300-500 MB).\nCowork: Heavier, with reported idle CPU issues.\nOpenClaw wins on resource efficiency.",
    ],
    [
        "Browser Automation",
        "Playwright-based.\n8 GB RAM hard requirement.\n3 nodes: 75-90% CPU.\nFull browser control.",
        "Chrome extension-based.\nLighter weight than Playwright.\nRequires Claude for Chrome extension.\nLess programmatic control.",
        "OpenClaw: More powerful (Playwright, full control).\nCowork: Lighter weight (Chrome extension).\nTrade-off: power vs. simplicity.",
    ],
    # ── Pricing ──
    ["— PRICING & COST —", "", "", ""],
    [
        "Software License",
        "FREE (MIT open source).",
        "Included in Claude subscription:\nPro: $20/mo | Max 5x: $100/mo |\nMax 20x: $200/mo | Team: $30/user/mo.",
        "OpenClaw: Free software.\nCowork: $20-200/mo subscription required.",
    ],
    [
        "Infrastructure Cost",
        "Self-hosted VPS: $4-24/mo.\nFree: Oracle Cloud Always Free.\nUser pays hosting.",
        "None – runs on user's Mac.\nNo separate infrastructure needed.",
        "OpenClaw: Requires separate hosting ($4-24/mo).\nCowork: Runs on existing Mac (no extra cost).",
    ],
    [
        "LLM / API Cost",
        "Variable: $5-200+/mo.\nDepends on model and usage.\nCheap models: $5-10/mo.\nClaude Opus heavy: ~$420/mo.",
        "Included in subscription.\nNo separate API charges.\nUsage limited by plan tier\n(resets every 5 hours).",
        "OpenClaw: Unbounded variable costs.\nCowork: Predictable fixed pricing.\nCowork wins on cost predictability.",
    ],
    [
        "TCO: Light Use",
        "$10-22/month\n(VPS + budget model APIs).",
        "$20/month (Pro plan).",
        "Comparable at light usage.\nOpenClaw slightly cheaper with budget models.",
    ],
    [
        "TCO: Moderate Use",
        "$27-72/month\n(VPS + moderate API usage).",
        "$100/month (Max 5x).",
        "OpenClaw cheaper at moderate use.\nCowork more predictable.",
    ],
    [
        "TCO: Heavy Use",
        "$62-224+/month per instance.\nEnterprise (100 people): $30-50K/mo.",
        "$200/month (Max 20x).\nTeam: $30/user/mo.\nEnterprise: custom.",
        "Cowork cheaper for individual heavy use ($200 flat).\nOpenClaw cheaper at team scale with budget models.",
    ],
    [
        "Hidden Costs",
        "System prompt token burn: $5-30/mo.\nHeartbeat checks: $0-90/mo.\nMaintenance: 2-5 hrs/mo.\nModel switching overhead.",
        "Cowork sessions burn quota faster\nthan regular chat (multi-step reasoning).\nRate limits reset every 5 hours.\nNo overage purchase option.",
        "OpenClaw: Unpredictable API spikes.\nCowork: Hard rate limits (can't buy more).\nDifferent risk profiles.",
    ],
    # ── Extensibility ──
    ["— EXTENSIBILITY & INTEGRATIONS —", "", "", ""],
    [
        "Plugin / Skill Ecosystem",
        "3,200+ MCP skills on ClawHub.\nCommunity-built, hot-reloading.\nCryptographic MCP signing (ECDSA P-256).",
        "11 open-source starter plugins.\nAgent Skills (markdown-based workflow templates).\nCommands (slash-shortcuts).\nPlugins stored locally (org distribution planned).",
        "OpenClaw: Massive existing ecosystem (3,200+).\nCowork: Early-stage but well-designed.\nOpenClaw wins on ecosystem size today.",
    ],
    [
        "Connectors / Integrations",
        "WhatsApp, Telegram, Slack, Discord,\nSignal, iMessage, email, calendar.\nAny MCP server.",
        "37+ native app connectors + Zapier.\nGitHub, Slack, Notion, Google Drive,\nJira, Salesforce pre-built.\nMost first-party connectors are read-only.",
        "OpenClaw: Messaging-channel focused.\nCowork: Enterprise SaaS app focused.\nDifferent integration philosophies.",
    ],
    [
        "MCP Support",
        "Full MCP support.\nEvery ClawHub skill is an MCP server.\nmcporter CLI for management.\nHot-reloading, no restart.",
        "Full MCP support.\nPre-built connectors are MCP-based.\nOAuth authentication flow.\nConnectors maintained by Anthropic + partners.",
        "Both fully MCP-compatible.\nOpenClaw: Community-driven MCP.\nCowork: Vendor-curated MCP.",
    ],
    [
        "Office File Handling",
        "Via community MCP skills.\nLess native office integration.\nPrimarily focused on messaging/automation.",
        "Native XLSX, PPTX, DOCX, PDF handling.\nDirect Excel and PowerPoint editing.\nProgressive disclosure for context management.\nAgent Skills spec for office formats.",
        "Cowork wins decisively for document/office work.\nOpenClaw is not designed for file processing.",
    ],
    # ── Data & Security ──
    ["— DATA SOVEREIGNTY & SECURITY —", "", "", ""],
    [
        "Data Location",
        "FULL sovereignty.\nAll data stays on user's infrastructure.\nNo data sent to third parties\n(except chosen LLM API calls).",
        "Partial sovereignty.\nFiles processed locally in VM.\nPrompts/responses go to Anthropic's cloud.\nTeam/Enterprise: not used for training.",
        "OpenClaw: Full data control.\nCowork: Files local, but prompts go to Anthropic.\nOpenClaw wins for data sovereignty.",
    ],
    [
        "Security Model",
        "Full system access by default.\nUser configures restrictions.\nNetwork policies via Kubernetes.\nCryptographic MCP signing.",
        "VM sandbox + bubblewrap + seccomp.\nFolder-specific permissions (read/write/create).\nExplicit directory approval required.\nKNOWN: Prompt injection vulnerability\n(file exfiltration via whitelisted APIs, unpatched Jan 2026).",
        "Cowork: Better default sandboxing.\nOpenClaw: Better production security (K8s).\nCowork has known unpatched vulnerability.",
    ],
    [
        "Enterprise Compliance",
        "No built-in compliance features.\nUser implements own controls.\nNo SSO/SCIM/audit logs natively.",
        "SSO/SCIM (Enterprise plan).\nAudit logs planned.\nConnector-level access controls.\nService account support.",
        "Cowork: Better enterprise compliance features.\nOpenClaw: Requires DIY compliance setup.",
    ],
    # ── Operational ──
    ["— OPERATIONAL CHARACTERISTICS —", "", "", ""],
    [
        "Always-On Availability",
        "YES – runs as persistent daemon 24/7.\nCron-scheduled autonomous tasks.\nHeartbeat monitoring.\nSurvives user disconnection.",
        "NO – session-based.\nTerminates if computer sleeps.\nRequires active Mac session.\nNo background persistence.",
        "OpenClaw wins decisively.\nAlways-on is OpenClaw's core strength.\nCowork requires active desktop session.",
    ],
    [
        "Multi-Channel Communication",
        "Native multi-channel:\nWhatsApp, Telegram, Slack, Discord,\nSignal, iMessage simultaneously.\nChannel-agnostic agent.",
        "Single interface (Claude app).\nConnectors for reading from external apps.\nNo native messaging channel support.",
        "OpenClaw wins decisively.\nMulti-channel messaging is its raison d'être.\nCowork is desktop-app only.",
    ],
    [
        "Collaboration Features",
        "Multi-agent coordination.\nHierarchical agent patterns.\nShared workspace via Kubernetes.",
        "No sharing or collaboration features.\nSingle-user desktop experience.\nTeam distribution planned (future).",
        "OpenClaw: Better for team/multi-agent use.\nCowork: Single-user only (for now).",
    ],
    [
        "Setup Complexity",
        "Moderate-High.\nRequires VPS provisioning, Docker/K8s,\nAPI key configuration, channel setup.\n2-5 hrs/mo maintenance.",
        "Low.\nBuilt into Claude app.\nOne-click folder permissions.\nNo infrastructure to manage.",
        "Cowork wins on ease of setup.\nOpenClaw requires technical expertise.\nCowork: minutes to start. OpenClaw: hours.",
    ],
    [
        "Maintenance Burden",
        "2-5 hours/month.\nUpdates, troubleshooting, config.\nAPI key rotation, model switching.\nInfra monitoring.",
        "Zero maintenance.\nManaged by Anthropic.\nAutomatic updates via Claude app.",
        "Cowork wins: zero maintenance.\nOpenClaw: ongoing operational overhead.",
    ],
    # ── Performance ──
    ["— PERFORMANCE —", "", "", ""],
    [
        "Task Completion Speed",
        "Depends on chosen model and infrastructure.\nNo published benchmarks.\nVariable based on API latency.",
        "92% reduction in task time vs. manual work\n(3.1 hrs → ~15 min, per Anthropic data).\nParallel execution reduces multi-task time\nto duration of longest single task.",
        "Cowork: Published performance data.\nOpenClaw: No comparable benchmarks.\nCowork has better documented productivity gains.",
    ],
    [
        "Scaling Ceiling",
        "Kubernetes: 1 to 100+ agents.\nPer-agent resource isolation.\nNetwork policies.\nProduction-grade multi-tenancy.",
        "Single Mac desktop.\nMultiple conversations share one VM.\nNo horizontal scaling.\nCapped by Mac hardware.",
        "OpenClaw wins on scaling.\nCowork is fundamentally single-machine.\nOpenClaw designed for enterprise scale.",
    ],
    # ── Summary ──
    ["— SUMMARY VERDICT —", "", "", ""],
    [
        "Best For",
        "Always-on automation.\nMulti-channel messaging agents.\nPrivacy-first deployments.\nBudget-conscious users.\nTeams wanting full control.",
        "Desktop knowledge work.\nDocument processing & organization.\nPolished single-user experience.\nUsers wanting zero-maintenance.\nmacOS-native workflows.",
        "NOT direct competitors.\nOpenClaw = persistent messaging automation platform.\nCowork = desktop productivity agent.\nComplementary tools for different workflows.",
    ],
    [
        "Biggest Advantage",
        "Open source, model-agnostic,\nalways-on, full data sovereignty,\nmassive plugin ecosystem,\nlowest possible TCO.",
        "Zero setup, polished UX,\nenterprise-grade Claude models,\nstrong sandboxing,\npredictable fixed pricing,\nnative document handling.",
        "OpenClaw: Control & flexibility.\nCowork: Simplicity & polish.",
    ],
    [
        "Biggest Weakness",
        "Operational overhead.\nSelf-hosting maintenance burden.\nNot designed for document/office work.\nNo native enterprise compliance.",
        "macOS only.\nSession-dependent (no 24/7 operation).\nClaude-only (vendor lock-in).\nKnown security vulnerability.\nNo multi-channel messaging.",
        "OpenClaw: Complexity.\nCowork: Limited platform & model choice.",
    ],
]

for r_idx, row_data in enumerate(cw_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws6.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        style_header(ws6, r_idx, len(cw_headers))
    elif row_data[0].startswith("—"):
        style_row(ws6, r_idx, len(cw_headers), fill=category_fill, font=category_font)
    else:
        style_row(ws6, r_idx, len(cw_headers))
        ws6.cell(row=r_idx, column=2).fill = openclaw_fill
        ws6.cell(row=r_idx, column=3).fill = cowork_fill
        ws6.cell(row=r_idx, column=4).fill = verdict_fill
        ws6.cell(row=r_idx, column=4).font = verdict_font

ws6.column_dimensions["A"].width = 28
ws6.column_dimensions["B"].width = 44
ws6.column_dimensions["C"].width = 44
ws6.column_dimensions["D"].width = 44

# ═══════════════════════════════════════════════════════════════
# SHEET 7 — NVIDIA GTC 2026: OpenClaw & NemoClaw
# ═══════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("NVIDIA GTC – OpenClaw")

nvidia_fill = PatternFill(start_color="76B900", end_color="76B900", fill_type="solid")
nvidia_header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")

gtc_headers = ["Topic", "Details"]

gtc_rows = [
    gtc_headers,
    # ── Jensen Huang Keynote ──
    ["— JENSEN HUANG KEYNOTE HIGHLIGHTS —", ""],
    [
        "Strategic Positioning",
        "Jensen Huang declared: \"Every single company in the world today has to have an OpenClaw strategy.\"\n"
        "Positioned OpenClaw as \"the operating system for personal AI\" — comparable to Windows, Linux, and HTML.\n"
        "NVIDIA views OpenClaw as a foundational layer for the emerging AI agent ecosystem.",
    ],
    [
        "Adoption Trajectory",
        "OpenClaw surpassed Linux's adoption curve — reaching equivalent milestones in 3 weeks vs. Linux's 30-year trajectory.\n"
        "250,000+ GitHub stars in 60 days (surpassing React's 10-year record).\n"
        "2.2 million weekly npm downloads.\n"
        "Described by Huang as \"the most popular open source project in the history of humanity.\"",
    ],
    [
        "Enterprise Penetration",
        "65% of OpenClaw users are in enterprise sectors.\n"
        "Gartner projects 40% of enterprise applications will embed task-specific AI agents by end of 2026 (up from <5% in 2025).\n"
        "Global agentic AI market: $7.6B (2025) → projected $199B by 2034.",
    ],
    # ── NemoClaw Stack ──
    ["— NEMOCLAW: NVIDIA'S OPENCLAW STACK —", ""],
    [
        "What Is NemoClaw",
        "An enterprise-grade, open-source (Apache 2.0) software stack that layers security, privacy, and runtime controls on top of OpenClaw.\n"
        "Pre-packaged: OpenClaw + NVIDIA Nemotron models + NVIDIA OpenShell runtime — installable with a single command.\n"
        "Enables production-ready agent deployment in under 1 hour.\n"
        "GitHub: github.com/NVIDIA/NemoClaw (alpha status as of March 2026).",
    ],
    [
        "NemoClaw Components",
        "1. OpenClaw: Core AI agent framework (task execution, messaging, MCP skills).\n"
        "2. NVIDIA OpenShell: Secure sandboxed runtime with kernel-level isolation (Landlock LSM, seccomp, bubblewrap).\n"
        "3. Nemotron Models: Open models optimized for agentic workloads (see below).\n"
        "4. Inference Routing: Intercepts all model calls to enforce privacy & cost constraints. Supports local vLLM, NIM, or NVIDIA cloud.",
    ],
    [
        "NemoClaw Inference Profiles",
        "Three profiles available:\n"
        "• default: NVIDIA cloud via integrate.api.nvidia.com (Nemotron-3-Super-120B)\n"
        "• nim-local: Local NVIDIA NIM service deployment\n"
        "• vllm: Local vLLM development option\n"
        "All profiles support policy-based cost and privacy constraints.",
    ],
    # ── OpenShell Security ──
    ["— NVIDIA OPENSHELL SECURITY RUNTIME —", ""],
    [
        "Architecture",
        "Open-sourced March 2026 under Apache 2.0.\n"
        "Provides sandboxed execution between AI agents and the host OS.\n"
        "Agents (OpenClaw, Claude Code, Codex) run unmodified inside OpenShell — zero code changes required.",
    ],
    [
        "Static Policies\n(locked at sandbox creation)",
        "• Filesystem: Read-only / read-write directory lists enforced via Landlock LSM kernel-level isolation.\n"
        "• Process: Agent runs as unprivileged user (defaults to 'sandbox'); rejects root. Seccomp filters block dangerous syscalls.\n"
        "• Landlock compatibility modes: 'best_effort' or 'hard_requirement'.",
    ],
    [
        "Dynamic Policies\n(hot-reloadable at runtime)",
        "• Network: Per-binary control (which executables can reach which endpoints). Protocol inspection for REST with TLS termination.\n"
        "• Per-method control: Governs specific API calls or shell functions.\n"
        "• Per-endpoint control: Limits traffic to specific IPs/domains.\n"
        "• Per-binary control: Restricts which executables (git, curl, python) agents can invoke.",
    ],
    [
        "Audit & Compliance",
        "All agent actions logged in audit trails for compliance and debugging.\n"
        "Developed in response to documented OpenClaw vulnerabilities (indirect prompt injection, malicious payloads in customization hub).\n"
        "Declarative YAML policy configuration.",
    ],
    # ── Nemotron Models ──
    ["— NEMOTRON MODEL FAMILY FOR OPENCLAW —", ""],
    [
        "Nemotron 3 Super 120B",
        "Parameters: 120B total, only 12B active (MoE architecture).\n"
        "Architecture: Hybrid Mamba-Transformer Mixture-of-Experts.\n"
        "Context Window: 1M tokens native.\n"
        "PinchBench Score: 85.6% (top open model; competitive with Claude Sonnet 4.6 at 86.9%).\n"
        "Throughput: 5× higher than predecessor; 42,855 tok/s prompt processing on DGX Spark.\n"
        "Generation: 18 tok/s on DGX Spark (128K input tokens).\n"
        "End-to-end latency: 99.4s for 128K token input on DGX Spark.\n"
        "License: Fully open (weights, datasets, recipes).",
    ],
    [
        "Nemotron 3 Nano 4B",
        "Parameters: 4B.\n"
        "Target: Resource-constrained RTX AI PCs, games, and apps.\n"
        "Strong instruction-following and tool use with minimal VRAM.\n"
        "Context Window: 131,072 tokens.\n"
        "Designed for local agents on consumer hardware.",
    ],
    [
        "Nemotron Ultra 253B",
        "Parameters: 253B.\n"
        "Context Window: 131,072 tokens; 4,096 max output.\n"
        "For enterprise-grade deep reasoning and complex multi-agent tasks.",
    ],
    [
        "Nemotron Super 49B v1.5",
        "Parameters: 49B.\n"
        "Context Window: 131,072 tokens; 4,096 max output.\n"
        "Mid-tier option balancing performance and resource needs.",
    ],
    # ── PinchBench ──
    ["— PINCHBENCH: OPENCLAW AGENT BENCHMARK —", ""],
    [
        "What Is PinchBench",
        "The first benchmarking system specifically designed for OpenClaw agents.\n"
        "23 tasks across 8 categories: long-term memory retrieval, inbox triage, spreadsheet processing, scheduling, and more.\n"
        "Tests real-world agentic task completion, not just code generation.",
    ],
    [
        "Leaderboard\n(Success Rate, 2026)",
        "1. Claude Sonnet 4.6:           86.9%\n"
        "2. GPT-5.4:                     86.4%\n"
        "3. Claude Opus 4.6:             86.3%\n"
        "4. Nemotron-3-Super-120B:       85.6%  ← Top open model\n"
        "5. Claude Opus 4.5:             85.4%\n"
        "(Gemini 3 Flash Preview reported at 95.1% in earlier runs)",
    ],
    [
        "Key Insight",
        "NVIDIA's open Nemotron-3-Super-120B trails the best proprietary model (Claude Sonnet 4.6) by only 1.3 percentage points.\n"
        "This near-parity enables fully local, private OpenClaw deployments without significant quality loss.\n"
        "Cost: $0 ongoing API fees vs. $5-420/mo for cloud models.",
    ],
    # ── Hardware Platforms ──
    ["— NVIDIA HARDWARE FOR OPENCLAW —", ""],
    [
        "DGX Spark\n(Desktop AI Supercomputer)",
        "GPU: NVIDIA Blackwell (5th gen Tensor Cores, 4th gen RT Cores).\n"
        "CPU: 20-core Arm (10× Cortex-X925 + 10× Cortex-A725).\n"
        "Memory: 128 GB LPDDR5x unified (273 GB/s bandwidth).\n"
        "Storage: 1 TB or 4 TB NVMe M.2.\n"
        "AI Compute: Up to 1,000 TOPS; 1 PFLOP at FP4 with sparsity.\n"
        "CUDA Cores: 6,144.\n"
        "Form Factor: 150mm × 150mm × 50.5mm, 1.2 kg.\n"
        "Power: 240W external (140W GPU TDP).\n"
        "Model Support: Up to 200B parameters (405B in dual-Spark config).\n"
        "Concurrency: 4 simultaneous subagents with ~3× throughput increase at 2.6× time.",
    ],
    [
        "DGX Station\n(Enterprise AI Supercomputer)",
        "GB300 variant: 748 GB coherent memory.\n"
        "Targets enterprise multi-agent workloads.\n"
        "Supports the full Nemotron model family including Ultra 253B.\n"
        "Designed for always-on, server-class OpenClaw deployments.",
    ],
    [
        "RTX PRO Workstations\n& GeForce RTX PCs",
        "Nemotron 3 Nano 4B runs on consumer RTX hardware.\n"
        "Dell Pro Max systems with GB10 (128 GB unified memory) available.\n"
        "Entry point for local OpenClaw agents without enterprise hardware.",
    ],
    # ── Enterprise Partners ──
    ["— ENTERPRISE PARTNERS (AGENT TOOLKIT) —", ""],
    [
        "Tier-1 Partners",
        "Adobe, Salesforce, SAP, ServiceNow, Siemens, CrowdStrike, Atlassian, Palantir.\n"
        "All building with the NVIDIA Agent Toolkit which includes NemoClaw.\n"
        "Focus: secure, enterprise-grade agentic AI deployments.",
    ],
    # ── Cost Impact ──
    ["— COST IMPACT —", ""],
    [
        "Vera Rubin Platform",
        "NVIDIA's next-gen inference platform delivers 40% lower cost-per-token vs. Blackwell deployments.\n"
        "Directly reduces OpenClaw operational costs for cloud-inference users.",
    ],
    [
        "Nemotron Cost Advantage",
        "Nemotron models reduce query costs by >50% vs. comparable proprietary models (per NVIDIA).\n"
        "Local inference: $0 ongoing API costs (hardware amortization only).\n"
        "Cloud Nemotron via NVIDIA API: significantly cheaper than Claude/GPT equivalents.",
    ],
    [
        "TCO Comparison:\nLocal NemoClaw on DGX Spark\nvs. Cloud OpenClaw",
        "Cloud OpenClaw (Claude Opus): $300-500/mo per agent in API fees.\n"
        "Local NemoClaw (DGX Spark): Hardware one-time cost; $0/mo API fees.\n"
        "Break-even: DGX Spark investment pays back in months for heavy users.\n"
        "Quality gap: Only 1.3% PinchBench difference (85.6% vs. 86.9%).",
    ],
    # ── Summary ──
    ["— SUMMARY: GTC 2026 SIGNIFICANCE FOR OPENCLAW —", ""],
    [
        "Strategic Takeaway",
        "NVIDIA's GTC 2026 announcements fundamentally elevate OpenClaw from a community project to an enterprise-grade platform:\n\n"
        "1. VALIDATION: Jensen Huang positioning OpenClaw as essential as Windows/Linux legitimizes it for enterprise.\n"
        "2. SECURITY: OpenShell/NemoClaw solves the #1 enterprise blocker (security concerns) with kernel-level sandboxing.\n"
        "3. LOCAL INFERENCE: Nemotron 3 Super 120B achieves near-parity with Claude/GPT at zero API cost.\n"
        "4. HARDWARE ECOSYSTEM: DGX Spark ($3K-$5K est.) to DGX Station creates a clear hardware upgrade path.\n"
        "5. ENTERPRISE PARTNERS: Adobe, Salesforce, SAP, ServiceNow adoption signals mainstream enterprise readiness.\n"
        "6. COST REDUCTION: Vera Rubin platform + Nemotron models compress costs by 40-50%+.\n\n"
        "Net effect: OpenClaw is now positioned as the default AI agent runtime with NVIDIA's full-stack backing, "
        "comparable to how Linux became the default server OS with Red Hat/IBM enterprise support.",
    ],
]

for r_idx, row_data in enumerate(gtc_rows, start=1):
    for c_idx, value in enumerate(row_data, start=1):
        ws7.cell(row=r_idx, column=c_idx, value=value)
    if r_idx == 1:
        for c in range(1, len(gtc_headers) + 1):
            cell = ws7.cell(row=r_idx, column=c)
            cell.font = nvidia_header_font
            cell.fill = nvidia_fill
            cell.alignment = center_alignment
            cell.border = thin_border
    elif row_data[0].startswith("—"):
        style_row(ws7, r_idx, len(gtc_headers), fill=category_fill, font=category_font)
    else:
        style_row(ws7, r_idx, len(gtc_headers))

ws7.column_dimensions["A"].width = 32
ws7.column_dimensions["B"].width = 110

# Move sheets into desired order:
# 1. OpenClaw vs Claude Cowork
# 2. NVIDIA GTC – OpenClaw  (new)
# 3-6. Original sheets
wb.move_sheet(ws6, offset=-5)  # OpenClaw vs Claude Cowork → first
wb.move_sheet(ws7, offset=-4)  # NVIDIA GTC → second

# ── Save ──
output_path = "/workspace/OpenClaw_Architecture_Comparison.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
