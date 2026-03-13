import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

HEADER_FONT = Font(bold=True, color="FFFFFF", size=12)
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUBHEADER_FONT = Font(bold=True, size=11)
BODY_FONT = Font(size=11)
WRAP = Alignment(wrap_text=True, vertical="top")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

def style_header(ws, row, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN_BORDER

def style_row(ws, row, num_cols, is_subheader=False):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = SUBHEADER_FONT if is_subheader else BODY_FONT
        if is_subheader:
            cell.fill = SUBHEADER_FILL
        cell.alignment = WRAP
        cell.border = THIN_BORDER

def auto_width(ws, num_cols, max_width=50):
    for col in range(1, num_cols + 1):
        letter = get_column_letter(col)
        max_len = 0
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=False):
            cell = row[0]
            if cell.value:
                lines = str(cell.value).split("\n")
                max_len = max(max_len, max(len(l) for l in lines))
        ws.column_dimensions[letter].width = min(max(max_len + 4, 15), max_width)


# ── TAB 1: Best Leaders in Gen AI ──────────────────────────────────────────────
ws1 = wb.active
ws1.title = "Best Leaders"
headers1 = ["Rank", "Leader", "Title / Company", "Key Achievements (Past 12 Months)", "Why They Stand Out", "Recognition"]
ws1.append(headers1)
style_header(ws1, 1, len(headers1))

leaders_data = [
    ["1", "Sam Altman", "CEO, OpenAI",
     "Led OpenAI to $110B funding round (Feb 2026) at $840B valuation; launched GPT-5.4 with native agentic capabilities; 800-900M weekly ChatGPT users; preparing potential IPO at $550-600B",
     "Transformed OpenAI into the fastest-growing business platform in history; championing AGI research incl. novel reasoning & scientific discovery; developing custom AI chips with Broadcom",
     "TIME 100 AI 2025; Gartner Emerging Leader"],
    ["2", "Jensen Huang", "CEO, NVIDIA",
     "NVIDIA market cap reached $4T; launched Nemotron 3 Super (120B-param open model, 5x throughput for agentic AI); GPUs power training/inference for virtually all leading AI models",
     "Made NVIDIA the indispensable foundation of modern AI infrastructure; drives the compute layer that enables the entire Gen AI ecosystem",
     "TIME 100 AI 2025"],
    ["3", "Dario Amodei", "CEO, Anthropic",
     "Raised $30B Series G at $380B valuation (Feb 2026); $14B annualized revenue with 10x annual growth for 3 consecutive years; Claude Code generates $2.5B+ run-rate revenue; 4% of all GitHub public commits authored by Claude Code",
     "Pioneered safety-first approach to frontier AI; Claude 5 achieved record 87.3% GPQA Diamond score; 300K+ business customers incl. 8 of Fortune 10",
     "TIME 100 AI 2025"],
    ["4", "Sundar Pichai", "CEO, Google / Alphabet",
     "Launched Gemini 3.1 Pro (Feb 2026) with advanced reasoning; expanded Google's AI ecosystem across search, cloud, and enterprise",
     "Integrating generative AI across Google's massive product surface area; competitive multi-model strategy",
     "TIME 100 AI 2025"],
    ["5", "Satya Nadella", "Chairman & CEO, Microsoft",
     "Deepened Microsoft-OpenAI partnership; Microsoft a lead investor in OpenAI's $110B round; integrating AI across Azure, Office 365, GitHub Copilot",
     "Enterprise AI adoption at unprecedented scale through Azure & Microsoft 365 Copilot ecosystem",
     "TIME 100 AI 2025"],
    ["6", "Mustafa Suleyman", "CEO, Microsoft AI",
     "Leading Microsoft's AI product strategy; shaping AI integration across consumer and enterprise products",
     "Brings DeepMind co-founder experience to Microsoft's AI ambitions; focused on practical AI deployment",
     "TIME 100 AI 2025"],
    ["7", "Elon Musk", "CEO, xAI / Tesla / SpaceX",
     "xAI reached ~$3.8B annualized revenue (38x YoY growth); raised $20B Series E at $230B valuation (Jan 2026); merged xAI with X platform",
     "Rapidly scaling Grok models; building massive compute clusters; ambitious AGI timeline",
     "TIME 100 AI 2025"],
    ["8", "Ravi Kumar S", "CEO, Cognizant",
     "Bet $1B on Gen AI development; launched Agent Foundry for enterprise AI agents; Synapse initiative to train 2M people in Gen AI",
     "Driving large-scale enterprise AI transformation; democratizing AI expertise across the workforce",
     "TIME 100 AI 2025"],
    ["9", "Yann LeCun", "VP & Chief AI Scientist, Meta; Founder, AMI Labs",
     "Founded AMI Labs (Paris); raised $1B seed round at $3.5B valuation — Europe's largest-ever seed; continues shaping Meta's open-source AI strategy (LLaMA models)",
     "Pioneering open-source AI research; Turing Award winner; bridging academia and industry",
     "Turing Award; TIME 100 AI"],
    ["10", "Liang Wenfeng", "CEO, DeepSeek",
     "Launched DeepSeek-V3.2 & V3.2-Speciale (Dec 2025) matching/exceeding GPT-5 on reasoning; DeepSeek-R1 demonstrated RL-only reasoning; all models open-sourced under MIT License",
     "Demonstrated frontier-level AI at dramatically lower cost ($2/M ops vs $60 US); gold-medal results in IMO, IOI, ICPC",
     "N/A"],
]

for i, row in enumerate(leaders_data, start=2):
    ws1.append(row)
    style_row(ws1, i, len(headers1))

auto_width(ws1, len(headers1))


# ── TAB 2: Companies with Most Momentum ────────────────────────────────────────
ws2 = wb.create_sheet("Most Momentum")
headers2 = ["Rank", "Company", "Valuation", "Latest Funding", "Revenue / Growth", "Key Momentum Indicators", "Category"]
ws2.append(headers2)
style_header(ws2, 1, len(headers2))

momentum_data = [
    ["1", "OpenAI", "$840B post-money", "$110B (Feb 2026) — led by Amazon ($50B), NVIDIA ($30B), SoftBank ($30B)",
     "800-900M weekly ChatGPT users; 2B daily prompts; 1M+ business customers; preparing $550-600B IPO",
     "Largest private funding round in history; fastest-growing business platform ever; GPT-5.4 agentic launch; Gartner Emerging Leader",
     "Foundation Model / Consumer + Enterprise AI"],
    ["2", "Anthropic", "$380B post-money", "$30B Series G (Feb 2026)",
     "$14B annualized revenue; 10x annual growth for 3 consecutive years; 300K+ business customers; 500+ customers at >$1M/yr",
     "Claude Code: $2.5B+ run-rate, 4% of all GitHub public commits; Claude 5 broke GPQA Diamond record (87.3%); 8 of Fortune 10 as customers",
     "Foundation Model / Enterprise AI"],
    ["3", "xAI", "$230B", "$20B Series E (Jan 2026)",
     "~$3.8B annualized revenue; 38x YoY growth; merged with X platform",
     "Grok models rapidly improving; massive compute investment; SuperGrok products ($30-$300/mo)",
     "Foundation Model / Consumer AI"],
    ["4", "Cursor (Anysphere)", "N/A (private)", "N/A",
     "$2B annualized revenue in 18 months — one of fastest revenue ramps in enterprise software history",
     "AI-powered code editor category leader; explosive developer adoption",
     "AI-Native Developer Tool"],
    ["5", "Perplexity", "$20B", "$200M (Sep 2025)",
     "$148M annualized revenue (Jun 2025); 800% YoY growth; targeting $656M by end-2026",
     "30M monthly users; 780M queries/month; expanding into advertising and commerce",
     "AI Search / Consumer AI"],
    ["6", "DeepSeek", "Private (China-based)", "Self-funded",
     "Rapidly growing user base and API adoption globally",
     "V3.2 matched/exceeded GPT-5 on reasoning at dramatically lower cost; open-source MIT license; gold-medal competition results",
     "Foundation Model / Open Source"],
    ["7", "Mistral AI", "~$7B (est.)", "Multiple rounds, EU-backed",
     "$50M+ in government AI contracts; 1.1B+ monthly queries; 6.2M active developers (up from 1.5M)",
     "40% of Europe's Fortune 500 deploying/piloting; sovereign AI leader; strong enterprise deals avg $1.2M/yr",
     "Foundation Model / Sovereign AI"],
    ["8", "Scale AI", "$14.3B+ valuation", "$14.3B (Jun 2025, Meta-led)",
     "Leading AI training data and model evaluation provider",
     "Critical infrastructure for training frontier models; Meta strategic partnership",
     "AI Infrastructure / Data"],
]

for i, row in enumerate(momentum_data, start=2):
    ws2.append(row)
    style_row(ws2, i, len(headers2))

auto_width(ws2, len(headers2))


# ── TAB 3: Biggest Breakthroughs ───────────────────────────────────────────────
ws3 = wb.create_sheet("Biggest Breakthroughs")
headers3 = ["Rank", "Breakthrough", "Company / Org", "Date", "Description", "Impact / Significance", "Technical Details"]
ws3.append(headers3)
style_header(ws3, 1, len(headers3))

breakthrough_data = [
    ["1", "DeepSeek-R1: RL-Only Reasoning", "DeepSeek",
     "Jan 2025",
     "First demonstration that reasoning capabilities can emerge purely through reinforcement learning without supervised fine-tuning",
     "Paradigm shift — proved AI can develop chain-of-thought reasoning from scratch via RL; frontier-level performance at a fraction of Western costs ($2/M ops vs $60)",
     "DeepSeek-R1-Zero trained purely via RL; R1 matches OpenAI o1; distilled 32B version outperforms o1-mini; open-sourced under MIT License"],
    ["2", "Claude 5 GPQA Diamond Record", "Anthropic",
     "Mar 2026",
     "Claude 5 achieved 87.3% on GPQA Diamond — first AI system to exceed 85% on this rigorous scientific reasoning benchmark",
     "8.1 percentage point improvement over previous record; demonstrates new frontier in scientific reasoning capability",
     "Achieved via inference-time reasoning optimization (not larger models); Extended Thinking mode improved performance by 15pp over standard mode"],
    ["3", "GPT-5.4 Autonomous Agents", "OpenAI",
     "Mar 2026",
     "Native computer-use capabilities enabling autonomous agent workflows — AI can operate computers, issue keyboard/mouse commands, complete tasks across applications",
     "Major step toward fully autonomous AI agents; 33% fewer false claims vs. earlier versions",
     "Multi-source info gathering; complex professional workflows (coding, spreadsheets, documents); enterprise-focused agentic AI"],
    ["4", "DeepSeek-V3.2-Speciale", "DeepSeek",
     "Dec 2025",
     "Open-source model matching/surpassing GPT-5 on reasoning; 'Thinking in Tool-Use' integrating CoT reasoning directly into tool execution",
     "Gold-medal results in IMO, IOI, ICPC, CMO — competitive math/coding olympiads; open-source frontier model under MIT License",
     "685B total params, ~37B active per token (MoE); DeepSeek Sparse Attention reducing complexity from O(L^2) to O(L); trained on 85K+ complex instructions across 1800+ simulated environments"],
    ["5", "Gemini 3.1 Pro", "Google DeepMind",
     "Feb 2026",
     "Advanced reasoning and complex problem-solving model with significant benchmark improvements for agentic workflows",
     "Keeps Google competitive in the frontier model race; advanced agentic capabilities",
     "Multi-modal reasoning; improved tool use and planning"],
    ["6", "NVIDIA Nemotron 3 Super", "NVIDIA",
     "2025-2026",
     "120B-parameter open model delivering 5x higher throughput for agentic AI with 1M-token context window",
     "Prevents goal drift in multi-agent workflows; major efficiency gain for production agentic systems",
     "Hybrid mixture-of-experts architecture; multi-token prediction for 3x faster inference; 1M-token context to maintain coherence"],
    ["7", "AMI Labs Foundation", "AMI Labs (Yann LeCun)",
     "Mar 2026",
     "Europe's largest-ever seed round ($1B at $3.5B valuation) to build next-generation AI foundation models",
     "Signals growing European competitiveness in frontier AI; open-science approach from Turing Award winner",
     "Backed by Cathay Innovation, Bezos Expeditions, NVIDIA, Eric Schmidt, Temasek, and others"],
]

for i, row in enumerate(breakthrough_data, start=2):
    ws3.append(row)
    style_row(ws3, i, len(headers3))

auto_width(ws3, len(headers3))


# ── TAB 4: Best Investors ──────────────────────────────────────────────────────
ws4 = wb.create_sheet("Best Investors")
headers4 = ["Rank", "Investor / Firm", "Type", "Notable Gen AI Investments", "Estimated AI Portfolio Value", "Investment Thesis / Strategy", "Standout Metric"]
ws4.append(headers4)
style_header(ws4, 1, len(headers4))

investor_data = [
    ["1", "SoftBank Vision Fund", "VC / Growth",
     "OpenAI ($30B in Feb 2026 round; $40B lead in Mar 2025); xAI",
     "$70B+ deployed in AI",
     "Mega-rounds in frontier AI companies; conviction bets on platform winners; global scale",
     "Led the two largest private funding rounds in history (OpenAI $40B + $110B)"],
    ["2", "Amazon", "Strategic / Corporate",
     "OpenAI ($50B in Feb 2026 round); AWS AI infrastructure",
     "$50B+ in AI investments",
     "Securing AI compute/infrastructure dominance via AWS; strategic partnerships with frontier labs",
     "Largest single investor in OpenAI's $110B round ($50B)"],
    ["3", "NVIDIA", "Strategic / Corporate",
     "OpenAI ($30B in Feb 2026 round); AMI Labs; broad AI ecosystem",
     "$30B+ in AI investments",
     "Investing in companies that drive GPU demand; building AI ecosystem moat around CUDA/hardware",
     "Co-lead of OpenAI $110B round; Nemotron open models"],
    ["4", "Sequoia Capital", "VC",
     "OpenAI; planning major Anthropic investment; broad AI portfolio",
     "Multi-billion AI portfolio",
     "Early-stage to growth; identifying AI platform companies; deep technical diligence",
     "Consistently backing winners across AI waves"],
    ["5", "Andreessen Horowitz (a16z)", "VC",
     "Broad AI portfolio across infrastructure, applications, and AGI",
     "Multi-billion AI portfolio",
     "AGI infrastructure thesis; AI-native application layer; developer tooling",
     "Dedicated AI fund; prolific deal flow"],
    ["6", "Coatue Management", "Hedge Fund / VC",
     "Hugging Face, Runway, Notion, Weights & Biases, Supabase, ClickHouse",
     "$70B AUM total",
     "AI-native software and developer infrastructure; uses proprietary AI data infra ($45M/yr) for deal sourcing",
     "Pioneering AI-augmented investment process"],
    ["7", "Founders Fund (Brian Singerman)", "VC",
     "OpenAI (2016 early investor), Anduril, Scale AI, Applied Intuition",
     "Portfolio companies valued >$150B combined",
     "Contrarian deep-tech bets; early conviction in frontier AI before mainstream",
     "18.3x average returns — 3x VC industry average"],
    ["8", "Elad Gil", "Angel / Individual",
     "Perplexity, Cognition, Scale AI — 36 AI deals since Jan 2024",
     "Diversified AI angel portfolio",
     "High-volume seed/early-stage investing in AI-native startups; application layer focus",
     "#1 ranked AI angel investor by deal count"],
    ["9", "Vinod Khosla / Khosla Ventures", "VC / Angel",
     "OpenAI ($50M early investment in 2019); robotics, medtech, AI startups",
     "Multi-billion AI portfolio",
     "Early conviction in frontier AI; patient capital through multiple hype cycles",
     "TIME 100 Most Influential in AI 2024"],
    ["10", "Iconiq Capital", "Growth",
     "Anthropic ($13B round lead, Sep 2025)",
     "Multi-billion AI growth portfolio",
     "Late-stage growth investing in proven AI platforms",
     "Led Anthropic's $13B round — one of the largest AI deals of 2025"],
]

for i, row in enumerate(investor_data, start=2):
    ws4.append(row)
    style_row(ws4, i, len(headers4))

auto_width(ws4, len(headers4))


# ── TAB 5: Best Hotspots ──────────────────────────────────────────────────────
ws5 = wb.create_sheet("Best Hotspots")
headers5 = ["Rank", "City / Region", "Country", "AI Funding Concentration", "Key Strengths", "Notable Companies / Orgs", "Recent Developments (Past 12 Months)"]
ws5.append(headers5)
style_header(ws5, 1, len(headers5))

hotspot_data = [
    ["1", "San Francisco / Silicon Valley", "United States",
     "62.4% of venture funding directed to AI; 65% of global AI-native funding",
     "Foundation models, AI applications, venture capital density, talent pool, compute infrastructure",
     "OpenAI, Anthropic, xAI, Scale AI, Cursor, Perplexity, NVIDIA (Santa Clara)",
     "Epicenter of $110B OpenAI round; Anthropic $30B round; continued dominance in frontier AI research and deployment"],
    ["2", "Beijing", "China",
     "66.2% of local startup funding goes to AI — highest concentration globally",
     "Foundation models at dramatically lower cost ($2/M ops vs $60 US); massive government support; large talent pool",
     "DeepSeek, Baidu, ByteDance, Zhipu AI, Moonshot AI",
     "DeepSeek's V3.2/R1 breakthroughs showed frontier-level AI at fraction of US cost; reshaping global AI cost economics"],
    ["3", "London", "United Kingdom",
     "Major and growing AI hub",
     "World-class AI research (DeepMind); regulatory leadership; strong talent from top universities",
     "Google DeepMind, Stability AI, Faculty AI",
     "OpenAI named London its largest research hub outside SF (Feb 2026); researchers own key components of frontier model development"],
    ["4", "Seattle", "United States",
     "Major North American AI center",
     "Cloud infrastructure (AWS, Azure); enterprise AI; strong academic research (UW)",
     "Microsoft, Amazon/AWS, Allen Institute for AI",
     "Microsoft AI division expansion under Mustafa Suleyman; AWS as backbone of AI compute"],
    ["5", "New York", "United States",
     "Major North American AI center",
     "AI in finance, media, advertising; enterprise applications; growing startup ecosystem",
     "Bloomberg AI, various AI startups, financial services AI",
     "Growing AI startup ecosystem; financial services AI deployment acceleration"],
    ["6", "Paris", "France",
     "4.3% of global AI-native funding",
     "Foundation models (Mistral AI); strong academic base (INRIA, ENS); EU sovereign AI initiatives",
     "Mistral AI, AMI Labs (Yann LeCun), Hugging Face",
     "AMI Labs raised $1B seed (Europe's largest-ever); Mistral deployed in 40% of Europe's Fortune 500; EU sovereign AI contracts"],
    ["7", "Singapore", "Singapore",
     "Growing AI investment hub",
     "World's best AI readiness score; government-wide AI implementation; AI governance leadership",
     "Government AI initiatives, AI Singapore, Southeast Asian hub",
     "Leading in practical AI governance frameworks; attracting global AI companies to Asia-Pacific hub"],
    ["8", "Toronto-Waterloo Corridor", "Canada",
     "50.3% AI funding concentration",
     "Pioneer AI research (Geoffrey Hinton, U of T); strong talent pipeline; affordable relative to SF",
     "Cohere, Vector Institute, U of Toronto",
     "Growing as alternative to SF for AI talent; Cohere enterprise traction"],
    ["9", "Shanghai / Greater Bay Area", "China",
     "Emerging AI powerhouse",
     "Hardware manufacturing ecosystem; application-layer AI; massive consumer market",
     "Alibaba AI, Tencent AI Lab, semiconductor ecosystem",
     "Growing investment in AI applications; connection to hardware supply chain (Shenzhen)"],
    ["10", "Zurich", "Switzerland",
     "Specialized AI hub",
     "Smart city applications; fintech AI; ETH Zurich research excellence; Google research presence",
     "ETH Zurich, Google Brain Zurich, fintech AI startups",
     "Leading in applied AI for financial services and urban tech; strong European AI research output"],
]

for i, row in enumerate(hotspot_data, start=2):
    ws5.append(row)
    style_row(ws5, i, len(headers5))

auto_width(ws5, len(headers5))


# ── TAB 6: Summary Dashboard ──────────────────────────────────────────────────
ws6 = wb.create_sheet("Summary Dashboard")
headers6 = ["Category", "Winner / Top Pick", "Runner-Up", "Key Data Point", "Source / Date"]
ws6.append(headers6)
style_header(ws6, 1, len(headers6))

summary_data = [
    ["Best Overall Leader", "Sam Altman (CEO, OpenAI)",
     "Dario Amodei (CEO, Anthropic)",
     "Led $110B raise at $840B valuation; 800-900M weekly users; preparing IPO",
     "TechCrunch, TIME 100 AI 2025 (Feb 2026)"],
    ["Most Influential Hardware Leader", "Jensen Huang (CEO, NVIDIA)",
     "N/A",
     "NVIDIA at $4T market cap; GPUs power virtually all frontier AI training",
     "TIME 100 AI 2025"],
    ["Company with Most Momentum", "OpenAI",
     "Anthropic",
     "OpenAI: $110B funding, 800-900M users; Anthropic: $30B funding, 10x growth 3 yrs straight",
     "TechCrunch, Anthropic Press (Feb 2026)"],
    ["Fastest Revenue Ramp (Startup)", "Cursor (Anysphere)",
     "Perplexity",
     "Cursor: $2B ARR in 18 months; Perplexity: 800% YoY growth",
     "AI Business Review (Mar 2026)"],
    ["Biggest Breakthrough", "DeepSeek-R1 (RL-Only Reasoning)",
     "Claude 5 GPQA Diamond Record (87.3%)",
     "Proved reasoning can emerge from pure RL; frontier AI at 1/30th the cost",
     "DeepSeek (Jan 2025), Anthropic (Mar 2026)"],
    ["Best Investor (Institutional)", "SoftBank Vision Fund",
     "Amazon",
     "Led two largest private funding rounds in history: OpenAI $40B + $110B",
     "TechCrunch (Feb 2026)"],
    ["Best Investor (Angel)", "Elad Gil",
     "Daniel Gross",
     "36 AI deals since Jan 2024; backed Perplexity, Cognition, Scale AI",
     "CB Insights (2025)"],
    ["Best Investor (Early Conviction)", "Founders Fund (Brian Singerman)",
     "Vinod Khosla / Khosla Ventures",
     "18.3x avg returns (3x industry avg); early OpenAI investor (2016)",
     "VC Analysis (2025)"],
    ["Best AI Hotspot", "San Francisco / Silicon Valley",
     "Beijing",
     "SF: 62-65% of global AI-native funding; Beijing: 66.2% local AI concentration, frontier AI at 1/30th cost",
     "10Think.AI, Visual Capitalist (2025)"],
    ["Fastest-Rising Hotspot", "Paris",
     "London",
     "Paris: AMI Labs $1B seed, Mistral in 40% EU Fortune 500; London: OpenAI's largest non-SF research hub",
     "Observer, UKTech News (2026)"],
]

for i, row in enumerate(summary_data, start=2):
    ws6.append(row)
    style_row(ws6, i, len(headers6))

auto_width(ws6, len(headers6))

# Freeze panes for all sheets
for ws in wb.worksheets:
    ws.freeze_panes = "A2"

output_path = "/workspace/generative_ai_market_insights.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
