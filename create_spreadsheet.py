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


# ── REGIONAL TABS ──────────────────────────────────────────────────────────────

SECTION_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
SECTION_FONT = Font(bold=True, color="FFFFFF", size=12)

def add_section_header(ws, row_num, label, num_cols):
    ws.cell(row=row_num, column=1, value=label)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=num_cols)

def add_sub_header(ws, row_num, headers):
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=row_num, column=col_idx, value=h)
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN_BORDER

def add_data_row(ws, row_num, data, num_cols):
    for col_idx, val in enumerate(data, start=1):
        ws.cell(row=row_num, column=col_idx, value=val)
    style_row(ws, row_num, num_cols)

NUM_COLS_REGIONAL = 7


# ── TAB 7: Americas ───────────────────────────────────────────────────────────
ws_am = wb.create_sheet("Americas")
r = 1

add_section_header(ws_am, r, "LEADERS", NUM_COLS_REGIONAL); r += 1
ldr_h = ["#", "Leader", "Title / Company", "HQ", "Key Achievements (Past 12 Months)", "Why They Stand Out", "Recognition"]
add_sub_header(ws_am, r, ldr_h); r += 1
am_leaders = [
    ["1", "Sam Altman", "CEO, OpenAI", "San Francisco, CA",
     "$110B funding at $840B valuation; GPT-5.4 agentic launch; 800-900M weekly users; preparing IPO",
     "Built the fastest-growing business platform in history; AGI research leader",
     "TIME 100 AI 2025"],
    ["2", "Jensen Huang", "CEO, NVIDIA", "Santa Clara, CA",
     "NVIDIA at $4T market cap; Nemotron 3 Super; GPUs power virtually all frontier AI",
     "Indispensable AI infrastructure — every major model trains on NVIDIA hardware",
     "TIME 100 AI 2025"],
    ["3", "Dario Amodei", "CEO, Anthropic", "San Francisco, CA",
     "$30B Series G at $380B; $14B revenue; 10x annual growth 3 yrs; Claude 5 GPQA record",
     "Safety-first frontier AI; Claude Code at $2.5B+ run-rate; 4% of GitHub public commits",
     "TIME 100 AI 2025"],
    ["4", "Satya Nadella", "Chairman & CEO, Microsoft", "Redmond, WA",
     "Lead investor in OpenAI $110B; Azure AI expansion; Microsoft 365 Copilot at scale",
     "Enterprise AI adoption at unprecedented scale",
     "TIME 100 AI 2025"],
    ["5", "Elon Musk", "CEO, xAI / Tesla / SpaceX", "Austin, TX",
     "xAI at ~$3.8B revenue (38x YoY); $20B Series E at $230B; merged with X platform",
     "Massive compute investment; Grok models; ambitious AGI timeline",
     "TIME 100 AI 2025"],
    ["6", "Mustafa Suleyman", "CEO, Microsoft AI", "Redmond, WA",
     "Leading Microsoft's AI product strategy; AI integration across consumer & enterprise",
     "DeepMind co-founder experience brought to Microsoft",
     "TIME 100 AI 2025"],
    ["7", "Sundar Pichai", "CEO, Google / Alphabet", "Mountain View, CA",
     "Gemini 3.1 Pro launch; AI integration across Search, Cloud, enterprise",
     "Multi-model strategy across Google's massive product surface",
     "TIME 100 AI 2025"],
]
for row in am_leaders:
    add_data_row(ws_am, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_am, r, "COMPANIES WITH MOST MOMENTUM", NUM_COLS_REGIONAL); r += 1
co_h = ["#", "Company", "HQ", "Valuation", "Revenue / Growth", "Key Momentum Indicators", "Category"]
add_sub_header(ws_am, r, co_h); r += 1
am_companies = [
    ["1", "OpenAI", "San Francisco, CA", "$840B",
     "800-900M weekly users; 2B daily prompts; 1M+ business customers",
     "$110B funding round — largest private round in history; GPT-5.4 agentic launch",
     "Foundation Model / Consumer + Enterprise"],
    ["2", "Anthropic", "San Francisco, CA", "$380B",
     "$14B annualized revenue; 10x growth 3 yrs; 300K+ business customers",
     "Claude Code $2.5B+ run-rate; Claude 5 GPQA record; 8 of Fortune 10",
     "Foundation Model / Enterprise AI"],
    ["3", "xAI", "Austin, TX", "$230B",
     "~$3.8B annualized revenue; 38x YoY",
     "$20B Series E; merged with X platform; SuperGrok products",
     "Foundation Model / Consumer AI"],
    ["4", "Cursor (Anysphere)", "San Francisco, CA", "N/A (private)",
     "$2B ARR in 18 months",
     "Fastest revenue ramp in enterprise software history",
     "AI-Native Developer Tool"],
    ["5", "Perplexity", "San Francisco, CA", "$20B",
     "$148M ARR; 800% YoY; targeting $656M by end-2026",
     "30M monthly users; expanding into advertising & commerce",
     "AI Search / Consumer AI"],
    ["6", "Scale AI", "San Francisco, CA", "$14.3B+",
     "Leading AI training data provider",
     "$14.3B Meta-led round; critical infrastructure for frontier models",
     "AI Infrastructure / Data"],
    ["7", "Cohere", "Toronto, Canada", "~$5.5B",
     "Growing enterprise customer base across financial services, healthcare",
     "Enterprise AI with private deployment; strong Canadian AI ecosystem",
     "Enterprise AI / NLP"],
    ["8", "Adept", "San Francisco, CA", "~$1B+",
     "Agentic AI for workforce automation",
     "Building enterprise agents for complex workflows across applications",
     "Agentic AI"],
]
for row in am_companies:
    add_data_row(ws_am, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_am, r, "BIGGEST BREAKTHROUGHS", NUM_COLS_REGIONAL); r += 1
bt_h = ["#", "Breakthrough", "Company", "Date", "Description", "Impact", "Technical Details"]
add_sub_header(ws_am, r, bt_h); r += 1
am_breakthroughs = [
    ["1", "GPT-5.4 Autonomous Agents", "OpenAI", "Mar 2026",
     "Native computer-use for autonomous agent workflows",
     "Major step toward fully autonomous AI agents; 33% fewer false claims",
     "Multi-source info gathering; agentic computer use; enterprise workflows"],
    ["2", "Claude 5 GPQA Diamond Record", "Anthropic", "Mar 2026",
     "87.3% GPQA Diamond — first AI system to exceed 85%",
     "8.1pp improvement; new frontier in scientific reasoning",
     "Inference-time reasoning optimization; Extended Thinking mode (+15pp)"],
    ["3", "NVIDIA Nemotron 3 Super", "NVIDIA", "2025-2026",
     "120B-param open model; 5x throughput for agentic AI; 1M-token context",
     "Prevents goal drift in multi-agent systems; major efficiency gain",
     "Hybrid MoE; multi-token prediction; 3x faster inference"],
    ["4", "Gemini 3.1 Pro", "Google DeepMind", "Feb 2026",
     "Advanced reasoning model for complex problem-solving & agentic workflows",
     "Keeps Google competitive in frontier model race",
     "Multi-modal reasoning; improved tool use and planning"],
]
for row in am_breakthroughs:
    add_data_row(ws_am, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_am, r, "TOP INVESTORS", NUM_COLS_REGIONAL); r += 1
inv_h = ["#", "Investor / Firm", "Type", "HQ", "Notable Gen AI Investments", "Investment Thesis", "Standout Metric"]
add_sub_header(ws_am, r, inv_h); r += 1
am_investors = [
    ["1", "Amazon", "Strategic / Corporate", "Seattle, WA",
     "OpenAI ($50B in Feb 2026 round); AWS AI infrastructure",
     "AI compute/infrastructure dominance via AWS",
     "Largest single investor in OpenAI's $110B round"],
    ["2", "NVIDIA", "Strategic / Corporate", "Santa Clara, CA",
     "OpenAI ($30B); AMI Labs; broad AI ecosystem investments",
     "Investing in companies that drive GPU demand",
     "Co-lead of OpenAI $110B round"],
    ["3", "Sequoia Capital", "VC", "Menlo Park, CA",
     "OpenAI; Anthropic; broad AI portfolio",
     "Early-stage to growth; identifying AI platform winners",
     "Consistently backing winners across AI waves"],
    ["4", "Andreessen Horowitz (a16z)", "VC", "Menlo Park, CA",
     "Broad AI portfolio: infrastructure, apps, AGI",
     "AGI infrastructure thesis; AI-native application layer",
     "Dedicated AI fund; prolific deal flow"],
    ["5", "Founders Fund (Brian Singerman)", "VC", "San Francisco, CA",
     "OpenAI (2016), Anduril, Scale AI, Applied Intuition",
     "Contrarian deep-tech bets; early conviction in frontier AI",
     "18.3x avg returns — 3x industry average"],
    ["6", "Coatue Management", "Hedge Fund / VC", "New York, NY",
     "Hugging Face, Runway, Notion, Weights & Biases",
     "AI-native software & developer infra; AI-augmented deal sourcing",
     "$70B AUM; $45M/yr on proprietary AI data infra"],
    ["7", "Iconiq Capital", "Growth", "San Francisco, CA",
     "Anthropic ($13B round lead, Sep 2025)",
     "Late-stage growth in proven AI platforms",
     "Led one of the largest AI deals of 2025"],
    ["8", "Elad Gil", "Angel / Individual", "San Francisco, CA",
     "Perplexity, Cognition, Scale AI — 36 AI deals since Jan 2024",
     "High-volume seed investing; application layer focus",
     "#1 ranked AI angel investor by deal count"],
    ["9", "Vinod Khosla / Khosla Ventures", "VC / Angel", "Menlo Park, CA",
     "OpenAI ($50M early, 2019); robotics, medtech, AI startups",
     "Early conviction; patient capital through hype cycles",
     "TIME 100 Most Influential in AI 2024"],
]
for row in am_investors:
    add_data_row(ws_am, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_am, r, "HOTSPOTS", NUM_COLS_REGIONAL); r += 1
hs_h = ["#", "City / Region", "Country", "AI Funding Concentration", "Key Strengths", "Notable Companies", "Recent Developments"]
add_sub_header(ws_am, r, hs_h); r += 1
am_hotspots = [
    ["1", "San Francisco / Silicon Valley", "United States",
     "62-65% of global AI-native funding",
     "Foundation models, VC density, talent pool, compute infrastructure",
     "OpenAI, Anthropic, xAI, Scale AI, Cursor, Perplexity, NVIDIA",
     "Epicenter of $110B OpenAI & $30B Anthropic rounds"],
    ["2", "Seattle", "United States",
     "Major North American AI center",
     "Cloud infrastructure (AWS, Azure); enterprise AI; UW research",
     "Microsoft, Amazon/AWS, Allen Institute for AI",
     "Microsoft AI division expansion; AWS as backbone of AI compute"],
    ["3", "New York", "United States",
     "Major North American AI center",
     "AI in finance, media, advertising; enterprise applications",
     "Bloomberg AI, Coatue, financial services AI startups",
     "Growing AI startup ecosystem; financial services AI acceleration"],
    ["4", "Toronto-Waterloo Corridor", "Canada",
     "50.3% AI funding concentration",
     "Pioneer research (Hinton, U of T); strong talent; affordable vs SF",
     "Cohere, Vector Institute, U of Toronto",
     "Growing alternative to SF for AI talent; Cohere enterprise traction"],
    ["5", "Austin, TX", "United States",
     "Growing AI hub",
     "xAI/Tesla AI presence; lower cost of living; tech migration from SF",
     "xAI, Tesla AI, growing startup ecosystem",
     "xAI HQ; emerging as secondary AI hub"],
    ["6", "Boston / Cambridge", "United States",
     "Major AI research center",
     "MIT, Harvard AI research; biotech AI; strong academic pipeline",
     "MIT CSAIL, Harvard AI, biotech AI startups",
     "AI-biotech convergence; strong research-to-startup pipeline"],
]
for row in am_hotspots:
    add_data_row(ws_am, r, row, NUM_COLS_REGIONAL); r += 1

auto_width(ws_am, NUM_COLS_REGIONAL)


# ── TAB 8: Europe ─────────────────────────────────────────────────────────────
ws_eu = wb.create_sheet("Europe")
r = 1

add_section_header(ws_eu, r, "LEADERS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_eu, r, ldr_h); r += 1
eu_leaders = [
    ["1", "Arthur Mensch", "CEO & Co-Founder, Mistral AI", "Paris, France",
     "Grew Mistral to ~$7B+ valuation; 1.1B+ monthly queries; 6.2M active developers; 40% of EU Fortune 500 deploying/piloting",
     "Leading European sovereign AI champion; open-source model leadership; €50M+ government contracts",
     "Top EMEA CEO"],
    ["2", "Yann LeCun", "VP & Chief AI Scientist, Meta; Founder, AMI Labs", "Paris, France",
     "Founded AMI Labs; raised $1B seed (Europe's largest-ever) at $3.5B valuation",
     "Turing Award winner; open-science AI research pioneer; bridging US tech & European innovation",
     "Turing Award; TIME 100 AI"],
    ["3", "Timothée Lacroix", "CTO & Co-Founder, Mistral AI", "Paris, France",
     "Led technical development of Mistral models incl. Mixtral MoE architecture",
     "Former Meta FAIR researcher; driving European model innovation",
     "Top EMEA CTO 2025"],
    ["4", "Jonas Andrulis", "CEO & Founder, Aleph Alpha", "Heidelberg, Germany",
     "Raised $500M Series B; built Alpha ONE (Europe's fastest commercial AI data center); ISO 27001 + EU AI Act alignment",
     "Leading sovereign, compliant AI for European governments & enterprise",
     "German AI leader"],
    ["5", "Demis Hassabis", "CEO, Google DeepMind", "London, UK",
     "Nobel Prize in Chemistry 2024; leads DeepMind's frontier research; AlphaFold impact continues",
     "World's most decorated AI researcher; bridging scientific discovery and AI",
     "Nobel Prize; TIME 100 AI"],
    ["6", "Clem Delangue", "CEO & Co-Founder, Hugging Face", "Paris, France",
     "Hugging Face as the GitHub of ML; millions of models hosted; critical open-source AI infrastructure",
     "Democratizing AI through open-source model sharing & collaboration",
     "Forbes 30 Under 30 alum"],
]
for row in eu_leaders:
    add_data_row(ws_eu, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_eu, r, "COMPANIES WITH MOST MOMENTUM", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_eu, r, co_h); r += 1
eu_companies = [
    ["1", "Mistral AI", "Paris, France", "~$7B+ (est.)",
     "$50M+ government AI contracts; 1.1B+ monthly queries; 6.2M developers (up from 1.5M)",
     "40% of Europe's Fortune 500 deploying/piloting; sovereign AI leader; avg enterprise deal $1.2M/yr",
     "Foundation Model / Sovereign AI"],
    ["2", "AMI Labs", "Paris, France", "$3.5B (pre-money)",
     "Just founded; targeting next-gen foundation models",
     "$1B seed — Europe's largest-ever; backed by NVIDIA, Bezos Expeditions, Temasek, Eric Schmidt",
     "Foundation Model / Open Science"],
    ["3", "Aleph Alpha", "Heidelberg, Germany", "~$2B+ (est.)",
     "$500M Series B; Luminous & PhariaAI models",
     "Alpha ONE data center (512 NVIDIA A100s); EU AI Act compliant; sovereign AI for governments",
     "Sovereign AI / Enterprise"],
    ["4", "Hugging Face", "Paris, France", "~$4.5B",
     "Millions of models hosted; critical open-source infrastructure",
     "GitHub of ML; growing enterprise tier; community of millions of developers",
     "AI Infrastructure / Open Source"],
    ["5", "Stability AI", "London, UK", "~$1B (restructured)",
     "Stable Diffusion models; restructuring under new leadership",
     "Pioneered open-source image generation; expanding into video & 3D",
     "Generative Media / Open Source"],
    ["6", "DeepL", "Cologne, Germany", "~$2B",
     "AI translation leader; expanding enterprise customer base",
     "Superior translation quality vs. Google Translate; strong enterprise moat",
     "AI Translation / Enterprise NLP"],
]
for row in eu_companies:
    add_data_row(ws_eu, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_eu, r, "BIGGEST BREAKTHROUGHS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_eu, r, bt_h); r += 1
eu_breakthroughs = [
    ["1", "AMI Labs $1B Seed Round", "AMI Labs (Yann LeCun)", "Mar 2026",
     "Europe's largest-ever seed round at $3.5B valuation for next-gen foundation models",
     "Signals growing European competitiveness in frontier AI; open-science approach",
     "Backed by Cathay Innovation, Bezos Expeditions, NVIDIA, Eric Schmidt, Temasek"],
    ["2", "Mistral Sovereign AI Adoption", "Mistral AI", "2025-2026",
     "40% of Europe's Fortune 500 deploying or piloting Mistral models; €50M+ government contracts",
     "Established European sovereign AI as a viable category; EU compliance leadership",
     "Open-source multilingual models; EU AI Act aligned; enterprise deals avg $1.2M/yr"],
    ["3", "Aleph Alpha PhariaAI Stack", "Aleph Alpha", "2025",
     "Enterprise-ready generative AI stack with transparency, compliance, and data sovereignty",
     "Proved sovereign AI can compete on quality; ISO 27001 + EU AI Act compliance",
     "Alpha ONE: Europe's fastest commercial AI data center (512 A100s)"],
    ["4", "Demis Hassabis Nobel Prize", "Google DeepMind (London)", "Oct 2024",
     "Nobel Prize in Chemistry for protein structure prediction (AlphaFold)",
     "AI's biggest scientific validation; demonstrates AI's potential for scientific discovery",
     "AlphaFold has been used by 2M+ researchers worldwide"],
]
for row in eu_breakthroughs:
    add_data_row(ws_eu, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_eu, r, "TOP INVESTORS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_eu, r, inv_h); r += 1
eu_investors = [
    ["1", "Bpifrance", "State Investment Bank", "Paris, France",
     "54 AI investments; €500K-€20M tickets; broad French AI ecosystem",
     "Government-backed AI development; sovereign AI support",
     "Most active European AI investor by deal count"],
    ["2", "Cathay Innovation", "VC", "Paris, France",
     "AMI Labs (co-led $1B seed); broad AI portfolio",
     "Co-led Europe's largest-ever seed round; deep tech focus",
     "€1B+ AI-focused fund"],
    ["3", "LocalGlobe", "VC", "London, UK",
     "43 AI investments; early-stage focus",
     "Seed to Series A; London-centric AI ecosystem development",
     "2nd most active European AI VC"],
    ["4", "Index Ventures", "VC", "London / SF",
     "Broad AI portfolio across Europe and US",
     "Cross-Atlantic AI investing; growth-stage focus",
     "Major European VC with global reach"],
    ["5", "Speedinvest", "VC", "Vienna, Austria",
     "40 AI investments; €200K-€3M tickets",
     "Pan-European early-stage AI; DACH region strength",
     "4th most active European AI VC"],
    ["6", "HTGF (High-Tech Gründerfonds)", "VC", "Bonn, Germany",
     "39 AI investments; €100K-€1M tickets",
     "Deep-tech seed investing; German AI ecosystem builder",
     "5th most active European AI VC"],
    ["7", "IQ Capital", "VC", "Cambridge / London, UK",
     "39 AI investments; £500K-£10M tickets",
     "Deep-tech AI and hard science; Oxbridge pipeline",
     "Strong research-to-startup pipeline"],
    ["8", "Balderton Capital", "VC", "London, UK",
     "Broad AI portfolio; growth-stage European AI",
     "Series A-B focused; European AI scaling",
     "Top-tier London VC"],
]
for row in eu_investors:
    add_data_row(ws_eu, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_eu, r, "HOTSPOTS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_eu, r, hs_h); r += 1
eu_hotspots = [
    ["1", "Paris", "France",
     "4.3% of global AI-native funding",
     "Foundation models (Mistral, AMI Labs); strong academia (INRIA, ENS); EU sovereign AI",
     "Mistral AI, AMI Labs, Hugging Face",
     "AMI Labs $1B seed; Mistral in 40% EU Fortune 500; EU sovereign AI contracts"],
    ["2", "London", "United Kingdom",
     "Major and growing AI hub",
     "World-class research (DeepMind); regulatory leadership; top universities",
     "Google DeepMind, Stability AI, Faculty AI",
     "OpenAI named London its largest research hub outside SF (Feb 2026)"],
    ["3", "Zurich", "Switzerland",
     "Specialized AI hub",
     "Smart city apps; fintech AI; ETH Zurich research; Google research presence",
     "ETH Zurich, Google Brain Zurich, fintech AI startups",
     "Leading in financial services AI; strong European research output"],
    ["4", "Berlin", "Germany",
     "Growing AI startup hub",
     "Strong startup ecosystem; enterprise AI; government support",
     "AI startups, enterprise SaaS with AI, research institutes",
     "Growing pool of AI talent; lower cost than London/Paris"],
    ["5", "Heidelberg", "Germany",
     "Sovereign AI hub",
     "Aleph Alpha HQ; research university strength",
     "Aleph Alpha, Heidelberg University, EMBL",
     "Alpha ONE data center; sovereign AI for German/EU government"],
    ["6", "Amsterdam-Delta", "Netherlands",
     "Emerging AI hub",
     "AI in logistics, fintech, agriculture; strong tech talent",
     "Various AI startups; Booking.com AI; Philips AI",
     "Growing AI investment; strategic location in Europe"],
]
for row in eu_hotspots:
    add_data_row(ws_eu, r, row, NUM_COLS_REGIONAL); r += 1

auto_width(ws_eu, NUM_COLS_REGIONAL)


# ── TAB 9: Asia ───────────────────────────────────────────────────────────────
ws_as = wb.create_sheet("Asia")
r = 1

add_section_header(ws_as, r, "LEADERS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_as, r, ldr_h); r += 1
as_leaders = [
    ["1", "Liang Wenfeng", "CEO, DeepSeek", "Hangzhou, China",
     "DeepSeek-V3.2 & V3.2-Speciale matched/exceeded GPT-5; R1 demonstrated RL-only reasoning; all open-sourced MIT",
     "Frontier AI at 1/30th US cost; gold-medal results in IMO, IOI, ICPC; reshaped global AI cost economics",
     "N/A"],
    ["2", "Robin Li (Li Yanhong)", "CEO, Baidu", "Beijing, China",
     "Leading Baidu's ERNIE Bot and enterprise AI platform in China",
     "China's leading search company pivoting to AI-first; massive Chinese language data advantage",
     "China AI leader"],
    ["3", "Masayoshi Son", "CEO, SoftBank Group", "Tokyo, Japan",
     "Led $40B OpenAI investment (Mar 2025); $30B follow-on (Feb 2026); cumulative $64.6B in OpenAI (~13% ownership)",
     "All-in bet on AI infrastructure; largest individual AI investor globally; Stargate data center initiative",
     "TIME 100 AI"],
    ["4", "Ravi Kumar S", "CEO, Cognizant", "Teaneck, NJ (HQ) / India ops",
     "Bet $1B on Gen AI; Agent Foundry for enterprise AI; Synapse to train 2M people in Gen AI",
     "Driving large-scale enterprise AI transformation with major India-based workforce",
     "TIME 100 AI 2025"],
    ["5", "Zhang Peng", "CEO, Zhipu AI", "Beijing, China",
     "GLM-5 reached top spot among open-source models on Artificial Analysis (Feb 2026); raised prices 30% on surging demand",
     "Leading China's open-source AI model race; shares rose 34%",
     "China AI leader"],
    ["6", "Yang Zhilin", "CEO & Founder, Moonshot AI", "Beijing, China",
     "Kimi K2 model launch (Jul 2025); briefly held top benchmarks before being surpassed",
     "China's fast-moving AI startup leader; $2.5B valuation; aggressive model iteration",
     "Forbes China 30 Under 30"],
    ["7", "Pratyush Kumar", "CEO & Co-Founder, Sarvam AI", "Bengaluru, India",
     "Sarvam Vision OCR at 84.3% accuracy (outperforming Gemini 3 Pro & ChatGPT); Bulbul V3 TTS in 11 Indian languages",
     "Building sovereign AI for Indian languages; democratizing AI for 1.4B people",
     "India AI pioneer"],
    ["8", "Ankush Sabharwal", "CEO & Founder, CoRover.ai", "India",
     "BharatGPT — India's first LLM-based Gen AI platform; 1B+ users served",
     "India's homegrown Gen AI leader; conversational AI at scale",
     "Top 10 global tech entrepreneur"],
]
for row in as_leaders:
    add_data_row(ws_as, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_as, r, "COMPANIES WITH MOST MOMENTUM", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_as, r, co_h); r += 1
as_companies = [
    ["1", "DeepSeek", "Hangzhou, China", "Private (self-funded)",
     "Rapidly growing user base and API adoption globally",
     "V3.2 matched/exceeded GPT-5 at dramatically lower cost; open-source MIT; gold-medal competition results",
     "Foundation Model / Open Source"],
    ["2", "Zhipu AI", "Beijing, China", "~$5B+ (est.)",
     "GLM-5 topped open-source benchmarks (Feb 2026); raised prices 30% on demand surge",
     "Top open-source model globally; shares up 34%; intense competitive pace",
     "Foundation Model / Open Source"],
    ["3", "ByteDance (Doubao/Seedance)", "Beijing, China", "Part of $220B+ parent",
     "Seedance 2.0 video model competing with OpenAI Sora; massive user base",
     "Video generation breakthrough; leveraging TikTok/Douyin distribution",
     "Generative Media / Video"],
    ["4", "Moonshot AI (Kimi)", "Beijing, China", "$2.5B+",
     "Kimi K2 model launch; aggressive model iteration cycle",
     "Fast-moving Chinese AI startup; briefly held top benchmarks",
     "Foundation Model / Consumer AI"],
    ["5", "Sarvam AI", "Bengaluru, India", "~$200M+ (est.)",
     "Industry-leading OCR (84.3%); TTS in 11 Indian languages",
     "Sovereign AI for India; outperforming Gemini & ChatGPT on Indian language tasks",
     "Sovereign AI / Indian Languages"],
    ["6", "CoRover.ai (BharatGPT)", "India", "~$100M+ (est.)",
     "1B+ users served; India's first LLM-based Gen AI platform",
     "Conversational AI at massive scale for Indian market",
     "Conversational AI / Indian Market"],
    ["7", "Alibaba (Qwen)", "Hangzhou, China", "Part of $200B+ parent",
     "Qwen models; Qwen-3.5 debut expected soon",
     "Leveraging Alibaba Cloud for AI distribution; open-source models",
     "Foundation Model / Cloud AI"],
    ["8", "Saltlux (LUXIA)", "Seoul, South Korea", "~$500M (est.)",
     "LUXIA 2.5 model; virtual human video generation (Plunit Studio)",
     "Korean AI leader competing with Chinese models; NLP specialist",
     "Foundation Model / NLP"],
]
for row in as_companies:
    add_data_row(ws_as, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_as, r, "BIGGEST BREAKTHROUGHS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_as, r, bt_h); r += 1
as_breakthroughs = [
    ["1", "DeepSeek-R1: RL-Only Reasoning", "DeepSeek (China)", "Jan 2025",
     "First demonstration of reasoning from pure reinforcement learning without supervised fine-tuning",
     "Paradigm shift; frontier AI at 1/30th US cost; open-sourced under MIT",
     "R1-Zero via pure RL; R1 matches OpenAI o1; distilled 32B outperforms o1-mini"],
    ["2", "DeepSeek-V3.2-Speciale", "DeepSeek (China)", "Dec 2025",
     "Open-source model matching/surpassing GPT-5; 'Thinking in Tool-Use' innovation",
     "Gold-medal results IMO, IOI, ICPC, CMO; frontier model under MIT License",
     "685B params, ~37B active (MoE); Sparse Attention O(L^2)→O(L); 85K+ training tasks"],
    ["3", "Zhipu GLM-5 Open-Source Top Rank", "Zhipu AI (China)", "Feb 2026",
     "GLM-5 claimed #1 open-source model on Artificial Analysis benchmarking",
     "Intensified China's AI model race; shares surged 34%",
     "Advanced model architecture; pricing raised 30% on demand"],
    ["4", "ByteDance Seedance 2.0", "ByteDance (China)", "Feb 2026",
     "Video generation model competing with OpenAI Sora; praised during testing",
     "China competitive in generative video; leverages massive distribution",
     "Advanced video generation; competing with Sora quality"],
    ["5", "Sarvam Vision OCR", "Sarvam AI (India)", "2025",
     "84.3% accuracy on OCR benchmarks — outperforming Gemini 3 Pro and ChatGPT",
     "Proved Indian-built AI can beat global leaders on specific tasks",
     "Specialized for Indian scripts; Bulbul V3 TTS in 11 Indian languages"],
]
for row in as_breakthroughs:
    add_data_row(ws_as, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_as, r, "TOP INVESTORS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_as, r, inv_h); r += 1
as_investors = [
    ["1", "SoftBank Group / Vision Fund", "VC / Growth", "Tokyo, Japan",
     "OpenAI ($64.6B cumulative ~13% ownership); xAI; Stargate initiative",
     "All-in bet on AI infrastructure; largest individual AI investor globally",
     "Led two largest private funding rounds in history"],
    ["2", "Temasek", "Sovereign Wealth Fund", "Singapore",
     "AMI Labs; Upscale AI ($200M round); bullish on China AI (B2B focus)",
     "AI infrastructure and B2B applications; True Light Capital for China innovation",
     "Invested across US, EU, and Asian AI ecosystems"],
    ["3", "Premji Invest", "Family Office", "Bengaluru, India",
     "Upscale AI ($200M Series A); Indian AI ecosystem investments",
     "Indian AI infrastructure and enterprise applications",
     "Co-investor in major Indian AI rounds"],
    ["4", "Alibaba Group", "Strategic / Corporate", "Hangzhou, China",
     "Qwen models; Alibaba Cloud AI; internal AI investments",
     "Cloud AI infrastructure; open-source model ecosystem",
     "Major corporate AI investor in China"],
    ["5", "Tencent", "Strategic / Corporate", "Shenzhen, China",
     "AI Lab investments; gaming AI; WeChat AI integrations",
     "Consumer AI via social/gaming platforms; enterprise AI emerging",
     "One of China's largest AI corporate investors"],
    ["6", "Shorooq Partners", "VC", "UAE / Middle East",
     "AMI Labs ($1B seed participant); MENA AI ecosystem",
     "Bridging Middle East capital with global AI ventures",
     "Expanding into frontier AI investments globally"],
    ["7", "Xora Innovation (Temasek-backed)", "VC", "Singapore",
     "Upscale AI ($200M Series A); Southeast Asian AI ecosystem",
     "AI infrastructure and B2B applications in Asia",
     "Active participant in large Asian AI rounds"],
]
for row in as_investors:
    add_data_row(ws_as, r, row, NUM_COLS_REGIONAL); r += 1
r += 1

add_section_header(ws_as, r, "HOTSPOTS", NUM_COLS_REGIONAL); r += 1
add_sub_header(ws_as, r, hs_h); r += 1
as_hotspots = [
    ["1", "Beijing", "China",
     "66.2% of local startup funding goes to AI — highest globally",
     "Foundation models at 1/30th US cost; massive government support; large talent pool",
     "DeepSeek, Baidu, ByteDance, Zhipu AI, Moonshot AI",
     "DeepSeek breakthroughs reshaped global AI cost economics"],
    ["2", "Shanghai / Greater Bay Area", "China",
     "Emerging AI powerhouse",
     "Hardware manufacturing ecosystem; application-layer AI; massive consumer market",
     "Alibaba AI, Tencent AI Lab, semiconductor ecosystem",
     "Growing AI applications; connection to hardware supply chain (Shenzhen)"],
    ["3", "Singapore", "Singapore",
     "Growing AI investment hub",
     "World's best AI readiness; government-wide implementation; governance leadership",
     "AI Singapore, Temasek AI investments, SEA hub",
     "Leading AI governance frameworks; attracting global AI companies"],
    ["4", "Bengaluru", "India",
     "India's #1 AI hub",
     "Massive talent pool; growing startup ecosystem; enterprise AI services",
     "Sarvam AI, CoRover.ai, Infosys AI, Wipro AI",
     "Indian sovereign AI movement; language-specific AI innovation"],
    ["5", "Tokyo", "Japan",
     "Rising AI investment center",
     "Robotics AI; industrial automation; SoftBank HQ; strong R&D culture",
     "SoftBank, NEC, Preferred Networks, Sakana AI",
     "SoftBank's massive AI investment strategy; growing AI startup scene"],
    ["6", "Seoul", "South Korea",
     "Specialized AI hub",
     "NLP and conversational AI; Samsung/LG AI divisions; gaming AI",
     "Saltlux, Samsung AI, Naver, Kakao",
     "LUXIA 2.5 model; growing Korean AI ecosystem"],
    ["7", "Hangzhou", "China",
     "DeepSeek's home city",
     "Alibaba ecosystem; DeepSeek research lab; e-commerce AI",
     "DeepSeek, Alibaba, e-commerce AI startups",
     "DeepSeek's breakthroughs put Hangzhou on the global AI map"],
]
for row in as_hotspots:
    add_data_row(ws_as, r, row, NUM_COLS_REGIONAL); r += 1

auto_width(ws_as, NUM_COLS_REGIONAL)


# ── TAB 10: Agentic AI & OpenClaw Ecosystem ───────────────────────────────────
ws_ag = wb.create_sheet("Agentic AI & OpenClaw")
NUM_COLS_AG = 8
r = 1

add_section_header(ws_ag, r, "OPENCLAW & OPEN-SOURCE AI AGENT RUNTIMES", NUM_COLS_AG); r += 1
oc_h = ["#", "Project / Company", "Type", "HQ / Origin", "GitHub Stars / Adoption", "Key Capabilities", "Funding / Business Model", "Status (Mar 2026)"]
add_sub_header(ws_ag, r, oc_h); r += 1
openclaw_data = [
    ["1", "OpenClaw", "Open-Source AI Agent Runtime", "Austria (founder); now OpenAI-affiliated",
     "234K+ GitHub stars; 1.5M+ agents running; #1 GitHub repo (early 2026)",
     "Autonomous agent execution via LLMs (GPT, Claude); shell commands, file management, email, messaging (WhatsApp/Slack/Telegram/Discord); web browsing; persistent memory; heartbeat daemon for continuous operation; 7,400+ MCP skills",
     "Free & open-source (MIT); LLM API costs ~$10-50/mo; founder Peter Steinberger joined OpenAI Feb 2026; project transitioned to independent foundation",
     "Fastest-growing open-source AI project; security vulnerability CVE-2026-25253 patched; foundation governance established"],
    ["2", "LangChain / LangGraph", "Open-Source Agent Framework", "San Francisco, CA (USA)",
     "90K+ GitHub stars (LangChain); wide enterprise adoption",
     "Agent orchestration with explicit state management; human-in-the-loop controls; 1,000+ integrations; durable runtime; streaming & memory persistence; best for complex pipelines with branching logic",
     "VC-backed (Sequoia-led $25M Series A); LangSmith commercial SaaS for observability/monitoring",
     "Preferred framework for production-grade agent complexity; Daytona ($24M) partnership"],
    ["3", "CrewAI", "Open-Source Multi-Agent Framework", "USA",
     "100K+ certified developers; fast-growing community",
     "Autonomous collaborative agents (Crews) + event-driven production workflows (Flows); lightweight Python framework; simple high-level abstractions for multi-agent teams",
     "VC-backed; CrewAI Enterprise for commercial use",
     "Standard for enterprise AI team automation; simpler alternative to LangGraph"],
    ["4", "AutoGen", "Open-Source Agent Framework", "Microsoft Research",
     "Strong adoption in research & enterprise",
     "Chat-based multi-agent conversations; best for two-agent dialogue patterns; Microsoft ecosystem integration",
     "Microsoft-backed open-source project",
     "Solid for chat-based agent patterns; less ideal for complex branching pipelines"],
    ["5", "n8n", "Open-Source Workflow Automation", "Berlin, Germany",
     "60K+ GitHub stars; 1,700+ community templates",
     "Visual node-based workflow builder; 400+ pre-built integrations; deterministic trigger-based pipelines; complements OpenClaw for structured workflows",
     "Free self-host; cloud from €24/mo; VC-backed",
     "Complementary to OpenClaw; used together for hybrid AI+deterministic automation"],
]
for row in openclaw_data:
    add_data_row(ws_ag, r, row, NUM_COLS_AG); r += 1
r += 1

add_section_header(ws_ag, r, "AGENTIC AI STARTUPS (VENTURE-BACKED)", NUM_COLS_AG); r += 1
ag_h = ["#", "Company", "Focus Area", "HQ", "Valuation / Funding", "Key Product / Capability", "Notable Customers / Metrics", "Momentum"]
add_sub_header(ws_ag, r, ag_h); r += 1
agentic_startups = [
    ["1", "Cognition (Devin)", "Autonomous Coding Agent", "San Francisco, CA",
     "$10.2B valuation; $400M raised (Sep 2025, Founders Fund-led)",
     "Devin: autonomous AI software engineer that plans, codes, debugs, and deploys; acquired Windsurf IDE (Jul 2025)",
     "Goldman Sachs, Citi, Dell, Cisco, Ramp, Palantir, Nubank; $73M ARR (Jun 2025, up from $1M in Sep 2024)",
     "140x ARR multiple; net burn under $20M since founding; combined ARR grew 30%+ in 7 weeks post-Windsurf acquisition"],
    ["2", "Temporal", "Durable Execution for AI Agents", "San Francisco, CA",
     "$5B valuation; $300M Series D (Feb 2026, a16z-led)",
     "Durable execution infrastructure for long-running AI agents; ensures reliability and fault tolerance at production scale",
     "20M+ monthly installs; 380% YoY revenue growth; 350% increase in weekly active usage",
     "Critical infrastructure layer for production agentic AI; addressing reliability bottleneck"],
    ["3", "Gumloop", "No-Code AI Agent Builder", "San Francisco, CA",
     "$50M Series B (Benchmark-led); Y Combinator, First Round, Shopify Ventures",
     "Platform enabling non-technical employees to build AI agents and automations",
     "Shopify, Ramp, Instacart",
     "Democratizing agent creation; enterprise adoption growing"],
    ["4", "Adept", "Enterprise Agentic AI", "San Francisco, CA",
     "~$1B+ valuation",
     "Enterprise agents that execute complex workflows across websites and software; proprietary models trained on web UI data",
     "Enterprise customers across multiple verticals",
     "Pioneered 'action model' approach; focused on workforce automation"],
    ["5", "Simular", "OS-Level Computer Use Agent", "USA",
     "$21.5M Series A (Felicis-led)",
     "Neuro-symbolic agents that control Mac/PC at OS level (mouse, keyboard, apps); not just browser-based",
     "Mac 1.0 released; Windows version via Microsoft Windows 365 for Agents program",
     "Going beyond browser to full OS control; Microsoft partnership"],
    ["6", "Relevance AI", "AI Agent Operating System", "Sydney, Australia",
     "$24M Series B (May 2025)",
     "Visual multi-agent system builder (Workforce); text-to-agent generator; enables non-technical users to build agents",
     "40,000 AI agents created on platform (Jan 2025)",
     "Bridging technical and non-technical agent creation"],
    ["7", "Daytona", "Agent Compute Infrastructure", "USA / Europe",
     "$24M Series A (Feb 2026)",
     "Programmatic, composable computers for AI agents that launch in milliseconds",
     "LangChain, Turing, Writer, SambaNova; $1M FRR in <3 months, doubled in 6 weeks",
     "Solving the 'where do agents run?' problem; rapid revenue growth"],
    ["8", "Sapiom", "Agent Financial Infrastructure", "USA",
     "$15M Seed (Accel-led)",
     "Financial infrastructure enabling agents to pay, authenticate, and audit their own API/cloud usage autonomously",
     "Early-stage; solving agent payment and identity",
     "Novel category: financial autonomy for AI agents"],
    ["9", "Lindy AI", "Personal AI Assistant / Phone Agent", "San Francisco, CA",
     "VC-backed (undisclosed)",
     "AI assistant for scheduling, email, meeting summaries; Gaia: fastest AI phone agent (<500ms response)",
     "Growing consumer and business user base",
     "Competing in personal AI assistant space; phone agent innovation"],
    ["10", "Genspark Claw", "Secure Cloud AI Agent", "USA",
     "Part of Genspark (VC-backed)",
     "Cloud-based alternative to OpenClaw addressing enterprise security and setup complexity",
     "Targeting enterprises wary of self-hosted agent security risks",
     "Launched Mar 2026 as direct response to OpenClaw security concerns"],
]
for row in agentic_startups:
    add_data_row(ws_ag, r, row, NUM_COLS_AG); r += 1
r += 1

add_section_header(ws_ag, r, "BIG TECH AGENTIC AI PRODUCTS", NUM_COLS_AG); r += 1
bt_h2 = ["#", "Company", "Product / Initiative", "Launch Date", "Key Capability", "Target Market", "How It Compares to OpenClaw", "Status"]
add_sub_header(ws_ag, r, bt_h2); r += 1
bigtech_agents = [
    ["1", "OpenAI", "GPT-5.4 Computer Use / Operator", "Mar 2026",
     "Native computer-use capabilities; autonomous agent workflows; keyboard/mouse commands across applications",
     "Enterprise & consumer",
     "Proprietary; cloud-based; integrated into ChatGPT; less customizable than OpenClaw but more polished",
     "Just launched; acquired OpenClaw founder Peter Steinberger"],
    ["2", "Anthropic", "Claude Computer Use / Claude Code", "2025-2026",
     "Claude can interact with computer interfaces; Claude Code autonomous coding ($2.5B+ run-rate); 4% of GitHub public commits",
     "Enterprise developers",
     "API-based; enterprise-focused; Claude Code is biggest revenue-generating agent product",
     "Claude Code is #1 revenue-generating AI agent product globally"],
    ["3", "Google", "Gemini Agents / Project Mariner", "2025-2026",
     "Gemini 3.1 Pro agentic workflows; browser/computer use capabilities",
     "Enterprise & consumer",
     "Cloud-based; integrated into Google ecosystem; Workspace, Search, Chrome",
     "Expanding agentic capabilities across Google product suite"],
    ["4", "Microsoft", "Copilot Agents / AutoGen", "2025-2026",
     "Microsoft 365 Copilot agents; AutoGen framework; Windows 365 for Agents",
     "Enterprise (M365 ecosystem)",
     "Deep enterprise integration; combines AutoGen open-source with proprietary Copilot",
     "Largest enterprise agent deployment via M365 Copilot"],
]
for row in bigtech_agents:
    add_data_row(ws_ag, r, row, NUM_COLS_AG); r += 1
r += 1

add_section_header(ws_ag, r, "AGENTIC AI MARKET LANDSCAPE SUMMARY", NUM_COLS_AG); r += 1
sum_h = ["Category", "Leader", "Runner-Up", "Key Metric", "Trend", "", "", ""]
add_sub_header(ws_ag, r, sum_h); r += 1
agent_summary = [
    ["Open-Source Agent Runtime", "OpenClaw",
     "LangChain / LangGraph",
     "234K GitHub stars; 1.5M+ agents running",
     "Open-source agentic AI is the fastest-growing category; security & governance are top concerns",
     "", "", ""],
    ["Venture-Backed Agent Startup", "Cognition (Devin)",
     "Temporal",
     "Cognition: $10.2B at $73M ARR; Temporal: $5B at 380% rev growth",
     "Autonomous coding and durable execution are the two hottest agentic sub-categories",
     "", "", ""],
    ["Enterprise Agent Platform", "Microsoft Copilot Agents",
     "Anthropic Claude Code",
     "Copilot: largest enterprise deployment; Claude Code: $2.5B+ run-rate",
     "Enterprise agentic AI is moving from pilot to production across Fortune 500",
     "", "", ""],
    ["Agent Infrastructure", "Temporal (durable execution)",
     "Daytona (compute)",
     "Temporal: $300M raise; Daytona: $1M FRR in 3 months",
     "Infrastructure for where and how agents run is an emerging critical layer",
     "", "", ""],
    ["Fastest-Growing Category", "AI Coding Agents",
     "AI Personal Assistants",
     "Devin $1M→$73M ARR in 9 months; Cursor $2B ARR in 18 months",
     "Developer tools are the highest-velocity segment of the agentic AI market",
     "", "", ""],
]
for row in agent_summary:
    add_data_row(ws_ag, r, row, NUM_COLS_AG); r += 1

auto_width(ws_ag, NUM_COLS_AG)


# Freeze panes for all sheets
for ws in wb.worksheets:
    ws.freeze_panes = "A2"

output_path = "/workspace/generative_ai_market_insights.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
