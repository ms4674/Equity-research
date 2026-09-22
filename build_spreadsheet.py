"""Builds ai_assistant_comparison.xlsx: MAUs, usage allowances, and
subscription pricing for Meta Muse, Grok, Claude (Cowork), and ChatGPT.
Data researched September 22, 2026."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
PRODUCT_FILLS = {
    "Meta Muse": PatternFill("solid", fgColor="DDEBF7"),
    "Grok": PatternFill("solid", fgColor="E2EFDA"),
    "Claude (Cowork)": PatternFill("solid", fgColor="FCE4D6"),
    "ChatGPT": PatternFill("solid", fgColor="FFF2CC"),
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def style_sheet(ws, widths, n_rows):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=1, max_row=n_rows, max_col=len(widths)):
        for cell in row:
            cell.border = BORDER
            if cell.row > 1:
                cell.alignment = WRAP
                fill = PRODUCT_FILLS.get(ws.cell(row=cell.row, column=1).value)
                if fill:
                    cell.fill = fill
    ws.freeze_panes = "A2"


wb = Workbook()

# ---------------------------------------------------------------- Overview
ws = wb.active
ws.title = "Overview"
overview_headers = [
    "Product", "Company", "Launched",
    "Monthly Active Users (MAU)", "MAU Source / Date",
    "Usage Allowance (summary)", "Subscription Pricing (summary)",
]
overview_rows = [
    [
        "Meta Muse", "Meta", "Sep 8, 2026 (limited testing)",
        "Not disclosed. Early proxies: ~642K US daily active users; "
        "~2.6-2.8M installs in first 12 days",
        "Apptopia / Sensor Tower estimates via TechCrunch & press, Sep 2026 "
        "(Meta has released no official figures)",
        "Metered in 'Muse tokens' per week: Free ~100M tokens/wk (per launch "
        "announcement), Power 500M/wk, Maximum 3B/wk. Wait for weekly refresh "
        "or upgrade when exhausted.",
        "Free $0; Power $20/mo; Maximum $100/mo. Paid tiers buy more usage, "
        "not extra features.",
    ],
    [
        "Grok", "xAI (SpaceX)", "Nov 2023",
        "~117 million (Mar 31, 2026); ~1.9M paid SuperGrok subscribers",
        "SpaceX IPO prospectus (SEC Form 424(b)(4), Jun 2026) - first "
        "audited Grok user number",
        "Single weekly usage pool per paid tier, spendable across Chat, "
        "Imagine, Voice and Build (since Jun 2026). Free tier has its own "
        "allowance. Extra usage credits purchasable from $5.",
        "Free $0; X Premium $8/mo; SuperGrok Lite $10/mo; SuperGrok $30/mo "
        "($300/yr = ~$25/mo); X Premium+ $40/mo; SuperGrok Plus $100/mo; "
        "SuperGrok Heavy $300/mo.",
    ],
    [
        "Claude (Cowork)", "Anthropic", "Claude: Mar 2023; Cowork bundled "
        "with paid plans",
        "Not disclosed. Third-party estimates: ~56M app MAU (Sensor Tower, "
        "Q2 2026); ~103.9M app MAU May 2026 (Similarweb); ~245M modelled "
        "web+app (Sensor Tower)",
        "Sensor Tower via Reuters (Jun 2026); Similarweb (Sep 2026 "
        "comparison). Anthropic publishes no official MAU.",
        "Shared usage pool across Claude chat, Desktop, Code and Cowork. "
        "Session limit resets every 5 hours plus weekly caps; no fixed "
        "message count (varies by model/complexity). Pro = 5x+ Free; Max = "
        "5x or 20x Pro. Overage via usage credits at standard API rates.",
        "Free $0 (no Cowork); Pro $20/mo (~$17/mo annual); Max 5x $100/mo; "
        "Max 20x $200/mo; Team/Enterprise per-seat. Cowork has no standalone "
        "price - included in all paid plans.",
    ],
    [
        "ChatGPT", "OpenAI", "Nov 2022",
        "~1 billion app MAU (crossed May 2026, fastest app ever to reach it); "
        "~1.1B modelled incl. web",
        "Sensor Tower estimates via Reuters, Jun 2, 2026. OpenAI publishes "
        "weekly-active figures, not MAU.",
        "Unlimited everyday text chats on all tiers (abuse guardrails). "
        "Caps on specific models/features: e.g. Pro $200 = 200 GPT-6 Pro "
        "msgs/wk; Pro $100 = 50/wk; Plus voice 3 hr/24h; Free 3 file "
        "uploads/day. Pro = 5x or 20x Plus usage.",
        "Free $0; Go $8/mo; Plus $20/mo; Pro $100/mo (5x Plus) or $200/mo "
        "(20x Plus); Business $25/user/mo annual ($30 monthly); Enterprise "
        "custom.",
    ],
]
ws.append(overview_headers)
for row in overview_rows:
    ws.append(row)
for r in range(2, len(overview_rows) + 2):
    ws.row_dimensions[r].height = 120
style_sheet(ws, [16, 14, 18, 34, 32, 48, 48], len(overview_rows) + 1)

# ----------------------------------------------------------- Pricing tiers
ws2 = wb.create_sheet("Pricing Tiers")
tier_headers = ["Product", "Plan", "Price (USD)", "Usage Allowance", "Notes"]
tier_rows = [
    ["Meta Muse", "Free", "$0",
     "Weekly usage limit (~100M Muse tokens/wk per launch announcement; "
     "Meta Help Center publishes no free-tier figure)",
     "Includes hosted agent and its own browser; wait for weekly refresh "
     "or upgrade"],
    ["Meta Muse", "Power", "$20/mo", "500M Muse tokens per week",
     "Renews monthly on purchase date; in limited testing, not all regions"],
    ["Meta Muse", "Maximum", "$100/mo", "3B Muse tokens per week",
     "5x the price of Power buys 6x the weekly allowance"],
    ["Grok", "Free", "$0", "Free Chat and Voice allowance, own reset cycle",
     "Grok 4 Mini, text only"],
    ["Grok", "X Premium", "$8/mo", "Higher than free",
     "Grok 4 inside X, plus X perks"],
    ["Grok", "SuperGrok Lite", "$10/mo", "Entry weekly usage pool",
     "Grok 4, basic Imagine (480p/6s video), 1 AI agent"],
    ["Grok", "SuperGrok", "$30/mo ($300/yr)", "Standard weekly usage pool",
     "Grok 4 + DeepSearch, Voice, Imagine; Grok 4.5 rolling out; most "
     "popular paid tier"],
    ["Grok", "X Premium+", "$40/mo", "Higher X limits",
     "Grok 4 + X Premium+ perks"],
    ["Grok", "SuperGrok Plus", "$100/mo",
     "Much higher usage across Chat, Imagine, Voice and Build",
     "Everything in SuperGrok plus priority access at peak times"],
    ["Grok", "SuperGrok Heavy", "$300/mo", "Largest weekly usage pool",
     "Full Grok 4.5, Grok 4 Heavy, multi-agent; aimed at researchers"],
    ["Claude (Cowork)", "Free", "$0",
     "Basic usage; 5-hour rolling session limits",
     "Does NOT include Cowork"],
    ["Claude (Cowork)", "Pro", "$20/mo (~$17/mo annual)",
     "At least 5x Free usage per 5-hour session, plus weekly limits",
     "Includes Cowork, Claude Code, Design, Science; entry point for Cowork"],
    ["Claude (Cowork)", "Max 5x", "$100/mo",
     "5x Pro usage per session, plus weekly limits",
     "Priority access to newest features and models"],
    ["Claude (Cowork)", "Max 20x", "$200/mo",
     "20x Pro usage per session, plus weekly limits",
     "Highest individual tier; overage via usage credits at API rates "
     "($2,000/day redemption cap)"],
    ["Claude (Cowork)", "Team / Enterprise", "Per-seat (varies)",
     "Standard seats > Pro; Premium seats 5x Standard",
     "Admin controls, SSO, centralized billing; Cowork availability by "
     "seat type"],
    ["ChatGPT", "Free", "$0",
     "Unlimited text chats (abuse guardrails); 3 file uploads/day",
     "GPT-5.6 Luna; ad-supported in US"],
    ["ChatGPT", "Go", "$8/mo", "Higher limits than Free",
     "Went global Jan 2026; ad-supported in US"],
    ["ChatGPT", "Plus", "$20/mo",
     "Unlimited text chats; voice 3 hr per rolling 24h; thinking-model "
     "allowances unpublished; 20 GB uploads",
     "Most popular tier; GPT-5.6 Sol access"],
    ["ChatGPT", "Pro ($100)", "$100/mo",
     "5x Plus usage; 50 GPT-6 Pro / Sol Pro msgs per week (shared); voice "
     "15 hr; 100 GB uploads",
     "Launched Apr 2026; 1M token context"],
    ["ChatGPT", "Pro ($200)", "$200/mo",
     "20x Plus usage; 200 GPT-6 Pro msgs/wk; Sol Pro 170/day; unlimited "
     "voice; 100 GB uploads",
     "Highest usage tier; unlimited Deep Research and Sora; new sign-ups "
     "paused as of Sep 2026"],
    ["ChatGPT", "Business", "$25/user/mo (annual) or $30 (monthly)",
     "Standard seats: 15 GPT-6 Pro msgs/mo; Premium seats: 50/wk",
     "Workspace credits purchasable for more access"],
    ["ChatGPT", "Enterprise", "Custom",
     "Custom allowances; credits purchasable",
     "Annual billing; custom pricing"],
]
ws2.append(tier_headers)
for row in tier_rows:
    ws2.append(row)
style_sheet(ws2, [16, 22, 26, 52, 52], len(tier_rows) + 1)

# ---------------------------------------------------------------- Sources
ws3 = wb.create_sheet("Sources & Notes")
src_headers = ["Product", "Data Point", "Source", "URL / Reference"]
src_rows = [
    ["Meta Muse", "Launch, early downloads & DAU",
     "Apptopia via TechCrunch; Sensor Tower via Indian Express (Sep 2026)",
     "techcrunch.com; indianexpress.com/article/technology/"
     "artificial-intelligence/meta-muse-ai-agent-what-makes-it-different-"
     "amazon-dispute-10889290/"],
    ["Meta Muse", "Subscription plans & token allowances",
     "Meta Help Center - About Muse subscriptions",
     "meta.com/help/subscriptions/1021145227643680/"],
    ["Meta Muse", "Free tier ~100M tokens/wk",
     "Mark Zuckerberg launch thread (Sep 8, 2026), via eesel.ai analysis",
     "eesel.ai/blog/meta-muse-agent-pricing"],
    ["Grok", "117M MAU (Mar 31, 2026); 1.9M paid subscribers; $365M 2025 "
     "subscription revenue",
     "SpaceX IPO prospectus, SEC Form 424(b)(4) (Jun 11, 2026)",
     "axis-intelligence.com/grok-ai-statistics/; pulse2.com"],
    ["Grok", "Plan pricing & weekly usage pools",
     "xAI pricing page; FelloAI Grok pricing guide (2026)",
     "x.ai/pricing; grok.com/plans; felloai.com/grok-pricing/"],
    ["Claude (Cowork)", "App MAU estimates (56M Q2 2026; 103.9M May 2026; "
     "245M modelled)",
     "Sensor Tower via Reuters (Jun 2, 2026); Similarweb (Sep 2026); "
     "Anthropic publishes no official MAU",
     "reuters.com/technology/chatgpt-app-hits-1-billion-monthly-active-"
     "users-record-time-data-shows-2026-06-02/; quashbugs.com/blog/"
     "claude-ai-statistics; memeburn.com/claude-usage-statistics-2026/"],
    ["Claude (Cowork)", "Plans, pricing, Cowork bundling, usage limits",
     "Anthropic pricing page & Help Center (Sep 2026)",
     "claude.com/pricing; support.claude.com/en/articles/11049741; "
     "support.claude.com/en/articles/11647753"],
    ["ChatGPT", "1B app MAU (May 2026)",
     "Sensor Tower via Reuters (Jun 2, 2026)",
     "reuters.com/technology/chatgpt-app-hits-1-billion-monthly-active-"
     "users-record-time-data-shows-2026-06-02/"],
    ["ChatGPT", "Plan pricing & usage limits",
     "OpenAI pricing page & Help Center; FelloAI and AI-Toolbox guides "
     "(Jul-Sep 2026)",
     "chatgpt.com/pricing; help.openai.com/en/articles/9793128; "
     "felloai.com; ai-toolbox.co"],
    ["All", "General note",
     "MAU figures for Meta Muse and Claude are third-party estimates; the "
     "companies have not published official numbers. Grok's figure is from "
     "an audited SEC filing. ChatGPT's is a Sensor Tower app estimate. "
     "Data compiled Sep 22, 2026; pricing subject to change.",
     "-"],
]
ws3.append(src_headers)
for row in src_rows:
    ws3.append(row)
style_sheet(ws3, [16, 42, 46, 60], len(src_rows) + 1)

wb.save("/workspace/ai_assistant_comparison.xlsx")
print("Saved /workspace/ai_assistant_comparison.xlsx")
