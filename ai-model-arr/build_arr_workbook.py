"""
Builds AI_Model_ARR_Gross_vs_Net.xlsx: gross vs. net ARR for frontier (closed-weight)
and open-weight model labs.

Gross ARR  = annualized end-customer billings for the lab's models, including sales that
             clear through cloud marketplaces / distribution partners at full value
             (Anthropic's reporting convention).
Net ARR    = annualized revenue retained by the lab after distribution-partner revenue
             share (OpenAI's reporting convention).

All figures are USD billions. Run:  python build_arr_workbook.py
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT_PATH = "AI_Model_ARR_Gross_vs_Net.xlsx"
AS_OF = "15-Oct-2026"

# ----------------------------------------------------------------------------- styles
NAVY = "1F3864"
LIGHT = "D9E1F2"
INPUT_BLUE = "0000FF"
GREY = "F2F2F2"
EST_YELLOW = "FFF2CC"

TITLE = Font(name="Calibri", size=14, bold=True, color=NAVY)
H2 = Font(name="Calibri", size=11, bold=True, color=NAVY)
HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=10)
BODY_B = Font(name="Calibri", size=10, bold=True)
INPUT = Font(name="Calibri", size=10, color=INPUT_BLUE)
NOTE = Font(name="Calibri", size=9, italic=True, color="595959")

HEAD_FILL = PatternFill("solid", fgColor=NAVY)
SUB_FILL = PatternFill("solid", fgColor=LIGHT)
TOTAL_FILL = PatternFill("solid", fgColor=GREY)
EST_FILL = PatternFill("solid", fgColor=EST_YELLOW)

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP_BOTTOM = Border(top=Side(style="thin", color=NAVY), bottom=Side(style="double", color=NAVY))

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_BN = '#,##0.00;[Red](#,##0.00);"–"'
FMT_PCT = '0.0%;[Red](0.0%);"–"'


def style_header_row(ws, row, ncols, height=42):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER
        cell.fill = HEAD_FILL
        cell.alignment = CENTER
        cell.border = BOX
    ws.row_dimensions[row].height = height


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------------- data
# Columns shared by both lab sheets. Blank ("") channel/take inputs default to assumption refs.
LAB_COLS = [
    "Company",
    "Flagship model(s)",
    "Weights",
    "Include in totals? (Y/N)",
    "As of",
    "Reported ARR ($B)",
    "Reported basis",
    "Channel share of billings",
    "Partner take rate",
    "Gross override ($B)",
    "Gross ARR ($B)",
    "Net ARR ($B)",
    "Gross-to-net haircut",
    "Confidence",
    "Source(s)",
    "Notes",
]

# (company, model, weights, include, as_of, reported, basis, channel, take, gross_override, confidence, source, notes)
FRONTIER_ROWS = [
    (
        "OpenAI",
        "GPT-5.6 family / ChatGPT",
        "Closed",
        "Y",
        "Sep-2026",
        50.0,
        "Net",
        "=Assumptions!$C$7",
        "=Assumptions!$C$6",
        "=Assumptions!$C$8",
        "High",
        "Reuters, CNBC, Axios, Bloomberg (8-Oct-2026)",
        "OpenAI reports NET of the ~20% Microsoft revenue share. $50B is the Sep-2026 run rate told to "
        "investors; the $68B figure circulated in Sep-2026 was a gross-up to Anthropic's methodology and is "
        "used as the gross override. Model-based gross at 20% share would be $62.5B (cross-check). Guided to "
        ">=$70B (net) by YE-2026. Q3 run-rate growth +77%; enterprise +107%. MSFT rev share continues to 2030 "
        "at the same % subject to a cap (Apr-2026 amended agreement).",
    ),
    (
        "Anthropic",
        "Claude (Fable 5 / Opus / Sonnet)",
        "Closed",
        "Y",
        "Jul-2026",
        65.0,
        "Gross",
        "=Assumptions!$C$12",
        "=Assumptions!$C$13",
        "",
        "High",
        "Reuters, Bloomberg (17-Aug-2026); Reuters on IPO prospectus; CNBC (15-Oct-2026)",
        "Anthropic reports GROSS: marketplace (AWS Bedrock, Google Vertex) sales booked at full value with the "
        "partner cut expensed in S&M. Prospectus: 2025 revenue ~$4.6B, 47% via marketplaces ($2.16B), "
        "distribution fees ~$351M (~16% take). Q3-2026 revenue $16.8B (x4 = $67.2B). Analysts see ~$110B run "
        "rate by YE-2026. BofA (Mar-2026) est. up to $6.4B of 2026 hyperscaler payouts - see sensitivity.",
    ),
    (
        "Google DeepMind",
        "Gemini 3 (Pro / Flash)",
        "Closed",
        "Y",
        "Q2-2026",
        12.0,
        "Gross",
        0.0,
        0.0,
        "",
        "Low",
        "SemiAnalysis Tokenomics (Gemini ARR est.); The Compute Chain ($7-10B est.)",
        "ESTIMATE. Alphabet does not disclose Gemini revenue. SemiAnalysis est. $12B 1P Gemini ARR in 2Q26; "
        "The Compute Chain est. $7-10B now / $10-14B by YE. First-party sales, so gross = net. Excludes GCP "
        "infrastructure and 3P model resale (Claude on Vertex). Broader 'Google AI' bundle incl. Workspace "
        "add-ons est. ~$28B (Presenc, May-2026) - not used.",
    ),
    (
        "xAI (SpaceX AI segment)",
        "Grok 4 / SuperGrok",
        "Closed",
        "Y",
        "Mid-2026",
        0.5,
        "Gross",
        0.0,
        0.0,
        "",
        "Medium",
        "Bloomberg via Sacra; SpaceX 10-Q (Q2-2026 AI segment)",
        "Standalone Grok ARR (subscriptions + API, ex-X advertising). SpaceX AI segment did $2.56B in Q2-2026 "
        "(~$10.2B annualized) but that includes X ads ($367M/qtr), compute-rental / AI infrastructure "
        "(~$1.6B/qtr) and data licensing, so it is not a model-revenue figure. Company 2026 target ~$2B.",
    ),
    (
        "Meta (Superintelligence Labs)",
        "Muse Spark 1.x (closed since Apr-2026)",
        "Closed",
        "Y",
        "Oct-2026",
        0.13,
        "Gross",
        0.0,
        0.0,
        "",
        "Low",
        "Vantage OpenRouter provider economics (13-Oct-2026); Bloomberg/CNBC on Meta Model API launch",
        "FLOOR ESTIMATE. Meta Model API launched Jul-2026 ($1.25/$4.25 per M tokens). Only OpenRouter-attributable "
        "billings are observable (~$131M annualized); direct API and Muse Code subscriptions not disclosed. "
        "Meta's open-weight Llama line is shown on the Open-Weight sheet at $0 licensing revenue.",
    ),
]

OPEN_ROWS = [
    (
        "Mistral AI",
        "Mistral Large 3 / Medium (Apache 2.0)",
        "Open (Apache 2.0) + some closed",
        "Y",
        "Sep-2026",
        1.0,
        "Gross",
        "=Assumptions!$C$17",
        "=Assumptions!$C$18",
        "",
        "High",
        "WSJ, Reuters (8-Sep-2026, CFO); CNBC Disruptor 50",
        "CFO: ARR 'slightly above $1B' at the Samsung-led EUR3B round. Was >$400M in Jan-2026, ~$20M a year "
        "earlier. ~60% of revenue in Europe; sells direct plus via Azure/AWS/GCP marketplaces (split not "
        "disclosed - generic channel assumption applied).",
    ),
    (
        "DeepSeek",
        "DeepSeek V4 Pro / Flash (MIT)",
        "Open (MIT)",
        "Y",
        "Sep-2026",
        1.0,
        "Gross",
        0.0,
        0.0,
        "",
        "Medium",
        "Epoch AI 'How do Chinese AI companies make money?' (Oct-2026); Reuters Breakingviews (Jul-2026)",
        "Epoch est. ~$1B run rate, mostly first-party API (Reuters had $400-500M in Jul-2026). Est. API gross "
        "margin 70-80%. Only ~8% of OpenRouter spend on DeepSeek models flows to DeepSeek -> implied total "
        "ecosystem spend on the model family ~$12B+ (memo). Third-party hosts pay DeepSeek nothing.",
    ),
    (
        "Alibaba (Qwen)",
        "Qwen3.x (small/mid open; Qwen3.7-Max proprietary)",
        "Mixed (open small tiers; closed flagship)",
        "Y",
        "Aug-2026",
        2.4,
        "Gross",
        0.0,
        0.0,
        "",
        "Medium",
        "Epoch AI (Model Studio / MaaS annualized); Alibaba Q2-FY27 results",
        "Model-and-application-services (MaaS) run rate; includes third-party model hosting on Model Studio "
        "(Bailian), so overstates Qwen-only revenue. Alibaba AI Cloud & Compute revenue was $7.1B in the Jun-2026 "
        "quarter (not model revenue). First-party cloud sale -> gross = net.",
    ),
    (
        "Z.ai (Zhipu)",
        "GLM-5 / GLM-5.2 (MIT)",
        "Open (MIT)",
        "Y",
        "Sep-2026",
        1.8,
        "Gross",
        0.0,
        0.0,
        "",
        "Low",
        "Epoch AI (management ARR); Zhipu H1-2026 interim report (HKEX)",
        "Management-cited $1.8B ARR conflicts with reported H1-2026 revenue of RMB954M (~$142M); Epoch's "
        "interim-based 2026 estimate is ~$0.87B (API $805M + enterprise $60M). API gross margin 24.6% in "
        "H1-2026. ~25% of GLM-5 tokens served via Z.ai's own API (rest via third-party hosts).",
    ),
    (
        "Moonshot AI (Kimi)",
        "Kimi K3 (Modified MIT)",
        "Open (Modified MIT)",
        "Y",
        "Aug-2026",
        1.0,
        "Gross",
        0.0,
        0.0,
        "",
        "Medium",
        "Epoch AI; Reuters on K3 hosting revenue-share talks",
        "~70% of revenue from model access (Jun-2026). ARR was >$200M in Apr and ~$300M in Jun-2026 before "
        "the K3 step-up. Seeking up to 30% revenue share on K3 hosted by Azure/AWS/GCP - would create a "
        "net-basis stream (Moonshot receives 30% of partner gross).",
    ),
    (
        "MiniMax",
        "MiniMax M3 (open weights, non-commercial)",
        "Open (non-commercial)",
        "Y",
        "Aug-2026",
        0.8,
        "Gross",
        0.0,
        0.0,
        "",
        "Low",
        "Epoch AI (annualized from H1-2026 interim report); MiniMax HKEX filings",
        "H1-2026 revenue $116.6M (+283% YoY). Consumer app (Talkie) gross margin ~5% vs ~69% for enterprise "
        "services. Management run rate used; interim-based estimate ~$0.3B.",
    ),
    (
        "Cohere",
        "Command A (open weights, CC-BY-NC)",
        "Open (non-commercial)",
        "Y",
        "May-2026",
        0.5,
        "Gross",
        "=Assumptions!$C$17",
        "=Assumptions!$C$18",
        "",
        "Low",
        "CNBC (13-Feb-2026 investor memo: $240M ARR FY2025); Presenc AI est. May-2026",
        "Disclosed $240M ARR for 2025 (beat $200M target; ~70% gross margin). $500M is a third-party May-2026 "
        "estimate. Sells direct and via Oracle/Azure/AWS marketplaces - generic channel assumption applied.",
    ),
    (
        "AI21 Labs",
        "Jamba (open weights)",
        "Open",
        "Y",
        "2026E",
        0.35,
        "Gross",
        "=Assumptions!$C$17",
        "=Assumptions!$C$18",
        "",
        "Low",
        "Presenc AI / Clink third-party estimates",
        "ESTIMATE - AI21 has never disclosed revenue. Generic channel assumption applied.",
    ),
    (
        "Meta (Llama)",
        "Llama 4 (community license)",
        "Open (community license, 700M MAU cap)",
        "Y",
        "2026",
        0.0,
        "Gross",
        0.0,
        0.0,
        "",
        "High",
        "Meta Llama Community License; Licenseware",
        "Royalty-free: $0 direct licensing ARR. Meta has undisclosed revenue-share arrangements with some Llama "
        "hosts. Shown so open-weight totals reflect the largest US open-weight line. Meta's paid Muse Spark API "
        "is on the Frontier sheet.",
    ),
    (
        "ByteDance (Doubao / Seed)",
        "Doubao (closed) / Seed-OSS (open)",
        "Mixed (flagship closed)",
        "N",
        "Jul-2026",
        4.0,
        "Gross",
        0.0,
        0.0,
        "",
        "Low",
        "Epoch AI (Volcano Engine MaaS ARR); IDC",
        "EXCLUDED from totals by default: flagship Doubao models are closed and the ~$4B Volcano Engine MaaS "
        "run rate includes third-party models (DeepSeek, Kimi, GLM). 2025 MaaS revenue was only ~$221M. Flip "
        "column D to 'Y' to include.",
    ),
]

# Memo: where open-weight inference spend accrues instead of the labs (not model ARR; excluded from totals).
HOST_ROWS = [
    ("Fireworks AI", 1.0, "Annualized run rate", "Jul-2026", "Company disclosed (>$1B; >40T tokens/day)"),
    ("Together AI", 1.0, "Annualized revenue (est.)", "Feb-2026", "Sacra estimate; company cites >$1.15B annual bookings (Q2-2026)"),
    ("Baseten", 0.6, "Annualized revenue (est.)", "2026", "Third-party estimate"),
    ("fal", 0.4, "Annualized revenue (est.)", "2026", "Third-party estimate"),
]

SOURCES = [
    ("OpenAI $50B net ARR; $68-70B gross-up", "Reuters via Investing.com", "https://www.investing.com/news/stock-market-news/openais-annualized-revenue-20-billion-less-than-previously-signaled-ft-reports-4939854"),
    ("OpenAI $50B vs $68B gross partner revenue", "CNBC", "https://www.cnbc.com/2026/10/08/open-ai-revenue-nvidia-oracle-coreweave.html"),
    ("OpenAI / Anthropic revenue accounting difference", "Axios", "https://www.axios.com/2026/10/08/openai-50-billion-arr-anthropic-revenue-accounting"),
    ("OpenAI expects >=$70B ARR by YE-2026", "Bloomberg via Yahoo Finance", "https://finance.yahoo.com/technology/ai/articles/openai-expects-70-billion-annualized-014338719.html"),
    ("OpenAI 20% Microsoft revenue share; net reporting", "TechCrunch", "https://techcrunch.com/2025/11/14/leaked-documents-shed-light-into-how-much-openai-pays-microsoft/"),
    ("OpenAI-Microsoft amended agreement (rev share to 2030, capped)", "OpenAI", "https://openai.com/index/next-phase-of-microsoft-partnership/"),
    ("Net vs gross reporting conventions; BofA $6.4B 2026 payout est.", "Forbes", "https://www.forbes.com/sites/josipamajic/2026/03/25/openai-and-anthropic-count-revenue-differently-and-investors-are-looking-into-it/"),
    ("Anthropic $65B run rate (end-Jul-2026)", "Reuters", "https://www.reuters.com/technology/anthropic-revenue-run-rate-tops-65-billion-source-says-2026-08-17/"),
    ("Anthropic Q2-2026 revenue $11.5B", "Axios", "https://www.axios.com/2026/08/17/anthropic-revenue-run-rate-ipo-openai"),
    ("Anthropic Q3-2026 revenue $16.8B; $110B YE target", "CNBC", "https://www.cnbc.com/2026/10/15/anthropics-third-quarter-revenue-hits-16point8-billion-ipo.html"),
    ("Anthropic prospectus: 47% via marketplaces, $351M fees", "Reuters via Investing.com", "https://www.investing.com/news/stock-market-news/exclusiveanthropic-ipo-prospectus-lays-bare-deep-dependence-on-big-tech-partners-4923585"),
    ("Anthropic Bedrock economics (~75% lab share)", "SemiAnalysis", "https://newsletter.semianalysis.com/p/anthropic-growth-and-bedrock-mix"),
    ("Gemini ARR est. $12B (2Q26)", "SemiAnalysis", "https://newsletter.semianalysis.com/p/gemini-is-cooked-but-gcp-is-cooking"),
    ("Gemini direct model revenue est. $7-10B", "The Compute Chain", "https://thecomputechain.substack.com/p/beyond-openai-and-anthropic-how-much"),
    ("Google AI bundle ~$28B; lab comparison table", "Presenc AI", "https://presenc.ai/research/ai-lab-revenue-and-valuation-2026"),
    ("xAI standalone ~$500M ARR; SpaceX AI segment", "Value Add VC", "https://valueaddvc.com/blog/xai-revenue-how-grok-and-the-api-actually-make-money"),
    ("xAI / X combined revenue; Grok standalone", "Sacra", "https://sacra.com/research/why-spacex-acquired-xai/"),
    ("Meta Muse Spark paid API launch", "Bloomberg via Yahoo Finance", "https://finance.yahoo.com/technology/ai/articles/meta-launches-paid-muse-spark-205823294.html"),
    ("Meta OpenRouter-attributable API billings", "Vantage", "https://routers.vantage.sh/providers/meta"),
    ("Llama royalty-free community license", "Licenseware", "https://wiki.licenseware.io/vendors/meta/"),
    ("Mistral ARR slightly above $1B (CFO)", "WSJ", "https://www.wsj.com/tech/ai/mistral-ai-exceeds-24-billion-valuation-after-samsung-led-investment-round-ea377b82"),
    ("Mistral EUR3B round; $1B ARR track", "Reuters", "https://www.reuters.com/world/europe/french-ai-company-mistral-hits-24-billion-valuation-funding-round-2026-09-08/"),
    ("Chinese labs: DeepSeek, Z.ai, Moonshot, MiniMax, Alibaba, ByteDance", "Epoch AI", "https://epoch.ai/publications/how-do-chinese-ai-companies-make-money"),
    ("Chinese lab revenue summary; DeepSeek 8% capture", "BigGo Finance (Epoch summary)", "https://finance.biggo.com/news/6f6fb5ef-6493-41f1-849c-d6fd0d4bfeab"),
    ("Open-weight revenue leakage to third-party hosts", "Opentechwire", "https://www.opentechwire.com/article/open-weight-paradox-chinese-ai-labs-revenue-third-party"),
    ("Zhipu / MiniMax H1-2026 revenue; Kimi run rate", "Nexuswild", "https://nexuswild.com/news/ai-revenue-openai-claude-gemini-grok-deepseek-kimi-competition/"),
    ("Cohere $240M ARR (FY2025)", "CNBC", "https://www.cnbc.com/2026/02/13/ai-startup-cohere-revenue-ipo.html"),
    ("AI21 Labs ~$350M est.", "Clink", "https://clinkbill.com/arr-leaderboard/ai21"),
    ("Fireworks / Together / Baseten / fal run rates", "New Market Pitch", "https://newmarketpitch.com/blogs/news/ai-infrastructure-top-revenue"),
    ("Together AI $1B annualized est.", "Sacra", "https://sacra.com/c/together-ai/"),
]


# ----------------------------------------------------------------------------- builders
def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions")
    set_widths(ws, [4, 46, 14, 70])
    ws["B2"] = "Assumptions & inputs"
    ws["B2"].font = TITLE
    ws["B3"] = f"Blue = hard-coded input (editable). All $ in billions. As of {AS_OF}."
    ws["B3"].font = NOTE

    rows = [
        # row, label, value, fmt, note
        (5, "OpenAI (reports NET)", None, None, None),
        (6, "Microsoft revenue share on OpenAI revenue", 0.20, FMT_PCT, "Widely reported ~20% of ChatGPT + API revenue; OpenAI reports net of this. Continues to 2030 subject to a cap."),
        (7, "Share of OpenAI billings subject to the MSFT share", 1.00, FMT_PCT, "Set to 100%: the share applies to OpenAI's total revenue, not just Azure channel sales."),
        (8, "OpenAI gross ARR per investor gross-up ($B)", 68.0, FMT_BN, "Sep-2026 figure shared with investors on Anthropic's gross methodology (CNBC/FT). Blank this cell to fall back to the model-based gross-up (= Net / (1 - share x rate))."),
        (9, "OpenAI model-based gross ARR ($B) - cross-check", "=Frontier!F5/(1-C7*C6)", FMT_BN, "Formula: reported net / (1 - channel share x take rate)."),
        (11, "Anthropic (reports GROSS)", None, None, None),
        (12, "Channel share of billings via cloud marketplaces", 0.47, FMT_PCT, "2025 actual from IPO prospectus ($2.16B of ~$4.6B). Trend: 11% (2023) -> 32% (2024) -> 47% (2025); likely higher in 2026."),
        (13, "Blended partner take rate on marketplace sales", 0.16, FMT_PCT, "2025 realized: ~$351M fees / $2.16B marketplace sales. SemiAnalysis puts Anthropic's Bedrock share at ~75% (25% take) on newer deals."),
        (14, "Implied gross-to-net haircut", "=C12*C13", FMT_PCT, "Channel share x take rate."),
        (16, "Generic open-weight lab channel assumption (Mistral, Cohere, AI21)", None, None, None),
        (17, "Assumed channel share of billings via marketplaces", 0.25, FMT_PCT, "ASSUMPTION - none of these labs disclose the marketplace share of sales. Set to 0% to treat as fully direct."),
        (18, "Assumed partner take rate", 0.20, FMT_PCT, "ASSUMPTION - typical hyperscaler marketplace resale cut is 20-30%; Anthropic realized ~16% in 2025."),
        (19, "Implied gross-to-net haircut", "=C17*C18", FMT_PCT, ""),
        (21, "Memo: BofA estimate of Anthropic hyperscaler payouts ($B)", None, None, None),
        (22, "2025E payouts", 1.9, FMT_BN, "BofA, Mar-2026. Note: broader than the $351M of distribution fees disclosed in the prospectus (likely includes other cost-sharing)."),
        (23, "2026E payouts", 6.4, FMT_BN, "BofA, Mar-2026 ('up to')."),
        (24, "2026E payouts as % of YE-2026 $110B run rate", "=C23/110", FMT_PCT, "Alternative haircut lens; compare with C14."),
    ]
    for r, label, val, fmt, note in rows:
        ws.cell(row=r, column=2, value=label)
        if val is None and fmt is None:
            ws.cell(row=r, column=2).font = H2
            ws.cell(row=r, column=2).fill = SUB_FILL
            ws.cell(row=r, column=3).fill = SUB_FILL
            ws.cell(row=r, column=4).fill = SUB_FILL
            continue
        ws.cell(row=r, column=2).font = BODY
        c = ws.cell(row=r, column=3, value=val)
        c.number_format = fmt
        c.border = BOX
        c.font = INPUT if not (isinstance(val, str) and val.startswith("=")) else BODY_B
        if isinstance(val, str) and val.startswith("="):
            c.fill = TOTAL_FILL
        n = ws.cell(row=r, column=4, value=note)
        n.font = NOTE
        n.alignment = WRAP
    ws.freeze_panes = "A5"
    return ws


def build_lab_sheet(wb, name, title, rows, start_row=5):
    ws = wb.create_sheet(name)
    set_widths(ws, [26, 32, 22, 11, 10, 11, 10, 11, 10, 11, 11, 11, 11, 11, 44, 90])
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = (
        "Gross ARR = annualized end-customer billings incl. cloud-marketplace / partner channel sales at full value. "
        "Net ARR = retained by the lab after partner revenue share. Blue = input; yellow fill = estimate or assumption. $B."
    )
    ws["A2"].font = NOTE
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=16)
    ws["A3"] = f"As of {AS_OF}. Run rates are from different months; see 'As of' column. Toggle column D to include/exclude a row from totals."
    ws["A3"].font = NOTE
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=16)

    hdr = start_row - 1
    for i, h in enumerate(LAB_COLS, start=1):
        ws.cell(row=hdr, column=i, value=h)
    style_header_row(ws, hdr, len(LAB_COLS))

    first = start_row
    for idx, row in enumerate(rows):
        r = start_row + idx
        company, model, weights, include, as_of, reported, basis, channel, take, override, conf, source, notes = row
        values = {
            1: company, 2: model, 3: weights, 4: include, 5: as_of, 6: reported, 7: basis,
            8: channel, 9: take, 10: override if override != "" else None,
            # Gross: override if given; else reported if basis is Gross; else gross-up of the net figure.
            11: f'=IF(J{r}<>"",J{r},IF(G{r}="Gross",F{r},F{r}/(1-H{r}*I{r})))',
            # Net: reported if basis is Net; else gross less channel share x take.
            12: f'=IF(G{r}="Net",F{r},K{r}*(1-H{r}*I{r}))',
            13: f'=IF(K{r}=0,0,1-L{r}/K{r})',
            14: conf, 15: source, 16: notes,
        }
        for c, v in values.items():
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BOX
            cell.font = BODY
            cell.alignment = WRAP if c in (1, 2, 3, 15, 16) else CENTER
        for c in (4, 5, 6, 7, 8, 9, 10):
            cell = ws.cell(row=r, column=c)
            is_formula = isinstance(cell.value, str) and str(cell.value).startswith("=")
            cell.font = BODY_B if is_formula else INPUT
        ws.cell(row=r, column=6).number_format = FMT_BN
        ws.cell(row=r, column=10).number_format = FMT_BN
        ws.cell(row=r, column=11).number_format = FMT_BN
        ws.cell(row=r, column=12).number_format = FMT_BN
        ws.cell(row=r, column=8).number_format = FMT_PCT
        ws.cell(row=r, column=9).number_format = FMT_PCT
        ws.cell(row=r, column=13).number_format = FMT_PCT
        ws.cell(row=r, column=11).font = BODY_B
        ws.cell(row=r, column=12).font = BODY_B
        if conf == "Low":
            for c in (6, 11, 12):
                ws.cell(row=r, column=c).fill = EST_FILL
        ws.row_dimensions[r].height = 92
    last = start_row + len(rows) - 1

    # Totals (only rows flagged Y).
    t = last + 1
    ws.cell(row=t, column=1, value="Total (included rows)")
    ws.cell(row=t, column=4, value=f'=COUNTIF(D{first}:D{last},"Y")')
    ws.cell(row=t, column=11, value=f'=SUMIF($D${first}:$D${last},"Y",K{first}:K{last})')
    ws.cell(row=t, column=12, value=f'=SUMIF($D${first}:$D${last},"Y",L{first}:L{last})')
    ws.cell(row=t, column=13, value=f'=IF(K{t}=0,0,1-L{t}/K{t})')
    for c in range(1, len(LAB_COLS) + 1):
        cell = ws.cell(row=t, column=c)
        cell.font = BODY_B
        cell.fill = TOTAL_FILL
        cell.border = TOP_BOTTOM
    ws.cell(row=t, column=11).number_format = FMT_BN
    ws.cell(row=t, column=12).number_format = FMT_BN
    ws.cell(row=t, column=13).number_format = FMT_PCT
    ws.cell(row=t, column=4).alignment = CENTER

    # Data validation for the include flag and basis.
    dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    dv_basis = DataValidation(type="list", formula1='"Gross,Net"', allow_blank=False)
    ws.add_data_validation(dv_yn)
    ws.add_data_validation(dv_basis)
    dv_yn.add(f"D{first}:D{last}")
    dv_basis.add(f"G{first}:G{last}")

    # Highlight excluded rows.
    ws.conditional_formatting.add(
        f"A{first}:P{last}",
        FormulaRule(formula=[f'$D{first}="N"'], font=Font(color="808080", italic=True)),
    )
    ws.freeze_panes = ws.cell(row=start_row, column=2)
    ws.auto_filter.ref = f"A{hdr}:P{last}"
    return ws, first, last, t


def add_host_memo(ws, start_row):
    r = start_row
    ws.cell(row=r, column=1, value="Memo: open-weight inference hosts (where third-party spend on open models accrues; NOT model-lab ARR, excluded from totals)").font = H2
    r += 1
    heads = ["Provider", "Annualized ($B)", "Metric", "As of", "Note"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, len(heads), height=20)
    r += 1
    first = r
    for prov, val, metric, asof, note in HOST_ROWS:
        ws.cell(row=r, column=1, value=prov).font = BODY
        c = ws.cell(row=r, column=2, value=val)
        c.number_format = FMT_BN
        c.font = INPUT
        ws.cell(row=r, column=3, value=metric).font = BODY
        ws.cell(row=r, column=4, value=asof).font = BODY
        ws.cell(row=r, column=5, value=note).font = NOTE
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = BOX
        r += 1
    ws.cell(row=r, column=1, value="Total hosts (memo)").font = BODY_B
    tot = ws.cell(row=r, column=2, value=f"=SUM(B{first}:B{r-1})")
    tot.number_format = FMT_BN
    tot.font = BODY_B
    for col in range(1, 6):
        ws.cell(row=r, column=col).fill = TOTAL_FILL
        ws.cell(row=r, column=col).border = TOP_BOTTOM
    r += 2
    ws.cell(row=r, column=1, value="Memo: DeepSeek ecosystem spend implied by its ~8% capture of OpenRouter spend on its models ($B)").font = BODY
    c = ws.cell(row=r, column=11, value="=Open_Weight!F6/0.08")
    c.number_format = FMT_BN
    c.font = BODY_B
    ws.cell(row=r, column=12, value="(Epoch AI; illustrative)").font = NOTE
    return r


def build_anthropic_sensitivity(wb):
    ws = wb.create_sheet("Anthropic_Sensitivity")
    set_widths(ws, [4, 30, 12, 12, 12, 12, 12, 12])
    ws["B2"] = "Anthropic net ARR sensitivity ($B)"
    ws["B2"].font = TITLE
    ws["B3"] = "Net ARR = Gross ARR x (1 - channel share x partner take rate). Gross ARR links to the Frontier sheet."
    ws["B3"].font = NOTE
    ws["B5"] = "Gross ARR ($B)"
    ws["B5"].font = BODY
    ws["C5"] = "=Frontier!K6"
    ws["C5"].number_format = FMT_BN
    ws["C5"].font = BODY_B
    ws["C5"].fill = TOTAL_FILL
    ws["C5"].border = BOX

    takes = [0.12, 0.16, 0.20, 0.25, 0.30]
    shares = [0.35, 0.47, 0.55, 0.60, 0.70]
    ws["B7"] = "Channel share  \\  Partner take rate"
    ws["B7"].font = HEADER
    ws["B7"].fill = HEAD_FILL
    ws["B7"].alignment = CENTER
    ws["B7"].border = BOX
    for j, t in enumerate(takes):
        c = ws.cell(row=7, column=3 + j, value=t)
        c.number_format = FMT_PCT
        c.font = HEADER
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BOX
    for i, s in enumerate(shares):
        r = 8 + i
        c = ws.cell(row=r, column=2, value=s)
        c.number_format = FMT_PCT
        c.font = BODY_B
        c.fill = SUB_FILL
        c.alignment = CENTER
        c.border = BOX
        for j in range(len(takes)):
            col = get_column_letter(3 + j)
            cell = ws.cell(row=r, column=3 + j, value=f"=$C$5*(1-$B{r}*{col}$7)")
            cell.number_format = FMT_BN
            cell.font = BODY
            cell.border = BOX
            cell.alignment = CENTER

    ws["B14"] = "Haircut (1 - Net/Gross)"
    ws["B14"].font = H2
    ws["B15"] = "Channel share  \\  Partner take rate"
    ws["B15"].font = HEADER
    ws["B15"].fill = HEAD_FILL
    ws["B15"].alignment = CENTER
    ws["B15"].border = BOX
    for j, t in enumerate(takes):
        c = ws.cell(row=15, column=3 + j, value=t)
        c.number_format = FMT_PCT
        c.font = HEADER
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BOX
    for i, s in enumerate(shares):
        r = 16 + i
        c = ws.cell(row=r, column=2, value=s)
        c.number_format = FMT_PCT
        c.font = BODY_B
        c.fill = SUB_FILL
        c.alignment = CENTER
        c.border = BOX
        for j in range(len(takes)):
            col = get_column_letter(3 + j)
            cell = ws.cell(row=r, column=3 + j, value=f"=$B{r}*{col}$15")
            cell.number_format = FMT_PCT
            cell.font = BODY
            cell.border = BOX
            cell.alignment = CENTER

    ws["B23"] = "Reference points"
    ws["B23"].font = H2
    refs = [
        ("Base case (prospectus 2025: 47% share x 16% take)", "=Frontier!L6"),
        ("Share of 2025 billings via marketplaces", "=Assumptions!C12"),
        ("2025 realized take rate", "=Assumptions!C13"),
        ("BofA 2026E payouts / $110B YE run rate", "=Assumptions!C24"),
        ("OpenAI implied haircut (1 - 50/68) for comparison", "=Frontier!M5"),
    ]
    for i, (label, f) in enumerate(refs):
        r = 24 + i
        ws.cell(row=r, column=2, value=label).font = BODY
        c = ws.cell(row=r, column=3, value=f)
        c.number_format = FMT_BN if i == 0 else FMT_PCT
        c.font = BODY_B
        c.fill = TOTAL_FILL
        c.border = BOX
    return ws


def build_summary(wb, fr, op):
    """fr / op: dicts with first,last,total row indices for the lab sheets."""
    ws = wb.create_sheet("Summary", 0)
    set_widths(ws, [4, 34, 15, 15, 15, 15, 15, 4, 60])
    ws["B2"] = "AI model labs: gross vs. net ARR - frontier (closed-weight) vs. open-weight"
    ws["B2"].font = TITLE
    ws["B3"] = (
        f"USD billions, latest reported run rates as of {AS_OF}. Gross = end-customer billings incl. partner channel; "
        "Net = retained by the lab after partner revenue share. Figures are run rates from different months and are not audited."
    )
    ws["B3"].font = NOTE
    ws.merge_cells("B3:I3")
    ws.row_dimensions[3].height = 28
    ws["B3"].alignment = WRAP

    heads = ["Segment", "# labs included", "Gross ARR ($B)", "Net ARR ($B)", "Gross-to-net haircut", "Share of total net"]
    for i, h in enumerate(heads, start=2):
        ws.cell(row=5, column=i, value=h)
    style_header_row(ws, 5, 7, height=32)
    ws.cell(row=5, column=1).fill = PatternFill(fill_type=None)
    ws.cell(row=5, column=1).border = Border()

    ws["B6"] = "Frontier (closed-weight) labs"
    ws["C6"] = f"=Frontier!D{fr['total']}"
    ws["D6"] = f"=Frontier!K{fr['total']}"
    ws["E6"] = f"=Frontier!L{fr['total']}"
    ws["F6"] = "=IF(D6=0,0,1-E6/D6)"
    ws["G6"] = "=IF($E$8=0,0,E6/$E$8)"
    ws["B7"] = "Open-weight labs"
    ws["C7"] = f"=Open_Weight!D{op['total']}"
    ws["D7"] = f"=Open_Weight!K{op['total']}"
    ws["E7"] = f"=Open_Weight!L{op['total']}"
    ws["F7"] = "=IF(D7=0,0,1-E7/D7)"
    ws["G7"] = "=IF($E$8=0,0,E7/$E$8)"
    ws["B8"] = "Total"
    ws["C8"] = "=C6+C7"
    ws["D8"] = "=D6+D7"
    ws["E8"] = "=E6+E7"
    ws["F8"] = "=IF(D8=0,0,1-E8/D8)"
    ws["G8"] = "=G6+G7"
    for r in (6, 7, 8):
        for c in range(2, 8):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            cell.font = BODY_B if r == 8 else BODY
            if r == 8:
                cell.fill = TOTAL_FILL
                cell.border = TOP_BOTTOM
        ws.cell(row=r, column=3).number_format = "0"
        ws.cell(row=r, column=4).number_format = FMT_BN
        ws.cell(row=r, column=5).number_format = FMT_BN
        ws.cell(row=r, column=6).number_format = FMT_PCT
        ws.cell(row=r, column=7).number_format = FMT_PCT
        for c in range(3, 8):
            ws.cell(row=r, column=c).alignment = CENTER

    ws["B10"] = "Open-weight net ARR as % of frontier net ARR"
    ws["C10"] = "=IF(E6=0,0,E7/E6)"
    ws["C10"].number_format = FMT_PCT
    ws["B11"] = "Open-weight gross ARR as % of frontier gross ARR"
    ws["C11"] = "=IF(D6=0,0,D7/D6)"
    ws["C11"].number_format = FMT_PCT
    ws["B12"] = "OpenAI + Anthropic share of total net ARR"
    ws["C12"] = "=IF(E8=0,0,(Frontier!L5+Frontier!L6)/E8)"
    ws["C12"].number_format = FMT_PCT
    for r in (10, 11, 12):
        ws.cell(row=r, column=2).font = BODY
        ws.cell(row=r, column=3).font = BODY_B
        ws.cell(row=r, column=3).fill = TOTAL_FILL
        ws.cell(row=r, column=3).border = BOX
        ws.cell(row=r, column=3).alignment = CENTER

    # Per-lab table feeding the chart.
    ws["B14"] = "By lab (included rows only)"
    ws["B14"].font = H2
    heads2 = ["Lab", "Segment", "Gross ARR ($B)", "Net ARR ($B)", "Haircut", "As of"]
    for i, h in enumerate(heads2, start=2):
        ws.cell(row=15, column=i, value=h)
    style_header_row(ws, 15, 7, height=22)
    ws.cell(row=15, column=1).fill = PatternFill(fill_type=None)
    ws.cell(row=15, column=1).border = Border()

    r = 16
    chart_first = r
    for sheet, seg, info in (("Frontier", "Frontier", fr), ("Open_Weight", "Open-weight", op)):
        for src in range(info["first"], info["last"] + 1):
            ws.cell(row=r, column=2, value=f"={sheet}!A{src}")
            ws.cell(row=r, column=3, value=seg)
            ws.cell(row=r, column=4, value=f'=IF({sheet}!D{src}="Y",{sheet}!K{src},0)')
            ws.cell(row=r, column=5, value=f'=IF({sheet}!D{src}="Y",{sheet}!L{src},0)')
            ws.cell(row=r, column=6, value=f"={sheet}!M{src}")
            ws.cell(row=r, column=7, value=f"={sheet}!E{src}")
            for c in range(2, 8):
                cell = ws.cell(row=r, column=c)
                cell.border = BOX
                cell.font = BODY
                cell.alignment = CENTER if c >= 3 else WRAP
            ws.cell(row=r, column=4).number_format = FMT_BN
            ws.cell(row=r, column=5).number_format = FMT_BN
            ws.cell(row=r, column=6).number_format = FMT_PCT
            r += 1
    chart_last = r - 1
    ws.cell(row=r, column=2, value="Check: sum of per-lab net = segment total").font = NOTE
    chk = ws.cell(row=r, column=5, value=f"=SUM(E{chart_first}:E{chart_last})-E8")
    chk.number_format = '0.00;[Red]-0.00;"OK"'
    chk.font = NOTE

    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "clustered"
    chart.title = "Gross vs. net ARR by lab ($B)"
    chart.y_axis.title = "$B"
    chart.height = 14
    chart.width = 24
    data = Reference(ws, min_col=4, max_col=5, min_row=15, max_row=chart_last)
    cats = Reference(ws, min_col=2, min_row=chart_first, max_row=chart_last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.y_axis.majorGridlines = None
    chart.x_axis.tickLblSkip = 1
    chart.x_axis.scaling.orientation = "maxMin"  # largest labs (top of the table) at the top of the chart
    chart.legend.position = "b"
    ws.add_chart(chart, "I5")

    notes = [
        "Reading the numbers",
        "1. The gross/net gap is concentrated in the two leaders: OpenAI reports net of Microsoft's ~20% share (the $68B 'gross' figure was an investor gross-up), while Anthropic books cloud-marketplace sales gross and expenses the partner cut.",
        "2. Google, xAI and the Chinese labs sell predominantly first-party, so gross = net for them in this model; their uncertainty is in the reported figure itself (yellow = estimate).",
        "3. Open-weight labs capture only part of the spend on their models: third-party hosts (Fireworks, Together, Baseten, fal) and clouds serve most tokens. DeepSeek is estimated to receive ~8% of OpenRouter spend on DeepSeek models. See memo on the Open_Weight sheet.",
        "4. ByteDance (~$4B MaaS) is excluded by default because its flagship Doubao models are closed and the figure includes third-party models; flip Open_Weight!D14 to 'Y' to include.",
        "5. Run rates are month-of-record figures and are not comparable across labs without the 'As of' date; Anthropic's Q3-2026 quarterly revenue ($16.8B) annualizes to $67.2B and YE-2026 targets are ~$110B (Anthropic) and >=$70B net (OpenAI).",
    ]
    nr = r + 2
    for i, line in enumerate(notes):
        cell = ws.cell(row=nr + i, column=2, value=line)
        cell.font = H2 if i == 0 else NOTE
        cell.alignment = WRAP
        ws.merge_cells(start_row=nr + i, start_column=2, end_row=nr + i, end_column=7)
        ws.row_dimensions[nr + i].height = 15 if i == 0 else 30
    ws.freeze_panes = "A6"
    return ws


def build_methodology(wb):
    ws = wb.create_sheet("Methodology")
    set_widths(ws, [4, 120])
    ws["B2"] = "Methodology & definitions"
    ws["B2"].font = TITLE
    lines = [
        ("Definitions", True),
        ("ARR / annualized run rate: the most recent month's (or quarter's) revenue extrapolated to a year. Labs disclose these to investors or the press; none are audited annual revenue. Where only quarterly revenue exists, x4 is shown in the notes, not in the ARR column.", False),
        ("Gross ARR: annualized value of end-customer billings for the lab's models, including sales that clear through cloud marketplaces (AWS Bedrock, Google Vertex / Gemini Enterprise Agent Platform, Azure Foundry) or other distribution partners at full value. This is Anthropic's GAAP presentation (lab is the principal; partner fee is an S&M expense).", False),
        ("Net ARR: annualized revenue retained by the lab after the distribution partner's revenue share. This is OpenAI's GAAP presentation (reports net of the ~20% Microsoft share).", False),
        ("Gross-to-net haircut: 1 - Net / Gross = channel share of billings x partner take rate.", False),
        ("", False),
        ("Calculation logic (Frontier and Open_Weight sheets)", True),
        ("Gross ARR = Gross override if provided; otherwise the reported figure if the basis is 'Gross'; otherwise reported net / (1 - channel share x take rate).", False),
        ("Net ARR = reported figure if the basis is 'Net'; otherwise Gross ARR x (1 - channel share x take rate).", False),
        ("Rows with 'N' in the Include column are excluded from totals and the Summary via SUMIF. Totals are not adjusted for differing 'As of' dates.", False),
        ("", False),
        ("Classification", True),
        ("Frontier (closed-weight): labs whose flagship models are API/subscription-only - OpenAI, Anthropic, Google DeepMind, xAI, Meta (Muse Spark, closed since Apr-2026).", False),
        ("Open-weight: labs whose current flagship or main commercial line ships downloadable weights - Mistral, DeepSeek, Alibaba Qwen (small/mid tiers open; Max proprietary), Z.ai, Moonshot, MiniMax (non-commercial license), Cohere (CC-BY-NC), AI21, Meta Llama. ByteDance is listed but excluded by default.", False),
        ("Open-weight artefacts from closed labs (OpenAI gpt-oss, Google Gemma) generate no separate revenue and are not listed.", False),
        ("Inference hosts (Fireworks, Together, Baseten, fal) are a memo item: they monetise open weights but are not model labs, so they are excluded from lab totals.", False),
        ("", False),
        ("Confidence flags", True),
        ("High: company disclosure or multiple major outlets citing investor materials. Medium: single credible report or management figure. Low: third-party estimate or a management figure that conflicts with filed financials (yellow fill).", False),
        ("", False),
        ("Known limitations", True),
        ("Anthropic's channel share and take rate are 2025 full-year actuals from the IPO prospectus; 2026 marketplace mix is likely higher and newer Bedrock terms reportedly carry ~25% take. Use the Anthropic_Sensitivity sheet.", False),
        ("OpenAI's gross figure is an investor construct; OpenAI itself has not published a gross number. Model-based gross at 20% is $62.5B.", False),
        ("Gemini, Grok and Muse Spark revenue are not disclosed by their parents; estimates vary widely (Gemini $7-14B).", False),
        ("Chinese lab run rates mix management claims and interim-report annualisations; Z.ai's $1.8B ARR vs ~$142M H1-2026 revenue is the largest discrepancy.", False),
        ("Generic 25% channel / 20% take assumption for Mistral, Cohere and AI21 is a placeholder; set Assumptions!C17 to 0% to treat them as fully direct.", False),
    ]
    r = 4
    for text, is_head in lines:
        cell = ws.cell(row=r, column=2, value=text)
        cell.font = H2 if is_head else BODY
        cell.alignment = WRAP
        if is_head:
            cell.fill = SUB_FILL
        elif text:
            ws.row_dimensions[r].height = 30
        r += 1
    return ws


def build_sources(wb):
    ws = wb.create_sheet("Sources")
    set_widths(ws, [4, 56, 30, 110])
    ws["B2"] = "Sources"
    ws["B2"].font = TITLE
    ws["B3"] = f"Accessed {AS_OF}."
    ws["B3"].font = NOTE
    for i, h in enumerate(["Data point", "Outlet", "URL"], start=2):
        ws.cell(row=5, column=i, value=h)
    style_header_row(ws, 5, 4, height=20)
    ws.cell(row=5, column=1).fill = PatternFill(fill_type=None)
    ws.cell(row=5, column=1).border = Border()
    for i, (dp, outlet, url) in enumerate(SOURCES):
        r = 6 + i
        ws.cell(row=r, column=2, value=dp).font = BODY
        ws.cell(row=r, column=3, value=outlet).font = BODY
        c = ws.cell(row=r, column=4, value=url)
        c.hyperlink = url
        c.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).border = BOX
            ws.cell(row=r, column=col).alignment = WRAP
    ws.freeze_panes = "A6"
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_assumptions(wb)
    _, f_first, f_last, f_total = build_lab_sheet(
        wb, "Frontier", "Frontier (closed-weight) model labs: gross vs. net ARR", FRONTIER_ROWS
    )
    ws_open, o_first, o_last, o_total = build_lab_sheet(
        wb, "Open_Weight", "Open-weight model labs: gross vs. net ARR", OPEN_ROWS
    )
    add_host_memo(ws_open, o_total + 3)
    build_anthropic_sensitivity(wb)
    build_summary(
        wb,
        {"first": f_first, "last": f_last, "total": f_total},
        {"first": o_first, "last": o_last, "total": o_total},
    )
    build_methodology(wb)
    build_sources(wb)

    # Sheet order: Summary, Frontier, Open_Weight, Assumptions, Anthropic_Sensitivity, Methodology, Sources
    order = ["Summary", "Frontier", "Open_Weight", "Assumptions", "Anthropic_Sensitivity", "Methodology", "Sources"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(OUT_PATH)
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
