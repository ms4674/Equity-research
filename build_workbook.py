"""Builds Q2_2026_Earnings_Preview.xlsx aggregating consensus estimates,
estimate-revision trends, beat/miss history, and upside/downside drivers
for GOOGL, META, SNAP, PINS, RDDT. Data retrieved July 10, 2026."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY = "1F3864"
LIGHT = "D9E2F3"
GREEN = "C6EFCE"
RED = "FFC7CE"
YELLOW = "FFEB9C"
GREY = "F2F2F2"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header(ws, row, ncols, fill=NAVY):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def add_title(ws, title, subtitle, ncols):
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14, color=NAVY)
    ws.cell(row=2, column=1, value=subtitle).font = Font(italic=True, size=9, color="595959")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)


def fill_rows(ws, start_row, rows, widths=None, wrap_cols=None, band=True):
    wrap_cols = wrap_cols or set()
    for i, row in enumerate(rows):
        r = start_row + i
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=j, value=val)
            cell.border = BORDER
            cell.font = Font(size=10, bold=(j == 1))
            halign = "left" if j == 1 or j in wrap_cols else "center"
            cell.alignment = Alignment(horizontal=halign, vertical="top",
                                       wrap_text=(j in wrap_cols or j == 1))
            if band and i % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=GREY)
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w


wb = Workbook()

# ----------------------------------------------------------------------
# 1. SUMMARY SCORECARD
# ----------------------------------------------------------------------
ws = wb.active
ws.title = "Summary Scorecard"
hdr = ["Metric", "GOOGL", "META", "SNAP", "PINS", "RDDT"]
rows = [
    ["Report date", "Jul 22 (AMC)", "Jul 29 (AMC)", "Jul 16 (AMC)", "early Aug (est.)", "Jul 29-30 (est.)"],
    ["Price (7/9/26 close)", "$358.89", "$631.48", "$4.70", "$22.70", "$200.31"],
    ["Q2'26 revenue consensus", "~$116.8B", "~$60.2-61.3B", "~$1.57B", "~$1.15B", "~$729M"],
    ["Company Q2 revenue guidance", "n/a (no guide)", "$58-61B", "$1.52-1.55B", "$1.133-1.153B", "$715-725M"],
    ["Consensus vs. guidance", "n/a", "At/above top of range", "ABOVE high end (risk)", "Upper half of range", "ABOVE high end (risk)"],
    ["Revenue growth YoY (cons.)", "+21%", "+27-29%", "+13-17%", "+15%", "+46%"],
    ["Q2'26 EPS consensus", "$2.86-2.89", "$7.17-7.32", "+$0.08 adj / -$0.12 GAAP", "$0.36 adj", "$0.96"],
    ["EPS growth YoY", "+24% (vs $2.31)", "~flat (vs $7.14)", "loss narrowing (vs -$0.16 GAAP)", "+9% (vs $0.33)", "+113% (vs $0.45)"],
    ["Q2 EBITDA guidance", "n/a", "n/a", "$175-200M", "$256-276M", "$285-295M"],
    ["FY26 revenue consensus", "~$488.4B (+21%)", "~$252.9B (+26%)", "~$6.69B (+13%)", "~$4.87B (+15%)", "~$3.23B (+47%)"],
    ["FY26 EPS consensus", "~$14.22", "~$30.0-32.3", "~$0.60 adj", "~$1.90", "~$4.97"],
    ["90-day EPS revision (Q2)", "+4.7% ($2.76 to $2.89)", "Rev revised +1.1% (3-mo)", "Mixed; PT cuts into print", "+16% ($0.31 to $0.36)", "+16% ($0.83 to $0.96, 60d)"],
    ["90-day EPS revision (FY26)", "+24% ($11.47 to $14.22)", "Higher post-Q1", "Up post-Q1 (cost cuts)", "+10% ($1.72 to $1.90)", "+21% ($4.12 to $4.97)"],
    ["Beat history (recent)", "13 of 15 EPS beats; median +7.5%", "Beat every recent qtr", "Rev in-line; EBITDA beats", "Missed adj EPS Q2'25 & Q3'25", "4 beats of +33% to +138%"],
    ["Analyst rating / avg PT", "Strong Buy / ~22% upside", "Strong Buy / $827 (+31%)", "Hold / $7.48 (30 of 43 Hold)", "Buy-skewed", "Mod. Buy / ~$225-233"],
    ["Primary upside source", "Cloud acceleration; AI ad yield (Gemini/AI Overviews)", "Ad pricing + AI targeting; conservative guide; Q3 outlook", "Cost cuts >$500M ann.; EBITDA beat; Snapchat+ subs +30%", "Performance+ adoption; 3P demand (Google/Magnite); margin leverage", "74% ad growth momentum; AI licensing renewal (WFC: ~$550M/yr vs $130M)"],
    ["Primary disappointment risk", "Capex fatigue / equity raise; AI Overviews click economics; antitrust", "Capex guide-up ($125-145B), FCF ~0; consensus above guide; EU regulation", "Rev miss vs consensus; Perplexity loss; Middle East; restructuring charges", "Retail/tariff ad softness; mid-year EPS miss pattern; ARPU dependence", "US DAU deceleration (AI answer engines); modest licensing terms; 50x P/E"],
    ["Setup skew", "Positive but crowded (+98% 12-mo)", "Positive; demanding bar", "Negative revisions; low expectations", "Constructive", "Best momentum; highest bar"],
]
add_title(ws, "2Q 2026 Earnings Preview - Digital Advertising & Social Media",
          "GOOGL, META, SNAP, PINS, RDDT | Estimates & revisions as of July 10, 2026 | See tabs for detail; companion write-up in Q2_2026_Earnings_Preview.md", 6)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 6)
fill_rows(ws, 5, rows, widths=[28, 30, 30, 30, 30, 30], wrap_cols={2, 3, 4, 5, 6})
for r in (9, 16, 20, 21):  # highlight rows: cons-vs-guide, revisions, drivers
    pass
# highlight the consensus-vs-guidance row and setup row
for c in range(2, 7):
    ws.cell(row=9, column=c).fill = PatternFill("solid", fgColor=YELLOW)
ws.cell(row=9, column=4).fill = PatternFill("solid", fgColor=RED)
ws.cell(row=9, column=6).fill = PatternFill("solid", fgColor=RED)
ws.freeze_panes = "B5"

# ----------------------------------------------------------------------
# 2. ESTIMATES DETAIL
# ----------------------------------------------------------------------
ws = wb.create_sheet("Estimates Detail")
hdr = ["Company", "Ticker", "Report Date", "Q2 Rev Consensus", "Q2 Rev Low", "Q2 Rev High",
       "# Rev Analysts", "Rev YoY", "Company Rev Guidance", "Q2 EPS Consensus", "Q2 EPS Low",
       "Q2 EPS High", "# EPS Analysts", "Year-Ago EPS", "Q2 EBITDA Guidance",
       "FY26 Rev Consensus", "FY26 Rev YoY", "FY26 EPS Consensus", "FY27 EPS Consensus"]
rows = [
    ["Alphabet", "GOOGL", "Jul 22, 2026", "$116.82B", "$113.62B", "$120.13B", 37, "+21.1%",
     "n/a", "$2.89", "$2.59", "$3.60", 42, "$2.31", "n/a", "$488.37B", "+21.2%", "$14.22", "$14.54"],
    ["Meta Platforms", "META", "Jul 29, 2026", "$60.2-61.3B", "n/a", "n/a", "n/a", "+27-29%",
     "$58.0-61.0B (~2% FX tailwind)", "$7.17-7.32", "$6.47", "$7.79", "n/a", "$7.14",
     "n/a (FY exp guide $162-169B)", "$252.89B", "+25.8%", "$30.0-32.3", "$36.8"],
    ["Snap", "SNAP", "Jul 16, 2026", "$1.572B", "n/a", "n/a", "~40-50", "+16.9% (cons) / +11-13% (guide)",
     "$1.52-1.55B (ex-Perplexity; MidEast headwind)", "+$0.08 adj / -$0.12 GAAP", "n/a", "n/a", "n/a", "-$0.16 GAAP",
     "$175-200M", "$6.69B", "+12.9%", "$0.60 adj", "$0.74 adj"],
    ["Pinterest", "PINS", "early Aug (est.)", "$1.15B", "$1.14B", "$1.16B", 30, "+15.0%",
     "$1.133-1.153B (+14-16%; ~1pt FX tailwind)", "$0.36 adj", "$0.31", "$0.42", 19, "$0.33",
     "$256-276M", "$4.87B", "+15.2%", "$1.90", "$2.24"],
    ["Reddit", "RDDT", "Jul 29-30 (est.)", "$729.3M", "$720.0M", "$782.8M", 26, "+46.0%",
     "$715-725M", "$0.96", "$0.80", "$1.26", 21, "$0.45",
     "$285-295M", "$3.23B", "+46.6%", "$4.97", "$6.47"],
]
add_title(ws, "Q2 2026 Consensus Estimates vs. Company Guidance",
          "Sources: Yahoo Finance, Benzinga, MarketBeat, ChartMill, TipRanks (Jul 10, 2026). Consensus varies by aggregator; ranges shown where sources disagree.", 19)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 19)
fill_rows(ws, 5, rows, widths=[16, 8, 15, 15, 12, 12, 10, 16, 30, 18, 10, 10, 10, 12, 22, 14, 10, 14, 14],
          wrap_cols={9, 15})
ws.freeze_panes = "C5"

# ----------------------------------------------------------------------
# 3. ESTIMATE REVISIONS
# ----------------------------------------------------------------------
ws = wb.create_sheet("Estimate Revisions")
hdr = ["Ticker", "Q2 EPS Now", "Q2 EPS 30d Ago", "Q2 EPS 60d Ago", "Q2 EPS 90d Ago",
       "Q2 EPS 90d Chg", "FY26 EPS Now", "FY26 EPS 90d Ago", "FY26 EPS 90d Chg",
       "Revisions Up (30d)", "Revisions Down (30d)", "Revision Commentary"]
rows = [
    ["GOOGL", 2.89, 2.87, 2.88, 2.76, "+4.7%", 14.22, 11.47, "+24.0%", 2, 4,
     "Huge post-Q1 reset (Q1 EPS $5.11 vs $2.67 est). FY26 up 7 vs down 4 over 30d; near-term Q2 drift mildly negative as analysts fine-tune. 13 of 15 qtrs beat."],
    ["META", "7.17-7.32", "~7.3", "~7.2", "~7.1", "Rev +1.1% (3-mo)", "30.0-32.3", "~29-30", "Higher post-Q1", "n/a", "n/a",
     "Q2 revenue consensus revised up ~1.1% in past 3 months; FY26 revenue to ~$253B (+26%). EPS growth compressed to ~+9% and FY26 FCF ~breakeven on $125-145B capex guide (raised from $120-135B)."],
    ["SNAP", "0.08 adj", "n/a", "n/a", "n/a", "Mixed", "0.60 adj", "n/a", "Up post-Q1", 8, "n/a",
     "8 analysts revised earnings UP post-Q1 (cost-cut driven: >$500M annualized reductions in H2'26). BUT pre-print PT cuts: GS $7->$6 (Jul 8), WFC $7->$5 (Jul 7), DADavidson init Hold/$5. Consensus rev $1.57B sits ABOVE $1.52-1.55B guide."],
    ["PINS", 0.36, 0.36, 0.35, 0.31, "+16.1%", 1.90, 1.72, "+10.5%", 0, 1,
     "Steady upward march post-Q1 blowout (rev $1.008B vs $966M est; stock +17%). Slight negative drift last 30d (1 down / 0 up on Q2) but stable. Caution: missed adj EPS in Q2'25 (-6%) and Q3'25 (-9%)."],
    ["RDDT", 0.96, 0.96, 0.83, 0.83, "+15.7% (60d)", 4.97, 4.12, "+20.6%", 1, 1,
     "Steepest upward revisions in group after +79% Q1 EPS surprise. FY26 EPS $4.12->$4.97 in 90d; FY27 $5.75->$6.47. 4 up / 1 down on FY over 30d. Avg PT trimmed -3.4% in 3 months (stock volatility, AI-traffic debate)."],
]
add_title(ws, "Estimate Revision Trends (as of July 10, 2026)",
          "EPS estimate progression per Yahoo Finance revision tables, ChartMill, InvestingPro; up/down = number of analysts revising in last 30 days.", 12)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 12)
fill_rows(ws, 5, rows, widths=[9, 11, 12, 12, 12, 12, 12, 13, 12, 11, 12, 80], wrap_cols={12})
# color the revision-change columns
for i, tk in enumerate(["GOOGL", "META", "SNAP", "PINS", "RDDT"]):
    r = 5 + i
    for col, positive in ((6, None), (9, None)):
        val = str(ws.cell(row=r, column=col).value)
        if val.startswith("+") or "Higher" in val or "Up" in val:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=GREEN)
        elif "Mixed" in val:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=YELLOW)
ws.freeze_panes = "B5"

# ----------------------------------------------------------------------
# 4. BEAT-MISS HISTORY
# ----------------------------------------------------------------------
ws = wb.create_sheet("Beat-Miss History")
hdr = ["Ticker", "Quarter", "EPS Est.", "EPS Actual", "EPS Surprise %", "Rev Est.", "Rev Actual", "Rev Surprise %"]
rows = [
    ["GOOGL", "Q1'26 (Mar)", "$2.63-2.67", "$5.11", "+92-94%", "$104.1B", "$109.9B", "+5.6%"],
    ["GOOGL", "Q4'25 (Dec)", "$2.62-2.64", "$2.82", "+7.0-7.6%", "$104.8B", "$113.8B", "+8.7%"],
    ["GOOGL", "Q3'25 (Sep)", "$2.26-2.32", "$2.87", "+23-27%", "$94.1B", "$102.4B", "+8.8%"],
    ["GOOGL", "Q2'25 (Jun)", "$2.16-2.19", "$2.31", "+5.7-6.9%", "$87.9B", "$96.4B", "+9.7%"],
    ["META", "Q1'26 (Mar)", "$6.66-6.70", "$10.44", "+56% (incl. tax benefit)", "$55.6B", "$56.31B", "+1.4%"],
    ["META", "Q4'25 (Dec)", "$8.16-8.21", "$8.88", "+8.8%", "$58.3B", "$59.89B", "+2.7%"],
    ["META", "Q3'25 (Sep)", "$6.72-6.74", "$7.25 adj ($1.05 GAAP)", "+7.6% adj", "$49.3B", "$51.24B", "+3.9%"],
    ["META", "Q2'25 (Jun)", "$5.75-5.88", "$7.14", "+21-24%", "$44.6B", "$47.52B", "+6.7%"],
    ["SNAP", "Q1'26 (Mar)", "-$0.07 GAAP", "-$0.05 GAAP", "beat; EBITDA $233M vs $215M (+8.6%)", "~$1.53B", "$1.529B", "in line"],
    ["SNAP", "Q4'25 (Dec)", "-$0.03", "+$0.03", "beat", "n/a", "n/a", "n/a"],
    ["SNAP", "Q3'25 (Sep)", "-$0.12", "-$0.06", "beat", "n/a", "n/a", "n/a"],
    ["SNAP", "Q2'25 (Jun)", "-$0.15", "-$0.16", "MISS", "n/a", "n/a", "n/a"],
    ["PINS", "Q1'26 (Mar)", "$0.22-0.23", "$0.27", "+25%", "$966M", "$1.008B", "+4.3%"],
    ["PINS", "Q4'25 (Dec)", "$0.67", "$0.67", "-0.7% (in line)", "n/a", "n/a", "n/a"],
    ["PINS", "Q3'25 (Sep)", "$0.42", "$0.38", "-8.8% MISS", "n/a", "n/a", "n/a"],
    ["PINS", "Q2'25 (Jun)", "$0.35", "$0.33", "-6.2% MISS", "n/a", "n/a", "n/a"],
    ["RDDT", "Q1'26 (Mar)", "$0.56-0.62", "$1.01", "+79%", "$608-611M", "$663.4M", "+8.6-9.2%"],
    ["RDDT", "Q4'25 (Dec)", "$0.93", "$1.24", "+33%", "n/a", "n/a", "n/a"],
    ["RDDT", "Q3'25 (Sep)", "$0.52", "$0.80", "+54%", "n/a", "n/a", "n/a"],
    ["RDDT", "Q2'25 (Jun)", "$0.19", "$0.45", "+138%", "n/a", "n/a", "n/a"],
]
add_title(ws, "Beat / Miss History - Last Four Reported Quarters",
          "Per MarketBeat, Yahoo Finance, TipRanks, Benzinga. GAAP vs adjusted conventions vary by source; large GOOGL/META Q1'26 beats include one-time items.", 8)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 8)
fill_rows(ws, 5, rows, widths=[9, 13, 13, 20, 26, 12, 12, 14], wrap_cols={5}, band=False)
for i, row in enumerate(rows):
    r = 5 + i
    surprise = str(row[4])
    fill = RED if "MISS" in surprise else GREEN if ("+" in surprise or "beat" in surprise) else GREY
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=fill if c == 5 else (GREY if i % 2 else "FFFFFF"))
ws.freeze_panes = "C5"

# ----------------------------------------------------------------------
# 5. UPSIDE-DOWNSIDE DRIVERS
# ----------------------------------------------------------------------
ws = wb.create_sheet("Upside-Downside Drivers")
hdr = ["Ticker", "Sources of Upside", "Sources of Disappointment", "Key Things to Watch on the Call"]
rows = [
    ["GOOGL",
     "1) Google Cloud re-acceleration (Q1 +63%, >$20B/qtr) - biggest swing factor. "
     "2) Search resilience: AI Overviews / Gemini lifting ad conversions with no paid-click cannibalization. "
     "3) YouTube strength; opex discipline driving double-digit OI growth. "
     "4) 13-of-15 beat history; revenue beats of +5.6% to +9.7% last 4 qtrs.",
     "1) Capex escalation / financing fatigue - stock already eased after equity raise for AI buildout; rising depreciation drag. "
     "2) Any sign AI answers are absorbing monetizable queries (paid click decel). "
     "3) Antitrust / remedies headlines. "
     "4) Positioning: +98% 12-mo run, 43 of 53 Strong Buy - in-line may not be enough.",
     "Cloud growth rate & backlog; capex guide; paid clicks / cost-per-click; AI Overviews monetization metrics; buyback/equity commentary."],
    ["META",
     "1) Ad pricing + AI targeting (Advantage+) momentum - Q1 +33% growth; ~2% FX tailwind embedded in guide. "
     "2) History of conservative guides -> beat-and-raise mechanics. "
     "3) Unchanged FY expense guide ($162-169B) + May headcount cuts = opex leverage. "
     "4) Q3 revenue guide above consensus is the bull trigger.",
     "1) Another capex guide-up (current $125-145B) without ROI proof; FY26 consensus FCF ~breakeven. "
     "2) Consensus rev at/above top of $58-61B guide - no room for in-line. "
     "3) EU regulatory friction on data privacy / AI content; KIDS Act. "
     "4) Operating margin slipping below ~40% as infra costs ramp.",
     "Q3 revenue guidance vs ~$64.4B consensus; capex trajectory & FY27 framing; depreciation commentary; ad price vs impression mix; AI assistant usage disclosures."],
    ["SNAP",
     "1) EBITDA/cost execution: Q1 EBITDA +116% YoY, beat by 8.6%; >$500M annualized cost cuts start flowing H2; beat above $200M guide-top is achievable. "
     "2) Snapchat+ subscriptions compounding >30% - highest-quality line. "
     "3) Middle East stabilization vs guide assumption of persistent headwind. "
     "4) Expectations rock-bottom: Hold-rated, $4.70 stock.",
     "1) Consensus $1.57B ABOVE the $1.52-1.55B guide: printing in-range = optical miss. "
     "2) Perplexity revenue (high-margin) removed from 2026. "
     "3) Brand budget share loss to Meta/TikTok; MidEast geopolitics flagged as uncertain. "
     "4) $95-130M restructuring charges hit mostly in Q2 (ugly GAAP); fresh PT cuts (GS, WFC) into the print.",
     "Revenue vs $1.52-1.55B guide; adj EBITDA vs $175-200M; DAU trajectory & NA engagement; subscription revenue run-rate; Q3 guide incl. cost-cut cadence."],
    ["PINS",
     "1) Performance+ / AI-driven search & recs lifting relevance and ARPU (Q1 ARPU +6%). "
     "2) Third-party demand (Google, Magnite) monetizing international users - largest untapped ARPU pool. "
     "3) Margin leverage from Jan restructuring (~15% workforce cut). "
     "4) Guide already de-risked for tariff/consumer noise per CFO; consensus achievable at upper half of range.",
     "1) Retail/e-commerce ad sensitivity to tariffs and consumer softness (Asia e-comm advertisers a past swing factor). "
     "2) Mature US MAUs -> growth increasingly ARPU-dependent. "
     "3) Pattern risk: missed adj EPS in Q2'25 and Q3'25 - mid-year quarters were the culprits. "
     "4) GAAP optics (restructuring) and EBITDA margin vs AI infra costs.",
     "Revenue vs $1.133-1.153B; adj EBITDA vs $256-276M; MAU (631M record in Q1) & intl ARPU; 3P demand contribution; Q3 guide."],
    ["RDDT",
     "1) Ad momentum: Q1 ads +74% on impressions + pricing + full-funnel demand; intl ARPU still fraction of US. "
     "2) AI data-licensing renewals - THE 2026 catalyst: Google (~$60M/yr) + OpenAI (~$70M/yr) deals reprice this year; WFC models ~$550M/yr combined (~4x) at 90%+ gross margin. "
     "3) Most-cited domain in AI Overviews/Perplexity = negotiating leverage. "
     "4) Street chronically under-models: beats of +33% to +138% last 4 qtrs.",
     "1) US DAU deceleration - AI answer engines synthesize Reddit content without the click; ~60% of traffic is organic search; WFC pegs ~$16B EV overhang. WATCH US DAU FIRST. "
     "2) Licensing renewal at modest terms (~$130M status quo) invalidates the bull catalyst (RBC: 5-8x unlikely). "
     "3) Meta Forum app competition (-6% day reaction in May); KIDS Act. "
     "4) Consensus ~$729M ABOVE $715-725M guide; ~50x P/E leaves no cushion.",
     "US vs intl DAU split (Q1: 126.8M total, +17%); licensing renewal terms/commentary; ad revenue growth rate vs +74% in Q1; Q3 guide; 'other revenue' line for licensing run-rate."],
]
add_title(ws, "Sources of Upside vs. Disappointment - Q2 2026 Prints",
          "Synthesized from company guidance, sell-side commentary (Wells Fargo, Goldman, RBC, BofA as reported), and Q1'26 results.", 4)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 4)
fill_rows(ws, 5, rows, widths=[9, 60, 60, 45], wrap_cols={2, 3, 4})
for i in range(len(rows)):
    ws.row_dimensions[5 + i].height = 120
    ws.cell(row=5 + i, column=2).fill = PatternFill("solid", fgColor=GREEN)
    ws.cell(row=5 + i, column=3).fill = PatternFill("solid", fgColor=RED)
ws.freeze_panes = "B5"

# ----------------------------------------------------------------------
# 6. CALENDAR & MACRO
# ----------------------------------------------------------------------
ws = wb.create_sheet("Calendar & Macro")
hdr = ["Date (2026)", "Event", "Relevance to Group"]
rows = [
    ["Jul 14-15", "Big banks kick off Q2 season; June CPI (Jul 15)", "Macro tone-setter; May CPI was 4.2% - ad budgets sensitive to rate path"],
    ["Jul 16", "SNAP reports (AMC); June PPI", "First read on social ad spend; sets tone for PINS/RDDT"],
    ["Jul 17", "Netflix (AMC); retail sales", "CTV ad commentary; consumer health read for retail ad vertical"],
    ["Jul 22", "GOOGL reports (AMC)", "Search/YouTube = broadest ad-demand signal; Cloud + capex read-through to all AI names"],
    ["Jul 29", "META reports (AMC); RDDT expected; FOMC decision", "Heaviest day: Meta Q3 guide + capex; Reddit DAU/licensing; Fed backdrop"],
    ["Jul 30", "Alt. RDDT date (some services); AAPL", "Reddit print timing varies by source (Jul 29 vs Jul 30)"],
    ["early Aug", "PINS expected (reported May 4 last qtr)", "Last of group; retail/shopping ad read"],
    ["Ongoing", "KIDS Act in Senate", "Sector-wide regulatory overhang; compliance cost risk - most cushion at RDDT (40% EBITDA margin), least at SNAP"],
    ["2026", "Reddit Google/OpenAI licensing renewals", "Group-relevant precedent for AI content licensing economics"],
    ["Context", "Hyperscaler AI capex >$220B trailing 4 qtrs", "'AI revenue / AI capex' is the season's most-watched ratio (GOOGL, META)"],
    ["Context", "Tech sector Q2 cons.: EPS +16.8%, revenue +11.2% (revised UP)", "Group enters season with positive sector-level revisions"],
]
add_title(ws, "Q2 2026 Reporting Calendar & Macro Context", "Report dates per company IR / aggregators; some are estimates and may shift.", 3)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 3)
fill_rows(ws, 5, rows, widths=[14, 50, 75], wrap_cols={2, 3})

# ----------------------------------------------------------------------
# 7. SOURCES & NOTES
# ----------------------------------------------------------------------
ws = wb.create_sheet("Sources & Notes")
hdr = ["#", "Source", "Used For"]
rows = [
    [1, "Yahoo Finance analyst estimates & revision tables (GOOGL, PINS, RDDT)", "Consensus, ranges, analyst counts, 7/30/60/90-day EPS trends, up/down revision counts"],
    [2, "Benzinga earnings pages (GOOG, RDDT)", "Q2 estimates, beat/miss history"],
    [3, "MarketBeat earnings pages (GOOGL, META, RDDT)", "Report dates, guidance vs consensus, historical surprises"],
    [4, "ChartMill analyst ratings (META, SNAP, RDDT)", "Alternate consensus figures, 3-month revenue revision (META +1.1%), rating actions"],
    [5, "TipRanks (META, SNAP)", "EPS forecasts, guidance summaries (Meta capex $125-145B; Snap cost/restructuring details)"],
    [6, "stockanalysis.com forecasts (META, SNAP)", "Price targets, rating distributions, FY26/27 estimates, pre-print PT cuts (GS/WFC/DADavidson on SNAP)"],
    [7, "Company sources: Reddit IR Q1'26 release; Pinterest 8-K/press; Snap Q1'26 call transcript (Investing.com)", "Official Q2 guidance ranges (revenue, EBITDA), management framing"],
    [8, "CNBC, StreetInsider, GuruFocus/TradingView, 24/7 Wall St, TIKR, tickeron", "Q1'26 recaps, guidance-vs-consensus at time of guide, stock reactions"],
    [9, "Wells Fargo (via TIKR/AL Capital summaries), Goldman, RBC as reported", "RDDT licensing renewal scenario (~$550M/yr), $16B EV traffic overhang, SNAP PT cuts"],
    [10, "StocksAnalyzer / ClaritX / WallStreet.AI season previews", "Macro backdrop, sector-level revision tone, calendar, AI-capex framing"],
]
add_title(ws, "Sources & Methodology Notes", "All data retrieved July 10, 2026.", 3)
ws.append([])
ws.append(hdr)
style_header(ws, 4, 3)
fill_rows(ws, 5, rows, widths=[5, 70, 80], wrap_cols={2, 3})
notes = [
    "NOTES:",
    "- Consensus figures differ across aggregators (LSEG vs FactSet vs Zacks pools; GAAP vs adjusted EPS conventions). Ranges are shown where sources disagree materially.",
    "- SNAP EPS is especially convention-dependent: ~+$0.08 adjusted consensus vs ~-$0.12 GAAP; management framed ~$0.09 for Q2.",
    "- META Q2'26 EPS is optically ~flat YoY because Q2'25 ($7.14) was a large beat; Q1'26's $10.44 included a significant tax benefit.",
    "- GOOGL Q1'26 EPS of $5.11 (vs $2.67 est) included one-time items; underlying FY26 consensus is ~$14.22.",
    "- PINS and RDDT report dates are estimates based on prior-year cadence; RDDT shown as Jul 29 or Jul 30 depending on the service.",
    "- This workbook is an informational summary compiled from public web sources, not investment advice.",
]
r = 5 + len(rows) + 2
for i, n in enumerate(notes):
    c = ws.cell(row=r + i, column=1, value=n)
    c.font = Font(size=9, italic=True, bold=(i == 0))
    ws.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=3)

wb.save("/workspace/Q2_2026_Earnings_Preview.xlsx")
print("Workbook saved.")
