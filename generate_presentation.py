"""Generate 'Sources of Alpha & the Bloomberg Terminal' PowerPoint deck.

Usage:
    python generate_presentation.py

Produces: alpha_sources_bloomberg.pptx (16:9, dark Bloomberg-terminal theme).

Text markup: segments wrapped in backticks (e.g. `EQS <GO>`) are rendered
as orange monospace runs, mimicking terminal function mnemonics.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette --
BG = RGBColor(0x0D, 0x11, 0x17)          # near-black background
PANEL = RGBColor(0x16, 0x1D, 0x27)       # card / panel fill
PANEL_EDGE = RGBColor(0x2A, 0x35, 0x44)  # card border
ORANGE = RGBColor(0xFF, 0x8C, 0x1A)      # Bloomberg-style amber/orange
ORANGE_DIM = RGBColor(0xC9, 0x6A, 0x0A)
WHITE = RGBColor(0xF2, 0xF4, 0xF7)
GREY = RGBColor(0xA8, 0xB2, 0xBF)
GREY_DIM = RGBColor(0x66, 0x71, 0x80)
CYAN = RGBColor(0x4F, 0xC3, 0xF7)        # secondary accent

BODY_FONT = "Calibri"
MONO_FONT = "Consolas"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

_slide_no = 0


# ---------------------------------------------------------------- helpers --
def add_slide():
    global _slide_no
    slide = prs.slides.add_slide(BLANK)
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG
    _slide_no += 1
    return slide


def txbox(slide, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    return box, tf


def set_runs(paragraph, text, size, color=WHITE, bold=False, mono_color=ORANGE):
    """Write text into paragraph; `code` segments become orange monospace."""
    parts = text.split("`")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run()
        run.text = part
        f = run.font
        f.size = Pt(size)
        if i % 2 == 1:  # inside backticks
            f.name = MONO_FONT
            f.color.rgb = mono_color
            f.bold = True
        else:
            f.name = BODY_FONT
            f.color.rgb = color
            f.bold = bold


def footer(slide, section):
    _, tf = txbox(slide, Inches(0.55), Inches(7.02), Inches(9.5), Inches(0.35))
    p = tf.paragraphs[0]
    set_runs(p, f"Sources of Alpha  |  {section}", 10, color=GREY_DIM)
    _, tf2 = txbox(slide, Inches(12.35), Inches(7.02), Inches(0.6), Inches(0.35))
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    set_runs(p2, f"{_slide_no:02d}", 10, color=GREY_DIM)


def accent_bar(slide, top=Inches(1.28), left=Inches(0.55), width=Inches(0.85)):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(3.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ORANGE
    bar.line.fill.background()
    bar.shadow.inherit = False


def title_block(slide, kicker, title):
    _, tf = txbox(slide, Inches(0.55), Inches(0.32), Inches(12.2), Inches(0.35))
    p = tf.paragraphs[0]
    set_runs(p, kicker.upper(), 12, color=ORANGE, bold=True)
    _, tf2 = txbox(slide, Inches(0.5), Inches(0.60), Inches(12.3), Inches(0.75))
    p2 = tf2.paragraphs[0]
    set_runs(p2, title, 30, color=WHITE, bold=True)
    accent_bar(slide)


def bullets(slide, items, left=Inches(0.6), top=Inches(1.62),
            width=Inches(12.1), height=Inches(5.2), size=15.5,
            space_after=9):
    """items: list of (text, level). Level 0 = bullet, 1 = sub-bullet,
    'H' = small orange sub-heading."""
    _, tf = txbox(slide, left, top, width, height)
    first = True
    for text, level in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(space_after)
        if level == "H":
            set_runs(p, text.upper(), size - 3, color=ORANGE, bold=True)
            p.space_before = Pt(6)
            continue
        prefix = "\u25aa  " if level == 0 else "\u2013  "
        indent_run = p.add_run()
        indent_run.text = ("      " * level) + prefix
        indent_run.font.size = Pt(size if level == 0 else size - 1.5)
        indent_run.font.color.rgb = ORANGE if level == 0 else GREY
        indent_run.font.name = BODY_FONT
        set_runs(p, text, size if level == 0 else size - 1.5,
                 color=WHITE if level == 0 else GREY)
    return tf


def card(slide, left, top, width, height, heading, body,
         head_color=ORANGE, body_size=11.5):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 left, top, width, height)
    shp.adjustments[0] = 0.055
    shp.fill.solid()
    shp.fill.fore_color.rgb = PANEL
    shp.line.color.rgb = PANEL_EDGE
    shp.line.width = Pt(1)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.14)
    tf.margin_right = Inches(0.14)
    tf.margin_top = Inches(0.10)
    tf.margin_bottom = Inches(0.08)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    set_runs(p, heading, 13.5, color=head_color, bold=True)
    p.space_after = Pt(4)
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    set_runs(p2, body, body_size, color=GREY)
    return shp


# ================================================================ SLIDE 1 ==
s = add_slide()
# thin top rule + terminal-style header
strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Pt(5))
strip.fill.solid(); strip.fill.fore_color.rgb = ORANGE
strip.line.fill.background(); strip.shadow.inherit = False

_, tf = txbox(s, Inches(0.75), Inches(0.75), Inches(11.8), Inches(0.4))
set_runs(tf.paragraphs[0], "`<HELP> ALPHA <GO>`", 15)

_, tf = txbox(s, Inches(0.7), Inches(2.15), Inches(11.9), Inches(1.9))
p = tf.paragraphs[0]
set_runs(p, "Sources of Alpha in the Market", 46, color=WHITE, bold=True)
p2 = tf.add_paragraph()
set_runs(p2, "\u2014 and how to discover them with the Bloomberg Terminal",
         26, color=ORANGE, bold=True)
p2.space_before = Pt(8)

_, tf = txbox(s, Inches(0.75), Inches(4.45), Inches(11.5), Inches(1.1))
p = tf.paragraphs[0]
set_runs(p, "A practitioner's map of where excess returns come from \u2014 "
            "factors, fundamentals, events, macro, carry, flows, sentiment "
            "and statistical arbitrage \u2014 with the terminal workflow to "
            "find, validate and monitor each one.", 15, color=GREY)

_, tf = txbox(s, Inches(0.75), Inches(6.35), Inches(11.5), Inches(0.5))
set_runs(tf.paragraphs[0],
         "Investment Research  \u00b7  Equities | Macro | Cross-Asset", 12,
         color=GREY_DIM)

# ================================================================ SLIDE 2 ==
s = add_slide()
title_block(s, "Foundations", "Alpha 101: What We Are Actually Hunting")
bullets(s, [
    ("Alpha is the slice of return NOT explained by market exposure \u2014 "
     "beta is cheap and abundant, alpha is scarce and roughly zero-sum "
     "after fees and costs.", 0),
    ("Jensen's alpha:  \u03b1 = Rp \u2212 [ Rf + \u03b2p \u00d7 (Rm \u2212 Rf) ]  "
     "\u2014 always risk-adjust before claiming skill.", 0),
    ("Much \u201calpha\u201d is hidden beta: static factor tilts, leverage, or "
     "illiquidity dressed up as stock-picking.", 0),
    ("Durable alpha requires a repeatable edge \u2014 informational, "
     "analytical, behavioral, or structural (next slide).", 0),
    ("On the Terminal", "H"),
    ("`PORT <GO>` \u2014 attribution decomposes returns into benchmark, factor "
     "tilts and true idiosyncratic selection.", 1),
    ("`BETA <GO>` and `HRA <GO>` \u2014 regress any security or portfolio "
     "series against a benchmark to isolate \u03b1 and \u03b2.", 1),
], top=Inches(1.65))
footer(s, "Foundations")

# ================================================================ SLIDE 3 ==
s = add_slide()
title_block(s, "Framework", "Four Edges, Eight Hunting Grounds")
_, tf = txbox(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.75))
p = tf.paragraphs[0]
set_runs(p, "Every persistent alpha source maps to an edge:  "
            "`Informational` (know it first)  \u00b7  `Analytical` (process it better)  \u00b7  "
            "`Behavioral` (exploit biases)  \u00b7  `Structural` (harvest others' constraints).",
         13.5, color=GREY, mono_color=CYAN)

grid = [
    ("1 \u00b7 Factor & Style Premia", "Value, momentum, quality, low-vol, size \u2014 systematic and harvestable."),
    ("2 \u00b7 Fundamental Selection", "Variant perception on earnings power vs. what consensus has priced."),
    ("3 \u00b7 Events & Special Situations", "M&A arbitrage, spin-offs, index adds/deletes, buybacks, restructurings."),
    ("4 \u00b7 Macro & Cross-Asset", "Rates, inflation, FX and growth surprises vs. what markets have discounted."),
    ("5 \u00b7 Carry & Term Structure", "FX carry, commodity roll yield, bond rolldown, volatility risk premium."),
    ("6 \u00b7 Flow & Positioning", "Crowding, short squeezes, forced sellers, ownership concentration."),
    ("7 \u00b7 Sentiment & Alt Data", "News momentum, social velocity, supply chains, web & transaction exhaust."),
    ("8 \u00b7 Statistical Arbitrage", "Mean reversion, pairs, cross-sectional signals \u2014 breadth over conviction."),
]
cw, ch = Inches(3.05), Inches(1.95)
gx, gy = Inches(0.6), Inches(2.45)
gap = Inches(0.12)
for idx, (h, b) in enumerate(grid):
    row, col = divmod(idx, 4)
    card(s,
         Emu(int(gx) + col * int(cw + gap)),
         Emu(int(gy) + row * int(ch + gap)),
         cw, ch, h, b)
footer(s, "Framework")

# ================================================================ SLIDE 4 ==
s = add_slide()
title_block(s, "Alpha Source 1", "Factor & Style Premia")
bullets(s, [
    ("The classics: Value (cheapness), Momentum (winners persist), Quality "
     "(profitability, low accruals), Low Volatility, Size.", 0),
    ("Why they pay: part compensation for real risk, part persistent "
     "behavioral bias (over-extrapolation, lottery-seeking, anchoring).", 0),
    ("The catch: factors are cyclical and get crowded \u2014 the entry "
     "valuation of the factor itself matters.", 0),
    ("Discover it on the Terminal", "H"),
    ("`EQS <GO>` \u2014 build multi-factor screens (e.g. cheap + high ROIC + "
     "positive 12-1M momentum) across any universe.", 1),
    ("`EQBT <GO>` \u2014 backtest an `EQS` screen historically: hit rate, "
     "turnover, factor decay.", 1),
    ("`PORT <GO>` \u2014 measure your live factor exposures and how much "
     "return each tilt actually contributed.", 1),
    ("`FTW <GO>` \u2014 Factors To Watch: which styles are working, by "
     "region and sector, right now.", 1),
], top=Inches(1.6))
footer(s, "Factor & Style Premia")

# ================================================================ SLIDE 5 ==
s = add_slide()
title_block(s, "Alpha Source 2", "Fundamental Stock Selection")
bullets(s, [
    ("The alpha is the gap between your model and consensus \u2014 a variant "
     "perception on revenue, margins, capital allocation or duration of "
     "competitive advantage.", 0),
    ("Edge is analytical: same public data, better interpretation \u2014 "
     "unit economics, cohort math, channel checks, management quality.", 0),
    ("Discover it on the Terminal", "H"),
    ("`MODL <GO>` \u2014 consensus line-item estimates: see exactly what "
     "the street models, then attack the weakest assumption.", 1),
    ("`FA <GO>` \u2014 two decades of financials: margin bridges, segment "
     "data, cash conversion, capital returns.", 1),
    ("`EE <GO>` / `SURP <GO>` \u2014 estimate revision trends and earnings "
     "surprise history: revisions drift, surprises cluster.", 1),
    ("`ANR <GO>` \u2014 analyst recs and price targets: measure how "
     "contrarian your view really is.", 1),
    ("`RV <GO>` \u2014 relative valuation vs. custom peer sets, through "
     "time, not just spot.", 1),
    ("`BI <GO>` \u2014 Bloomberg Intelligence primers to get to industry "
     "fluency fast.", 1),
], top=Inches(1.6), size=14.5, space_after=7)
footer(s, "Fundamental Selection")

# ================================================================ SLIDE 6 ==
s = add_slide()
title_block(s, "Alpha Source 3", "Event-Driven & Special Situations")
bullets(s, [
    ("Merger arbitrage: earn the spread between market price and deal "
     "terms; alpha = better handicapping of regulatory and financing risk.", 0),
    ("Index rebalances, spin-offs, forced conversions: uneconomic, "
     "predictable flows from constrained investors create the mispricing.", 0),
    ("Earnings events: positioning + implied move vs. your estimate of the "
     "distribution.", 0),
    ("Discover it on the Terminal", "H"),
    ("`MA <GO>` \u2014 searchable deal database; `MARB <GO>` \u2014 live "
     "merger-arb spreads, annualized returns, deal timelines.", 1),
    ("`EVTS <GO>` / `CACS <GO>` \u2014 company event calendars and corporate "
     "actions (splits, tenders, spin-offs).", 1),
    ("`MEMB <GO>` \u2014 index membership; anticipate adds/deletes and the "
     "passive flow they force.", 1),
    ("`ECDR <GO>` \u2014 IPO/secondary calendar; `DVD <GO>` \u2014 dividend "
     "projections vs. announced.", 1),
], top=Inches(1.6), size=15, space_after=8)
footer(s, "Event-Driven")

# ================================================================ SLIDE 7 ==
s = add_slide()
title_block(s, "Alpha Source 4", "Macro & Cross-Asset")
bullets(s, [
    ("Alpha is forecasting the surprise, not the level: growth, inflation, "
     "and policy vs. what is already priced in.", 0),
    ("Express views where they are cheapest \u2014 rates, FX, equity "
     "indices, or vol \u2014 cross-asset choice is itself a source of edge.", 0),
    ("Discover it on the Terminal", "H"),
    ("`ECO <GO>` \u2014 economic calendar with consensus, actual, and "
     "surprise history; build a surprise-reaction playbook.", 1),
    ("`WIRP <GO>` / `MIPR <GO>` \u2014 market-implied central-bank paths: "
     "the hurdle your macro view must beat.", 1),
    ("`ECFC <GO>` \u2014 economist forecasts by country and quarter; find "
     "where consensus is stale.", 1),
    ("`FXFC <GO>` \u2014 FX forecasts vs. forwards; `GP <GO>` / `GC <GO>` "
     "\u2014 chart any curve or cross-asset relationship.", 1),
    ("`ECST <GO>` \u2014 deep historical macro statistics for building "
     "regime indicators.", 1),
], top=Inches(1.6))
footer(s, "Macro & Cross-Asset")

# ================================================================ SLIDE 8 ==
s = add_slide()
title_block(s, "Alpha Source 5", "Carry & Term Structure")
bullets(s, [
    ("Carry: get paid for holding \u2014 FX interest differentials, bond "
     "coupon + rolldown, commodity backwardation, selling richly-priced "
     "implied volatility.", 0),
    ("Why it persists: it is compensation for tail risk and for providing "
     "liquidity/insurance that others must buy.", 0),
    ("The discipline: carry strategies earn steadily and drawdown "
     "violently \u2014 sizing and diversification are the real edge.", 0),
    ("Discover it on the Terminal", "H"),
    ("`CCRV <GO>` \u2014 commodity futures curves: spot backwardation/"
     "contango and roll yield by contract.", 1),
    ("`FWCV <GO>` \u2014 projected forward rate curves; measure rolldown "
     "along any sovereign curve.", 1),
    ("`FXFA <GO>` \u2014 FX forwards vs. spot: the carry you capture, "
     "hedged or unhedged.", 1),
    ("`GV <GO>` / `SKEW <GO>` \u2014 implied vs. realized volatility and "
     "skew richness for vol-premium harvesting.", 1),
], top=Inches(1.6))
footer(s, "Carry & Term Structure")

# ================================================================ SLIDE 9 ==
s = add_slide()
title_block(s, "Alpha Source 6", "Flow, Positioning & Ownership")
bullets(s, [
    ("Prices move when constrained investors must trade \u2014 anticipate "
     "the flow, or at least never be its exit liquidity.", 0),
    ("Signals: crowded longs (fragility), high short interest + catalyst "
     "(squeeze), fund-flow momentum, dealer hedging pressure.", 0),
    ("Discover it on the Terminal", "H"),
    ("`HDS <GO>` / `OWN <GO>` \u2014 who holds it, concentration, and "
     "recent accumulation or distribution by holder type.", 1),
    ("`SI <GO>` \u2014 short interest, days-to-cover, borrow trends: "
     "squeeze fuel and sentiment in one screen.", 1),
    ("`OMON <GO>` \u2014 options chain: open interest and skew reveal how "
     "the market is positioned around strikes and events.", 1),
    ("`CFTC <GO>` \u2014 Commitments of Traders: speculative positioning "
     "extremes in futures markets.", 1),
    ("`FLNG <GO>` \u2014 13F/13D filings aggregated: follow (or fade) "
     "hedge-fund clustering.", 1),
], top=Inches(1.6))
footer(s, "Flow & Positioning")

# =============================================================== SLIDE 10 ==
s = add_slide()
title_block(s, "Alpha Source 7", "Sentiment, News & Alternative Data")
bullets(s, [
    ("Informational edge today is mostly speed + coverage: machine-readable "
     "news, social velocity, supply-chain links, web and transaction "
     "exhaust.", 0),
    ("Sentiment works both ways: news momentum short-term, contrarian "
     "fading at extremes \u2014 define your horizon first.", 0),
    ("Discover it on the Terminal", "H"),
    ("`NSE <GO>` \u2014 advanced news search: filter by topic, source and "
     "sentiment tags across millions of stories.", 1),
    ("`TREN <GO>` \u2014 news & social trend velocity: unusual chatter "
     "before it hits the tape.", 1),
    ("`SPLC <GO>` \u2014 supply-chain map: trade a company's guidance "
     "through its suppliers and customers first.", 1),
    ("`BI <GO>` \u2014 analyst-curated industry data libraries (e.g. unit "
     "shipments, pricing trackers) updated continuously.", 1),
    ("Enterprise: Bloomberg Data License + event-driven feeds pipe "
     "point-in-time news sentiment into your own models.", 1),
], top=Inches(1.6))
footer(s, "Sentiment & Alt Data")

# =============================================================== SLIDE 11 ==
s = add_slide()
title_block(s, "Alpha Source 8", "Statistical Arbitrage & Quant Research")
bullets(s, [
    ("Edge = breadth: many small, diversified, risk-controlled bets \u2014 "
     "pairs mean-reversion, cross-sectional signal ranking, ML on "
     "fundamentals + alt data.", 0),
    ("The law of active management: IR \u2248 IC \u00d7 \u221abreadth \u2014 "
     "a tiny, repeatable signal beats a big, rare one.", 0),
    ("Discover it on the Terminal", "H"),
    ("`HRA <GO>` / `CORR <GO>` \u2014 regressions and correlation matrices "
     "for pair discovery and hedge ratios.", 1),
    ("`BQNT <GO>` \u2014 BQuant: hosted Python notebooks with full data "
     "access for signal research and backtesting.", 1),
    ("`BQL` \u2014 Bloomberg Query Language: screen, aggregate and compute "
     "point-in-time factors server-side.", 1),
    ("`FLDS <GO>` \u2014 discover the 30,000+ data fields available for "
     "any security \u2014 raw material for features.", 1),
    ("Excel `BDH/BDP/BDS` + `XLTP <GO>` templates for rapid prototyping "
     "before you productionize.", 1),
], top=Inches(1.6))
footer(s, "Stat Arb & Quant")

# =============================================================== SLIDE 12 ==
s = add_slide()
title_block(s, "Workflow", "From Idea to Alpha: The Terminal Pipeline")

steps = [
    ("DISCOVER", "Idea generation", "`BI` primers \u00b7 `NSE` news \u00b7 `FTW` factors \u00b7 `TREN` chatter"),
    ("SCREEN", "Narrow the universe", "`EQS` multi-factor screens \u00b7 `MA` deals \u00b7 `ECO` surprises"),
    ("ANALYZE", "Build the variant view", "`FA` financials \u00b7 `MODL` consensus \u00b7 `RV` valuation \u00b7 `SPLC`"),
    ("VALIDATE", "Prove it historically", "`EQBT` backtests \u00b7 `HRA` regressions \u00b7 `BQNT` research"),
    ("CONSTRUCT", "Size and hedge", "`PORT` optimizer \u00b7 factor-neutral weights \u00b7 `BTCA` cost model"),
    ("MONITOR", "Defend the alpha", "`ALRT` alerts \u00b7 `PORT` attribution \u00b7 `HDS`/`SI` crowding"),
]
cw2, ch2 = Inches(4.05), Inches(1.62)
gx2, gy2 = Inches(0.6), Inches(1.75)
gapx, gapy = Inches(0.14), Inches(0.16)
for idx, (stage, sub, fns) in enumerate(steps):
    row, col = divmod(idx, 3)
    left = Emu(int(gx2) + col * int(cw2 + gapx))
    top = Emu(int(gy2) + row * int(ch2 + gapy))
    shp = slide_card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          left, top, cw2, ch2)
    shp.adjustments[0] = 0.07
    shp.fill.solid(); shp.fill.fore_color.rgb = PANEL
    shp.line.color.rgb = PANEL_EDGE; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    set_runs(p, f"{idx + 1}  {stage}", 15, color=ORANGE, bold=True)
    p.space_after = Pt(2)
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    set_runs(p2, sub, 12, color=WHITE, bold=True)
    p2.space_after = Pt(3)
    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.LEFT
    set_runs(p3, fns, 11, color=GREY, mono_color=CYAN)

_, tf = txbox(s, Inches(0.6), Inches(5.55), Inches(12.1), Inches(1.1))
p = tf.paragraphs[0]
set_runs(p, "The loop matters more than any single step: alpha that is not "
            "re-validated and monitored for crowding is alpha you used to "
            "have.", 14, color=GREY)
footer(s, "Workflow")

# =============================================================== SLIDE 13 ==
s = add_slide()
title_block(s, "Reference", "Bloomberg Function Cheat Sheet")

rows = [
    ("Function", "What it shows", "Alpha use"),
    ("EQS / EQBT", "Equity screening + historical backtest", "Factor & fundamental idea generation, validation"),
    ("PORT", "Portfolio analytics, risk, attribution", "Separate real alpha from hidden factor beta"),
    ("FA / MODL / EE", "Financials, consensus models, estimates", "Find the street's weakest assumption"),
    ("SURP / ANR", "Earnings surprises, analyst recs", "Revision drift; measure your contrarianism"),
    ("MA / MARB / CACS", "Deals, merger-arb spreads, corp actions", "Event-driven and special situations"),
    ("ECO / WIRP / ECFC", "Macro calendar, implied policy, forecasts", "Trade the surprise, not the level"),
    ("CCRV / FWCV / FXFA", "Commodity, rates and FX curves", "Carry, roll yield and rolldown"),
    ("HDS / SI / OMON / CFTC", "Holders, shorts, options, COT", "Crowding, squeezes, positioning extremes"),
    ("NSE / TREN / SPLC", "News search, trends, supply chains", "Sentiment and information-chain alpha"),
    ("HRA / CORR / BQNT / BQL", "Regression, correlation, Python research", "Stat-arb signals and systematic backtests"),
    ("ALRT / BTCA", "Alerts, transaction cost analysis", "Monitor theses; keep paper alpha net of costs"),
]
tbl_shape = s.shapes.add_table(len(rows), 3, Inches(0.6), Inches(1.6),
                               Inches(12.1), Inches(5.1))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(2.9)
tbl.columns[1].width = Inches(4.3)
tbl.columns[2].width = Inches(4.9)

# strip default table style banding
tbl_pr = tbl_shape._element.graphic.graphicData.tbl
tbl_pr[0][-1].text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # no-style grid

for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.fill.solid()
        if r == 0:
            cell.fill.fore_color.rgb = RGBColor(0x25, 0x1C, 0x0E)
        else:
            cell.fill.fore_color.rgb = PANEL if r % 2 else BG
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run(); run.text = val
        f = run.font
        f.size = Pt(11.5 if r else 12.5)
        if r == 0:
            f.bold = True; f.color.rgb = ORANGE; f.name = BODY_FONT
        elif c == 0:
            f.bold = True; f.color.rgb = ORANGE; f.name = MONO_FONT
            f.size = Pt(11)
        else:
            f.color.rgb = WHITE if c == 1 else GREY
            f.name = BODY_FONT
footer(s, "Cheat Sheet")

# =============================================================== SLIDE 14 ==
s = add_slide()
title_block(s, "Risk Management", "Why Alpha Dies \u2014 and How to Keep Yours Alive")
bullets(s, [
    ("Overfitting: with enough parameters any backtest \u201cworks\u201d \u2014 "
     "demand an economic reason the edge exists, out-of-sample proof, and "
     "point-in-time data.", 0),
    ("Crowding: published anomalies decay ~50%+ post-publication; monitor "
     "ownership (`HDS`), shorts (`SI`) and factor valuations before entry.", 0),
    ("Costs: turnover, borrow fees and market impact convert paper alpha "
     "into real losses \u2014 model them with `BTCA <GO>` before scaling.", 0),
    ("Hidden beta: re-run `PORT` attribution regularly; if the \u201calpha\u201d "
     "is a static value tilt, own it deliberately and cheaply instead.", 0),
    ("Regime change: carry and mean-reversion invert in stress \u2014 "
     "pre-define the kill-switch, size for the drawdown you haven't seen.", 0),
    ("Survivorship & look-ahead bias: backtest on the universe as it "
     "existed then \u2014 `EQBT` and `BQNT` support point-in-time universes.", 0),
], top=Inches(1.65), size=15, space_after=11)
footer(s, "Risk Management")

# =============================================================== SLIDE 15 ==
s = add_slide()
strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Pt(5))
strip.fill.solid(); strip.fill.fore_color.rgb = ORANGE
strip.line.fill.background(); strip.shadow.inherit = False

title_block(s, "Summary", "Key Takeaways")
bullets(s, [
    ("Alpha is scarce, risk-adjusted, and zero-sum \u2014 classify every "
     "idea by WHY the edge exists: informational, analytical, behavioral "
     "or structural.", 0),
    ("Eight durable hunting grounds: factors, fundamentals, events, macro, "
     "carry, flows, sentiment/alt-data, and statistical arbitrage.", 0),
    ("The Terminal is an alpha pipeline, not a quote screen: "
     "`discover \u2192 screen \u2192 analyze \u2192 validate \u2192 construct \u2192 monitor`, "
     "each step has a function.", 0, ),
    ("Defend what you find: measure crowding, model costs, re-attribute "
     "returns \u2014 assume every edge is decaying and re-earn it.", 0),
], top=Inches(1.7), size=17, space_after=14)

_, tf = txbox(s, Inches(0.6), Inches(5.7), Inches(12.1), Inches(0.9))
p = tf.paragraphs[0]
set_runs(p, "\u201cThe terminal doesn't hand you alpha \u2014 it compresses "
            "the time from hypothesis to evidence.\u201d", 17, color=ORANGE,
         bold=True)
footer(s, "Summary")

# ------------------------------------------------------------------- save --
OUT = "alpha_sources_bloomberg.pptx"
prs.save(OUT)
print(f"Saved {OUT} with {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
