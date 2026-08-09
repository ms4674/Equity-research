#!/usr/bin/env python3
"""Generate 'Sell-Side Research in the Age of LLMs' PowerPoint deck.

Usage: python3 generate_presentation.py
Output: sell-side-research-in-the-age-of-llms.pptx
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ----------------------------------------------------------------------------
# Design system
# ----------------------------------------------------------------------------
NAVY = RGBColor(0x0B, 0x22, 0x39)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
SLATE = RGBColor(0x44, 0x52, 0x5E)
GRAY = RGBColor(0x6B, 0x77, 0x84)
LIGHT = RGBColor(0xF4, 0xF6, 0xF8)
LIGHT2 = RGBColor(0xE9, 0xED, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x7D, 0x52)
AMBER = RGBColor(0xB0, 0x77, 0x0A)
RED = RGBColor(0xA8, 0x3A, 0x2E)
GREEN_BG = RGBColor(0xE3, 0xF0, 0xE8)
AMBER_BG = RGBColor(0xFB, 0xF0, 0xD9)
RED_BG = RGBColor(0xF7, 0xE4, 0xE1)

FONT = "Calibri"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)
CONTENT_W = Inches(12.133)

DECK_FOOTER = "Sell-Side Research in the Age of LLMs"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

_page = [0]


def _solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_rect(slide, x, y, w, h, color, rounded=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, x, y, w, h)
    if rounded:
        shape.adjustments[0] = 0.045
    return _solid(shape, color)


def add_text(slide, x, y, w, h, runs_by_para, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, wrap=True):
    """runs_by_para: list of paragraph dicts:
    {runs: [(text, size, color, bold, italic)], space_after, space_before,
     align, line}
    """
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs_by_para):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get("align", align)
        if para.get("space_after") is not None:
            p.space_after = para["space_after"]
        if para.get("space_before") is not None:
            p.space_before = para["space_before"]
        if para.get("line") is not None:
            p.line_spacing = para["line"]
        for text, size, color, bold, italic in para["runs"]:
            r = p.add_run()
            r.text = text
            r.font.name = FONT
            r.font.size = size
            r.font.color.rgb = color
            r.font.bold = bold
            r.font.italic = italic
    return box


def para(text, size, color, bold=False, italic=False, **kw):
    return {"runs": [(text, size, color, bold, italic)], **kw}


def bullets_para(items, size=Pt(13), color=SLATE, lead_color=NAVY,
                 space=Pt(8), line=1.08):
    """items: list of strings or (lead, rest) tuples; sub-bullets prefix '>'."""
    out = []
    for item in items:
        if isinstance(item, tuple):
            lead, rest = item
            out.append({
                "runs": [("\u25aa  ", size, GOLD, True, False),
                         (lead, size, lead_color, True, False),
                         (rest, size, color, False, False)],
                "space_after": space, "line": line})
        elif item.startswith(">"):
            out.append({
                "runs": [("      \u2013  ", size - Pt(1), GRAY, False, False),
                         (item[1:].strip(), size - Pt(1), color, False, False)],
                "space_after": Pt(int(space.pt * 0.7)), "line": line})
        else:
            out.append({
                "runs": [("\u25aa  ", size, GOLD, True, False),
                         (item, size, color, False, False)],
                "space_after": space, "line": line})
    return out


def content_slide(kicker, title, subtitle=None):
    _page[0] += 1
    slide = prs.slides.add_slide(BLANK)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, NAVY)
    add_text(slide, MARGIN, Inches(0.38), CONTENT_W, Inches(0.3),
             [para(kicker.upper(), Pt(11), GOLD, bold=True)])
    add_text(slide, MARGIN, Inches(0.66), CONTENT_W, Inches(0.75),
             [para(title, Pt(26), NAVY, bold=True)])
    y_rule = Inches(1.32)
    if subtitle:
        add_text(slide, MARGIN, Inches(1.28), CONTENT_W, Inches(0.32),
                 [para(subtitle, Pt(13), GRAY, italic=True)])
        y_rule = Inches(1.66)
    add_rect(slide, MARGIN, y_rule, CONTENT_W, Pt(2.2), GOLD)
    # Footer
    add_rect(slide, MARGIN, Inches(7.08), CONTENT_W, Pt(1), LIGHT2)
    add_text(slide, MARGIN, Inches(7.14), Inches(8), Inches(0.28),
             [para(DECK_FOOTER, Pt(9), GRAY)])
    add_text(slide, SLIDE_W - MARGIN - Inches(1.2), Inches(7.14), Inches(1.2),
             Inches(0.28), [para(str(_page[0]), Pt(9), GRAY)],
             align=PP_ALIGN.RIGHT)
    return slide, (y_rule + Inches(0.24))


def add_card(slide, x, y, w, h, header, body_paras, accent=GOLD,
             header_size=Pt(13.5), fill=LIGHT):
    add_rect(slide, x, y, w, h, fill, rounded=True)
    add_rect(slide, x, y, w, Pt(3.2), accent)
    inner_x = x + Inches(0.18)
    inner_w = w - Inches(0.36)
    add_text(slide, inner_x, y + Inches(0.14), inner_w, Inches(0.5),
             [para(header, header_size, NAVY, bold=True)])
    add_text(slide, inner_x, y + Inches(0.56), inner_w, h - Inches(0.7),
             body_paras)


def add_stat(slide, x, y, w, value, label, value_color=NAVY):
    add_rect(slide, x, y, w, Inches(1.5), LIGHT, rounded=True)
    add_text(slide, x + Inches(0.12), y + Inches(0.14), w - Inches(0.24),
             Inches(0.62), [para(value, Pt(27), value_color, bold=True)],
             align=PP_ALIGN.CENTER)
    add_text(slide, x + Inches(0.14), y + Inches(0.78), w - Inches(0.28),
             Inches(0.66), [para(label, Pt(10.5), SLATE, line=1.05)],
             align=PP_ALIGN.CENTER)


# ----------------------------------------------------------------------------
# Slide 1 — Title
# ----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
add_rect(slide, 0, Inches(7.02), SLIDE_W, Inches(0.48), GOLD)
add_text(slide, MARGIN, Inches(1.7), CONTENT_W, Inches(0.4),
         [para("INDUSTRY PERSPECTIVE", Pt(14), GOLD, bold=True)])
add_text(slide, MARGIN, Inches(2.15), Inches(11.6), Inches(2.2), [
    para("Sell-Side Research in the Age of LLMs",
         Pt(44), WHITE, bold=True, space_after=Pt(10)),
    para("The value proposition, what large language models commoditize, "
         "and how the business must adapt", Pt(20),
         RGBColor(0xC9, 0xD6, 0xE3))])
add_rect(slide, MARGIN, Inches(4.35), Inches(2.4), Pt(3), GOLD)
add_text(slide, MARGIN, Inches(5.9), Inches(11), Inches(0.9), [
    para("Strategy discussion document", Pt(13), RGBColor(0xC9, 0xD6, 0xE3),
         space_after=Pt(4)),
    para("August 2026  \u2022  Figures cited are directional industry "
         "estimates for discussion", Pt(11), RGBColor(0x8F, 0xA3, 0xB8))])

# ----------------------------------------------------------------------------
# Slide 2 — Executive summary
# ----------------------------------------------------------------------------
slide, y0 = content_slide("Executive summary",
                          "The report is dying; the analyst is not")
msgs = [
    ("Research was structurally challenged before LLMs arrived.",
     "MiFID II unbundling, shrinking commission wallets, the passive shift "
     "and falling readership had already compressed economics and headcount "
     "for a decade."),
    ("LLMs commoditize the written report.",
     "Summaries, recaps, primers and maintenance notes \u2014 anything "
     "derived from public information \u2014 can now be produced by a "
     "client's own tools at near-zero marginal cost, on demand."),
    ("The durable value was never the PDF.",
     "Proprietary data and primary research, non-consensus judgment with "
     "named accountability, corporate access and trusted relationships stay "
     "scarce \u2014 and become relatively more valuable."),
    ("The product must be rebuilt, not just accelerated.",
     "Machine-readable, queryable, agent-accessible research; analysts "
     "redeployed from maintenance output to differentiated insight and "
     "primary information gathering."),
    ("The business model must follow.",
     "From bundled content toward data and IP licensing, priced analyst "
     "access and smaller, more senior teams \u2014 while deliberately "
     "solving the apprenticeship problem automation creates."),
]
y = y0 + Inches(0.12)
for i, (lead, rest) in enumerate(msgs, start=1):
    add_rect(slide, MARGIN, y, Inches(0.52), Inches(0.86), NAVY, rounded=True)
    add_text(slide, MARGIN, y + Inches(0.17), Inches(0.52), Inches(0.5),
             [para(str(i), Pt(20), GOLD, bold=True)], align=PP_ALIGN.CENTER)
    add_text(slide, MARGIN + Inches(0.75), y + Inches(0.02), Inches(11.35),
             Inches(0.95), [
                 para(lead, Pt(14), NAVY, bold=True, space_after=Pt(2)),
                 para(rest, Pt(12), SLATE, line=1.05)])
    y += Inches(1.03)

# ----------------------------------------------------------------------------
# Slide 3 — The business today
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Context", "A business already under structural pressure",
    "LLMs land on an industry whose economics were repriced well before "
    "generative AI")
stats = [
    ("~50%", "Decline in global equity research spend since pre-MiFID II "
     "peaks (industry estimates)"),
    ("<5%", "Share of published notes meaningfully read by intended "
     "clients (industry estimates)"),
    ("-30%+", "Reduction in senior analyst headcount at major banks over "
     "the past decade (est.)"),
    ("~90%", "Share of research output that is maintenance coverage: "
     "recaps, updates, previews (est.)"),
]
sw = Inches(2.88)
gap = Inches(0.2)
x = MARGIN
for value, label in stats:
    add_stat(slide, x, y0 + Inches(0.1), sw, value, label)
    x += sw + gap
items = [
    ("Unbundling repriced the product. ",
     "MiFID II (2018) forced explicit payment for research; budgets fell "
     "sharply and never recovered, and US practice converged toward the "
     "same discipline."),
    ("Consumption shifted away from documents. ",
     "Clients triage torrents of notes via aggregators and internal "
     "platforms; the marginal PDF competes with every other PDF for "
     "seconds of attention."),
    ("Cross-subsidy sustains the economics. ",
     "Research is rarely profitable standalone; it is carried by trading "
     "relationships and the investment-banking franchise it supports."),
    ("Coverage concentrates where the wallet is. ",
     "Large caps are over-covered; small and mid caps increasingly "
     "orphaned \u2014 a gap both a risk and an opening for AI-scaled "
     "coverage."),
]
add_text(slide, MARGIN, y0 + Inches(1.85), CONTENT_W, Inches(3.2),
         bullets_para(items, size=Pt(13), space=Pt(10)))

# ----------------------------------------------------------------------------
# Slide 4 — Value proposition decomposed
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Value proposition", "What clients actually buy \u2014 five pillars",
    "The written report is the container for several distinct services; "
    "LLMs affect each pillar very differently")
cards = [
    ("1. Information & synthesis",
     ["Collecting, cleaning and contextualizing public and industry "
      "information at speed \u2014 filings, transcripts, data, news \u2014 "
      "especially around earnings and events."]),
    ("2. Analysis & forecasting",
     ["Maintained financial models and estimates that anchor consensus; "
      "scenario and valuation work clients benchmark against."]),
    ("3. Judgment & \u201cthe call\u201d",
     ["Ratings and non-consensus views from named analysts with public "
      "track records; pattern recognition built across cycles."]),
    ("4. Corporate access",
     ["Management meetings, non-deal roadshows, conferences and expert "
      "networks \u2014 physical-world scarcity that content cannot "
      "replicate."]),
    ("5. Service & relationships",
     ["Bespoke analysis on request, market color and positioning, being "
      "the first call \u2014 trust reinforced by accountability "
      "(e.g., Reg AC)."]),
]
cw = Inches(3.92)
ch = Inches(1.98)
positions = [
    (MARGIN, y0 + Inches(0.12)),
    (MARGIN + cw + Inches(0.18), y0 + Inches(0.12)),
    (MARGIN + 2 * (cw + Inches(0.18)), y0 + Inches(0.12)),
    (MARGIN + Inches(2.05), y0 + Inches(0.12) + ch + Inches(0.22)),
    (MARGIN + Inches(2.05) + cw + Inches(0.18),
     y0 + Inches(0.12) + ch + Inches(0.22)),
]
for (header, body), (cx, cy) in zip(cards, positions):
    add_card(slide, cx, cy, cw, ch, header,
             [para(body[0], Pt(11.5), SLATE, line=1.12)])
add_text(slide, MARGIN, y0 + Inches(4.65), CONTENT_W, Inches(0.45),
         [para("Pillars 1\u20132 are content businesses exposed to "
               "automation; pillars 3\u20135 are people-and-trust businesses "
               "that are far harder to replicate.",
               Pt(12.5), NAVY, bold=True, italic=True)],
         align=PP_ALIGN.CENTER)

# ----------------------------------------------------------------------------
# Slide 5 — How research gets paid
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Economics", "Clients pay for the analyst, not the PDF",
    "Understanding the payment mechanics reframes the LLM threat")
items = [
    ("Broker vote. ", "Buy-side firms allocate research payments via "
     "periodic votes that reward service, responsiveness and access "
     "at least as much as written output."),
    ("Explicit research payments (CSA / RPA / hard dollar). ",
     "Post-unbundling budgets are negotiated, scrutinized and benchmarked "
     "\u2014 every line item must justify itself annually."),
    ("Execution alignment. ", "Research strengthens trading relationships "
     "and market share; the report is marketing for the platform as much "
     "as a product in itself."),
    ("Banking halo. ", "Coverage supports ECM/IPO franchises within strict "
     "independence rules; issuers value credible, visible analysts."),
    ("Corporate access allocations. ", "A large share of the vote in many "
     "accounts is effectively payment for access, not documents."),
]
add_text(slide, MARGIN, y0 + Inches(0.12), Inches(7.3), Inches(4.6),
         bullets_para(items, size=Pt(13), space=Pt(12)))
add_card(slide, MARGIN + Inches(7.7), y0 + Inches(0.12), Inches(4.43),
         Inches(4.5), "Why this matters for LLMs",
         [para("Because payment already attaches to relationships, "
               "judgment and access rather than documents, LLMs do not "
               "destroy the whole franchise \u2014 they hollow out the "
               "container.", Pt(12.5), SLATE, line=1.15,
               space_after=Pt(10)),
          para("But the container is what justified headcount, coverage "
               "breadth and much of the visible activity the vote "
               "rewards. When machines write the notes, firms must "
               "re-anchor the vote to what machines cannot do.",
               Pt(12.5), SLATE, line=1.15)],
         accent=NAVY)

# ----------------------------------------------------------------------------
# Slide 6 — What LLMs already do well
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "The technology", "What LLMs already do well in the research workflow",
    "Frontier models plus agentic workflows now chain these tasks "
    "end-to-end with limited supervision")
rows = [
    ("Ingest & summarize", "Filings, transcripts, news and prior research "
     "digested in seconds; instant earnings first-takes and 10-K/10-Q "
     "change detection."),
    ("Extract & structure", "KPIs, guidance and segment data lifted from "
     "PDFs and calls into clean, comparable tables across a full "
     "coverage universe."),
    ("Draft", "First-pass maintenance notes, previews and recaps in the "
     "analyst's voice, ready for editing and compliance review."),
    ("Screen & answer", "Natural-language queries across corpora: thematic "
     "exposure screens, cross-company comparisons, \u201cwho said what, "
     "when\u201d retrieval."),
    ("Model & update", "Agents populate and refresh models after prints, "
     "run flux analysis and flag estimate divergence versus consensus."),
    ("Translate & distribute", "Global-language versions, client-specific "
     "summaries and format conversion at effectively zero marginal cost."),
]
col_w = Inches(5.95)
row_h = Inches(1.36)
for i, (head, body) in enumerate(rows):
    cx = MARGIN + (i % 2) * (col_w + Inches(0.24))
    cy = y0 + Inches(0.12) + (i // 2) * (row_h + Inches(0.18))
    add_card(slide, cx, cy, col_w, row_h, head,
             [para(body, Pt(11), SLATE, line=1.1)], accent=BLUE,
             header_size=Pt(12.5))

# ----------------------------------------------------------------------------
# Slide 7 — Commoditization map
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Exposure", "The commoditization map",
    "Substitutability today, by research output \u2014 public-information "
    "synthesis is most exposed")
table_rows = [
    ("Earnings recaps / first takes", "HIGH",
     "Clients' own agents produce these from the transcript in minutes"),
    ("Maintenance notes & estimate changes", "HIGH",
     "Largest share of output volume; automate or abandon"),
    ("Industry primers & initiation background", "HIGH",
     "One-off reference content is exactly what LLMs synthesize best"),
    ("Comp sheets & consensus aggregation", "HIGH",
     "Structured-data problem; already ceded to platforms and agents"),
    ("Maintained models & scenario tools", "MEDIUM",
     "Mechanics automate; assumption-setting and credibility do not"),
    ("Proprietary surveys & channel checks", "LOW",
     "Primary information LLMs cannot scrape; value rises"),
    ("Non-consensus stock calls & track record", "LOW",
     "Judgment plus named accountability remains the franchise core"),
    ("Corporate access & relationships", "LOW",
     "Physical-world scarcity; not a content task at all"),
]
tbl_x, tbl_y = MARGIN, y0 + Inches(0.1)
tbl = slide.shapes.add_table(len(table_rows) + 1, 3, tbl_x, tbl_y,
                             CONTENT_W, Inches(4.55)).table
tbl.columns[0].width = Inches(4.35)
tbl.columns[1].width = Inches(1.55)
tbl.columns[2].width = Inches(6.23)
headers = ("Research output", "LLM substitutability", "Implication")
for j, h in enumerate(headers):
    cell = tbl.cell(0, j)
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY
    cell.margin_top = cell.margin_bottom = Pt(3)
    tfp = cell.text_frame.paragraphs[0]
    r = tfp.add_run()
    r.text = h
    r.font.name = FONT
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = WHITE
badge = {"HIGH": (RED, RED_BG), "MEDIUM": (AMBER, AMBER_BG),
         "LOW": (GREEN, GREEN_BG)}
for i, (output, level, implication) in enumerate(table_rows, start=1):
    for j, text in enumerate((output, level, implication)):
        cell = tbl.cell(i, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE if i % 2 else LIGHT
        cell.margin_top = cell.margin_bottom = Pt(2)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfp = cell.text_frame.paragraphs[0]
        r = tfp.add_run()
        r.text = text
        r.font.name = FONT
        r.font.size = Pt(11)
        if j == 0:
            r.font.color.rgb = NAVY
            r.font.bold = True
        elif j == 1:
            fg, bg = badge[level]
            r.font.color.rgb = fg
            r.font.bold = True
            tfp.alignment = PP_ALIGN.CENTER
            cell.fill.fore_color.rgb = bg
        else:
            r.font.color.rgb = SLATE

# ----------------------------------------------------------------------------
# Slide 8 — Buy-side disintermediation
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Demand shift", "The buy side is building its own synthesis layer",
    "The client's agent now sits between the analyst's PDF and the "
    "portfolio manager")
items = [
    ("Clients read primary sources directly. ", "Internal copilots ingest "
     "filings, transcripts and alternative data the moment they publish "
     "\u2014 ahead of, and instead of, the sell-side note."),
    ("Research is consumed by machines. ", "Notes are parsed into internal "
     "knowledge bases; the analyst's carefully crafted document becomes "
     "one retrieval source among many, stripped of format and branding."),
    ("Attribution weakens, so the vote weakens. ", "When insight arrives "
     "via an internal agent, clients struggle to know \u2014 or care \u2014 "
     "whose research contributed, eroding vote-based monetization."),
    ("The competitive set widens. ", "Alt-data vendors, expert networks, "
     "aggregators and AI-native research startups all compete for the "
     "same shrinking, better-scrutinized wallet."),
]
add_text(slide, MARGIN, y0 + Inches(0.12), Inches(7.4), Inches(4.4),
         bullets_para(items, size=Pt(13), space=Pt(13)))
add_rect(slide, MARGIN + Inches(7.8), y0 + Inches(0.3), Inches(4.33),
         Inches(3.5), NAVY, rounded=True)
add_text(slide, MARGIN + Inches(8.05), y0 + Inches(0.62), Inches(3.83),
         Inches(3.0), [
             para("The uncomfortable test", Pt(14), GOLD, bold=True,
                  space_after=Pt(10)),
             para("\u201cIf the client's agent has already read the 10-K, "
                  "what is the analyst's summary of the 10-K worth?\u201d",
                  Pt(15), WHITE, italic=True, line=1.2,
                  space_after=Pt(10)),
             para("Every research product should be re-underwritten "
                  "against this question.", Pt(11.5),
                  RGBColor(0xC9, 0xD6, 0xE3), line=1.15)])

# ----------------------------------------------------------------------------
# Slide 9 — What remains defensible
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Durable value", "What remains scarce when content is free",
    "Five assets LLMs strengthen rather than erode")
cards = [
    ("Primary information",
     "Proprietary surveys, channel checks, expert relationships and "
     "datasets that do not exist on the public internet \u2014 the raw "
     "material client models cannot scrape."),
    ("Judgment & accountability",
     "A named analyst with a public track record, willing to be "
     "non-consensus and wrong in public. Regulatory attestation (Reg AC) "
     "means someone stands behind the view."),
    ("Corporate access",
     "Management trust, non-deal roadshows and conferences. Scarcity is "
     "physical and relational; AI increases the premium on genuine "
     "face-to-face access."),
    ("Trusted, liable synthesis",
     "Compliance-reviewed, source-verified, auditable content that a "
     "fiduciary can rely on and cite \u2014 provenance becomes a product "
     "feature, not a footnote."),
    ("Flows & market color",
     "Positioning, liquidity and sentiment read from the trading floor in "
     "real time \u2014 information that exists only inside the "
     "franchise."),
]
cw = Inches(3.92)
ch = Inches(2.15)
positions = [
    (MARGIN, y0 + Inches(0.12)),
    (MARGIN + cw + Inches(0.18), y0 + Inches(0.12)),
    (MARGIN + 2 * (cw + Inches(0.18)), y0 + Inches(0.12)),
    (MARGIN + Inches(2.05), y0 + Inches(0.12) + ch + Inches(0.22)),
    (MARGIN + Inches(2.05) + cw + Inches(0.18),
     y0 + Inches(0.12) + ch + Inches(0.22)),
]
for (header, body), (cx, cy) in zip(cards, positions):
    add_card(slide, cx, cy, cw, ch, header,
             [para(body, Pt(11.5), SLATE, line=1.12)], accent=GREEN)

# ----------------------------------------------------------------------------
# Slide 10 — The value shift
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "The value shift", "The report was the container, not the product",
    "LLMs collapse the value of the document \u2014 and raise the premium "
    "on what feeds it and what sits above it")
col_w = Inches(5.9)
add_card(slide, MARGIN, y0 + Inches(0.15), col_w, Inches(3.6),
         "What clients stop paying for",
         bullets_para([
             "Summaries and recaps of public information",
             "Maintenance notes and mechanical estimate updates",
             "Background primers, initiations-as-reference",
             "Aggregation: comp sheets, consensus tables",
             "Volume and speed as differentiators in themselves",
         ], size=Pt(12.5), space=Pt(9)), accent=RED)
add_card(slide, MARGIN + col_w + Inches(0.33), y0 + Inches(0.15), col_w,
         Inches(3.6), "What clients pay more for",
         bullets_para([
             "Proprietary data and primary research inputs",
             "Differentiated, accountable, non-consensus judgment",
             "Corporate access and trusted relationships",
             "Verified, auditable, citable synthesis (provenance)",
             "Direct access to the analyst \u2014 the scarce human hour",
         ], size=Pt(12.5), space=Pt(9)), accent=GREEN)
add_rect(slide, MARGIN, y0 + Inches(4.0), CONTENT_W, Inches(0.75), NAVY,
         rounded=True)
add_text(slide, MARGIN + Inches(0.3), y0 + Inches(4.15), CONTENT_W -
         Inches(0.6), Inches(0.5),
         [para("Value migrates along the chain: from content \u2192 to "
               "data, access and judgment. Firms that own the inputs and "
               "the trust win; firms that own only the document do not.",
               Pt(13.5), WHITE, bold=True)], align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)

# ----------------------------------------------------------------------------
# Slide 11 — Adaptation: efficiency
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Adaptation \u2014 horizon 1", "Automate the commoditized layer",
    "Efficiency is the entry ticket, not the strategy \u2014 but it funds "
    "everything else")
items = [
    ("Deploy analyst copilots across coverage. ", "Summarization, "
     "extraction, drafting and model maintenance embedded in the research "
     "workflow, tuned to each franchise's voice and templates."),
    ("Automate the first take \u2014 with sign-off. ", "Machine-drafted "
     "earnings notes published minutes after the print, reviewed and "
     "signed by the named analyst; speed with accountability."),
    ("Industrialize model maintenance. ", "Agents refresh models, run flux "
     "analysis and flag anomalies; analysts adjudicate assumptions rather "
     "than re-key data."),
    ("Build compliance into the pipeline. ", "Draft \u2192 automated "
     "checks (sourcing, disclosures, restricted lists) \u2192 human review "
     "\u2192 publish; auditable end to end."),
    ("Redeploy the time saved. ", "The point is not fewer analysts "
     "producing the same notes \u2014 it is the same analysts producing "
     "primary research, calls and client engagement instead."),
]
add_text(slide, MARGIN, y0 + Inches(0.12), Inches(7.9), Inches(4.6),
         bullets_para(items, size=Pt(13), space=Pt(12)))
add_card(slide, MARGIN + Inches(8.2), y0 + Inches(0.12), Inches(3.93),
         Inches(4.35), "Metrics that matter",
         bullets_para([
             "Time-to-publish at earnings",
             "Coverage breadth at constant cost",
             "Share of analyst time on primary work",
             "Client interactions per analyst",
             "Error and correction rates vs. baseline",
         ], size=Pt(11.5), space=Pt(9)), accent=BLUE)

# ----------------------------------------------------------------------------
# Slide 12 — Adaptation: product redesign
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Adaptation \u2014 horizon 2", "Publish for two readers: humans and "
    "machines",
    "If clients' agents are the new readership, the product must be "
    "designed for them")
rows = [
    ("Research as API", "Structured estimates, ratings, theses and model "
     "outputs delivered as entitled data feeds \u2014 not locked inside "
     "PDFs."),
    ("Licensed knowledge base", "The research corpus made available, with "
     "entitlements and usage tracking, for retrieval inside clients' own "
     "LLM stacks."),
    ("Interactive models", "Live scenario tools replacing static "
     "exhibits: clients stress the analyst's assumptions instead of "
     "reading about them."),
    ("\u201cAsk the franchise\u201d", "A conversational interface over the "
     "analyst's models and notes, answering instantly and escalating to "
     "the human when it matters."),
    ("Provenance as a feature", "Every claim linked to its source; "
     "citation metadata and audit trails that clients' compliance "
     "functions can verify."),
    ("Attribution telemetry", "Usage signals from machine consumption "
     "feeding the vote \u2014 proving value even when no human opens the "
     "note."),
]
col_w = Inches(5.95)
row_h = Inches(1.36)
for i, (head, body) in enumerate(rows):
    cx = MARGIN + (i % 2) * (col_w + Inches(0.24))
    cy = y0 + Inches(0.12) + (i // 2) * (row_h + Inches(0.18))
    add_card(slide, cx, cy, col_w, row_h, head,
             [para(body, Pt(11), SLATE, line=1.1)], accent=GOLD,
             header_size=Pt(12.5))

# ----------------------------------------------------------------------------
# Slide 13 — Adaptation: business model
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Adaptation \u2014 horizon 3", "Reprice around scarcity",
    "The revenue model follows the value: data, access and judgment "
    "\u2014 not documents")
items = [
    ("From subscriptions to data & IP licensing. ", "Structured research "
     "content, estimates and proprietary datasets licensed for machine "
     "consumption, with usage-based and query-based pricing."),
    ("Tier the human. ", "Analyst time becomes the explicit premium "
     "product \u2014 priced calls, bespoke work and advisory retainers "
     "\u2014 while the machine-served layer scales cheaply underneath."),
    ("Price corporate access properly. ", "A compliant, transparent "
     "access product rather than an implicit subsidy buried in the "
     "vote."),
    ("Fewer, more senior, more leveraged teams. ", "Each franchise "
     "supported by agents and shared data infrastructure; coverage "
     "breadth via automation, differentiation via people."),
    ("Decide the platform question deliberately. ", "Build proprietary "
     "AI infrastructure, partner with LLM and data platforms, or license "
     "out the corpus \u2014 each path has different margin and control "
     "trade-offs."),
]
add_text(slide, MARGIN, y0 + Inches(0.12), Inches(7.9), Inches(4.6),
         bullets_para(items, size=Pt(13), space=Pt(12)))
add_card(slide, MARGIN + Inches(8.2), y0 + Inches(0.12), Inches(3.93),
         Inches(4.35), "The moat, restated",
         [para("The defensible asset is the franchise: named analysts, "
               "proprietary inputs and institutional trust.",
               Pt(12), SLATE, line=1.18, space_after=Pt(10)),
          para("LLMs are distribution and leverage for that franchise "
               "\u2014 they are only a threat to firms whose product was "
               "the document itself.", Pt(12), SLATE, line=1.18)],
         accent=NAVY)

# ----------------------------------------------------------------------------
# Slide 14 — Talent & organization
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "People", "The apprenticeship problem is the hardest problem",
    "Automation removes the work juniors learned on \u2014 the pipeline "
    "must be redesigned, not left to atrophy")
items = [
    ("The pyramid inverts. ", "Associates historically learned by doing "
     "what LLMs now do: modeling, note drafting, data work. Teams get "
     "smaller and more senior \u2014 but seniors have to come from "
     "somewhere."),
    ("Design deliberate training. ", "Rotations through primary research, "
     "client interaction and live calls; simulation and review-based "
     "learning replacing grunt-work osmosis."),
    ("New roles enter the team. ", "Research engineers, agent and prompt "
     "operations, data curators and provenance stewards sit alongside "
     "analysts as first-class functions."),
    ("The analyst skill mix shifts. ", "Sourcing proprietary information, "
     "making judgments, communicating and building relationships dominate; "
     "production mechanics fade."),
    ("Compensation follows scarcity. ", "Franchise analysts and the "
     "engineers who leverage them capture more of the pool; "
     "undifferentiated production roles compress."),
]
add_text(slide, MARGIN, y0 + Inches(0.12), CONTENT_W, Inches(4.7),
         bullets_para(items, size=Pt(13.5), space=Pt(14)))

# ----------------------------------------------------------------------------
# Slide 15 — Risks & constraints
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Risks", "Constraints on the transition",
    "Research is regulated, liable content \u2014 adoption must be "
    "engineered, not improvised")
cards = [
    ("Hallucination & liability",
     "Errors in published research carry fiduciary and reputational "
     "consequences. Human sign-off, automated fact-checking and "
     "source-linking are non-negotiable."),
    ("Regulatory perimeter",
     "Reg AC attestations, MiFID inducement rules, research independence "
     "and fair-distribution obligations all apply to machine-assisted "
     "output exactly as to human output."),
    ("Information barriers & MNPI",
     "Data products and agent pipelines must respect walls; aggregation "
     "of client flows or private signals can create MNPI risk at scale."),
    ("IP leakage",
     "Research used to train or ground external models without "
     "entitlement destroys the licensing opportunity; contracts and "
     "telemetry must protect the corpus."),
    ("Willingness-to-pay compression",
     "Clients' own AI keeps deflating the content layer; cost savings "
     "alone cannot outrun a shrinking wallet \u2014 new products must "
     "land before old revenue erodes."),
    ("Trust asymmetry",
     "One prominent AI-generated error can damage a franchise built over "
     "decades; the compliance bar for automation is rightly higher than "
     "for humans."),
]
cw = Inches(3.92)
ch = Inches(2.1)
for i, (header, body) in enumerate(cards):
    cx = MARGIN + (i % 3) * (cw + Inches(0.18))
    cy = y0 + Inches(0.12) + (i // 3) * (ch + Inches(0.22))
    add_card(slide, cx, cy, cw, ch, header,
             [para(body, Pt(10.8), SLATE, line=1.1)], accent=RED,
             header_size=Pt(12.5))

# ----------------------------------------------------------------------------
# Slide 16 — Scenarios
# ----------------------------------------------------------------------------
slide, y0 = content_slide(
    "Scenarios", "Three futures for sell-side research by 2030",
    "The middle of the industry is most at risk in every scenario")
scenarios = [
    ("Managed decline", RED,
     ["Research shrinks toward a banking-support cost center",
      "Coverage narrows; maintenance output fully automated",
      "Wallet keeps compressing; juniors not replaced",
      "Value migrates to alt-data vendors and AI-native entrants"]),
    ("Barbell (base case)", AMBER,
     ["Scaled platforms win on proprietary data + distribution",
      "Boutiques win on depth, access and named judgment",
      "The undifferentiated middle exits or consolidates",
      "Vote re-anchors to access, data and analyst time"]),
    ("Reinvention", GREEN,
     ["Research becomes a data & IP business with API distribution",
      "Machine consumption is measured, attributed and monetized",
      "Margins improve as production cost collapses",
      "Analyst franchises monetized directly \u2014 media-like economics"]),
]
cw = Inches(3.92)
for i, (header, accent, points) in enumerate(scenarios):
    cx = MARGIN + i * (cw + Inches(0.18))
    add_card(slide, cx, y0 + Inches(0.12), cw, Inches(3.85), header,
             bullets_para(points, size=Pt(11.5), space=Pt(9), line=1.1),
             accent=accent, header_size=Pt(14))
add_text(slide, MARGIN, y0 + Inches(4.2), CONTENT_W, Inches(0.5),
         [para("Which scenario a firm lands in is a choice: the "
               "determining variables are proprietary data, franchise "
               "analysts and machine-era distribution \u2014 all "
               "investable today.", Pt(12.5), NAVY, bold=True,
               italic=True)], align=PP_ALIGN.CENTER)

# ----------------------------------------------------------------------------
# Slide 17 — Strategic imperatives
# ----------------------------------------------------------------------------
slide, y0 = content_slide("Conclusion", "Five strategic imperatives")
imps = [
    ("Automate the commoditized layer ruthlessly",
     "\u2014 and reinvest every saved hour in primary research and client "
     "engagement, not in more notes."),
    ("Rebuild distribution for machines",
     "\u2014 APIs, structured data, licensed corpora and provenance "
     "metadata; the client's agent is now a reader."),
    ("Reprice around scarcity",
     "\u2014 data, corporate access and analyst judgment carry the "
     "P&L; documents do not."),
    ("Protect trust as the core asset",
     "\u2014 compliance-by-design, verified sourcing and named "
     "accountability on everything published."),
    ("Rebuild the talent pipeline deliberately",
     "\u2014 juniors no longer learn by grinding; apprenticeship must be "
     "designed, or the senior bench of 2035 will not exist."),
]
y = y0 + Inches(0.15)
for i, (lead, rest) in enumerate(imps, start=1):
    add_rect(slide, MARGIN, y, Inches(0.5), Inches(0.72), GOLD, rounded=True)
    add_text(slide, MARGIN, y + Inches(0.12), Inches(0.5), Inches(0.5),
             [para(str(i), Pt(19), NAVY, bold=True)], align=PP_ALIGN.CENTER)
    add_text(slide, MARGIN + Inches(0.72), y + Inches(0.05), Inches(11.4),
             Inches(0.75), [{
                 "runs": [(lead + " ", Pt(14.5), NAVY, True, False),
                          (rest, Pt(13), SLATE, False, False)],
                 "line": 1.08}])
    y += Inches(0.88)
add_rect(slide, MARGIN, y + Inches(0.08), CONTENT_W, Inches(0.78), NAVY,
         rounded=True)
add_text(slide, MARGIN + Inches(0.3), y + Inches(0.08), CONTENT_W -
         Inches(0.6), Inches(0.78),
         [para("The PDF is dying; the analyst is not. Franchises that "
               "treat LLMs as leverage and distribution \u2014 rather "
               "than as a threat to the document \u2014 will take share.",
               Pt(14.5), WHITE, bold=True)],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ----------------------------------------------------------------------------
OUT = "sell-side-research-in-the-age-of-llms.pptx"
prs.save(OUT)
print(f"Saved {OUT} with {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
