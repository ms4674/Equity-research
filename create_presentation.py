from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData

DARK_BG = RGBColor(0x0F, 0x17, 0x2A)
ACCENT_BLUE = RGBColor(0x00, 0x7A, 0xFF)
ACCENT_CYAN = RGBColor(0x00, 0xD4, 0xFF)
ACCENT_GREEN = RGBColor(0x00, 0xE6, 0x96)
ACCENT_ORANGE = RGBColor(0xFF, 0x8C, 0x00)
ACCENT_RED = RGBColor(0xFF, 0x45, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xB0, 0xB8, 0xC8)
MID_GRAY = RGBColor(0x60, 0x68, 0x78)
CARD_BG = RGBColor(0x1A, 0x24, 0x3B)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

W = prs.slide_width
H = prs.slide_height


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, left, top, width, height, fill_color, border_color=None, border_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width or Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=14, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_paragraph(text_frame, text, font_size=14, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT, space_before=Pt(4), space_after=Pt(2), font_name="Calibri"):
    p = text_frame.add_paragraph()
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    p.space_before = space_before
    p.space_after = space_after
    return p


def add_accent_line(slide, left, top, width, color=ACCENT_BLUE):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_stat_card(slide, left, top, width, height, number, label, accent_color=ACCENT_BLUE):
    card = add_shape(slide, left, top, width, height, CARD_BG, accent_color, Pt(1.5))

    add_accent_line(slide, left + Inches(0.15), top + Inches(0.08), width - Inches(0.3), accent_color)

    add_text_box(slide, left + Inches(0.2), top + Inches(0.2), width - Inches(0.4), Inches(0.7),
                 number, font_size=28, color=accent_color, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left + Inches(0.2), top + Inches(0.85), width - Inches(0.4), Inches(0.6),
                 label, font_size=11, color=LIGHT_GRAY, bold=False, alignment=PP_ALIGN.CENTER)
    return card


def add_slide_number(slide, num, total):
    add_text_box(slide, Inches(12.3), Inches(7.05), Inches(0.9), Inches(0.35),
                 f"{num}/{total}", font_size=10, color=MID_GRAY, alignment=PP_ALIGN.RIGHT)


def add_section_header(slide, title, subtitle=""):
    add_accent_line(slide, Inches(0.6), Inches(0.55), Inches(1.5), ACCENT_BLUE)
    add_text_box(slide, Inches(0.6), Inches(0.7), Inches(10), Inches(0.6),
                 title, font_size=30, color=WHITE, bold=True)
    if subtitle:
        add_text_box(slide, Inches(0.6), Inches(1.25), Inches(10), Inches(0.4),
                     subtitle, font_size=14, color=LIGHT_GRAY)


TOTAL_SLIDES = 12

# =============================================================================
# SLIDE 1: Title Slide
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

gradient_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
gradient_shape.fill.solid()
gradient_shape.fill.fore_color.rgb = RGBColor(0x0A, 0x12, 0x28)
gradient_shape.line.fill.background()
gradient_shape.shadow.inherit = False

for i, (left_pct, top_pct, size, color, alpha) in enumerate([
    (0.7, 0.2, 3.0, ACCENT_BLUE, 0.08),
    (0.1, 0.6, 2.5, ACCENT_CYAN, 0.06),
    (0.85, 0.7, 2.0, ACCENT_GREEN, 0.05),
]):
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        int(W * left_pct), int(H * top_pct), Inches(size), Inches(size))
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.fill.background()
    circle.shadow.inherit = False

add_accent_line(slide, Inches(0.8), Inches(2.5), Inches(2), ACCENT_CYAN)

add_text_box(slide, Inches(0.8), Inches(2.7), Inches(11), Inches(1.2),
             "Corporate AI Adoption & Agent Deployments",
             font_size=42, color=WHITE, bold=True)
add_text_box(slide, Inches(0.8), Inches(3.9), Inches(10), Inches(0.7),
             "S&P 500 & Global 2000 Intelligence Report",
             font_size=24, color=ACCENT_CYAN, bold=False)
add_text_box(slide, Inches(0.8), Inches(4.7), Inches(10), Inches(0.5),
             "Data-Driven Analysis of Enterprise AI Investment, Deployment, and Impact | February 2026",
             font_size=14, color=LIGHT_GRAY)

add_shape(slide, Inches(0.8), Inches(5.8), Inches(3.5), Inches(0.05), ACCENT_BLUE)

sources_box = add_text_box(slide, Inches(0.8), Inches(5.95), Inches(10), Inches(0.4),
             "Sources: Microsoft, Deloitte, McKinsey, Gartner, BCG, PwC, KPMG, Bloomberg, UBS",
             font_size=10, color=MID_GRAY)

add_slide_number(slide, 1, TOTAL_SLIDES)

# =============================================================================
# SLIDE 2: Executive Summary
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Executive Summary", "Key findings from S&P 500 and Global 2000 AI adoption analysis")

stats = [
    ("80%", "Fortune 500 companies\nnow deploy active\nAI agents", ACCENT_BLUE),
    ("$650B", "Big Tech AI infra\nspending in 2026\n(4 companies alone)", ACCENT_CYAN),
    ("215%", "Agentic AI market\ngrowth rate, reaching\n$89.6B in 2026", ACCENT_GREEN),
    ("540%", "Average enterprise\nROI within 18 months\nof AI agent deployment", ACCENT_ORANGE),
]

for i, (number, label, color) in enumerate(stats):
    add_stat_card(slide, Inches(0.6 + i * 3.15), Inches(1.9), Inches(2.85), Inches(1.6), number, label, color)

bullet_data = [
    "AI has crossed the threshold from experimentation to enterprise-scale deployment across S&P 500 firms",
    "Global AI spending is projected to reach $2.53 trillion in 2026, with infrastructure exceeding $1.3 trillion",
    "40% of enterprise applications will embed AI agents by end of 2026 (Gartner), up from 5% in 2025",
    "Critical challenges remain: $890K avg. implementation costs, 340K global talent shortage, governance gaps",
    "CEOs now directly own AI investment decisions at 75% of large enterprises, signaling strategic priority"
]

for i, bullet in enumerate(bullet_data):
    box = add_text_box(slide, Inches(0.8), Inches(3.85 + i * 0.55), Inches(11.5), Inches(0.5),
                       f"\u2022  {bullet}", font_size=13, color=LIGHT_GRAY)

add_slide_number(slide, 2, TOTAL_SLIDES)

# =============================================================================
# SLIDE 3: AI Agent Adoption Rates
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "AI Agent Adoption Rates", "Fortune 500 deployment trajectory 2023-2026")

chart_data = CategoryChartData()
chart_data.categories = ['2023', '2024', '2025', '2026']
chart_data.add_series('Deployment Rate (%)', (25, 45, 67, 80))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.6), Inches(1.8), Inches(5.5), Inches(4.5),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = False

plot = chart.plots[0]
plot.gap_width = 120
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = ACCENT_BLUE
series.format.line.fill.background()
series.data_labels.font.size = Pt(12)
series.data_labels.font.color.rgb = WHITE
series.data_labels.font.bold = True
series.data_labels.number_format = '0"%"'
series.data_labels.show_value = True

cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(12)
cat_axis.tick_labels.font.color.rgb = LIGHT_GRAY
cat_axis.format.line.fill.background()
cat_axis.major_tick_mark = 2

val_axis = chart.value_axis
val_axis.maximum_scale = 100
val_axis.minimum_scale = 0
val_axis.tick_labels.font.size = Pt(10)
val_axis.tick_labels.font.color.rgb = MID_GRAY
val_axis.format.line.fill.background()
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x2A, 0x34, 0x4A)

chart_area = chart.chart_style
chart.element.attrib.pop('{http://schemas.openxmlformats.org/drawingml/2006/chart}style', None)

milestones = [
    ("80%", "of Fortune 500 deploy\nactive AI agents (2026)", ACCENT_BLUE),
    ("78%", "projected deployment\nrate per analyst consensus", ACCENT_CYAN),
    ("8x", "jump in enterprise apps\nembedding AI agents\n(5% to 40%, Gartner)", ACCENT_GREEN),
    ("14", "Fortune 500 \"AI Leaders\"\nwith enterprise-scale\nAI integration", ACCENT_ORANGE),
]

for i, (num, label, color) in enumerate(milestones):
    add_stat_card(slide, Inches(6.8), Inches(1.8 + i * 1.3), Inches(3.0), Inches(1.15), num, label, color)

card = add_shape(slide, Inches(10.3), Inches(1.8), Inches(2.6), Inches(5.1), CARD_BG, ACCENT_BLUE, Pt(1))
tb = add_text_box(slide, Inches(10.5), Inches(1.95), Inches(2.2), Inches(0.35),
                  "KEY INSIGHT", font_size=12, color=ACCENT_BLUE, bold=True, alignment=PP_ALIGN.CENTER)
insight_text = ("AI agent adoption has shifted from "
                "experimentation to production. "
                "Microsoft reports 80% of Fortune 500 "
                "now deploy active agents using "
                "low-code tools, marking a decisive "
                "inflection point in enterprise AI maturity.")
add_text_box(slide, Inches(10.45), Inches(2.35), Inches(2.3), Inches(4.2),
             insight_text, font_size=11, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(6.6), Inches(12), Inches(0.3),
             "Sources: Microsoft Cyber Pulse Report (Feb 2026), Axis Intelligence, Gartner (Feb 2026)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 3, TOTAL_SLIDES)

# =============================================================================
# SLIDE 4: Global AI Spending Landscape
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Global AI Spending Landscape", "Big Tech capex and enterprise investment trajectories")

companies = [
    ("Amazon", "$200B", "+60% YoY", ACCENT_BLUE, 200),
    ("Alphabet", "$175-185B", "+97% YoY", ACCENT_CYAN, 180),
    ("Meta", "$115-135B", "+73% YoY", ACCENT_GREEN, 125),
    ("Microsoft", "~$105B", "+41% YoY", ACCENT_ORANGE, 105),
]

max_val = 200
bar_area_left = Inches(2.5)
bar_area_width = Inches(6.5)

for i, (company, amount, growth, color, val) in enumerate(companies):
    y = Inches(2.0 + i * 1.15)

    add_text_box(slide, Inches(0.6), y + Inches(0.05), Inches(1.8), Inches(0.35),
                 company, font_size=16, color=WHITE, bold=True, alignment=PP_ALIGN.RIGHT)

    bar_width = int(bar_area_width * val / max_val)
    bar = add_shape(slide, bar_area_left, y + Inches(0.05), bar_width, Inches(0.4), color)

    add_text_box(slide, bar_area_left + bar_width + Inches(0.15), y, Inches(1.5), Inches(0.3),
                 amount, font_size=15, color=WHITE, bold=True)
    add_text_box(slide, bar_area_left + bar_width + Inches(0.15), y + Inches(0.28), Inches(1.2), Inches(0.25),
                 growth, font_size=11, color=color)

add_text_box(slide, Inches(0.6), Inches(6.0), Inches(9), Inches(0.35),
             "Combined 2026 AI Infrastructure Spending: ~$650 Billion",
             font_size=18, color=ACCENT_CYAN, bold=True)
add_text_box(slide, Inches(0.6), Inches(6.4), Inches(9), Inches(0.3),
             "Equivalent to the annual GDP of Sweden or Israel",
             font_size=12, color=LIGHT_GRAY)

right_stats = [
    ("$2.53T", "Total global AI\nspending in 2026", ACCENT_BLUE),
    ("$480B", "Global AI spend\n(UBS estimate)", ACCENT_CYAN),
    ("1.7%", "Corp. AI spend as\n% of revenue (2x YoY)", ACCENT_GREEN),
    ("33%", "YoY growth in\nglobal AI spending", ACCENT_ORANGE),
]

for i, (num, label, color) in enumerate(right_stats):
    add_stat_card(slide, Inches(10.2), Inches(1.8 + i * 1.3), Inches(2.7), Inches(1.15), num, label, color)

add_text_box(slide, Inches(0.6), Inches(6.85), Inches(12), Inches(0.3),
             "Sources: Bloomberg (Feb 2026), UBS Wealth Management, BCG (2026)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 4, TOTAL_SLIDES)

# =============================================================================
# SLIDE 5: Agentic AI Market Growth
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Agentic AI Market Trajectory", "From $7B to $93B: the fastest-growing enterprise AI segment")

chart_data = CategoryChartData()
chart_data.categories = ['2024', '2025', '2026', '2027', '2028', '2029', '2030']
chart_data.add_series('Market Size ($B)', (3.0, 7.5, 15.0, 25.0, 40.0, 62.0, 93.0))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.AREA,
    Inches(0.6), Inches(1.8), Inches(7.5), Inches(4.8),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = False

plot = chart.plots[0]
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = RGBColor(0x00, 0x7A, 0xFF)
series.format.line.color.rgb = ACCENT_CYAN
series.format.line.width = Pt(2.5)
series.smooth = True

cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(12)
cat_axis.tick_labels.font.color.rgb = LIGHT_GRAY
cat_axis.format.line.color.rgb = MID_GRAY
cat_axis.major_tick_mark = 2

val_axis = chart.value_axis
val_axis.tick_labels.font.size = Pt(10)
val_axis.tick_labels.font.color.rgb = MID_GRAY
val_axis.format.line.fill.background()
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x2A, 0x34, 0x4A)
val_axis.tick_labels.number_format = '$#,##0"B"'

right_items = [
    ("$89.6B", "Agentic AI market\nsize in 2026 (est.)", ACCENT_BLUE),
    ("44.6%", "CAGR through 2032\nfor agentic AI market", ACCENT_CYAN),
    ("$450-650B", "Additional annual revenue\nby 2030 (McKinsey est.)", ACCENT_GREEN),
    ("30%", "Of enterprise software\nsales driven by agentic\nAI by 2035", ACCENT_ORANGE),
]

for i, (num, label, color) in enumerate(right_items):
    add_stat_card(slide, Inches(8.6), Inches(1.8 + i * 1.3), Inches(3.0), Inches(1.15), num, label, color)

add_text_box(slide, Inches(0.6), Inches(6.85), Inches(12), Inches(0.3),
             "Sources: Axis Intelligence, McKinsey, Gartner (Feb 2026), Synvestable",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 5, TOTAL_SLIDES)

# =============================================================================
# SLIDE 6: Fortune 500 AI Leaders
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Fortune 500 AI Leaders", "14 companies classified as enterprise-scale AI leaders")

leaders = [
    ("Amazon", "Multi-modal AI across retail, cloud (AWS),\nlogistics; Nova Act agent platform", ACCENT_BLUE),
    ("Microsoft", "Copilot ecosystem, Azure AI, 80% of\nFortune 500 on AI platform", ACCENT_CYAN),
    ("Alphabet", "Gemini models, Google Cloud AI, agentic\npartnerships with Walmart/Home Depot", ACCENT_GREEN),
    ("JPMorgan Chase", "400+ AI use cases in production,\n200K daily LLM Suite users", ACCENT_ORANGE),
    ("Walmart", "AI-driven supply chain, pricing, inventory;\nGoogle Universal Commerce Protocol", ACCENT_BLUE),
    ("Meta", "LLaMA open-source models, $115-135B\nAI infrastructure investment", ACCENT_CYAN),
    ("Visa", "AI-powered fraud detection, transaction\nanalytics at global scale", ACCENT_GREEN),
    ("Mastercard", "AI-driven risk management, real-time\npayment intelligence", ACCENT_ORANGE),
]

for i, (company, desc, color) in enumerate(leaders):
    col = i % 2
    row = i // 2
    left = Inches(0.6 + col * 6.3)
    top = Inches(1.8 + row * 1.3)

    card = add_shape(slide, left, top, Inches(5.9), Inches(1.15), CARD_BG, color, Pt(1))

    add_accent_line(slide, left + Inches(0.15), top + Inches(0.08), Inches(1.0), color)

    add_text_box(slide, left + Inches(0.2), top + Inches(0.15), Inches(2.0), Inches(0.35),
                 company, font_size=16, color=color, bold=True)
    add_text_box(slide, left + Inches(0.2), top + Inches(0.5), Inches(5.4), Inches(0.55),
                 desc, font_size=11, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(7.05), Inches(5), Inches(0.3),
             "Additional leaders: Apple, Tesla, UnitedHealth, NVIDIA, Salesforce, Oracle",
             font_size=11, color=MID_GRAY)

add_text_box(slide, Inches(7), Inches(7.05), Inches(6), Inches(0.3),
             "Source: SaltTechno Fortune 500 AI Adoption Tracker 2026",
             font_size=9, color=MID_GRAY, alignment=PP_ALIGN.RIGHT)
add_slide_number(slide, 6, TOTAL_SLIDES)

# =============================================================================
# SLIDE 7: AI Agent Use Cases by Industry
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "AI Agent Use Cases by Industry", "Production deployments across key sectors")

industries = [
    ("Financial Services", ACCENT_BLUE, [
        "Transaction reconciliation & verification",
        "Client onboarding with KYC automation",
        "Fraud detection & risk management",
        "Goldman Sachs: autonomous reconciliation agents"
    ]),
    ("Healthcare", ACCENT_CYAN, [
        "Clinical documentation automation",
        "Patient care workflow optimization",
        "Administrative burden reduction (2hr doc per 1hr patient)",
        "Drug discovery & clinical trial management"
    ]),
    ("Retail & Consumer", ACCENT_GREEN, [
        "Inventory & supply chain automation",
        "Customer insight from reviews & tickets",
        "Kroger: AI meal planning & shopping",
        "Home Depot: aisle-level project guidance"
    ]),
    ("Manufacturing", ACCENT_ORANGE, [
        "Production workflow optimization",
        "Predictive maintenance & quality control",
        "Factory fine-tuning & operational efficiency",
        "OpenAI Frontier: millions of daily transactions"
    ]),
]

for i, (industry, color, items) in enumerate(industries):
    left = Inches(0.6 + i * 3.15)
    top = Inches(1.8)
    width = Inches(2.9)
    height = Inches(4.5)

    card = add_shape(slide, left, top, width, height, CARD_BG, color, Pt(1.5))
    add_accent_line(slide, left + Inches(0.15), top + Inches(0.08), width - Inches(0.3), color)

    add_text_box(slide, left + Inches(0.2), top + Inches(0.2), width - Inches(0.4), Inches(0.35),
                 industry, font_size=15, color=color, bold=True, alignment=PP_ALIGN.CENTER)

    for j, item in enumerate(items):
        add_text_box(slide, left + Inches(0.2), top + Inches(0.7 + j * 0.8), width - Inches(0.4), Inches(0.7),
                     f"\u2022 {item}", font_size=10.5, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(6.55), Inches(12), Inches(0.3),
             "57% of companies now run AI agents operationally; only 14% have production-ready solutions across all use cases",
             font_size=12, color=ACCENT_CYAN, bold=True)

add_text_box(slide, Inches(0.6), Inches(6.95), Inches(12), Inches(0.3),
             "Sources: Context Studios, Ampcome, Beam AI, PYMNTS (2026)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 7, TOTAL_SLIDES)

# =============================================================================
# SLIDE 8: Enterprise AI Platforms & Ecosystems
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Enterprise AI Platforms & Ecosystems", "Major platform plays driving corporate AI consumption")

platforms = [
    ("Microsoft / Azure AI", ACCENT_BLUE,
     "Copilot + Azure OpenAI Service",
     "80% Fortune 500 using AI tools; Copilot integrated across Office 365, GitHub, Dynamics. Low-code agent deployment enabling rapid enterprise adoption."),
    ("Google Cloud / Gemini", ACCENT_CYAN,
     "Gemini + Vertex AI + Universal Commerce Protocol",
     "Partnerships with Walmart, Kroger, Home Depot for agentic commerce. Gemini Enterprise for CX. Vertex AI agent builder for custom deployments."),
    ("Amazon / AWS", ACCENT_GREEN,
     "Bedrock + Nova Models + Nova Act",
     "Nova Act achieving 90% reliability for browser automation. Amazon Bedrock for enterprise model hosting. $200B infrastructure investment in 2026."),
    ("Salesforce / Agentforce", ACCENT_ORANGE,
     "Agentforce + Data Cloud + MuleSoft",
     "3.2 trillion tokens delivered to customers. Informatica acquisition for AI data integration. Agentforce platform for autonomous CRM agents."),
    ("OpenAI / Frontier", ACCENT_RED,
     "Frontier Platform + Enterprise API",
     "Production deployments at HP, Intuit, Oracle, State Farm, Thermo Fisher, Uber. Pilots at BBVA, Cisco, T-Mobile. Millions of daily transactions processed."),
]

for i, (name, color, subtitle, desc) in enumerate(platforms):
    top = Inches(1.8 + i * 1.05)
    card = add_shape(slide, Inches(0.6), top, Inches(12.1), Inches(0.9), CARD_BG, color, Pt(1))

    add_text_box(slide, Inches(0.8), top + Inches(0.08), Inches(2.8), Inches(0.35),
                 name, font_size=15, color=color, bold=True)
    add_text_box(slide, Inches(3.8), top + Inches(0.08), Inches(3.5), Inches(0.3),
                 subtitle, font_size=11, color=WHITE, bold=True)
    add_text_box(slide, Inches(0.8), top + Inches(0.42), Inches(11.6), Inches(0.45),
                 desc, font_size=10.5, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(7.05), Inches(12), Inches(0.3),
             "Sources: Microsoft, Google Cloud, Amazon, Salesforce, OpenAI official announcements (2025-2026)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 8, TOTAL_SLIDES)

# =============================================================================
# SLIDE 9: AI Spending by Corporate Function
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "AI Investment by Sector & Function", "How corporations allocate AI budgets across industries")

chart_data = CategoryChartData()
chart_data.categories = ['Technology', 'Financial\nServices', 'Healthcare', 'Retail', 'Manufacturing', 'Energy', 'Real Estate']
chart_data.add_series('AI Spend as % of Revenue', (2.1, 2.0, 1.5, 1.4, 0.9, 0.8, 0.5))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED,
    Inches(0.6), Inches(1.8), Inches(7.0), Inches(4.5),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = False

plot = chart.plots[0]
plot.gap_width = 80
series = plot.series[0]

colors_for_bars = [ACCENT_BLUE, ACCENT_BLUE, ACCENT_CYAN, ACCENT_CYAN, ACCENT_GREEN, ACCENT_GREEN, ACCENT_ORANGE]
series.format.fill.solid()
series.format.fill.fore_color.rgb = ACCENT_BLUE

series.data_labels.font.size = Pt(11)
series.data_labels.font.color.rgb = WHITE
series.data_labels.font.bold = True
series.data_labels.number_format = '0.0"%"'
series.data_labels.show_value = True

cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(11)
cat_axis.tick_labels.font.color.rgb = LIGHT_GRAY
cat_axis.format.line.fill.background()

val_axis = chart.value_axis
val_axis.maximum_scale = 2.5
val_axis.tick_labels.font.size = Pt(10)
val_axis.tick_labels.font.color.rgb = MID_GRAY
val_axis.format.line.fill.background()
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x2A, 0x34, 0x4A)

insights = [
    ("1.7%", "Average corporate AI\nspend as % of revenue\nin 2026 (2x vs 2025)", ACCENT_BLUE),
    ("75%", "Of CEOs now directly\nown AI investment\ndecisions", ACCENT_CYAN),
    ("80%", "Of C-suite more\noptimistic about AI\nROI vs. last year", ACCENT_GREEN),
    ("88%", "Plan to increase\nAI budgets due to\nagentic AI capabilities", ACCENT_ORANGE),
]

for i, (num, label, color) in enumerate(insights):
    add_stat_card(slide, Inches(8.2), Inches(1.8 + i * 1.3), Inches(3.0), Inches(1.15), num, label, color)

add_text_box(slide, Inches(0.6), Inches(6.55), Inches(12), Inches(0.3),
             "Tech and financial services lead AI spending at ~2% of revenues; industrials and real estate under 1%",
             font_size=12, color=ACCENT_CYAN, bold=True)
add_text_box(slide, Inches(0.6), Inches(6.95), Inches(12), Inches(0.3),
             "Sources: BCG (2026), PwC AI Agent Survey, Deloitte State of AI in Enterprise",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 9, TOTAL_SLIDES)

# =============================================================================
# SLIDE 10: Challenges & Risk Landscape
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Challenges & Risk Landscape", "Critical barriers to enterprise AI scaling")

challenges = [
    ("Implementation Cost", "$890K", "Average cost to deploy enterprise\nAI agent systems", ACCENT_RED),
    ("Talent Shortage", "340K", "Global AI talent gap;\nonly 24-27% report adequate\nskilled workforce", ACCENT_ORANGE),
    ("Governance Gap", "1 in 5", "Companies have mature\ngovernance models for\nautonomous AI agents", ACCENT_BLUE),
    ("Shadow AI Risk", "29%", "Of employees used\nunsanctioned AI tools,\ncreating security risks", ACCENT_CYAN),
    ("Production Gap", "<33%", "Of AI initiatives in large\nenterprises reach\nfull production", ACCENT_GREEN),
    ("Security Concerns", "72%", "Lack confidence in their\norg's ability to secure\nAI systems", ACCENT_RED),
]

for i, (title, stat, desc, color) in enumerate(challenges):
    col = i % 3
    row = i // 3
    left = Inches(0.6 + col * 4.15)
    top = Inches(1.8 + row * 2.5)

    card = add_shape(slide, left, top, Inches(3.85), Inches(2.2), CARD_BG, color, Pt(1.5))
    add_accent_line(slide, left + Inches(0.15), top + Inches(0.08), Inches(1.2), color)

    add_text_box(slide, left + Inches(0.2), top + Inches(0.2), Inches(3.4), Inches(0.3),
                 title, font_size=14, color=WHITE, bold=True)
    add_text_box(slide, left + Inches(0.2), top + Inches(0.55), Inches(3.4), Inches(0.45),
                 stat, font_size=30, color=color, bold=True)
    add_text_box(slide, left + Inches(0.2), top + Inches(1.15), Inches(3.4), Inches(0.8),
                 desc, font_size=11, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(6.95), Inches(12), Inches(0.3),
             "Sources: Axis Intelligence, Microsoft, Deloitte, CSA/Google Cloud, KPMG (2025-2026)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 10, TOTAL_SLIDES)

# =============================================================================
# SLIDE 11: S&P 500 Earnings Call AI Signals
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "S&P 500 Earnings Call AI Signals", "How market leaders are framing AI in investor communications")

companies_data = [
    ("NVIDIA", "Q4 2026", ACCENT_GREEN, [
        "$68B quarterly revenue (+73% YoY)",
        "Data Center: $62B quarterly / $194B annual (13x since ChatGPT)",
        "$10B investment in Anthropic; partnerships with OpenAI, Groq",
        "Vera Rubin platform: 10x lower inference costs vs Blackwell",
        "Physical AI: $6B+ annual revenue contribution",
    ]),
    ("Oracle", "Q1 FY26", ACCENT_BLUE, [
        "RPO surged 359% YoY to $455B on AI contracts",
        "Record deals with OpenAI, xAI for cloud infrastructure",
        "Cloud revenue +28% to $7.2B; IaaS up 55%",
        "AI Database: run Gemini, ChatGPT, Grok, Llama natively",
        "OCI revenue expected to grow 77% to $18B in FY26",
    ]),
    ("Salesforce", "Q3-Q4 FY26", ACCENT_CYAN, [
        "Agentforce platform: autonomous CRM agents at scale",
        "3.2 trillion tokens delivered to customers",
        "Informatica acquisition for AI + data integration",
        "Data Cloud + MuleSoft as agentic AI foundation",
        "AI-first strategy driving subscription revenue growth",
    ]),
]

for i, (company, period, color, bullets) in enumerate(companies_data):
    left = Inches(0.6 + i * 4.15)
    top = Inches(1.8)
    width = Inches(3.85)

    card = add_shape(slide, left, top, width, Inches(4.8), CARD_BG, color, Pt(1.5))
    add_accent_line(slide, left + Inches(0.15), top + Inches(0.08), width - Inches(0.3), color)

    add_text_box(slide, left + Inches(0.2), top + Inches(0.18), Inches(2.0), Inches(0.35),
                 company, font_size=18, color=color, bold=True)
    add_text_box(slide, left + Inches(2.2), top + Inches(0.22), Inches(1.4), Inches(0.3),
                 period, font_size=11, color=MID_GRAY, alignment=PP_ALIGN.RIGHT)

    for j, bullet in enumerate(bullets):
        add_text_box(slide, left + Inches(0.2), top + Inches(0.65 + j * 0.75), width - Inches(0.4), Inches(0.7),
                     f"\u2022 {bullet}", font_size=10.5, color=LIGHT_GRAY)

add_text_box(slide, Inches(0.6), Inches(6.95), Inches(12), Inches(0.3),
             "Sources: NVIDIA, Oracle, Salesforce earnings call transcripts (2025-2026 fiscal year)",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 11, TOTAL_SLIDES)

# =============================================================================
# SLIDE 12: Key Takeaways & Outlook
# =============================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_section_header(slide, "Key Takeaways & 2026-2027 Outlook", "Strategic implications for enterprise decision-makers")

takeaways = [
    ("AI agents are now mainstream in the Fortune 500",
     "80% of Fortune 500 companies deploy active AI agents. The shift from conversational AI to autonomous, action-oriented agents marks a decisive inflection point. Companies not deploying agents risk competitive disadvantage.",
     ACCENT_BLUE),
    ("AI spending is accelerating exponentially",
     "Big Tech alone will invest $650B in AI infrastructure in 2026. Corporations are doubling AI spend to 1.7% of revenues. Global AI spending is on track to reach $2.53T in 2026.",
     ACCENT_CYAN),
    ("The governance gap is the critical risk",
     "Only 1 in 5 companies has mature AI governance. 29% of employees use shadow AI. 72% lack confidence in AI security. Organizations that solve governance will gain decisive advantages.",
     ACCENT_RED),
    ("Agentic AI is the next growth vector",
     "40% of enterprise apps will embed agents by end of 2026 (Gartner). The agentic AI market is growing at 44.6% CAGR. By 2035, agentic AI will drive 30% of all enterprise software sales.",
     ACCENT_GREEN),
]

for i, (title, desc, color) in enumerate(takeaways):
    top = Inches(1.8 + i * 1.3)
    card = add_shape(slide, Inches(0.6), top, Inches(12.1), Inches(1.15), CARD_BG, color, Pt(1.5))
    add_accent_line(slide, Inches(0.75), top + Inches(0.08), Inches(1.5), color)

    num_box = add_shape(slide, Inches(0.8), top + Inches(0.25), Inches(0.55), Inches(0.55), color)
    add_text_box(slide, Inches(0.8), top + Inches(0.25), Inches(0.55), Inches(0.55),
                 str(i + 1), font_size=22, color=DARK_BG, bold=True, alignment=PP_ALIGN.CENTER)

    add_text_box(slide, Inches(1.55), top + Inches(0.18), Inches(10.8), Inches(0.35),
                 title, font_size=15, color=WHITE, bold=True)
    add_text_box(slide, Inches(1.55), top + Inches(0.55), Inches(10.8), Inches(0.5),
                 desc, font_size=11, color=LIGHT_GRAY)

add_shape(slide, Inches(0.6), Inches(7.0), Inches(12.1), Inches(0.04), ACCENT_BLUE)

add_text_box(slide, Inches(0.6), Inches(6.5), Inches(12), Inches(0.4),
             "Outlook: By 2027, AI orchestration market will triple to $30B. Enterprise AI software spending will reach $270B by 2031.",
             font_size=13, color=ACCENT_CYAN, bold=True)

add_text_box(slide, Inches(0.6), Inches(6.95), Inches(12), Inches(0.3),
             "Compiled from: Microsoft, Deloitte, McKinsey, Gartner, BCG, PwC, KPMG, Bloomberg, UBS | February 2026",
             font_size=9, color=MID_GRAY)
add_slide_number(slide, 12, TOTAL_SLIDES)

# =============================================================================
# Save
# =============================================================================
output_path = "/workspace/Corporate_AI_Adoption_SP500_Global2000.pptx"
prs.save(output_path)
print(f"Presentation saved to {output_path}")
print(f"Total slides: {len(prs.slides)}")
