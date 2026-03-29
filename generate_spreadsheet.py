import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Styles ──────────────────────────────────────────────────────────────
header_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
title_font = Font(name="Calibri", bold=True, size=14, color="1F4E79")
section_font = Font(name="Calibri", bold=True, size=11, color="1F4E79")
normal_font = Font(name="Calibri", size=11)
source_font = Font(name="Calibri", size=9, color="555555")
note_font = Font(name="Calibri", size=10, italic=True, color="333333")

header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
section_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
alt_fill = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")
highlight_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

wrap_align = Alignment(wrap_text=True, vertical="top")
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_row(ws, row, font=normal_font, fill=None, alignment=wrap_align):
    for col in range(1, 7):
        cell = ws.cell(row=row, column=col)
        cell.font = font
        cell.alignment = alignment
        cell.border = thin_border
        if fill:
            cell.fill = fill


# ═══════════════════════════════════════════════════════════════════════
# Sheet 1: Cost Comparison
# ═══════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Cost Comparison"
ws.sheet_properties.tabColor = "1F4E79"

# Column widths
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 18
ws.column_dimensions["E"].width = 55
ws.column_dimensions["F"].width = 55

# Title
ws.merge_cells("A1:F1")
title_cell = ws["A1"]
title_cell.value = "TSMC Fab Operational Costs: Phoenix, AZ vs. Taipei/Hsinchu, Taiwan"
title_cell.font = title_font
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 35

# Subtitle
ws.merge_cells("A2:F2")
ws["A2"].value = "Data compiled from public sources (2024–2025). All USD conversions at ~NT$32.5 = US$1."
ws["A2"].font = note_font
ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 22

# Headers (row 4)
headers = [
    "Cost Category",
    "Phoenix, AZ (USA)",
    "Taipei / Hsinchu (Taiwan)",
    "Δ (AZ vs TW)",
    "Notes",
    "Sources",
]
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = thin_border
ws.row_dimensions[4].height = 30

# ── Data rows ───────────────────────────────────────────────────────────
rows = [
    # Section: Electricity
    ("SECTION", "ELECTRICITY", "", "", "", "", ""),
    (
        "Industrial Electricity Rate",
        "~$0.079/kWh",
        "~$0.136/kWh (NT$4.29/kWh after Oct 2024 hike)",
        "AZ ~42% cheaper",
        "AZ rate is Phoenix industrial avg. Taiwan rate is post-Oct 2024 industrial rate after 12.5% increase. Taiwan semiconductor companies classified as heavy users face up to 14% surcharge.",
        "EIA / Electricity Local (Phoenix industrial avg); Taipei Times Oct 2024; TrendForce Mar 2024",
    ),
    (
        "Electricity per Wafer (est.)",
        "~$3.20/wafer (40.5 kWh × $0.079)",
        "~$5.51/wafer (40.5 kWh × $0.136)",
        "AZ ~42% cheaper",
        "Per-wafer electricity is 40.5 kWh (TSMC 2023 data for 3nm-class). AZ benefits from lower electricity rates despite same consumption.",
        "WCCFTech (TSMC 2023 ESG data: 40.5 kWh/wafer); rate sources above",
    ),
    (
        "Grid Reliability / Renewables",
        "APS/SRP grid; solar potential high; ~12% renewable",
        "Taipower; ~8–10% renewable; green power deficit concern",
        "Comparable",
        "Both regions face grid strain from fab expansion. Taiwan has green power deficit threatening net-zero goals. AZ has abundant solar but limited current adoption.",
        "Reuters Jun 2024 (Taiwan green power deficit); Stand.earth Feb 2024 (AZ grid mix)",
    ),
    # Section: Water
    ("SECTION", "WATER", "", "", "", "", ""),
    (
        "Industrial Water Rate",
        "~$8.25/1,000 gal ($4.93–$6.13/CCF by season + env. charge; ~$2.18/m³)",
        "~$0.37/m³ (NT$12/m³ residential tier; industrial estimated NT$11–14/m³)",
        "AZ ~5–6× more expensive",
        "Phoenix uses tiered seasonal pricing (low/med/high season). Taiwan water has been essentially unchanged for 31 years; Taiwater loses NT$2.45/m³ sold. Industrial rates artificially low.",
        "City of Phoenix Water Rates Mar 2025; Focus Taiwan Jan 2024; Taipei Times Mar 2025 (Taiwater)",
    ),
    (
        "Water Consumption per Wafer",
        "~3,000–5,000 liters/wafer start (leading edge ≤5nm)",
        "~3,000–5,000 liters/wafer start (leading edge ≤5nm)",
        "Same",
        "Industry benchmark for advanced 300mm fabs. TSMC Taiwan achieves ~90% recycling; AZ currently ~65%, targeting 85–90% by 2028 with new reclamation plant.",
        "Energy Solutions 2026; AZ Central Jun 2024; AZ Central Aug 2025 (reclamation plant)",
    ),
    (
        "Water Cost per Wafer (est.)",
        "~$6.54–$10.90/wafer",
        "~$1.11–$1.85/wafer",
        "AZ ~5–6× more expensive",
        "Based on 3,000–5,000 L/wafer × local rate. Water is a small fraction of total wafer cost (<0.1%) but important for sustainability and permitting.",
        "Calculated from rates and consumption above",
    ),
    (
        "Total Fab Water Usage",
        "~4.75M gallons/day (18M liters/day) at full capacity",
        "~10–20M gallons/day per leading-edge fab",
        "Similar scale",
        "TSMC Phoenix Fab 21 at full capacity. Taiwan mega-fabs may be larger. Hsinchu faces water supply challenges with MOEA targeting no industrial cuts before Jun 2026.",
        "AZ Central Jun 2024; Digitimes Mar 2026; Business Insider Jul 2024",
    ),
    # Section: Natural Gas
    ("SECTION", "NATURAL GAS", "", "", "", "", ""),
    (
        "Industrial Natural Gas Rate",
        "~$5.33/MCF ($0.53/therm; ~$0.19/m³)",
        "~$0.35/m³ (NT$11.35/m³ for NG(1) industrial, Aug 2024)",
        "AZ ~46% cheaper",
        "AZ is ranked #1 nationally for affordable natural gas. Taiwan gas prices governed by CPC Corporation formula with 3%/month adjustment cap.",
        "EIA 2024 (AZ industrial NG); CPC Corporation Taiwan Aug 2024",
    ),
    # Section: Labor
    ("SECTION", "LABOR", "", "", "", "", ""),
    (
        "Process Engineer Salary (median)",
        "~$118,000/yr ($96K–$147K total comp)",
        "~$72,000/yr (NT$2.32M–2.62M median; range NT$868K–4.41M)",
        "AZ ~64% higher",
        "US wages ~3× Taiwan on average, but engineer gap is smaller due to TSMC Taiwan's above-market pay. Labor is <2% of total wafer cost due to automation.",
        "Glassdoor (TSMC Phoenix Engineer); Levels.fyi (TSMC Taiwan Engineer); TechInsights Mar 2025",
    ),
    (
        "Technician/Operator Salary",
        "~$45,000–$65,000/yr",
        "~$18,000–$30,000/yr (NT$600K–1M)",
        "AZ ~2–3× higher",
        "Technician roles show larger gap than engineers. AZ requires significant training investment for new workforce.",
        "Industry estimates; TSMC Taiwan salary data (blog.salary.tw)",
    ),
    (
        "Labor as % of Wafer Cost",
        "<2%",
        "<2%",
        "Negligible impact",
        "Modern fabs are highly automated. Despite ~3× wage differential, labor impact on per-wafer cost is minimal. This is the key insight from TechInsights analysis.",
        "TechInsights Mar 2025 (Chip Insider); multiple corroborating sources",
    ),
    # Section: Equipment & Depreciation
    ("SECTION", "EQUIPMENT & DEPRECIATION", "", "", "", "", ""),
    (
        "Equipment Cost (% of wafer cost)",
        ">67% of total wafer cost",
        ">67% of total wafer cost",
        "Same",
        "Equipment from ASML, Applied Materials, KLA, Lam Research, Tokyo Electron is priced the same globally. This is the dominant cost equalizer.",
        "TechInsights Mar 2025; SemiWiki; TechPowerUp Mar 2025",
    ),
    (
        "EUV Lithography System",
        "~$200–350M per tool",
        "~$200–350M per tool",
        "Same",
        "ASML EUV/High-NA EUV systems are globally priced. Lithography alone is ~25% of total wafer fabrication cost.",
        "ASML pricing data; costdata.de (IC Fabrication Cost Breakdown)",
    ),
    (
        "Fab Construction Cost",
        "~$20B per fab (4–5× Taiwan)",
        "~$4–5B per equivalent fab",
        "AZ 4–5× higher (capex only)",
        "Construction premium driven by US labor costs, regulatory/permitting, and first-greenfield-site complexity. TSMC total AZ commitment: $65B for 3 fabs.",
        "Wikipedia (TSMC Arizona); TechOvedas; BlackRidge Research",
    ),
    # Section: Materials & Chemicals
    ("SECTION", "MATERIALS & CHEMICALS", "", "", "", "", ""),
    (
        "Process Chemicals & Gases",
        "Higher (import from Asia + US sourcing)",
        "Lower (proximity to Asian supply chain)",
        "AZ ~10–15% higher",
        "Global specialty gas market ~$5.7B (2025). Modern fab uses 50–100 gas species, 30–60 liquid chemical precursors. Materials are a significant but not dominant cost component.",
        "AppIT Software (Semiconductor fab gas mgmt); Silicon Analysts 2026; costdata.de",
    ),
    (
        "Silicon Wafer Blanks (300mm)",
        "~$130–$150/wafer",
        "~$130–$150/wafer",
        "Same",
        "Raw 300mm silicon wafer blanks are globally commoditized. Price is same regardless of fab location.",
        "Industry pricing data; Silicon Analysts 2026",
    ),
    (
        "Photomask Sets",
        "~$5–15M per set (advanced nodes)",
        "~$5–15M per set (advanced nodes)",
        "Same",
        "Photomask costs are node-dependent, not location-dependent. 3nm masks cost ~$5–10M, 2nm can exceed $15M.",
        "costdata.de; industry estimates",
    ),
    # Section: Taxes & Incentives
    ("SECTION", "TAXES & GOVERNMENT INCENTIVES", "", "", "", "", ""),
    (
        "Corporate Income Tax Rate",
        "21% federal + 6.968% AZ state = ~28% combined",
        "20% + 5% surtax on undistributed earnings",
        "AZ ~3–8% higher nominal",
        "Effective rates differ significantly due to credits and deductions. Both jurisdictions offer substantial semiconductor-specific incentives.",
        "Tax Foundation 2024; AZ JBC; Deloitte Taiwan Highlights 2024",
    ),
    (
        "Semiconductor Tax Credits (R&D)",
        "Federal: 25% AMTC or 20% regular R&D credit; AZ: Qualified Facility credit up to $30K/job",
        "25% deduction on R&D spending (≥NT$6B R&D, 6% intensity); 5% deduction on advanced equipment (≥NT$10B)",
        "Both generous",
        "Taiwan CHIPS Act (Feb 2024) offers historic tax breaks. US CHIPS Act + AZ state incentives provide significant offsets. AZ: up to $125M/yr in Qualified Facility credits through 2030.",
        "TrendForce Jan 2024; AZ Commerce Authority; US CHIPS Act",
    ),
    (
        "Federal/National Subsidies",
        "$6.6B direct CHIPS Act funding + $5B in loans",
        "Indirect via tax incentives; NT$350B+ in various programs",
        "AZ has larger direct grants",
        "TSMC receives up to $6.6B direct + $5B loans. Conditions: 5-yr buyback restrictions, upside sharing. Taiwan subsidies are primarily via tax breaks rather than direct grants.",
        "Reuters Nov 2024; AZ Commerce Authority Apr 2024; NIST (TSMC Arizona)",
    ),
    (
        "Property Tax Rate",
        "~0.44–0.47% (Maricopa County effective rate)",
        "~0.2–1% (Hsinchu; science park concessions available)",
        "AZ slightly higher",
        "Maricopa County rate is below AZ state avg (0.55%) and national avg (0.91%). Hsinchu Science Park offers various property-related tax concessions.",
        "PropertyTaxByState 2024; Hsinchu Local Tax Bureau; Kaizen (HSP incentives)",
    ),
    # Section: Land & Real Estate
    ("SECTION", "LAND & REAL ESTATE", "", "", "", "", ""),
    (
        "Industrial Land Cost",
        "~$24,000/acre (2024 auction near TSMC site)",
        "Hsinchu Science Park: government-subsidized lease rates",
        "Variable",
        "AZ land purchased at state auction. TSMC AZ site is 1,129 acres. Taiwan science parks offer long-term leases at subsidized rates rather than outright purchase.",
        "AZ Central May 2024; AZ State Land Dept auction; InvestTaiwan (science park incentives)",
    ),
    # Section: Logistics & Supply Chain
    ("SECTION", "LOGISTICS & SUPPLY CHAIN", "", "", "", "", ""),
    (
        "Wafer Post-Processing Logistics",
        "Ship back to Taiwan for dicing/test/packaging (adds cost & time)",
        "On-site or nearby OSAT facilities",
        "AZ disadvantage",
        "Current AZ production requires returning wafers to Taiwan for back-end processing. Adds logistics cost and 2–4 weeks transit. Long-term OSAT buildout in AZ planned.",
        "TechInsights Mar 2025; TechPowerUp Mar 2025; inkl.com Mar 2025",
    ),
    (
        "Supply Chain Proximity",
        "Limited local ecosystem; most materials imported",
        "Mature local ecosystem; dense supplier network",
        "TW advantage",
        "Taiwan has decades of semiconductor supply chain density. AZ ecosystem is nascent but growing with adjacent development (Halo Vista, etc.).",
        "General industry analysis; AZ Central Mar 2026 (Halo Vista development)",
    ),
    # Section: Overall
    ("SECTION", "OVERALL WAFER PROCESSING COST", "", "", "", "", ""),
    (
        "Total Wafer Processing Cost (300mm, N4/N3)",
        "~$17,000–$22,000/wafer",
        "~$15,500–$20,000/wafer",
        "AZ ~10% higher overall",
        "Per TechInsights Strategic Cost & Price Model (Scotten Jones). Equipment dominance (~67%+ of cost) equalizes locations. Total difference is <10%.",
        "TechInsights Mar 2025 (Chip Insider, G. Dan Hutcheson); Silicon Analysts 2026; SemiWiki",
    ),
    (
        "Operating Cost Premium (all-in)",
        "Baseline + ~10%",
        "Baseline",
        "+10%",
        "Final TechInsights assessment. Earlier estimates of 30–50%+ premiums were based on construction costs and flawed labor assumptions. Equipment parity is decisive.",
        "TechInsights Mar 2025; TechPowerUp Feb 2025 & Mar 2025; TechSpot Mar 2025",
    ),
]

row_num = 5
for entry in rows:
    if entry[0] == "SECTION":
        ws.merge_cells(
            start_row=row_num, start_column=1, end_row=row_num, end_column=6
        )
        cell = ws.cell(row=row_num, column=1, value=entry[1])
        cell.font = section_font
        cell.fill = section_fill
        cell.alignment = Alignment(vertical="center")
        cell.border = thin_border
        for c in range(2, 7):
            ws.cell(row=row_num, column=c).fill = section_fill
            ws.cell(row=row_num, column=c).border = thin_border
        ws.row_dimensions[row_num].height = 25
    else:
        ws.cell(row=row_num, column=1, value=entry[0])
        ws.cell(row=row_num, column=2, value=entry[1])
        ws.cell(row=row_num, column=3, value=entry[2])
        ws.cell(row=row_num, column=4, value=entry[3])
        ws.cell(row=row_num, column=5, value=entry[4])
        ws.cell(row=row_num, column=6, value=entry[5])

        use_alt = (row_num % 2 == 0)
        fill = alt_fill if use_alt else None
        style_row(ws, row_num, fill=fill)

        delta_cell = ws.cell(row=row_num, column=4)
        delta_cell.alignment = center_align
        if "cheaper" in str(entry[3]).lower() or "lower" in str(entry[3]).lower():
            delta_cell.font = Font(name="Calibri", size=11, color="006100")
        elif "higher" in str(entry[3]).lower() or "expensive" in str(entry[3]).lower() or "disadvantage" in str(entry[3]).lower():
            delta_cell.font = Font(name="Calibri", size=11, color="9C0006")

        ws.row_dimensions[row_num].height = 65

    row_num += 1

# Freeze panes
ws.freeze_panes = "A5"

# ═══════════════════════════════════════════════════════════════════════
# Sheet 2: Sources
# ═══════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Sources & References")
ws2.sheet_properties.tabColor = "2E75B6"

ws2.column_dimensions["A"].width = 8
ws2.column_dimensions["B"].width = 35
ws2.column_dimensions["C"].width = 90
ws2.column_dimensions["D"].width = 20

ws2.merge_cells("A1:D1")
ws2["A1"].value = "Sources & References"
ws2["A1"].font = title_font
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 30

src_headers = ["#", "Source", "URL / Reference", "Date"]
for col_idx, h in enumerate(src_headers, 1):
    cell = ws2.cell(row=3, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = thin_border

sources = [
    ("1", "TechInsights – Chip Insider: TSMC's True Cost", "https://www.techinsights.com/blog/chip-insider-tsmcs-true-cost-arizona-versus-taiwan", "Mar 2025"),
    ("2", "TechPowerUp – TSMC AZ Only 10% More Expensive", "https://www.techpowerup.com/334634/tsmc-arizona-operations-only-10-more-expensive-than-taiwanese-fab-operations", "Mar 2025"),
    ("3", "TechPowerUp – TSMC AZ 30% More (earlier est.)", "https://www.techpowerup.com/330349/tsmc-arizona-plant-operations-will-reportedly-cost-30-more-than-taiwan-sites", "Feb 2025"),
    ("4", "NIST – TSMC Arizona CHIPS Act", "https://www.nist.gov/chips/tsmc-arizona-phoenix", "2024"),
    ("5", "Reuters – US finalizes $6.6B CHIPS award", "https://www.reuters.com/technology/us-finalizes-66-billion-chips-award-tsmc-ahead-trump-return-2024-11-15/", "Nov 2024"),
    ("6", "AZ Commerce Authority – TSMC CHIPS Funding", "https://azcommerce.com/news-events/news/2024/4/tsmc-arizona-and-us-department-of-commerce-announce-up-to-66-billion-in-proposed-chips-act-direct-funding", "Apr 2024"),
    ("7", "Electricity Local – Phoenix Industrial Rates", "https://www.electricitylocal.com/states/arizona/phoenix/", "2024"),
    ("8", "Taipei Times – Industrial power rates +12.5%", "https://www.taipeitimes.com/News/front/archives/2024/10/01/2003824615", "Oct 2024"),
    ("9", "TrendForce – Taiwan Electricity Price Increase", "https://www.trendforce.com/news/2024/03/11/news-taiwans-electricity-prices-to-increase-in-april-semiconductor-industry-among-top-consumers-facing-largest-hikes/", "Mar 2024"),
    ("10", "Taiwan News – Electricity prices +0.71%", "https://taiwannews.com.tw/news/6204295", "Sep 2025"),
    ("11", "City of Phoenix – Water & Sewer Rates", "https://phoenix.gov/waterrates", "Mar 2025"),
    ("12", "Focus Taiwan – Taiwater water price evaluation", "https://focustaiwan.tw/society/202401260017", "Jan 2024"),
    ("13", "Taipei Times – Taiwater water prices", "https://www.taipeitimes.com/News/taiwan/archives/2025/03/17/2003833563", "Mar 2025"),
    ("14", "AZ Central – TSMC water use less than expected", "https://www.azcentral.com/story/opinion/op-ed/joannaallhands/2024/06/12/tsmc-arizona-water-use-recycling/74059522007/", "Jun 2024"),
    ("15", "AZ Central – TSMC builds water reclamation plant", "https://www.azcentral.com/story/money/business/tech/2025/08/27/tsmc-builds-water-plant-to-bolster-sustainability-at-phoenix-campus/85841094007/", "Aug 2025"),
    ("16", "Business Insider – TSMC & Intel water innovation", "https://www.businessinsider.com/chip-production-water-usage-tsmc-intel-2024-7", "Jul 2024"),
    ("17", "EIA – Arizona Natural Gas Prices", "https://www.eia.gov/dnav/ng/ng_pri_sum_dcu_saz_m.htm", "2024"),
    ("18", "CPC Corporation Taiwan – NG Pricing (Aug 2024)", "https://www.energy-omni.com/en/news/article/0A5Zz37IsdSs3Uh0", "Aug 2024"),
    ("19", "Glassdoor – TSMC Engineer Phoenix Salaries", "https://www.glassdoor.com/Salary/TSMC-Engineer-Phoenix-Salaries-EJI_IE4130.0,4_KO5,13_IL.14,21_IM678.htm", "2024"),
    ("20", "Levels.fyi – TSMC Engineer Salaries (Taiwan)", "https://www.levels.fyi/companies/tsmc/salaries/engineer", "2024"),
    ("21", "blog.salary.tw – TSMC Taiwan salary data", "https://blog.salary.tw/article/265d13b4-ba23-41a9-b44c-b3c547ba6024", "Nov 2024"),
    ("22", "Tax Foundation – State Corporate Income Tax Rates", "https://taxfoundation.org/data/all/state/state-corporate-income-tax-rates-brackets-2024", "2024"),
    ("23", "Deloitte – Taiwan Tax Highlights 2024", "https://www2.deloitte.com/content/dam/Deloitte/global/Documents/Tax/dttl-tax-taiwanhighlights-2024.pdf", "2024"),
    ("24", "TrendForce – Taiwan CHIPS Act Tax Incentives", "https://www.trendforce.com/news/2024/01/16/news-taiwans-chip-act-takes-effect-in-february-tsmc-to-benefit-from-historic-tax-incentives/", "Jan 2024"),
    ("25", "AZ Commerce Authority – Qualified Facility Credit", "https://www.azcommerce.com/incentives/qualified-facility-tax-credit/", "2024"),
    ("26", "PropertyTaxByState – Maricopa County", "https://propertytaxbystate.com/arizona/maricopa-county", "2024"),
    ("27", "Hsinchu Local Tax Bureau – Land Value Tax", "https://www.hcct.gov.tw/en/home.jsp?id=24&parentpath=0%2C3", "2024"),
    ("28", "AZ Central – Land auction near TSMC", "https://www.azcentral.com/story/news/local/phoenix/2024/05/29/7b-development-planned-for-land-near-taiwan-semiconductor-manufacturing-company-phoenix/73841274007/", "May 2024"),
    ("29", "WCCFTech – TSMC electricity demand per wafer", "https://wccftech.com/tsmcs-growing-electricity-demand-could-stress-credit-in-2030-warns-sp/", "2024"),
    ("30", "Energy Solutions – Semiconductor Foundries 2026", "https://energy-solutions.co/articles/sub/semiconductor-foundries-managing-extreme-power-density-water-risks", "2026"),
    ("31", "Silicon Analysts – Semiconductor Market Data 2026", "https://siliconanalysts.com/market", "2026"),
    ("32", "Wikipedia – TSMC Arizona", "https://en.wikipedia.org/wiki/TSMC_Arizona", "2025"),
    ("33", "BlackRidge Research – TSMC Arizona Project Profile", "https://blackridgeresearch.com/project-profiles/tsmc-arizona-fab-united-states-us-details-cost-expansion-latest-update", "2025"),
    ("34", "Digitimes – Hsinchu water supply issue", "https://www.digitimes.com/news/a20260327PD213/water-taiwan-moea-industrial-2026.html", "Mar 2026"),
    ("35", "InvestTaiwan – Science Park Incentives", "https://investtaiwan.nat.gov.tw/showPageeng10310015?lang=eng&search=10310015", "2024"),
    ("36", "AppIT Software – Semiconductor Fab Gas Management", "https://www.appitsoftware.com/blog/specialty-gas-chemical-precursor-semiconductor-fabs", "2025"),
    ("37", "Resto NYC – Water per chip", "https://www.restonyc.com/how-many-gallons-of-water-does-it-take-to-make-a-chip/", "2024"),
    ("38", "inkl.com – TSMC AZ 10% more expensive", "https://www.inkl.com/news/producing-wafers-at-tsmc-arizona-is-only-10-more-expensive-than-in-taiwan-techinsights", "Mar 2025"),
    ("39", "Tom's Hardware – TSMC EUV power reduction", "https://www.tomshardware.com/tech-industry/semiconductors/tsmc-reduces-peak-power-consumption-of-euv-tools-by-44-percent-company-to-save-190-million-kilowatt-hours-of-electricity-by-2030", "2024"),
    ("40", "Reuters – Taiwan green power deficit", "https://www.reuters.com/sustainability/climate-energy/how-taiwans-green-power-deficit-threatens-tech-industrys-bid-net-zero-2024-06-04/", "Jun 2024"),
    ("41", "AZ Central – Halo Vista development near TSMC", "https://www.azcentral.com/picture-gallery/news/local/phoenix/2026/03/27/halo-vista-development-near-tsmc-in-phoenix-breaks-ground/89341509007/", "Mar 2026"),
    ("42", "costdata.de – IC Fabrication Cost Breakdown", "https://www.costdata.de/en/blog/ic-fabrication-cost-breakdown", "2024"),
]

for i, (num, name, url, date) in enumerate(sources):
    r = i + 4
    ws2.cell(row=r, column=1, value=int(num))
    ws2.cell(row=r, column=2, value=name)
    ws2.cell(row=r, column=3, value=url)
    ws2.cell(row=r, column=4, value=date)
    for c in range(1, 5):
        cell = ws2.cell(row=r, column=c)
        cell.font = normal_font
        cell.border = thin_border
        cell.alignment = wrap_align
        if i % 2 == 0:
            cell.fill = alt_fill
    ws2.row_dimensions[r].height = 22

ws2.freeze_panes = "A4"

# ═══════════════════════════════════════════════════════════════════════
# Sheet 3: Key Findings Summary
# ═══════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Key Findings")
ws3.sheet_properties.tabColor = "548235"

ws3.column_dimensions["A"].width = 5
ws3.column_dimensions["B"].width = 100

ws3.merge_cells("A1:B1")
ws3["A1"].value = "Key Findings: TSMC Phoenix vs Taiwan Fab Operational Costs"
ws3["A1"].font = title_font
ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 35

findings = [
    "Overall operating cost difference is less than 10%, per TechInsights' Strategic Cost and Price Model (Mar 2025). Earlier estimates of 30–50%+ were based on construction costs and labor misconceptions.",
    "Equipment costs dominate (>67% of wafer cost) and are globally uniform from ASML, Applied Materials, KLA, Lam Research, and Tokyo Electron — this is the single biggest cost equalizer between locations.",
    "Labor accounts for less than 2% of total wafer cost despite US wages being ~3× Taiwan wages. Modern fabs are highly automated, making wage differentials largely irrelevant to per-wafer economics.",
    "Electricity is significantly cheaper in Phoenix (~$0.079/kWh) vs Taiwan (~$0.136/kWh), a ~42% advantage for Arizona. Taiwan has been raising industrial electricity rates aggressively (4 increases in 3 years).",
    "Water is 5–6× more expensive in Phoenix vs Taiwan, but water represents <0.1% of total wafer cost. The bigger concern is water availability and recycling (AZ: 65% recycling now, targeting 90%).",
    "Natural gas is ~46% cheaper in Arizona than Taiwan. Arizona ranks #1 nationally for affordable natural gas.",
    "Construction/capex costs are 4–5× higher in Arizona vs Taiwan, driven by US labor costs, regulatory requirements, and greenfield site complexity. This is distinct from operating costs.",
    "TSMC receives $6.6B in direct CHIPS Act funding plus $5B in loans for its Arizona operations. Taiwan offers tax-based incentives (25% R&D deduction, 5% equipment deduction) rather than direct grants.",
    "Supply chain logistics remain an AZ disadvantage — wafers currently ship back to Taiwan for dicing, testing, and packaging. Long-term OSAT buildout in Arizona is planned.",
    "Both locations face infrastructure challenges: Arizona faces water scarcity in the desert Southwest; Taiwan faces green power deficits and water supply tension in Hsinchu.",
]

for i, finding in enumerate(findings):
    r = i + 3
    ws3.cell(row=r, column=1, value=i + 1)
    ws3.cell(row=r, column=1).font = Font(name="Calibri", bold=True, size=11, color="1F4E79")
    ws3.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="top")
    ws3.cell(row=r, column=2, value=finding)
    ws3.cell(row=r, column=2).font = normal_font
    ws3.cell(row=r, column=2).alignment = wrap_align
    ws3.row_dimensions[r].height = 45
    if i % 2 == 0:
        ws3.cell(row=r, column=1).fill = alt_fill
        ws3.cell(row=r, column=2).fill = alt_fill

# Save
output_path = "/workspace/TSMC_Phoenix_vs_Taiwan_Fab_Costs.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
