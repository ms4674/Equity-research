#!/usr/bin/env python3
"""
Generate a PDF explaining the Snap Inc. financial model forecast methodology.
"""

from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 8, "Snap Inc. (SNAP US Equity) - Financial Model: Forecast Methodology & Sources", align="C")
        self.ln(4)
        self.set_draw_color(47, 84, 150)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

    def section_title(self, title):
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(47, 84, 150)
        self.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(47, 84, 150)
        self.set_line_width(0.4)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def subsection_title(self, title):
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(47, 84, 150)
        self.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def sub_subsection_title(self, title):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(60, 60, 60)
        self.cell(0, 7, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def body_text(self, text):
        self.set_font("Helvetica", "", 10)
        self.set_text_color(30, 30, 30)
        self.multi_cell(0, 5.5, text)
        self.ln(2)

    def bullet(self, text, indent=10):
        x = self.get_x()
        self.set_font("Helvetica", "", 10)
        self.set_text_color(30, 30, 30)
        self.cell(indent, 5.5, "")
        self.set_font("Helvetica", "B", 10)
        self.cell(3, 5.5, "- ")
        self.set_font("Helvetica", "", 10)
        self.multi_cell(0, 5.5, text)
        self.ln(1)

    def bold_bullet(self, label, text, indent=10):
        self.cell(indent, 5.5, "")
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(30, 30, 30)
        self.cell(self.get_string_width("- " + label + ": "), 5.5, "- " + label + ": ")
        self.set_font("Helvetica", "", 10)
        self.multi_cell(0, 5.5, text)
        self.ln(1)

    def code_block(self, text):
        self.set_font("Courier", "", 8.5)
        self.set_fill_color(240, 240, 245)
        self.set_text_color(40, 40, 40)
        x = self.get_x()
        # Calculate height needed
        lines = text.split("\n")
        block_height = len(lines) * 4.5 + 4
        if self.get_y() + block_height > 270:
            self.add_page()
        self.set_x(15)
        y_start = self.get_y()
        self.rect(14, y_start - 1, 182, block_height, "F")
        for line in lines:
            self.set_x(16)
            self.cell(0, 4.5, line, new_x="LMARGIN", new_y="NEXT")
        self.ln(4)
        self.set_font("Helvetica", "", 10)

    def callout_box(self, text):
        self.set_font("Helvetica", "I", 9.5)
        self.set_fill_color(255, 242, 204)
        self.set_text_color(100, 70, 0)
        self.set_draw_color(191, 143, 0)
        y = self.get_y()
        self.rect(12, y, 186, 14, "DF")
        self.set_xy(15, y + 2)
        self.multi_cell(180, 5, text)
        self.ln(6)
        self.set_text_color(30, 30, 30)


pdf = PDF()
pdf.alias_nb_pages()
pdf.set_auto_page_break(auto=True, margin=20)
pdf.add_page()

# ── Title page content ────────────────────────────────────────────────────────
pdf.set_font("Helvetica", "B", 22)
pdf.set_text_color(47, 84, 150)
pdf.ln(15)
pdf.cell(0, 12, "Snap Inc. (SNAP US Equity)", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "", 16)
pdf.set_text_color(80, 80, 80)
pdf.cell(0, 10, "Three-Statement Financial Model", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.cell(0, 10, "Forecast Methodology & Sources", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.ln(8)
pdf.set_draw_color(47, 84, 150)
pdf.set_line_width(0.8)
pdf.line(60, pdf.get_y(), 150, pdf.get_y())
pdf.ln(10)

pdf.set_font("Helvetica", "", 11)
pdf.set_text_color(60, 60, 60)
pdf.cell(0, 8, "Model Coverage:", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "", 10)
items = [
    "Historical: FY2019 - FY2024 (6 years annual + 8 quarters)",
    "Forecast: FY2025E - FY2028E (4 years annual + 8 quarters forward)",
    "9 Excel tabs: Revenue Build, Income Statement, Balance Sheet,",
    "Cash Flow & FCF Bridge, Quarterly Detail, Share Count & Dilution,",
    "Sensitivity Analysis, Valuation Cross-Checks, Assumptions & Drivers",
]
for item in items:
    pdf.cell(0, 6, item, align="C", new_x="LMARGIN", new_y="NEXT")

pdf.ln(20)
pdf.set_font("Helvetica", "I", 9)
pdf.set_text_color(128, 128, 128)
pdf.cell(0, 6, "February 2026", align="C", new_x="LMARGIN", new_y="NEXT")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1: SOURCES
# ══════════════════════════════════════════════════════════════════════════════
pdf.add_page()
pdf.section_title("1. Data Sources")

pdf.subsection_title("Historical Data (FY2019 - FY2024)")

pdf.body_text(
    "Historical data was populated from Snap's public SEC filings (10-K annual reports and "
    "10-Q quarterly reports). These figures are sourced from the model author's knowledge of "
    "reported financials and represent approximate values that are directionally accurate and "
    "reflect the correct magnitudes, trends, and inflection points in Snap's financial history."
)

pdf.body_text("Key historical data points and their basis:")

pdf.bold_bullet("Revenue", "$1.72B (2019) to $5.36B (2024) -- reflects the actual reported trajectory of Snap's advertising-driven topline.")
pdf.bold_bullet("DAU", "218M (2019) to 443M (2024) -- consistent with quarterly DAU figures reported in earnings releases.")
pdf.bold_bullet("Long-Term Debt", "$3.745B reflects the convertible note issuances completed in 2021.")
pdf.bold_bullet("Geographic Splits", "NA ~64%, EU ~20%, ROW ~16% -- based on disclosed geographic segment data in 10-K filings.")

pdf.ln(2)
pdf.callout_box(
    "IMPORTANT: There is no live data feed in this model. The historical figures were not pulled from Bloomberg, "
    "FactSet, Capital IQ, or SEC EDGAR. They should be cross-referenced against Snap's actual 10-K filings "
    "before use in any investment decision."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2: REVENUE FORECAST
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("2. Revenue Forecast Methodology")

pdf.body_text(
    "Revenue is modeled as a mechanical identity: Revenue = DAU x ARPU. This decomposition "
    "allows separate stress-testing of user growth and monetization assumptions. All revenue "
    "forecasts flow from these two input drivers."
)

pdf.subsection_title("2.1 DAU Growth (6.5% to 4.0%, decelerating)")

pdf.code_block(
    "FY2025E: 6.5%  ->  472M DAU\n"
    "FY2026E: 5.5%  ->  498M DAU\n"
    "FY2027E: 4.5%  ->  520M DAU\n"
    "FY2028E: 4.0%  ->  541M DAU"
)

pdf.body_text("Rationale:")

pdf.bullet("Snap's DAU growth has been structurally decelerating: ~22% (2020), ~20% (2021), ~18% (2022), ~10% (2023), ~7% (2024). The forecast continues this well-established trend.")
pdf.bullet("At ~443M DAU vs. an estimated 800M total addressable market, Snap is at ~59% penetration -- consistent with the deceleration phase of a logistic S-curve adoption model.")
pdf.bullet("Growth is increasingly concentrated in Rest of World markets where monetization is lower, further supporting deceleration in DAU's revenue contribution.")
pdf.bullet("The terminal 4% growth rate implies Snap continues adding ~20M+ DAU/year, driven by emerging markets and Snapchat+ subscriber growth.")

pdf.subsection_title("2.2 ARPU Growth (10% to 7%, decelerating)")

pdf.code_block(
    "FY2025E: 10.0%  ->  $13.30/user\n"
    "FY2026E:  9.0%  ->  $14.50/user\n"
    "FY2027E:  8.0%  ->  $15.66/user\n"
    "FY2028E:  7.0%  ->  $16.76/user"
)

pdf.body_text("Rationale:")

pdf.bullet("ARPU growth reflects continued improvements in Snap's advertising platform: better direct-response ad products, ML-based targeting and bidding optimization, and expansion of ad surfaces (Spotlight, AR Lenses, Snap Map).")
pdf.bullet("Management has consistently highlighted ad-platform improvements as the primary monetization lever, and FY2024 showed renewed ARPU acceleration after the 2022-2023 ad recession.")
pdf.bullet("Snap's ~$12/user ARPU remains far below Meta's ~$50+, providing a large theoretical runway. The 7% terminal growth reflects that closing the gap fully is unlikely given Snap's younger, less commercially-intent user base.")
pdf.bullet("The deceleration from 10% to 7% reflects diminishing marginal returns on ad load increases and competition for digital ad dollars.")

pdf.subsection_title("2.3 Geographic Mix")

pdf.body_text(
    "North America's share gradually declines from ~64% to ~60% as ROW markets grow faster in DAU terms. "
    "Europe stays roughly stable at ~20%. This reflects the empirical pattern of Snap growing users faster "
    "in emerging markets while monetization remains concentrated in NA."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3: COST STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("3. Cost Structure & Margin Assumptions")

pdf.subsection_title("3.1 Gross Margin (55.5% to 60.0%)")

pdf.code_block(
    "FY2024A: ~54.1%  (actual)\n"
    "FY2025E: 55.5%\n"
    "FY2026E: 57.0%\n"
    "FY2027E: 58.5%\n"
    "FY2028E: 60.0%"
)

pdf.body_text("Rationale:")

pdf.bullet("Gross margin improved from ~49% (2019) to ~54% (2024), driven by cloud infrastructure cost optimization (renegotiated contracts with Google Cloud/AWS) and higher-margin ad formats.")
pdf.bullet("The forecast assumes continued but decelerating improvement (~150bps/year) as Snap benefits from: (1) infrastructure scale efficiencies, (2) higher share of direct-response ads vs. brand, and (3) content cost discipline.")
pdf.bullet("The terminal 60% gross margin is conservative vs. Meta (~82%) but reflects Snap's structurally higher infrastructure costs for camera-first features, AR compute, and video-heavy content delivery.")

pdf.subsection_title("3.2 Operating Expenses (as % of Revenue)")

pdf.code_block(
    "                    FY2025E  FY2026E  FY2027E  FY2028E\n"
    "R&D % of Revenue:    30.0%    27.0%    25.0%    23.0%\n"
    "S&M % of Revenue:    21.0%    19.0%    18.0%    17.0%\n"
    "G&A % of Revenue:     9.5%     8.8%     8.2%     7.8%"
)

pdf.body_text("This is the core operating leverage thesis. Rationale by line:")

pdf.bold_bullet("R&D (30% to 23%)", "Snap cut absolute R&D from $2.0B (2022) to $1.8B (2024) through layoffs and restructuring. The forecast assumes headcount is flat to slightly growing, but as a percentage of a growing revenue base, it naturally declines. The 23% terminal ratio is still elevated vs. scaled peers, reflecting Snap's continued investment in AR/ML technology.")

pdf.bold_bullet("S&M (21% to 17%)", "Declining as brand awareness saturates in core markets and self-serve ad tools (Snap Ads Manager) reduce direct sales costs. Historical trend: fell from 35% of revenue in 2019 to 24% in 2024.")

pdf.bold_bullet("G&A (9.5% to 7.8%)", "Standard back-office leverage. G&A is largely fixed (legal, finance, facilities) and grows slower than revenue. Historical trend: fell from 21% (2019) to 11% (2024).")

pdf.subsection_title("3.3 Stock-Based Compensation (20% to 13% of Revenue)")

pdf.code_block(
    "FY2022A: ~33% of revenue (peak)\n"
    "FY2024A: ~24% of revenue\n"
    "FY2025E: 20.0%\n"
    "FY2026E: 17.0%\n"
    "FY2027E: 15.0%\n"
    "FY2028E: 13.0%"
)

pdf.body_text(
    "SBC peaked at ~$1.54B / ~33% of revenue in 2022 and has been declining both in absolute dollars and as a "
    "percentage of revenue. The forecast reflects continued discipline (fewer new grants, lower headcount growth, "
    "more performance-based vesting) but acknowledges Snap will remain a relatively heavy SBC user vs. non-tech "
    "companies. The 13% terminal rate is comparable to scaled software companies."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4: BELOW THE LINE
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("4. Below-the-Line Items")

pdf.subsection_title("4.1 Interest Expense ($90M to $75M, declining)")

pdf.body_text(
    "Snap carries ~$3.745B in convertible notes. The declining interest expense reflects the low coupon rates "
    "on these instruments (0.00% to 0.75%) with most of the expense being non-cash amortization of debt discount "
    "that runs off over time. No new debt issuances or maturities are assumed before 2028."
)

pdf.subsection_title("4.2 Tax Rate (5% to 12%, gradually rising)")

pdf.body_text(
    "Snap has accumulated ~$8B+ in federal net operating losses (NOLs). As the company turns GAAP-profitable, "
    "these NOLs will shield near-term income from cash taxes. The effective tax rate ramps gradually from 5% (2025E) "
    "to 12% (2028E) as NOLs are consumed, but remains well below the 21% statutory US corporate rate through the "
    "forecast horizon. This is a standard approach for companies transitioning from cumulative losses to profitability."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5: BALANCE SHEET
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("5. Balance Sheet Methodology")

pdf.body_text(
    "Balance sheet items are projected using historical ratio analysis rather than bottom-up builds, "
    "which is standard for a financial model at this level of granularity."
)

pdf.bold_bullet("Accounts Receivable", "Projected at ~80 days sales outstanding (DSO), consistent with FY2024 implied DSO of ~84 days. This reflects the standard payment terms in the digital advertising industry.")
pdf.bold_bullet("Accounts Payable", "Projected at ~42 days payable outstanding (DPO) on cost of revenue, consistent with historical ratios.")
pdf.bold_bullet("Accrued Liabilities", "Held at ~12% of revenue, reflecting employee compensation accruals, content partner payables, and other operating accruals.")
pdf.bold_bullet("PP&E", "Roughly flat at ~$375-390M because annual CapEx (~$95-110M) approximately equals depreciation. Snap is in maintenance-mode on physical assets, with most infrastructure leased from cloud providers.")
pdf.bold_bullet("Long-Term Debt", "Held constant at $3.745B. No maturities fall within the forecast period, and no new issuances are assumed.")
pdf.bold_bullet("Intangibles", "Declining from $360M to $280M reflecting amortization of acquired intangibles with no major M&A assumed.")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6: CASH FLOW
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("6. Cash Flow & FCF Bridge")

pdf.body_text(
    "Cash from operations is built as: Net Income + D&A + SBC +/- Working Capital Changes. "
    "Free Cash Flow = CFO - CapEx. The GAAP Operating Loss to FCF bridge reconciles as follows:"
)

pdf.code_block(
    "  GAAP Operating Income (Loss)\n"
    "  (+) Depreciation & Amortization\n"
    "  (+) Stock-Based Compensation\n"
    "  = Adjusted EBITDA\n"
    "  (+/-) Working Capital Changes\n"
    "  (-) Cash Interest (net)\n"
    "  (-) Cash Taxes\n"
    "  (+/-) Other Adjustments\n"
    "  = Cash from Operations (CFO)\n"
    "  (-) Capital Expenditures\n"
    "  = Free Cash Flow (FCF)"
)

pdf.body_text(
    "Working capital changes are assumed to be a modest drag ($5-20M/year) as revenue growth drives "
    "incremental AR build. CapEx is modeled at $95-110M/year, reflecting maintenance-level infrastructure "
    "spending and modest investment in data center and office facilities."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7: SHARE COUNT
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("7. Share Count & Dilution")

pdf.body_text(
    "Diluted shares grow at ~1.0-1.2% annually on a net basis. The quarterly dilution schedule models:"
)

pdf.bold_bullet("RSU Vesting", "~11-13M shares per quarter, based on Snap's disclosed unvested RSU pool of ~120M+ shares with a typical 4-year vesting schedule.")
pdf.bold_bullet("Share Buybacks", "~5M shares per quarter (~$200M/year), reflecting a modest but consistent repurchase program. Snap initiated buybacks in 2023.")
pdf.bold_bullet("Option Exercises", "Declining pool of vested unexercised options (~15M shares, declining to ~8M by Q4-26E).")
pdf.bold_bullet("Net Dilution", "~1.2% annual dilution, which is moderate for a high-growth tech company but reflects the heavy SBC noted above.")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8: SENSITIVITY
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("8. Sensitivity Tables")

pdf.body_text(
    "Three two-way sensitivity tables are provided to stress-test the model on the most material drivers:"
)

pdf.bold_bullet("Table 1 -- DAU Growth x ARPU Growth to FY2026E Revenue", "These are the two direct revenue drivers. The table shows revenue outcomes ranging from conservative (2% DAU growth, 4% ARPU growth) to aggressive (8.5% DAU growth, 15% ARPU growth). The base case (5.5% DAU, 9% ARPU) is highlighted in green.")

pdf.bold_bullet("Table 2 -- Revenue Growth x EBITDA Margin to FY2026E EBITDA", "This captures the interplay between topline momentum and operating leverage. It answers: 'If revenue growth disappoints but margins expand faster (or vice versa), what does EBITDA look like?'")

pdf.bold_bullet("Table 3 -- WACC x Terminal Growth to Implied Share Price", "A standard DCF sensitivity grid. WACC ranges from 9-13% and terminal growth from 2-4%. This is the most important table for equity valuation, as terminal value typically represents 70-80% of DCF enterprise value.")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9: VALUATION
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("9. Valuation Cross-Checks")

pdf.subsection_title("9.1 DCF (Primary)")
pdf.body_text(
    "A standard 4-year explicit forecast + terminal value DCF using an 11% WACC (reflecting Snap's equity risk "
    "given its unprofitable history and competitive positioning) and a 3% terminal growth rate (slightly above "
    "long-term GDP growth, reflecting Snap's position in the secular-growth digital ad market)."
)

pdf.subsection_title("9.2 Cohort NPV (User-Level)")
pdf.body_text(
    "Each user cohort is valued using an 8-year discounted gross profit stream with 20% annual churn, 5% ARPU growth, "
    "56% gross margin, and $3.50 customer acquisition cost. This produces a Lifetime Value (LTV) and LTV/CAC ratio "
    "that can be multiplied by total DAU to derive an implied enterprise value. This cross-check validates whether "
    "the DCF output is consistent with bottom-up user economics."
)

pdf.subsection_title("9.3 Adoption S-Curve")
pdf.body_text(
    "DAU growth is mapped against an estimated 800M total addressable market. At ~59% penetration, Snap sits in the "
    "deceleration phase of a logistic S-curve, which is consistent with the 6.5% to 4.0% DAU growth assumptions "
    "in the forecast. This provides a structural check that the growth assumptions are not overly aggressive."
)

pdf.subsection_title("9.4 Unit Economics to EV Sanity Check")
pdf.body_text(
    "The model benchmarks Snap's EV/DAU, EV/Revenue, and EV/EBITDA multiples against social media peers "
    "(Meta, Pinterest, formerly Twitter). This provides a market-based reasonableness check on the implied "
    "valuation from the DCF and cohort NPV approaches."
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10: LIMITATIONS
# ══════════════════════════════════════════════════════════════════════════════
pdf.section_title("10. Key Limitations & Caveats")

pdf.body_text("Users of this model should be aware of the following limitations:")

pdf.bullet("Historical figures are approximate. They are sourced from training knowledge of Snap's SEC filings, not from a live data pull from Bloomberg, FactSet, Capital IQ, or SEC EDGAR. They should be validated against actual 10-K/10-Q filings before use in any investment decision.")

pdf.bullet("All forecasts are assumption-driven, not formula-linked. The yellow-highlighted cells in the Excel model are the key inputs. Changing them will not auto-recalculate downstream values because the model is generated by a Python script, not built with live Excel formulas. To make it dynamic, the forecast columns would need to be rebuilt with cell references in Excel.")

pdf.bullet("The model reflects a base-case scenario with moderate growth, margin expansion, and disciplined spending. It does not explicitly model downside scenarios (ad market recession, intensified TikTok competition, regulatory risk) or upside scenarios (AR/Spectacles inflection, Snapchat+ subscription breakthrough), though the sensitivity tables allow stress-testing.")

pdf.bullet("Geographic and product-level revenue splits are approximations. Snap does not break out revenue by product (Spotlight vs. Stories vs. Map) in its filings, so the model uses geography as the primary segmentation dimension.")

pdf.bullet("The convertible debt structure is simplified. Snap's converts have complex conversion features, make-whole provisions, and hedge/warrant overlays that affect diluted share count calculation. The model uses a simplified dilution approach.")

pdf.bullet("No scenario or Monte Carlo analysis is included. For a production-grade model, bull/bear/base scenarios and probability-weighted outcomes would strengthen the framework.")

# Save
output_path = "/workspace/Snap_Model_Forecast_Methodology.pdf"
pdf.output(output_path)
print(f"PDF saved to: {output_path}")
