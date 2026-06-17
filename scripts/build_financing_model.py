"""
Build an Excel model of AI-infrastructure financing (equity + debt) through 2030
for the major hyperscalers and neocloud / AI-native infrastructure companies.

The workbook breaks total debt financing into mutually-exclusive instrument
buckets (corporate bonds, bank & syndicated loans, private credit, insurance-linked
lending, securitization / ABS, convertibles, vendor & equipment-lease financing)
plus a memo line for the off-balance-sheet SPV / JV portion that those instruments
fund.

Figures are research estimates for the cumulative 2025-2030 buildout. They blend
disclosed transactions (2023-early 2026) with forward projections and are intended
for analytical/illustrative use, not as audited figures. Sources are listed on the
"Key Deals & Sources" sheet and in the methodology notes.

Run:  python scripts/build_financing_model.py
Output: AI_Infrastructure_Financing_2025-2030.xlsx  (repo root)
"""

from __future__ import annotations

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
NAVY = "1F2A44"
BLUE = "2E5A9C"
LIGHT_BLUE = "D6E1F2"
TEAL = "1B6E6E"
LIGHT_TEAL = "D2E8E8"
AMBER = "B5760A"
GREY = "5A5A5A"
LIGHT_GREY = "EEEEEE"
WHITE = "FFFFFF"
TOTAL_FILL = "FCEFC7"
HYPER_FILL = "E8EEF7"
NEO_FILL = "E6F2EF"

THIN = Side(style="thin", color="BFBFBF")
MED = Side(style="medium", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP_BORDER = Border(top=MED)

# ---------------------------------------------------------------------------
# DEBT INSTRUMENT BUCKETS  (mutually exclusive -> sum to total debt)
# ---------------------------------------------------------------------------
DEBT_COLS = [
    "Corporate bonds (public IG/HY)",
    "Bank & syndicated loans",
    "Alternate private credit",
    "Insurance-linked lending",
    "Securitization / ABS / CMBS",
    "Convertibles / equity-linked",
    "Vendor & equipment-lease financing",
]

# ---------------------------------------------------------------------------
# DATA  (cumulative 2025-2030, US$ billions)
#   each entry: segment, equity/internal funding, {debt bucket: $B}, SPV memo
# ---------------------------------------------------------------------------
# fmt: off
COMPANIES = [
    # ---------------- HYPERSCALERS ----------------
    ("Microsoft",  "Hyperscaler", 750,
        {"Corporate bonds (public IG/HY)": 130, "Bank & syndicated loans": 25,
         "Alternate private credit": 30, "Insurance-linked lending": 20,
         "Securitization / ABS / CMBS": 10, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 35}, 70),
    ("Amazon (AWS)", "Hyperscaler", 950,
        {"Corporate bonds (public IG/HY)": 200, "Bank & syndicated loans": 25,
         "Alternate private credit": 35, "Insurance-linked lending": 20,
         "Securitization / ABS / CMBS": 10, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 30}, 50),
    ("Alphabet (Google)", "Hyperscaler", 931,
        {"Corporate bonds (public IG/HY)": 160, "Bank & syndicated loans": 15,
         "Alternate private credit": 25, "Insurance-linked lending": 12,
         "Securitization / ABS / CMBS": 8, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 10}, 30),
    ("Meta Platforms", "Hyperscaler", 487,
        {"Corporate bonds (public IG/HY)": 150, "Bank & syndicated loans": 15,
         "Alternate private credit": 110, "Insurance-linked lending": 70,
         "Securitization / ABS / CMBS": 35, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 20}, 180),
    ("Oracle (OCI)", "Hyperscaler", 175,
        {"Corporate bonds (public IG/HY)": 120, "Bank & syndicated loans": 30,
         "Alternate private credit": 35, "Insurance-linked lending": 15,
         "Securitization / ABS / CMBS": 5, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 10}, 45),
    # ---------------- NEOCLOUDS / AI-NATIVE INFRA ----------------
    ("CoreWeave", "Neocloud / AI-native", 25,
        {"Corporate bonds (public IG/HY)": 10, "Bank & syndicated loans": 50,
         "Alternate private credit": 30, "Insurance-linked lending": 12,
         "Securitization / ABS / CMBS": 8, "Convertibles / equity-linked": 10,
         "Vendor & equipment-lease financing": 0}, 95),
    ("SpaceX / xAI (Colossus)", "Neocloud / AI-native", 140,
        {"Corporate bonds (public IG/HY)": 15, "Bank & syndicated loans": 25,
         "Alternate private credit": 30, "Insurance-linked lending": 5,
         "Securitization / ABS / CMBS": 3, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 27}, 70),
    ("Nebius", "Neocloud / AI-native", 18,
        {"Corporate bonds (public IG/HY)": 2, "Bank & syndicated loans": 12,
         "Alternate private credit": 8, "Insurance-linked lending": 3,
         "Securitization / ABS / CMBS": 2, "Convertibles / equity-linked": 6,
         "Vendor & equipment-lease financing": 2}, 10),
    ("IREN", "Neocloud / AI-native", 10,
        {"Corporate bonds (public IG/HY)": 3, "Bank & syndicated loans": 8,
         "Alternate private credit": 8, "Insurance-linked lending": 3,
         "Securitization / ABS / CMBS": 3, "Convertibles / equity-linked": 5,
         "Vendor & equipment-lease financing": 0}, 8),
    ("Crusoe", "Neocloud / AI-native", 8,
        {"Corporate bonds (public IG/HY)": 2, "Bank & syndicated loans": 6,
         "Alternate private credit": 9, "Insurance-linked lending": 3,
         "Securitization / ABS / CMBS": 2, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 3}, 12),
    ("Lambda", "Neocloud / AI-native", 8,
        {"Corporate bonds (public IG/HY)": 1, "Bank & syndicated loans": 8,
         "Alternate private credit": 5, "Insurance-linked lending": 2,
         "Securitization / ABS / CMBS": 1, "Convertibles / equity-linked": 1,
         "Vendor & equipment-lease financing": 0}, 10),
    ("Fluidstack", "Neocloud / AI-native", 4,
        {"Corporate bonds (public IG/HY)": 1, "Bank & syndicated loans": 6,
         "Alternate private credit": 8, "Insurance-linked lending": 2,
         "Securitization / ABS / CMBS": 1, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 2}, 14),
    ("Nscale", "Neocloud / AI-native", 5,
        {"Corporate bonds (public IG/HY)": 1, "Bank & syndicated loans": 6,
         "Alternate private credit": 5, "Insurance-linked lending": 1,
         "Securitization / ABS / CMBS": 1, "Convertibles / equity-linked": 0,
         "Vendor & equipment-lease financing": 1}, 8),
    ("Other neoclouds & AI-native DCs", "Neocloud / AI-native", 20,
        {"Corporate bonds (public IG/HY)": 8, "Bank & syndicated loans": 18,
         "Alternate private credit": 18, "Insurance-linked lending": 6,
         "Securitization / ABS / CMBS": 6, "Convertibles / equity-linked": 2,
         "Vendor & equipment-lease financing": 2}, 25),
]

# Annual AI-infrastructure capex projections (US$B); cumulative ~= total financing
CAPEX = {
    "Microsoft":          [95, 140, 165, 185, 200, 215],
    "Amazon (AWS)":       [100, 200, 225, 240, 250, 255],
    "Alphabet (Google)":  [91, 180, 200, 215, 230, 245],
    "Meta Platforms":     [72, 125, 150, 165, 180, 195],
    "Oracle (OCI)":       [25, 50, 65, 75, 85, 90],
    "Neoclouds / AI-native (aggregate)": [70, 110, 125, 115, 120, 126],
}
CAPEX_YEARS = [2025, 2026, 2027, 2028, 2029, 2030]

# Notable transactions feeding the estimates
DEALS = [
    ("Meta Platforms", "Hyperion JV (Beignet Investor LLC) — Blue Owl 80% / Meta 20%; $27.3bn A+ private bond led by PIMCO (~$18bn) & BlackRock; off-B/S SPV w/ residual-value guarantee", "2025", "$27.3bn debt", "SPV / private credit / insurance", "Meta IR; FT; S&P"),
    ("Meta Platforms", "$30bn multi-tranche IG corporate bond (largest 2025 IG deal), maturities to 2063", "2025", "$30bn", "Corporate bonds", "BofA; MUFG"),
    ("Amazon (AWS)", "~$54bn USD + €14.5bn (~$16.8bn) IG bonds (near-record), ~4x oversubscribed", "2026", "~$71bn", "Corporate bonds", "CNA; Reuters"),
    ("Amazon (AWS)", "$15bn IG bond", "2025", "$15bn", "Corporate bonds", "CNA"),
    ("Alphabet (Google)", "~$32bn multi-currency bond incl. rare 100-yr tranche; plus $17.5bn Nov 2025", "2025-26", "~$50bn", "Corporate bonds", "ConstructConnect; CNA"),
    ("Oracle (OCI)", "$25bn 8-part IG bond (Feb 2026) within $45-50bn CY26 debt+equity plan; $20bn ATM + mandatory convertible preferred (equity)", "2026", "$25bn debt", "Corporate bonds + equity", "Oracle IR; IFR"),
    ("Oracle (OCI)", "$18bn 6-part IG bond (Sep 2025)", "2025", "$18bn", "Corporate bonds", "Reuters/CNA"),
    ("Oracle (OCI)", "Project-level SPV financing: ~$16.3bn Michigan campus; further Texas / Wisconsin deals", "2025-26", "~$25bn+", "SPV / project finance", "Press reports"),
    ("Microsoft", "Finance-lease driven build (~$108bn uncommenced finance leases as of 9/30/24; commence FY25-FY30)", "2024-30", "$100bn+ leases", "Vendor / lease / SPV", "MSFT 10-Q; CNBC"),
    ("CoreWeave", "DDTL 4.0 — $8.5bn delayed-draw term loan; first IG-rated (Moody's A3 / DBRS A-low) GPU-backed financing; SOFR+225 / ~5.9%", "2026", "$8.5bn", "Bank / GPU-backed loan", "CoreWeave IR"),
    ("CoreWeave", "$7.5bn debt facility led by Blackstone & Magnetar (largest private debt deal at the time)", "2024", "$7.5bn", "Private credit / loan", "Blackstone"),
    ("CoreWeave", "$2.7bn 1.75% convertible notes; ~$30bn total debt by early 2026; ~$28bn debt+equity raised in 12 mos", "2025-26", "$2.7bn conv.", "Convertibles", "thedig; CoreWeave"),
    ("CoreWeave", "$2bn strategic equity from NVIDIA; IPO Mar 2025 raised ~$1.5bn", "2025-26", "$3.5bn equity", "Equity", "CNBC; CoreWeave"),
    ("SpaceX / xAI", "SpaceX acquired xAI effective Feb 2, 2026; Colossus I & II now SpaceX-owned (vertically-integrated AI platform; Terafab chip JV w/ Tesla & Intel)", "2026", "merger", "Equity / M&A", "SpaceX S-1"),
    ("SpaceX / xAI", "SpaceX AI capex $12.7bn in 2025 (~61% of total) + $7.7bn in Q1 2026; AI buildout funded by Starlink FCF, private raises, planned IPO", "2025-26", "$20bn+ capex", "Internal / equity", "SpaceX S-1; The Verge"),
    ("SpaceX / xAI", "Anthropic compute contract: $1.25bn/month (~$45bn) through May 2029 for Colossus I & II — anchor revenue underwriting infra financing", "2026-29", "~$45bn", "Take-or-pay (demand)", "SpaceX S-1; The Verge"),
    ("SpaceX / xAI", "xAI $10bn raise ($5bn secured notes/term loans incl. Apollo $5bn facility + $5bn equity) via Morgan Stanley", "2025", "$10bn", "Loans + equity", "CNBC"),
    ("SpaceX / xAI", "Colossus 2 SPV (Valor Equity): ~$12.5bn debt + ~$7.5bn equity (NVIDIA up to $2bn) to buy & lease GPUs", "2025-26", "~$20bn", "SPV / vendor lease", "The Information; Yahoo"),
    ("Nebius", "$3bn+ convertibles + bank lines; lighter balance sheet; Yandex spin-off cash ~$2.5bn; NVIDIA ~$2bn equity; Microsoft ~$7bn prepay", "2025-26", "$3bn+", "Convertibles", "frankk research"),
    ("Lambda", "$500m GPU-backed facility from Macquarie", "2025", "$0.5bn", "Bank / GPU-backed loan", "theaiinsider"),
    ("Nscale", "$1.4bn term loan", "2026", "$1.4bn", "Bank loan", "theaiinsider"),
    ("Fluidstack", "Facility backed by ~$6.7bn Google-contracted revenue (Google lease backstop)", "2026", "~$6.7bn", "SPV / private credit", "theaiinsider"),
    ("Applied Digital", "$1.59bn high-yield bond (7%) for CoreWeave-leased campus (Polaris Forge 1)", "2026", "$1.59bn", "Securitization / HY bond", "TNW"),
    ("Sector-wide", "HY/unrated AI-infra issuers raised ~$107bn funded+committed debt by ~May 2026 ($68.7bn neoclouds / $38.7bn AI-native DCs)", "to 2026", "~$107bn", "All debt", "Octus"),
    ("Sector-wide", "Insurance-linked platforms deployed ~$180bn into private credit in 2025 (~25% of global private credit AUM)", "2025", "~$180bn", "Insurance / private credit", "McKinsey; ABF Journal"),
    ("Sector-wide", "Goldman Sachs: ~$5.3tn AI + data-center capex 2025-2030; Morgan Stanley: ~$800bn private credit for AI DCs 2025-28", "2025-30", "$5.3tn capex", "Context", "Goldman; Morgan Stanley"),
]
# fmt: on


# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------
def title_cell(ws, cell, text, size=14):
    ws[cell] = text
    ws[cell].font = Font(bold=True, size=size, color=NAVY)


def header_row(ws, row, start_col, labels, fill=BLUE, color=WHITE, wrap=True):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=label)
        c.font = Font(bold=True, color=color, size=10)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = BORDER


def num_fmt(cell, fmt="#,##0"):
    cell.number_format = fmt
    cell.alignment = Alignment(horizontal="right")


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Sheet 1: Cover & Methodology
# ---------------------------------------------------------------------------
def build_cover(wb):
    ws = wb.active
    ws.title = "Cover & Methodology"
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 100})

    ws.merge_cells("B2:B2")
    title_cell(ws, "B2", "AI Infrastructure Financing — Hyperscalers & Neoclouds", 18)
    ws["B3"] = "Equity & debt funding to build AI infrastructure, cumulative 2025-2030"
    ws["B3"].font = Font(size=12, italic=True, color=GREY)
    ws["B4"] = "All figures in US$ billions unless noted."
    ws["B4"].font = Font(size=10, color=GREY)

    blocks = [
        ("What this workbook contains", BLUE, [
            "• Summary (by Company): equity / internal funding vs. total debt, with % debt-funded.",
            "• Debt by Type: total debt split into mutually-exclusive instrument buckets.",
            "• Capex Projections: annual AI-infrastructure capex 2025-2030 by company.",
            "• Key Deals & Sources: notable disclosed transactions underpinning the estimates.",
        ]),
        ("Scope & definitions", TEAL, [
            "• Hyperscalers: Microsoft, Amazon (AWS), Alphabet (Google), Meta, Oracle (OCI).",
            "• Neoclouds / AI-native infra: GPU-cloud & AI-native data-center builders (CoreWeave,",
            "   SpaceX/xAI, Nebius, IREN, Crusoe, Lambda, Fluidstack, Nscale, + an 'Other' aggregate).",
            "• SpaceX acquired xAI (effective 2 Feb 2026), so Colossus I & II are now SpaceX-owned; the two",
            "   are shown as one consolidated 'SpaceX / xAI' line to avoid double-counting the same clusters.",
            "• 'Equity / internal funding' = operating cash flow / free cash flow allocated to AI capex",
            "   PLUS external equity (IPOs, strategic stakes e.g. NVIDIA, ATM programs, mandatory",
            "   convertible PREFERRED treated as equity). Hyperscalers self-fund the majority via FCF.",
            "• 'Debt financing' = external borrowing, split by instrument (see below).",
        ]),
        ("Debt instrument buckets (mutually exclusive — they sum to total debt)", AMBER, [
            "1. Corporate bonds (public IG/HY) — on-balance-sheet senior notes sold in public markets.",
            "2. Bank & syndicated loans — term loans, revolvers, delayed-draw term loans (incl. GPU-backed DDTLs).",
            "3. Alternate private credit — direct lending / private placements by credit funds (Blackstone,",
            "   Apollo, Blue Owl, PIMCO, Magnetar, etc.), including privately-placed SPV/project debt.",
            "4. Insurance-linked lending — debt funded off insurer/annuity balance sheets (Athene/Apollo,",
            "   Global Atlantic/KKR, MetLife, Brookfield, etc.). NOTE: large overlap with private credit &",
            "   bonds; shown separately to size the insurance channel — see double-counting note below.",
            "5. Securitization / ABS / CMBS — secured data-center bonds & GPU-backed asset-backed securities.",
            "6. Convertibles / equity-linked debt — convertible notes (excludes mandatory convert. PREFERRED).",
            "7. Vendor & equipment-lease financing — chip/equipment leases & sale-leasebacks (e.g. NVIDIA",
            "   chip-lease SPVs, finance leases). Microsoft's large finance-lease book sits largely here.",
            "Memo: 'Off-B/S SPV / JV (memo)' flags the portion routed through special-purpose vehicles /",
            "   joint ventures (e.g. Meta-Hyperion, xAI Colossus 2, Oracle project SPVs). It OVERLAPS the",
            "   buckets above and is therefore NOT added into the debt total.",
        ]),
        ("Methodology & caveats", GREY, [
            "• Figures are RESEARCH ESTIMATES for the cumulative 2025-2030 buildout. They blend disclosed",
            "  transactions (2023-early 2026) with forward projections; they are illustrative, not audited.",
            "• Aggregate cross-check: total financing ≈ $5.25tn, consistent with Goldman Sachs' ~$5.3tn",
            "  2025-2030 AI + data-center capex estimate; aggregate debt ≈ $1.8tn, consistent with Morgan",
            "  Stanley / BofA projections of $1-2tn of debt financing for the buildout.",
            "• Insurance-linked lending double counts with private credit/bonds at the capital-source level;",
            "  to avoid inflating totals, each dollar of debt is classified ONCE by its primary instrument and",
            "  the insurance bucket captures only directly insurer-originated / insurance-balance-sheet debt.",
            "• Company splits reflect observed behavior: Microsoft/Alphabet/Amazon are FCF-heavy with",
            "  selective bonds; Meta & Oracle lean more on debt & SPVs; neoclouds are debt-heavy (asset-backed).",
            "• Built " + "with openpyxl. See Key Deals & Sources sheet for citations.",
        ]),
    ]
    r = 6
    for heading, color, lines in blocks:
        c = ws.cell(row=r, column=2, value=heading)
        c.font = Font(bold=True, size=12, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 20
        r += 1
        for line in lines:
            cc = ws.cell(row=r, column=2, value=line)
            cc.font = Font(size=10, color="333333")
            cc.alignment = Alignment(wrap_text=False, indent=1)
            r += 1
        r += 1

    disclaimer = ("Disclaimer: For analytical/illustrative purposes only. Not investment advice and not a "
                  "representation of any company's actual financing. Verify against primary filings before use.")
    c = ws.cell(row=r + 1, column=2, value=disclaimer)
    c.font = Font(size=9, italic=True, color="999999")
    c.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r + 1].height = 30


# ---------------------------------------------------------------------------
# Sheet 2: Summary (by Company)
# ---------------------------------------------------------------------------
def total_debt(debt: dict) -> int:
    return sum(debt.values())


def build_summary(wb):
    ws = wb.create_sheet("Summary (by Company)")
    ws.sheet_view.showGridLines = False
    title_cell(ws, "A1", "Summary — Equity vs. Debt Financing, Cumulative 2025-2030 (US$bn)")
    ws["A2"] = "Sorted within segment by total financing. % Debt = total debt / total financing."
    ws["A2"].font = Font(italic=True, size=9, color=GREY)

    cols = ["Company", "Segment", "Equity / internal\nfunding", "Total debt\nfinancing",
            "Total financing", "% Debt-\nfunded", "Off-B/S SPV/JV\n(memo)"]
    hdr = 4
    header_row(ws, hdr, 1, cols)
    ws.row_dimensions[hdr].height = 30

    r = hdr + 1
    seg_rows = {"Hyperscaler": [], "Neocloud / AI-native": []}

    def write_company(name, seg, equity, debt, spv, fill):
        nonlocal r
        td = total_debt(debt)
        tot = equity + td
        ws.cell(row=r, column=1, value=name).font = Font(bold=False, size=10)
        ws.cell(row=r, column=2, value=seg)
        for col, val in [(3, equity), (4, td), (5, tot), (7, spv)]:
            cc = ws.cell(row=r, column=col, value=val)
            num_fmt(cc)
        pct = ws.cell(row=r, column=6, value=td / tot if tot else 0)
        pct.number_format = "0%"
        pct.alignment = Alignment(horizontal="right")
        for col in range(1, 8):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDER
            if cell.fill.patternType is None:
                cell.fill = PatternFill("solid", fgColor=fill)
        seg_rows[seg].append(r)
        r += 1

    # group + section headers
    for seg, fill in [("Hyperscaler", HYPER_FILL), ("Neocloud / AI-native", NEO_FILL)]:
        sect = ws.cell(row=r, column=1, value=seg.upper())
        sect.font = Font(bold=True, color=WHITE, size=10)
        for col in range(1, 8):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=NAVY)
        r += 1
        members = [c for c in COMPANIES if c[1] == seg]
        members.sort(key=lambda c: c[2] + total_debt(c[3]), reverse=True)
        for name, s, equity, debt, spv in members:
            write_company(name, s, equity, debt, spv, fill)
        # subtotal
        sub = ws.cell(row=r, column=1, value=f"{seg} subtotal")
        sub.font = Font(bold=True, size=10, color=NAVY)
        rows = seg_rows[seg]
        for col, letter in [(3, "C"), (4, "D"), (5, "E"), (7, "G")]:
            cc = ws.cell(row=r, column=col,
                         value=f"=SUM({letter}{rows[0]}:{letter}{rows[-1]})")
            num_fmt(cc)
            cc.font = Font(bold=True, color=NAVY)
        pc = ws.cell(row=r, column=6, value=f"=D{r}/E{r}")
        pc.number_format = "0%"
        pc.font = Font(bold=True, color=NAVY)
        pc.alignment = Alignment(horizontal="right")
        for col in range(1, 8):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=LIGHT_GREY)
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    # grand total
    gt = ws.cell(row=r, column=1, value="GRAND TOTAL — all companies")
    gt.font = Font(bold=True, size=11, color=WHITE)
    first_company = hdr + 2
    for col, letter in [(3, "C"), (4, "D"), (5, "E"), (7, "G")]:
        # sum only company rows (exclude section/subtotal rows) via explicit cell list
        comp_rows = [rr for lst in seg_rows.values() for rr in lst]
        ref = ",".join(f"{letter}{rr}" for rr in comp_rows)
        cc = ws.cell(row=r, column=col, value=f"=SUM({ref})")
        num_fmt(cc)
        cc.font = Font(bold=True, color=WHITE)
    pc = ws.cell(row=r, column=6, value=f"=D{r}/E{r}")
    pc.number_format = "0%"
    pc.font = Font(bold=True, color=WHITE)
    pc.alignment = Alignment(horizontal="right")
    for col in range(1, 8):
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=BLUE)
        ws.cell(row=r, column=col).border = BORDER
    grand_row = r

    set_widths(ws, {"A": 32, "B": 20, "C": 14, "D": 13, "E": 14, "F": 11, "G": 15})
    ws.freeze_panes = "A5"

    # Chart: equity vs debt by company
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Equity vs. Debt Financing by Company (2025-2030, $bn)"
    chart.height = 9
    chart.width = 26
    comp_rows = [rr for lst in seg_rows.values() for rr in lst]
    # build contiguous data range: company rows are contiguous except section/subtotal rows.
    # Simpler: reference the whole block and rely on user; instead place a helper table.
    # We'll create a clean helper table to the right for the chart.
    hx = 9  # column I
    ws.cell(row=hdr, column=hx, value="Company").font = Font(bold=True, size=9, color=GREY)
    ws.cell(row=hdr, column=hx + 1, value="Equity").font = Font(bold=True, size=9, color=GREY)
    ws.cell(row=hdr, column=hx + 2, value="Debt").font = Font(bold=True, size=9, color=GREY)
    hr = hdr + 1
    for name, s, equity, debt, spv in COMPANIES:
        ws.cell(row=hr, column=hx, value=name)
        ws.cell(row=hr, column=hx + 1, value=equity).number_format = "#,##0"
        ws.cell(row=hr, column=hx + 2, value=total_debt(debt)).number_format = "#,##0"
        hr += 1
    data = Reference(ws, min_col=hx + 1, max_col=hx + 2, min_row=hdr, max_row=hr - 1)
    cats = Reference(ws, min_col=hx, max_col=hx, min_row=hdr + 1, max_row=hr - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.title = "US$bn"
    ws.add_chart(chart, f"A{grand_row + 3}")
    # de-emphasize helper table (data rows only; leave header labels as set)
    for row in ws.iter_rows(min_row=hdr + 1, max_row=hr - 1, min_col=hx, max_col=hx + 2):
        for cell in row:
            cell.font = Font(size=8, color="BBBBBB")


# ---------------------------------------------------------------------------
# Sheet 3: Debt by Type (the main breakdown)
# ---------------------------------------------------------------------------
def build_debt_breakdown(wb):
    ws = wb.create_sheet("Debt by Type")
    ws.sheet_view.showGridLines = False
    title_cell(ws, "A1", "Debt Financing by Instrument Type — Cumulative 2025-2030 (US$bn)")
    ws["A2"] = ("Instrument buckets are mutually exclusive and sum to 'Total debt'. "
                "'Off-B/S SPV/JV' is a memo (overlaps the buckets; not added).")
    ws["A2"].font = Font(italic=True, size=9, color=GREY)

    cols = ["Company", "Segment"] + DEBT_COLS + ["Total debt", "Off-B/S SPV/JV\n(memo)", "Equity /\ninternal"]
    hdr = 4
    header_row(ws, hdr, 1, cols)
    ws.row_dimensions[hdr].height = 46

    n_debt = len(DEBT_COLS)
    first_debt_col = 3
    last_debt_col = first_debt_col + n_debt - 1
    total_col = last_debt_col + 1            # Total debt
    spv_col = total_col + 1                   # SPV memo
    equity_col = spv_col + 1                  # Equity

    r = hdr + 1
    seg_rows = {"Hyperscaler": [], "Neocloud / AI-native": []}

    for seg, fill in [("Hyperscaler", HYPER_FILL), ("Neocloud / AI-native", NEO_FILL)]:
        sect = ws.cell(row=r, column=1, value=seg.upper())
        sect.font = Font(bold=True, color=WHITE, size=10)
        for col in range(1, equity_col + 1):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=NAVY)
        r += 1
        members = [c for c in COMPANIES if c[1] == seg]
        members.sort(key=lambda c: total_debt(c[3]), reverse=True)
        for name, s, equity, debt, spv in members:
            ws.cell(row=r, column=1, value=name).font = Font(size=10)
            ws.cell(row=r, column=2, value=s).font = Font(size=9, color=GREY)
            for i, key in enumerate(DEBT_COLS):
                cc = ws.cell(row=r, column=first_debt_col + i, value=debt[key])
                num_fmt(cc)
            tcell = ws.cell(row=r, column=total_col,
                            value=f"=SUM({get_column_letter(first_debt_col)}{r}:{get_column_letter(last_debt_col)}{r})")
            num_fmt(tcell)
            tcell.font = Font(bold=True)
            num_fmt(ws.cell(row=r, column=spv_col, value=spv))
            ws.cell(row=r, column=spv_col).font = Font(italic=True, color=GREY)
            num_fmt(ws.cell(row=r, column=equity_col, value=equity))
            ws.cell(row=r, column=equity_col).font = Font(color=TEAL)
            for col in range(1, equity_col + 1):
                cell = ws.cell(row=r, column=col)
                cell.border = BORDER
                if cell.fill.patternType is None:
                    cell.fill = PatternFill("solid", fgColor=fill)
            seg_rows[seg].append(r)
            r += 1
        # subtotal
        ws.cell(row=r, column=1, value=f"{seg} subtotal").font = Font(bold=True, color=NAVY)
        rows = seg_rows[seg]
        for col in range(first_debt_col, equity_col + 1):
            letter = get_column_letter(col)
            cc = ws.cell(row=r, column=col, value=f"=SUM({letter}{rows[0]}:{letter}{rows[-1]})")
            num_fmt(cc)
            cc.font = Font(bold=True, color=NAVY)
        for col in range(1, equity_col + 1):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=LIGHT_GREY)
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    # grand total
    ws.cell(row=r, column=1, value="GRAND TOTAL").font = Font(bold=True, color=WHITE, size=11)
    comp_rows = [rr for lst in seg_rows.values() for rr in lst]
    for col in range(first_debt_col, equity_col + 1):
        letter = get_column_letter(col)
        ref = ",".join(f"{letter}{rr}" for rr in comp_rows)
        cc = ws.cell(row=r, column=col, value=f"=SUM({ref})")
        num_fmt(cc)
        cc.font = Font(bold=True, color=WHITE)
    for col in range(1, equity_col + 1):
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=BLUE)
        ws.cell(row=r, column=col).border = BORDER
    grand_row = r

    # % of total debt row
    r += 1
    ws.cell(row=r, column=1, value="% of total debt").font = Font(bold=True, italic=True, color=AMBER)
    for col in range(first_debt_col, total_col):
        letter = get_column_letter(col)
        cc = ws.cell(row=r, column=col,
                     value=f"={letter}{grand_row}/${get_column_letter(total_col)}${grand_row}")
        cc.number_format = "0%"
        cc.font = Font(italic=True, color=AMBER)
        cc.alignment = Alignment(horizontal="right")
    cc = ws.cell(row=r, column=total_col, value=f"={get_column_letter(total_col)}{grand_row}/{get_column_letter(total_col)}{grand_row}")
    cc.number_format = "0%"
    cc.font = Font(italic=True, color=AMBER, bold=True)
    cc.alignment = Alignment(horizontal="right")

    widths = {"A": 30, "B": 19}
    for i in range(n_debt):
        widths[get_column_letter(first_debt_col + i)] = 15
    widths[get_column_letter(total_col)] = 11
    widths[get_column_letter(spv_col)] = 13
    widths[get_column_letter(equity_col)] = 11
    set_widths(ws, widths)
    ws.freeze_panes = "C5"

    # Chart: debt mix by instrument (grand total)
    chart = BarChart()
    chart.type = "bar"
    chart.title = "Aggregate Debt Mix by Instrument (2025-2030, $bn)"
    chart.height = 8
    chart.width = 22
    data = Reference(ws, min_col=first_debt_col, max_col=last_debt_col,
                     min_row=grand_row, max_row=grand_row)
    cats = Reference(ws, min_col=first_debt_col, max_col=last_debt_col, min_row=hdr, max_row=hdr)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, f"A{r + 3}")


# ---------------------------------------------------------------------------
# Sheet 4: Capex Projections
# ---------------------------------------------------------------------------
def build_capex(wb):
    ws = wb.create_sheet("Capex Projections")
    ws.sheet_view.showGridLines = False
    title_cell(ws, "A1", "AI-Infrastructure Capex Projections, 2025-2030 (US$bn)")
    ws["A2"] = ("Estimated annual capital expenditure directed at AI infrastructure (data centers, GPUs, "
                "networking, power). Cumulative ≈ total financing on the Summary sheet.")
    ws["A2"].font = Font(italic=True, size=9, color=GREY)

    cols = ["Company"] + [str(y) for y in CAPEX_YEARS] + ["Cumulative\n2025-30"]
    hdr = 4
    header_row(ws, hdr, 1, cols)
    ws.row_dimensions[hdr].height = 28

    r = hdr + 1
    hyper = ["Microsoft", "Amazon (AWS)", "Alphabet (Google)", "Meta Platforms", "Oracle (OCI)"]
    data_rows = []
    for name, vals in CAPEX.items():
        ws.cell(row=r, column=1, value=name).font = Font(size=10)
        for i, v in enumerate(vals):
            num_fmt(ws.cell(row=r, column=2 + i, value=v))
        last = 2 + len(vals) - 1
        cc = ws.cell(row=r, column=last + 1,
                     value=f"=SUM({get_column_letter(2)}{r}:{get_column_letter(last)}{r})")
        num_fmt(cc)
        cc.font = Font(bold=True)
        fill = HYPER_FILL if name in hyper else NEO_FILL
        for col in range(1, last + 2):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(row=r, column=col).border = BORDER
        data_rows.append(r)
        r += 1

    # total
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True, color=WHITE)
    last = 1 + len(CAPEX_YEARS) + 1
    for col in range(2, last + 1):
        letter = get_column_letter(col)
        cc = ws.cell(row=r, column=col, value=f"=SUM({letter}{data_rows[0]}:{letter}{data_rows[-1]})")
        num_fmt(cc)
        cc.font = Font(bold=True, color=WHITE)
    for col in range(1, last + 1):
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=BLUE)
        ws.cell(row=r, column=col).border = BORDER
    total_row = r

    set_widths(ws, {"A": 34, "B": 9, "C": 9, "D": 9, "E": 9, "F": 9, "G": 9, "H": 13})
    ws.freeze_panes = "B5"

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Annual AI-Infrastructure Capex by Company ($bn)"
    chart.height = 9
    chart.width = 24
    # companies as series (title from col A), years as categories -> from_rows
    data = Reference(ws, min_col=1, max_col=1 + len(CAPEX_YEARS),
                     min_row=data_rows[0], max_row=data_rows[-1])
    chart.add_data(data, titles_from_data=True, from_rows=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=1 + len(CAPEX_YEARS), min_row=hdr, max_row=hdr))
    chart.y_axis.title = "US$bn"
    ws.add_chart(chart, f"A{total_row + 3}")


# ---------------------------------------------------------------------------
# Sheet 5: Key Deals & Sources
# ---------------------------------------------------------------------------
def build_deals(wb):
    ws = wb.create_sheet("Key Deals & Sources")
    ws.sheet_view.showGridLines = False
    title_cell(ws, "A1", "Key Disclosed Transactions & Sources")
    ws["A2"] = ("Representative deals underpinning the estimates. Amounts as reported; '~' denotes "
                "approximate/announced figures. Not exhaustive.")
    ws["A2"].font = Font(italic=True, size=9, color=GREY)

    cols = ["Company / Group", "Transaction", "Period", "Amount", "Primary instrument", "Source(s)"]
    hdr = 4
    header_row(ws, hdr, 1, cols)
    ws.row_dimensions[hdr].height = 18

    r = hdr + 1
    for i, (co, desc, period, amt, instr, src) in enumerate(DEALS):
        ws.cell(row=r, column=1, value=co).font = Font(size=9, bold=True)
        ws.cell(row=r, column=2, value=desc).font = Font(size=9)
        ws.cell(row=r, column=3, value=period).font = Font(size=9)
        ws.cell(row=r, column=4, value=amt).font = Font(size=9)
        ws.cell(row=r, column=5, value=instr).font = Font(size=9)
        ws.cell(row=r, column=6, value=src).font = Font(size=9, color=BLUE)
        fill = WHITE if i % 2 == 0 else LIGHT_GREY
        for col in range(1, 7):
            cell = ws.cell(row=r, column=col)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = BORDER
        ws.row_dimensions[r].height = 30
        r += 1

    set_widths(ws, {"A": 20, "B": 64, "C": 10, "D": 14, "E": 22, "F": 22})
    ws.freeze_panes = "A5"

    # source list footer
    r += 1
    ws.cell(row=r, column=1, value="Selected sources:").font = Font(bold=True, size=10, color=NAVY)
    r += 1
    sources = [
        "Goldman Sachs Research — ~$5.3tn AI + data-center capex 2025-2030; private markets in DC financing.",
        "Morgan Stanley — ~$2tn capex 2025-28 with >$1tn debt; ~$800bn private credit for AI DCs.",
        "Bank of America / BofA Securities — hyperscaler bond issuance ($121bn 2025; ~$175bn 2026E).",
        "Moody's Ratings — hyperscaler capex (~$785bn 2026, ~$1tn 2027); $662bn off-B/S DC leases.",
        "McKinsey — ~$2.7tn US (5.2tn global) data-center capex by 2030; insurance-linked private credit.",
        "Octus — HY/unrated AI-infra debt (~$107bn to ~May 2026); ~14GW unfunded ≈ $344bn need.",
        "Company disclosures & press: Meta IR / FT / S&P (Hyperion); Oracle IR / IFR / Reuters; CoreWeave IR;",
        "  CNBC / WSJ / The Information (xAI); Blackstone; Apollo Academy; MUFG; CreditSights; CNA.",
    ]
    for s in sources:
        ws.cell(row=r, column=1, value=s).font = Font(size=9, color="444444")
        r += 1


# ---------------------------------------------------------------------------
def main():
    wb = Workbook()
    # force Excel/Sheets to recalculate all formulas when the file is opened
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    build_cover(wb)
    build_summary(wb)
    build_debt_breakdown(wb)
    build_capex(wb)
    build_deals(wb)

    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "AI_Infrastructure_Financing_2025-2030.xlsx")
    wb.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
