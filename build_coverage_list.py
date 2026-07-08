"""Generate the Semiconductor & Data Center coverage universe spreadsheet.

Produces:
  - semi_datacenter_coverage.xlsx  (formatted, multi-sheet)
  - semi_datacenter_coverage.csv   (flat, portable)

Market caps are approximate and rounded, as of ~Jul 8, 2026. They are meant for
coverage-scoping purposes only and are NOT investment advice. Refresh from a live
data source before use.
"""

import csv
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

AS_OF = "2026-07-08"

# Columns:
# name, ticker, listing, subsector, region, approx market cap ($B),
# recent IPO/listing, IPO date, notes
COVERAGE = [
    # ---- Data Center: Colocation / REITs ----
    ("GDS Holdings", "GDS / 9698.HK", "Nasdaq / HKEX", "Data Center - Colocation", "China", 6.7, "No", "2016-11 (Nasdaq)",
     "Carrier-neutral hyperscale colo leader in China; AI-driven wholesale demand; DayOne overseas spin."),
    ("VNET Group (21Vianet)", "VNET", "Nasdaq", "Data Center - Colocation", "China", 2.4, "No", "2011-04",
     "Carrier-neutral IDC; wholesale (Hyperscale 2.0) growth + retail; REIT/asset monetization optionality."),
    ("Equinix", "EQIX", "Nasdaq", "Data Center - REIT", "Global/US", 75.0, "No", "2000-08",
     "Largest global interconnection/colo REIT; xScale hyperscale JV; AI inference footprint."),
    ("Digital Realty Trust", "DLR", "NYSE", "Data Center - REIT", "Global/US", 55.0, "No", "2004-11",
     "Global wholesale + interconnection (PlatformDIGITAL); large hyperscale backlog."),
    ("Iron Mountain", "IRM", "NYSE", "Data Center - REIT", "Global/US", 30.0, "No", "1996-02",
     "Records/storage pivoting to fast-growing data center (ALM + colo) business."),
    # ---- Data Center: AI Neoclouds / GPU cloud ----
    ("CoreWeave", "CRWV", "Nasdaq", "Data Center - AI Neocloud", "US", 45.0, "Yes", "2025-03",
     "GPU cloud specialist; large hyperscaler backlog (~$99B); 'between the models and the silicon'."),
    ("Nebius Group", "NBIS", "Nasdaq", "Data Center - AI Neocloud", "Europe/US", 51.0, "Yes", "2024-10 (relisting)",
     "AI-focused cloud spun out of Yandex; hyperscale AI capacity; ~684% YoY revenue growth."),
    ("IREN Limited", "IREN", "Nasdaq", "Data Center - AI Neocloud", "Australia/US", 15.0, "No", "2021-11",
     "Vertically integrated AI datacenter/HPC (ex-bitcoin miner) with owned power sites."),
    ("Applied Digital", "APLD", "Nasdaq", "Data Center - AI Neocloud", "US", 6.0, "No", "2022-04",
     "HPC/AI datacenter developer leasing capacity to hyperscalers/neoclouds."),
    # ---- Data Center Infrastructure (power/cooling/interconnect) ----
    ("Vertiv Holdings", "VRT", "NYSE", "Data Center - Infrastructure", "Global/US", 50.0, "No", "2020-02 (SPAC)",
     "Power, thermal (liquid cooling), and rack infrastructure for AI data centers."),
    # ---- Semis: Logic / Foundry ----
    ("Taiwan Semiconductor (TSMC)", "TSM / 2330.TW", "NYSE / TWSE", "Semis - Foundry", "Taiwan", 1200.0, "No", "1997-10 (ADR)",
     "World's leading-edge foundry; CoWoS/advanced packaging bottleneck for AI."),
    ("Samsung Electronics", "005930.KS", "KRX", "Semis - Memory/Foundry (IDM)", "South Korea", 450.0, "No", "1975-06",
     "Memory (DRAM/NAND/HBM) + foundry; HBM ramp key AI catalyst."),
    ("SK hynix", "000660.KS", "KRX", "Semis - Memory (IDM)", "South Korea", 800.0, "No", "1996-12",
     "HBM leadership into AI supercycle; DRAM/NAND; NVIDIA HBM supplier."),
    ("Intel", "INTC", "Nasdaq", "Semis - IDM/Foundry", "US", 120.0, "No", "1971-10",
     "IDM turnaround; 18A ramp + Intel Foundry; datacenter CPU/AI accelerators."),
    ("GlobalFoundries", "GFS", "Nasdaq", "Semis - Foundry", "US", 25.0, "No", "2021-10",
     "Specialty (mature-node) foundry; auto/industrial/RF; US/EU capacity."),
    ("United Microelectronics (UMC)", "UMC / 2303.TW", "NYSE / TWSE", "Semis - Foundry", "Taiwan", 18.0, "No", "2000-09 (ADR)",
     "Mature/specialty foundry; 22/28nm; US-Japan-Taiwan footprint."),
    ("SMIC", "0981.HK / 688981.SS", "HKEX / STAR", "Semis - Foundry", "China", 80.0, "No", "2004-03 / 2020-07 (STAR)",
     "China's largest foundry; advanced-node localization amid export controls."),
    ("Hua Hong Semiconductor", "1347.HK / 688347.SS", "HKEX / STAR", "Semis - Foundry", "China", 15.0, "No", "2014-10 / 2023-08 (STAR)",
     "China #2 foundry; specialty/mature nodes; power/analog."),
    # ---- Semis: Fabless / AI compute / connectivity ----
    ("NVIDIA", "NVDA", "Nasdaq", "Semis - AI Accelerators", "US", 4200.0, "No", "1999-01",
     "Dominant AI GPU/systems + CUDA; primary beneficiary of AI datacenter capex."),
    ("Advanced Micro Devices (AMD)", "AMD", "Nasdaq", "Semis - CPU/GPU", "US", 300.0, "No", "1972 (NYSE)",
     "MI-series AI GPUs + EPYC datacenter CPUs; #2 AI accelerator challenger."),
    ("Broadcom", "AVGO", "Nasdaq", "Semis - Networking/Custom ASIC", "US", 1500.0, "No", "2009-08 (Avago)",
     "Custom AI ASIC (XPU) + networking (Tomahawk/Jericho); hyperscaler silicon partner."),
    ("Qualcomm", "QCOM", "Nasdaq", "Semis - Fabless", "US", 180.0, "No", "1991-12",
     "Mobile/edge SoC leader; auto + datacenter AI inference push."),
    ("MediaTek", "2454.TW", "TWSE", "Semis - Fabless", "Taiwan", 90.0, "No", "2001-07",
     "Mobile/connectivity SoCs; custom ASIC ambitions for AI datacenter."),
    ("Marvell Technology", "MRVL", "Nasdaq", "Semis - Networking/Custom ASIC", "US", 100.0, "No", "2000-06",
     "Custom AI silicon + optical DSP/networking for data centers."),
    ("Arm Holdings", "ARM", "Nasdaq", "Semis - IP/Design", "UK", 160.0, "Yes", "2023-09",
     "CPU IP; datacenter (Neoverse) traction; recent IPO."),
    ("Astera Labs", "ALAB", "Nasdaq", "Semis - Connectivity", "US", 63.0, "Yes", "2024-03",
     "Connectivity semis (PCIe/CXL/Ethernet retimers, Scorpio fabric) for AI racks; Nasdaq-100 add."),
    ("Credo Technology", "CRDO", "Nasdaq", "Semis - Connectivity", "US", 25.0, "No", "2022-01",
     "High-speed connectivity (AECs/SerDes/optical DSP) for AI data centers."),
    # ---- Semis: Memory ----
    ("Micron Technology", "MU", "Nasdaq", "Semis - Memory", "US", 180.0, "No", "1984-06",
     "Only US-based DRAM/NAND maker; HBM3E ramp; AI memory supercycle."),
    ("CXMT (ChangXin Memory)", "STAR (pending)", "SSE STAR (pending)", "Semis - Memory", "China", 150.0, "Yes", "2026-H2 (pending)",
     "China's largest DRAM maker; ~$4.3B STAR IPO approved Jun-2026 (2nd-largest STAR listing); HBM ambitions."),
    ("Kioxia Holdings", "285A.T", "TSE", "Semis - Memory (NAND)", "Japan", 20.0, "Yes", "2024-12",
     "NAND flash pure-play (ex-Toshiba Memory); recent Tokyo IPO."),
    ("SanDisk", "SNDK", "Nasdaq", "Semis - Memory (NAND)", "US", 25.0, "Yes", "2025-02 (spin-off)",
     "NAND/flash pure-play spun off from Western Digital in Feb-2025."),
    # ---- Semis: Equipment / Materials (WFE) ----
    ("ASML Holding", "ASML", "Nasdaq / Euronext", "Semis - Equipment (WFE)", "Netherlands", 400.0, "No", "1995-03",
     "Monopoly on EUV litho; critical enabler of leading-edge scaling."),
    ("Applied Materials", "AMAT", "Nasdaq", "Semis - Equipment (WFE)", "US", 180.0, "No", "1972-10",
     "Broadest WFE portfolio (deposition/etch/CMP); DRAM/HBM leverage."),
    ("Lam Research", "LRCX", "Nasdaq", "Semis - Equipment (WFE)", "US", 140.0, "No", "1984-05",
     "Etch/deposition leader; NAND/DRAM 3D scaling leverage."),
    ("KLA Corporation", "KLAC", "Nasdaq", "Semis - Equipment (WFE)", "US", 140.0, "No", "1980 (KLA)",
     "Process control/metrology & inspection near-monopoly; advanced packaging."),
    ("Tokyo Electron", "8035.T", "TSE", "Semis - Equipment (WFE)", "Japan", 150.0, "No", "1980s",
     "Top-tier WFE (coater/developer, etch, deposition); Japan flagship."),
    ("NAURA Technology", "002371.SZ", "SZSE", "Semis - Equipment (WFE)", "China", 55.0, "No", "2010-06",
     "China's largest domestic WFE maker; etch/deposition/thermal; localization play."),
    ("AMEC (Advanced Micro-Fabrication)", "688012.SS", "SSE STAR", "Semis - Equipment (WFE)", "China", 30.0, "No", "2019-07 (STAR)",
     "China etch equipment leader; MOCVD; leading-edge localization."),
]

HEADERS = [
    "Company", "Ticker", "Listing", "Subsector", "Region",
    "Approx. Market Cap ($B)", "Recent IPO/Listing", "IPO/Listing Date", "Notes / Thesis",
]


def write_csv(path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HEADERS)
        for row in COVERAGE:
            w.writerow(row)


def write_xlsx(path):
    wb = openpyxl.Workbook()

    # ---- Sheet 1: Coverage universe ----
    ws = wb.active
    ws.title = "Coverage Universe"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    ipo_fill = PatternFill("solid", fgColor="E2EFDA")  # highlight recent IPOs
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title_font = Font(bold=True, size=14)
    ws["A1"] = "Semiconductor & Data Center - Proposed Coverage Universe"
    ws["A1"].font = title_font
    ws["A2"] = (f"As of {AS_OF}. Market caps are approximate/rounded for scoping only "
                "(not investment advice). Recent IPOs/listings highlighted in green.")
    ws["A2"].font = Font(italic=True, size=9, color="808080")
    ws.append([])  # row 3 spacer

    header_row_idx = 4
    ws.append(HEADERS)
    for c, _ in enumerate(HEADERS, start=1):
        cell = ws.cell(row=header_row_idx, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        cell.border = border

    for row in sorted(COVERAGE, key=lambda r: (r[3], -r[5])):
        ws.append(row)
        r = ws.max_row
        recent = row[6] == "Yes"
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 9))
            if recent:
                cell.fill = ipo_fill
        ws.cell(row=r, column=6).number_format = "#,##0"
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="right", vertical="top")

    ws.freeze_panes = f"A{header_row_idx + 1}"
    ws.auto_filter.ref = f"A{header_row_idx}:I{ws.max_row}"

    widths = [30, 20, 20, 34, 14, 18, 14, 18, 70]
    for i, wdt in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = wdt

    # ---- Sheet 2: Subsector summary ----
    ws2 = wb.create_sheet("Subsector Summary")
    subsectors = {}
    for row in COVERAGE:
        sub = row[3]
        subsectors.setdefault(sub, {"count": 0, "cap": 0.0, "ipos": 0})
        subsectors[sub]["count"] += 1
        subsectors[sub]["cap"] += row[5]
        if row[6] == "Yes":
            subsectors[sub]["ipos"] += 1

    ws2.append(["Subsector", "# Names", "Aggregate Approx. Mkt Cap ($B)", "# Recent IPOs"])
    for c in range(1, 5):
        cell = ws2.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border
    for sub in sorted(subsectors):
        d = subsectors[sub]
        ws2.append([sub, d["count"], round(d["cap"], 1), d["ipos"]])
        r = ws2.max_row
        for c in range(1, 5):
            ws2.cell(row=r, column=c).border = border
        ws2.cell(row=r, column=3).number_format = "#,##0"

    total_names = len(COVERAGE)
    total_cap = round(sum(r[5] for r in COVERAGE), 1)
    total_ipos = sum(1 for r in COVERAGE if r[6] == "Yes")
    ws2.append(["TOTAL", total_names, total_cap, total_ipos])
    r = ws2.max_row
    for c in range(1, 5):
        cell = ws2.cell(row=r, column=c)
        cell.font = Font(bold=True)
        cell.border = border
    ws2.cell(row=r, column=3).number_format = "#,##0"

    for i, wdt in enumerate([34, 12, 30, 16], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = wdt
    ws2.freeze_panes = "A2"

    wb.save(path)


if __name__ == "__main__":
    write_csv("semi_datacenter_coverage.csv")
    write_xlsx("semi_datacenter_coverage.xlsx")
    print(f"Coverage names: {len(COVERAGE)}")
    print(f"Recent IPOs/listings: {sum(1 for r in COVERAGE if r[6] == 'Yes')}")
    print(f"Names with mkt cap > $1B: {sum(1 for r in COVERAGE if r[5] > 1)}")
    print(f"Generated on {date.today().isoformat()}")
