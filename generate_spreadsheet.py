import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

header_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
subheader_font = Font(name="Calibri", bold=True, size=11, color="1F4E79")
subheader_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
title_font = Font(name="Calibri", bold=True, size=14, color="1F4E79")
data_font = Font(name="Calibri", size=11)
money_font = Font(name="Calibri", size=11, bold=True, color="006100")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ── Sheet 1: AV Investments ──
ws1 = wb.active
ws1.title = "AV Investments"

ws1.merge_cells("A1:G1")
ws1["A1"].value = "Uber Autonomous Vehicle (AV) Investments — Past 5 Years (2021–2025)"
ws1["A1"].font = title_font
ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

inv_headers = [
    "Partner / Company",
    "Year",
    "Investment Amount",
    "Investment Type",
    "Uber Stake",
    "Deal Description",
    "Status",
]
for col, header in enumerate(inv_headers, 1):
    cell = ws1.cell(row=3, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border

investments = [
    [
        "Aurora Innovation",
        "2021 (closed Q1)",
        "$400 million",
        "Equity (part of ATG sale)",
        "~26%",
        "Uber sold its ATG self-driving unit to Aurora and simultaneously invested $400M. Uber CEO joined Aurora's board.",
        "Active — Aurora stake valued at ~$2.05B as of Q4 2024",
    ],
    [
        "Aurora Innovation (additional)",
        "2024 (Q4)",
        "$2.05 billion (market value of holding)",
        "Equity holding (325.9M shares)",
        "~19%",
        "Aurora represents ~42% of Uber's total equity holdings; second-largest position.",
        "Active — publicly traded (AUR)",
    ],
    [
        "Wayve",
        "2024 (August)",
        "Undisclosed (part of $1.05B Series C)",
        "Strategic equity investment",
        "Minority stake",
        "Investment into Wayve's Series C led by SoftBank. Partnership to integrate Wayve's AI driving tech into vehicles on Uber's platform.",
        "Active — L4 trials in London planned spring 2026",
    ],
    [
        "Avride",
        "2025 (October)",
        "Part of up to $375 million (with Nebius)",
        "Strategic investment + commercial commitments",
        "Undisclosed",
        "Funding to scale Avride's autonomous delivery robots and robotaxi fleet. Deployed on Uber Eats and UberX.",
        "Active — robotaxi launched in Dallas Dec 2025",
    ],
    [
        "Lucid Motors",
        "2025 (July)",
        "$300 million",
        "Strategic equity investment",
        "Undisclosed",
        "Investment to acquire 20,000+ Lucid Gravity SUVs for robotaxi deployment with Nuro's L4 tech over 6 years.",
        "Active — first launch planned late 2026",
    ],
    [
        "Nuro",
        "2025 (July)",
        "Multi-hundred million (>$300M, exact undisclosed)",
        "Strategic equity investment",
        "Board seat acquired",
        "Investment for Nuro's L4 self-driving system to be integrated into Lucid vehicles deployed on Uber's network.",
        "Active — first launch planned late 2026",
    ],
]

for row_idx, row_data in enumerate(investments, 4):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=value)
        cell.font = data_font
        cell.alignment = left_align
        cell.border = thin_border
        if col_idx == 3:
            cell.font = money_font

total_row = len(investments) + 5
ws1.merge_cells(f"A{total_row}:B{total_row}")
ws1.cell(row=total_row, column=1, value="ESTIMATED TOTAL DISCLOSED INVESTMENTS").font = Font(
    name="Calibri", bold=True, size=11
)
ws1.cell(row=total_row, column=3, value="~$3.1 billion+ (disclosed amounts)").font = Font(
    name="Calibri", bold=True, size=11, color="006100"
)
for c in range(1, 8):
    ws1.cell(row=total_row, column=c).border = thin_border
    ws1.cell(row=total_row, column=c).fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )

note_row = total_row + 2
ws1.merge_cells(f"A{note_row}:G{note_row}")
ws1.cell(
    row=note_row,
    column=1,
    value="Note: Uber sold its in-house AV unit (ATG) to Aurora in Dec 2020 / Q1 2021. Since then, Uber has pursued a platform strategy — partnering with and investing in multiple AV companies rather than developing its own self-driving technology.",
).font = Font(name="Calibri", size=10, italic=True, color="666666")
ws1.cell(row=note_row, column=1).alignment = left_align

col_widths_1 = [28, 18, 38, 32, 18, 65, 42]
for i, w in enumerate(col_widths_1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w


# ── Sheet 2: Alliances Without Investments ──
ws2 = wb.create_sheet(title="Alliances (No Investment)")

ws2.merge_cells("A1:F1")
ws2["A1"].value = "Uber AV Alliances & Partnerships — No Direct Financial Investment"
ws2["A1"].font = title_font
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")

alliance_headers = [
    "Alliance Partner",
    "Year Announced",
    "Deal Type",
    "Deal Description",
    "Markets / Cities",
    "Current Status",
]
for col, header in enumerate(alliance_headers, 1):
    cell = ws2.cell(row=3, column=col, value=header)
    cell.font = header_font
    cell.fill = PatternFill(start_color="C05028", end_color="C05028", fill_type="solid")
    cell.alignment = center_align
    cell.border = thin_border

alliances = [
    [
        "Motional (Hyundai/Aptiv JV)",
        "2022 (October)",
        "10-year commercial agreement",
        "Non-exclusive deal for autonomous ride-hail (UberX, Uber Comfort Electric) and Uber Eats delivery using Hyundai IONIQ 5 robotaxis.",
        "Las Vegas (launched Dec 2022); Los Angeles planned",
        "Active — though Motional paused operations in 2024 amid restructuring",
    ],
    [
        "Waymo (Alphabet)",
        "2023 (May); expanded 2024 (Sept)",
        "Multi-year strategic partnership",
        "Integration of Waymo One autonomous rides into Uber app. Uber manages fleet operations (cleaning, repair, depot). No financial investment by Uber.",
        "Phoenix (2023); Austin & Atlanta (early 2025)",
        "Active — Waymo provides 100K+ trips/week in Phoenix",
    ],
    [
        "May Mobility",
        "2025 (May)",
        "Strategic partnership",
        "Deployment of Toyota Sienna vehicles with May Mobility's MPDM autonomous tech on Uber's platform.",
        "Arlington, TX (end of 2025); more U.S. cities in 2026",
        "Active",
    ],
    [
        "Momenta",
        "2025 (May)",
        "Strategic agreement",
        "First European robotaxi deployment using Momenta's autonomous driving tech on Uber's network. Initially with safety operators.",
        "Europe (early 2026)",
        "Active — first international AV deployment outside US/China",
    ],
    [
        "Pony.ai",
        "2025 (May)",
        "Strategic partnership",
        "Deployment of Pony.ai's 7th-gen autonomous vehicles on Uber. Initial pilots with safety operators before fully autonomous operation.",
        "Middle East (late 2025); additional international markets",
        "Active",
    ],
    [
        "WeRide",
        "2025 (May)",
        "Strategic partnership (expanded)",
        "Expansion of existing partnership to 15 additional cities globally over 5 years. Already operating in Abu Dhabi.",
        "Abu Dhabi (live); Dubai; 15 more cities in Europe & beyond",
        "Active",
    ],
    [
        "Baidu (Apollo Go)",
        "2025 (July)",
        "Multi-year strategic partnership",
        "Deployment of thousands of Baidu autonomous vehicles on Uber globally (outside US & mainland China).",
        "Asia & Middle East (late 2025 onward)",
        "Active",
    ],
    [
        "NVIDIA",
        "2025 (October)",
        "Technology partnership",
        "Partnership to build an L4-ready mobility network supporting up to 100,000 AVs by 2027. NVIDIA providing compute platform for AV ecosystem.",
        "Global — targeting 28 cities by 2028",
        "Active",
    ],
    [
        "Zoox (Amazon)",
        "2026 (March)",
        "Strategic partnership",
        "Deployment of Zoox purpose-built robotaxis on Uber's platform.",
        "Las Vegas (Summer 2026); Los Angeles (mid-2027)",
        "Announced",
    ],
    [
        "Wayve & Nissan (Tokyo)",
        "2026 (March)",
        "MOU / Partnership",
        "Robotaxi service in Tokyo using Wayve's AI Driver in Nissan LEAF vehicles. (Note: Uber has a separate equity investment in Wayve.)",
        "Tokyo (pilot late 2026); London (planned)",
        "Announced — MOU signed",
    ],
]

for row_idx, row_data in enumerate(alliances, 4):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=value)
        cell.font = data_font
        cell.alignment = left_align
        cell.border = thin_border

summary_row = len(alliances) + 5
ws2.merge_cells(f"A{summary_row}:F{summary_row}")
ws2.cell(
    row=summary_row,
    column=1,
    value="Total Alliance Partners Without Direct Uber Investment: 10 (listed above). These are commercial / technology partnerships where Uber has not made a disclosed equity investment.",
).font = Font(name="Calibri", bold=True, size=11, color="C05028")
ws2.cell(row=summary_row, column=1).alignment = left_align

note2_row = summary_row + 2
ws2.merge_cells(f"A{note2_row}:F{note2_row}")
ws2.cell(
    row=note2_row,
    column=1,
    value="Note: Wayve appears in both sheets — Uber has a separate equity investment in Wayve (Sheet 1), while the Wayve & Nissan Tokyo robotaxi partnership is a distinct operational alliance listed here. Uber has 25+ total AV partnerships as of early 2026.",
).font = Font(name="Calibri", size=10, italic=True, color="666666")
ws2.cell(row=note2_row, column=1).alignment = left_align

col_widths_2 = [28, 22, 30, 70, 45, 50]
for i, w in enumerate(col_widths_2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# ── Sheet 3: Summary ──
ws3 = wb.create_sheet(title="Summary")

ws3.merge_cells("A1:D1")
ws3["A1"].value = "Uber AV Strategy Summary (2021–2025)"
ws3["A1"].font = title_font
ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")

summary_data = [
    ["Metric", "Value", "Notes", ""],
    [
        "Total Disclosed AV Investments",
        "~$3.1 billion+",
        "Includes Aurora ($400M initial + $2.05B holding), Lucid ($300M), Nuro (>$300M), Avride (part of $375M), Wayve (undisclosed)",
        "",
    ],
    [
        "Number of Investment Partners",
        "5",
        "Aurora, Wayve, Avride, Lucid Motors, Nuro",
        "",
    ],
    [
        "Number of Alliance Partners (no investment)",
        "10",
        "Motional, Waymo, May Mobility, Momenta, Pony.ai, WeRide, Baidu, NVIDIA, Zoox, Wayve/Nissan",
        "",
    ],
    [
        "Total AV Partnerships",
        "25+",
        "Per Uber disclosure as of early 2026",
        "",
    ],
    [
        "AV Cities Target (end of 2026)",
        "15 cities",
        "Autonomous rides available across major markets",
        "",
    ],
    [
        "Largest Single Investment",
        "Aurora Innovation (~$2.05B holding)",
        "~19% stake; 42% of Uber's total equity holdings",
        "",
    ],
    [
        "Strategic Shift",
        "Platform model (since 2021)",
        "Sold ATG to Aurora in Q1 2021; now partners with AV developers rather than building in-house",
        "",
    ],
]

for row_idx, row_data in enumerate(summary_data, 3):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws3.cell(row=row_idx, column=col_idx, value=value)
        cell.border = thin_border
        cell.alignment = left_align
        if row_idx == 3:
            cell.font = header_font
            cell.fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
        else:
            cell.font = data_font
            if col_idx == 2:
                cell.font = money_font

ws3.column_dimensions["A"].width = 38
ws3.column_dimensions["B"].width = 35
ws3.column_dimensions["C"].width = 80
ws3.column_dimensions["D"].width = 5

output_path = "/workspace/uber_av_investments_and_alliances.xlsx"
wb.save(output_path)
print(f"Spreadsheet saved to {output_path}")
