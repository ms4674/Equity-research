#!/usr/bin/env python3
"""
Meta Hyperscale & Neocloud Commitments — Excel Workbook Generator
Reads CSV data files and produces a formatted multi-tab Excel workbook.
"""

import csv
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, LineChart

DATA_DIR = "data"
OUTPUT_FILE = "meta_hyperscale_neocloud_commitments.xlsx"

HEADER_FILL = PatternFill(start_color="1B3A5C", end_color="1B3A5C", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUBHEADER_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="1B3A5C")
SUBTITLE_FONT = Font(bold=True, size=12, color="2E75B6")
TOTAL_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
TOTAL_FONT = Font(bold=True, size=11)
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MONEY_FMT = '#,##0.0"B"'
PCT_FMT = "0.0%"


def read_csv(filename):
    path = os.path.join(DATA_DIR, filename)
    with open(path, newline="") as f:
        reader = csv.reader(f)
        return list(reader)


def apply_header_style(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def apply_data_style(ws, start_row, end_row, max_col, money_cols=None):
    money_cols = money_cols or []
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FMT


def auto_width(ws, max_col, min_width=12, max_width=45):
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        longest = min_width
        for row in ws.iter_rows(min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    longest = max(longest, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[letter].width = longest


def write_rows(ws, data, start_row=1, money_cols=None):
    money_cols = money_cols or []
    for r_idx, row in enumerate(data):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=start_row + r_idx, column=c_idx + 1)
            if c_idx in money_cols and r_idx > 0:
                try:
                    cell.value = float(val)
                    cell.number_format = MONEY_FMT
                except (ValueError, TypeError):
                    cell.value = val
            elif r_idx > 0:
                try:
                    cell.value = float(val)
                except (ValueError, TypeError):
                    cell.value = val
            else:
                cell.value = val
    return start_row + len(data)


def create_capex_sheet(wb):
    ws = wb.create_sheet("Capex Trajectory", 0)
    data = read_csv("meta_capex_trajectory.csv")

    ws.cell(row=1, column=1, value="Meta Platforms — Capital Expenditure Trajectory").font = TITLE_FONT
    ws.merge_cells("A1:F1")

    end_row = write_rows(ws, data, start_row=3, money_cols=[1, 3])
    apply_header_style(ws, 3, len(data[0]))
    apply_data_style(ws, 4, end_row - 1, len(data[0]), money_cols=[2, 4])

    for r in range(4, end_row):
        pct_cell = ws.cell(row=r, column=3)
        if isinstance(pct_cell.value, (int, float)):
            pct_cell.value = pct_cell.value / 100
            pct_cell.number_format = PCT_FMT
        rev_pct = ws.cell(row=r, column=5)
        if isinstance(rev_pct.value, (int, float)):
            rev_pct.value = rev_pct.value / 100
            rev_pct.number_format = PCT_FMT

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Meta Capex ($B) — 2018 to 2028E"
    chart.y_axis.title = "Capex ($B)"
    chart.x_axis.title = "Year"
    cats = Reference(ws, min_col=1, min_row=4, max_row=end_row - 1)
    vals = Reference(ws, min_col=2, min_row=3, max_row=end_row - 1)
    chart.add_data(vals, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 22
    chart.height = 14
    ws.add_chart(chart, "A" + str(end_row + 2))

    auto_width(ws, len(data[0]))
    return ws


def create_neocloud_sheet(wb):
    ws = wb.create_sheet("Neocloud Commitments", 1)
    data = read_csv("meta_neocloud_commitments.csv")

    ws.cell(row=1, column=1, value="Meta — Third-Party Neocloud & Cloud Commitments").font = TITLE_FONT
    ws.merge_cells("A1:H1")

    end_row = write_rows(ws, data, start_row=3, money_cols=[1])
    apply_header_style(ws, 3, len(data[0]))
    apply_data_style(ws, 4, end_row - 1, len(data[0]), money_cols=[2])

    total_row = end_row
    ws.cell(row=total_row, column=1, value="TOTAL (if all signed/exercised)").font = TOTAL_FONT
    ws.cell(row=total_row, column=1).fill = TOTAL_FILL
    total_val = 0
    for r in range(4, end_row):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, (int, float)):
            total_val += v
    total_cell = ws.cell(row=total_row, column=2, value=total_val)
    total_cell.font = TOTAL_FONT
    total_cell.fill = TOTAL_FILL
    total_cell.number_format = MONEY_FMT
    for c in range(1, len(data[0]) + 1):
        ws.cell(row=total_row, column=c).border = THIN_BORDER
        ws.cell(row=total_row, column=c).fill = TOTAL_FILL

    notes_row = total_row + 2
    ws.cell(row=notes_row, column=1, value="Notes:").font = SUBTITLE_FONT
    ws.cell(row=notes_row + 1, column=1, value="• Nebius Phase 2 ($27B) supersedes Phase 1 ($3B); total Nebius = $27B, not $30B")
    ws.cell(row=notes_row + 2, column=1, value="• Oracle deal reported at ~$20B but still in negotiation as of Apr 2026")
    ws.cell(row=notes_row + 3, column=1, value="• Google Cloud amount is estimated; exact terms not publicly disclosed")
    ws.cell(row=notes_row + 4, column=1, value="• Adjusted total (netting Nebius overlap): ~$64-74B")

    auto_width(ws, len(data[0]))
    return ws


def create_datacenter_sheet(wb):
    ws = wb.create_sheet("Data Center Projects", 2)
    data = read_csv("meta_datacenter_projects.csv")

    ws.cell(row=1, column=1, value="Meta — Major Hyperscale Data Center Projects").font = TITLE_FONT
    ws.merge_cells("A1:I1")

    end_row = write_rows(ws, data, start_row=3, money_cols=[2, 3])
    apply_header_style(ws, 3, len(data[0]))
    apply_data_style(ws, 4, end_row - 1, len(data[0]), money_cols=[3])

    total_row = end_row
    ws.cell(row=total_row, column=1, value="TOTAL DISCLOSED").font = TOTAL_FONT
    ws.cell(row=total_row, column=1).fill = TOTAL_FILL
    total_inv = 0
    for r in range(4, end_row):
        v = ws.cell(row=r, column=3).value
        if isinstance(v, (int, float)):
            total_inv += v
    total_cell = ws.cell(row=total_row, column=3, value=total_inv)
    total_cell.font = TOTAL_FONT
    total_cell.fill = TOTAL_FILL
    total_cell.number_format = MONEY_FMT
    for c in range(1, len(data[0]) + 1):
        ws.cell(row=total_row, column=c).border = THIN_BORDER
        ws.cell(row=total_row, column=c).fill = TOTAL_FILL

    auto_width(ws, len(data[0]))
    return ws


def create_timeline_sheet(wb):
    ws = wb.create_sheet("Commitment Timeline", 3)
    data = read_csv("meta_commitment_timeline.csv")

    ws.cell(row=1, column=1, value="Meta — Chronological Commitment Timeline").font = TITLE_FONT
    ws.merge_cells("A1:F1")

    end_row = write_rows(ws, data, start_row=3, money_cols=[2])
    apply_header_style(ws, 3, len(data[0]))
    apply_data_style(ws, 4, end_row - 1, len(data[0]), money_cols=[3])

    category_fills = {
        "Owned Capex": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),
        "Neocloud": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
        "Third-Party Cloud": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
        "Off-Balance-Sheet": PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid"),
        "Data Center": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
    }

    for r in range(4, end_row):
        cat = ws.cell(row=r, column=5).value
        if cat in category_fills:
            for c in range(1, len(data[0]) + 1):
                ws.cell(row=r, column=c).fill = category_fills[cat]

    legend_row = end_row + 2
    ws.cell(row=legend_row, column=1, value="Color Legend:").font = SUBTITLE_FONT
    for i, (cat, fill) in enumerate(category_fills.items()):
        cell = ws.cell(row=legend_row + 1 + i, column=1, value=cat)
        cell.fill = fill
        cell.border = THIN_BORDER

    auto_width(ws, len(data[0]))
    return ws


def create_balance_sheet_sheet(wb):
    ws = wb.create_sheet("Balance Sheet Snapshot", 4)
    data = read_csv("meta_balance_sheet_snapshot.csv")

    ws.cell(row=1, column=1, value="Meta — Balance Sheet & Financial Snapshot").font = TITLE_FONT
    ws.merge_cells("A1:D1")

    end_row = write_rows(ws, data, start_row=3, money_cols=[1])
    apply_header_style(ws, 3, len(data[0]))
    apply_data_style(ws, 4, end_row - 1, len(data[0]), money_cols=[2])

    auto_width(ws, len(data[0]))
    return ws


def create_summary_sheet(wb):
    ws = wb.create_sheet("Consolidated Summary", 5)

    ws.cell(row=1, column=1, value="Meta — Consolidated Commitment Summary").font = TITLE_FONT
    ws.merge_cells("A1:E1")

    headers = ["Category", "Amount ($B)", "Timeframe", "Type", "Notes"]
    summary_data = [
        ["Owned Capex (2025 actual)", 72.2, "2025", "Capital expenditure", "Includes finance lease payments"],
        ["Owned Capex (2026 guided)", 125.0, "2026", "Capital expenditure", "Midpoint of $115-135B guidance"],
        ["Owned Capex (cumulative to 2028)", 600.0, "2024-2028", "Capital expenditure", "Zuckerberg commitment"],
        ["CoreWeave Contract", 14.2, "Through 2031", "Neocloud", "Option extends to 2032"],
        ["Nebius Contract (total)", 27.0, "5 years (2025-2031)", "Neocloud", "$12B guaranteed + $15B option"],
        ["Oracle (reported)", 20.0, "Multi-year", "Cloud", "Still in negotiation"],
        ["Google Cloud", 10.0, "Multi-year", "Cloud", "Estimated amount"],
        ["Blue Owl / Hyperion Financing", 30.0, "Construction period", "Off-balance-sheet", "Meta retains 20% ownership"],
        ["Operating Lease Obligations", 25.2, "Multi-year", "On-balance-sheet", "As of Dec 31, 2025"],
        ["Renewable Energy PPAs", 16.5, "Long-term", "Energy", "Majority due beyond 5 years"],
    ]

    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c_idx, value=h)
    apply_header_style(ws, 3, len(headers))

    for r_idx, row in enumerate(summary_data, 4):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c_idx == 2 and isinstance(val, (int, float)):
                cell.number_format = MONEY_FMT

    total_row = 4 + len(summary_data)
    ws.cell(row=total_row, column=1, value="GRAND TOTAL (estimated all-in)").font = TOTAL_FONT
    ws.cell(row=total_row, column=1).fill = TOTAL_FILL
    grand_total = sum(row[1] for row in summary_data)
    total_cell = ws.cell(row=total_row, column=2, value=grand_total)
    total_cell.font = TOTAL_FONT
    total_cell.fill = TOTAL_FILL
    total_cell.number_format = MONEY_FMT
    for c in range(1, len(headers) + 1):
        ws.cell(row=total_row, column=c).border = THIN_BORDER
        ws.cell(row=total_row, column=c).fill = TOTAL_FILL

    ws.cell(row=total_row, column=4, value="Note: Some overlap exists between capex and lease categories; Blue Owl structure reduces Meta direct capex by ~$24B").alignment = Alignment(wrap_text=True)

    note_row = total_row + 2
    ws.cell(row=note_row, column=1, value="Adjusted Total (removing overlap):").font = SUBTITLE_FONT
    ws.cell(row=note_row + 1, column=1, value="Owned Capex 2024-2028 cumulative")
    ws.cell(row=note_row + 1, column=2, value=600).number_format = MONEY_FMT
    ws.cell(row=note_row + 2, column=1, value="Third-Party Neocloud/Cloud (net)")
    ws.cell(row=note_row + 2, column=2, value=64).number_format = MONEY_FMT
    ws.cell(row=note_row + 3, column=1, value="Blue Owl Off-Balance-Sheet")
    ws.cell(row=note_row + 3, column=2, value=30).number_format = MONEY_FMT
    ws.cell(row=note_row + 4, column=1, value="Operating Leases")
    ws.cell(row=note_row + 4, column=2, value=25).number_format = MONEY_FMT
    ws.cell(row=note_row + 5, column=1, value="Renewable Energy PPAs")
    ws.cell(row=note_row + 5, column=2, value=17).number_format = MONEY_FMT
    ws.cell(row=note_row + 6, column=1, value="Estimated All-In Commitment").font = TOTAL_FONT
    ws.cell(row=note_row + 6, column=2, value=736).number_format = MONEY_FMT
    ws.cell(row=note_row + 6, column=2).font = TOTAL_FONT

    auto_width(ws, len(headers))
    return ws


def create_peer_comparison_sheet(wb):
    ws = wb.create_sheet("Peer Comparison", 6)

    ws.cell(row=1, column=1, value="Hyperscaler Capex & Neocloud Comparison (2025-2026)").font = TITLE_FONT
    ws.merge_cells("A1:F1")

    headers = ["Company", "2025 Capex ($B)", "2026E Capex ($B)", "YoY Growth", "2025 Revenue ($B)", "Capex/Revenue"]
    peer_data = [
        ["Meta", 72.2, 125.0, 0.73, 201.0, 0.36],
        ["Amazon", 125.0, 200.0, 0.60, 638.0, 0.20],
        ["Microsoft", 80.0, 145.0, 0.81, 262.0, 0.31],
        ["Alphabet/Google", 91.4, 180.0, 0.97, 382.0, 0.24],
        ["Oracle", 20.0, 50.0, 1.50, 59.0, 0.34],
    ]

    for c_idx, h in enumerate(headers, 1):
        ws.cell(row=3, column=c_idx, value=h)
    apply_header_style(ws, 3, len(headers))

    for r_idx, row in enumerate(peer_data, 4):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="top")
            if c_idx in (2, 3, 5) and isinstance(val, (int, float)):
                cell.number_format = MONEY_FMT
            if c_idx in (4, 6) and isinstance(val, (int, float)):
                cell.number_format = PCT_FMT
        if row[0] == "Meta":
            for c in range(1, len(headers) + 1):
                ws.cell(row=r_idx, column=c).fill = PatternFill(
                    start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"
                )

    ws.cell(row=11, column=1, value="Neocloud / Third-Party Cloud Dependency").font = SUBTITLE_FONT
    ws.merge_cells("A11:F11")

    neo_headers = ["Hyperscaler", "Neocloud/Third-Party Contracts ($B)", "% of 2026E Capex Equivalent", "Key Partners"]
    neo_data = [
        ["Meta", 64, 0.51, "CoreWeave, Nebius, Oracle, Google Cloud"],
        ["Microsoft", 17.4, 0.12, "Nebius; CoreWeave investor"],
        ["Amazon", "<5", None, "Limited disclosed neocloud contracts"],
        ["Alphabet/Google", "N/A", None, "Wiz acquisition ($32B); limited neocloud"],
        ["Oracle", "N/A", None, "Provides neocloud capacity (not buyer)"],
    ]

    for c_idx, h in enumerate(neo_headers, 1):
        ws.cell(row=13, column=c_idx, value=h)
    apply_header_style(ws, 13, len(neo_headers))

    for r_idx, row in enumerate(neo_data, 14):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c_idx == 2 and isinstance(val, (int, float)):
                cell.number_format = MONEY_FMT
            if c_idx == 3 and isinstance(val, (int, float)):
                cell.number_format = PCT_FMT

    auto_width(ws, max(len(headers), len(neo_headers)))
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)

    create_capex_sheet(wb)
    create_neocloud_sheet(wb)
    create_datacenter_sheet(wb)
    create_timeline_sheet(wb)
    create_balance_sheet_sheet(wb)
    create_summary_sheet(wb)
    create_peer_comparison_sheet(wb)

    wb.save(OUTPUT_FILE)
    print(f"Workbook saved to {OUTPUT_FILE}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    main()
