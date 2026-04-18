"""
Generate an Excel workbook containing a monthly token-consumption time series
for three AI agent verticals (Coding, Customer Service, Banking & Financial
Services) along with a clustered bar chart visualising the data.

Token figures are expressed in millions of tokens (M tokens) per month and
represent an illustrative 12-month trajectory.  The Coding agent shows the
highest per-interaction consumption (large source-code context windows); the
Banking & Financial Services agent shows moderate consumption (document and
report analysis); the Customer Service agent shows the lowest per-interaction
consumption (short conversational turns) but a steady, high-volume curve.
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


OUTPUT_PATH = "token_consumption_by_agent.xlsx"


MONTHS = [
    "2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10",
    "2025-11", "2025-12", "2026-01", "2026-02", "2026-03", "2026-04",
]

# Monthly token consumption in millions of tokens (M tokens).
CODING_AGENT = [
    420, 465, 530, 612, 690, 755,
    810, 845, 905, 980, 1060, 1145,
]

CUSTOMER_SERVICE = [
    150, 162, 175, 188, 205, 220,
    232, 245, 260, 278, 295, 312,
]

BANKING_FINANCIAL = [
    240, 258, 285, 310, 345, 372,
    398, 420, 455, 490, 528, 565,
]


def style_header(cell):
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def style_border(cell):
    thin = Side(border_style="thin", color="BFBFBF")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Token Consumption"

    title_cell = ws.cell(row=1, column=1,
                         value="Monthly Token Consumption by AI Agent Vertical (M tokens)")
    title_cell.font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    subtitle = ws.cell(row=2, column=1,
                       value="Values shown in millions of tokens consumed per month.")
    subtitle.font = Font(italic=True, size=10, color="595959")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    headers = [
        "Month",
        "Coding Agent",
        "Customer Service Agent",
        "Banking & Financial Services Agent",
        "Monthly Total",
    ]

    header_row = 4
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        style_header(cell)
        style_border(cell)

    for i, month in enumerate(MONTHS):
        row = header_row + 1 + i
        ws.cell(row=row, column=1, value=month).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2, value=CODING_AGENT[i])
        ws.cell(row=row, column=3, value=CUSTOMER_SERVICE[i])
        ws.cell(row=row, column=4, value=BANKING_FINANCIAL[i])
        ws.cell(row=row, column=5,
                value=f"=SUM(B{row}:D{row})")

        for col_idx in range(1, 6):
            c = ws.cell(row=row, column=col_idx)
            style_border(c)
            if col_idx >= 2:
                c.number_format = "#,##0"

    totals_row = header_row + 1 + len(MONTHS)
    ws.cell(row=totals_row, column=1, value="12-Month Total").font = Font(bold=True)
    for col_idx in range(2, 6):
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=totals_row, column=col_idx,
                    value=f"=SUM({col_letter}{header_row + 1}:{col_letter}{totals_row - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0"
        c.fill = PatternFill("solid", fgColor="D9E1F2")
    for col_idx in range(1, 6):
        style_border(ws.cell(row=totals_row, column=col_idx))
    ws.cell(row=totals_row, column=1).fill = PatternFill("solid", fgColor="D9E1F2")
    ws.cell(row=totals_row, column=1).alignment = Alignment(horizontal="center")

    avg_row = totals_row + 1
    ws.cell(row=avg_row, column=1, value="Monthly Average").font = Font(bold=True, italic=True)
    for col_idx in range(2, 6):
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=avg_row, column=col_idx,
                    value=f"=AVERAGE({col_letter}{header_row + 1}:{col_letter}{totals_row - 1})")
        c.font = Font(bold=True, italic=True)
        c.number_format = "#,##0"
        c.fill = PatternFill("solid", fgColor="EDEDED")
    for col_idx in range(1, 6):
        style_border(ws.cell(row=avg_row, column=col_idx))
    ws.cell(row=avg_row, column=1).fill = PatternFill("solid", fgColor="EDEDED")
    ws.cell(row=avg_row, column=1).alignment = Alignment(horizontal="center")

    column_widths = {1: 14, 2: 16, 3: 24, 4: 36, 5: 18}
    for col_idx, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "B5"

    chart = BarChart()
    chart.type = "col"
    chart.style = 11
    chart.grouping = "clustered"
    chart.title = "Monthly Token Consumption by AI Agent Vertical"
    chart.y_axis.title = "Tokens Consumed (Millions)"
    chart.x_axis.title = "Month"
    chart.height = 12
    chart.width = 24
    chart.gapWidth = 80

    data_ref = Reference(
        ws,
        min_col=2,
        max_col=4,
        min_row=header_row,
        max_row=header_row + len(MONTHS),
    )
    categories_ref = Reference(
        ws,
        min_col=1,
        max_col=1,
        min_row=header_row + 1,
        max_row=header_row + len(MONTHS),
    )
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(categories_ref)
    chart.legend.position = "b"
    chart.dataLabels = DataLabelList(showVal=False)

    ws.add_chart(chart, "G4")

    notes_row = avg_row + 3
    ws.cell(row=notes_row, column=1, value="Notes & Assumptions").font = Font(bold=True, size=11, color="1F4E78")
    notes = [
        "1. Token volumes are illustrative monthly totals across all interactions for each agent vertical.",
        "2. Coding Agent: large source-code context windows and multi-file edits drive higher per-task token cost.",
        "3. Customer Service Agent: short conversational turns; high interaction volume but low tokens per turn.",
        "4. Banking & Financial Services Agent: moderate volume; analysis of statements, reports, and disclosures.",
        "5. Growth trend reflects expanding adoption and richer multi-turn workflows over the 12-month window.",
    ]
    for i, note in enumerate(notes, start=1):
        c = ws.cell(row=notes_row + i, column=1, value=note)
        c.font = Font(size=10, color="404040")
        ws.merge_cells(start_row=notes_row + i, start_column=1,
                       end_row=notes_row + i, end_column=5)

    wb.save(OUTPUT_PATH)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
