"""Generate `AI_Token_Consumption_Model.xlsx` — an equity-research style
spreadsheet that forecasts global AI inference token consumption across
text, voice, and video modalities, with time-series and bar charts.

The workbook is fully formula-driven from a single Assumptions sheet so an
analyst can flex drivers (user counts, queries/user/day, tokens/query,
adoption growth) and see the model and charts update.
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT_PATH = "AI_Token_Consumption_Model.xlsx"

YEARS = [2024, 2025, 2026, 2027, 2028]
QUARTERS = [f"Q{q} {y}" for y in YEARS for q in (1, 2, 3, 4)]
N_PERIODS = len(QUARTERS)

MODALITIES = ["Text", "Voice", "Video"]


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
ASSUMP_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
WHITE = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")


def style_header(cell, fill=HEADER_FILL):
    cell.font = WHITE
    cell.fill = fill
    cell.alignment = CENTER
    cell.border = BORDER


def style_assump(cell):
    cell.fill = ASSUMP_FILL
    cell.border = BORDER
    cell.alignment = RIGHT


def style_total(cell):
    cell.fill = TOTAL_FILL
    cell.font = BOLD
    cell.border = BORDER
    cell.alignment = RIGHT


def widen(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_assumptions(ws):
    ws.title = "Assumptions"

    ws["A1"] = "AI Inference Token Consumption — Model Assumptions"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:D1")

    ws["A2"] = "All figures are illustrative; edit the yellow cells to flex the model."
    ws["A2"].font = Font(italic=True, color="595959")
    ws.merge_cells("A2:D2")

    ws["A4"] = "Driver"
    ws["B4"] = "Text"
    ws["C4"] = "Voice"
    ws["D4"] = "Video"
    for col in "ABCD":
        style_header(ws[f"{col}4"])

    rows = [
        ("Initial monthly active users (M, Q1 2024)", 500, 60, 12),
        ("User CAGR (annualized)",                    0.45, 0.85, 1.20),
        ("Avg sessions per user per day",             8, 2, 0.5),
        ("Avg tokens per session",                    750, 6_000, 60_000),
        ("Days per quarter",                          91, 91, 91),
    ]

    for i, (label, t, v, vid) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=label).font = BOLD
        for j, val in enumerate((t, v, vid), start=2):
            c = ws.cell(row=i, column=j, value=val)
            style_assump(c)
            if "CAGR" in label:
                c.number_format = "0.0%"
            elif "tokens" in label or "users" in label:
                c.number_format = "#,##0"
            else:
                c.number_format = "#,##0.00"

    ws["A12"] = "Notes"
    ws["A12"].font = BOLD
    notes = [
        "Tokens-per-session benchmarks reflect typical 2025 multimodal models:",
        "  Text  ≈ 500-1,000 tokens (prompt + completion)",
        "  Voice ≈ 3,000 tokens / minute (ASR + LLM + TTS), ~2-min session",
        "  Video ≈ 50,000 tokens / minute (frame patch tokens), ~1-min clip",
        "Quarterly user growth = (1 + CAGR) ^ (1/4) compounded from Q1 2024.",
    ]
    for i, n in enumerate(notes, start=13):
        ws.cell(row=i, column=1, value=n).alignment = LEFT
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)

    widen(ws, [46, 14, 14, 14])


def build_forecast(ws):
    ws.title = "Forecast"

    ws["A1"] = "Quarterly Token Consumption Forecast (Trillions of Tokens)"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N_PERIODS + 2)

    ws["A3"] = "Modality"
    style_header(ws["A3"])
    for i, q in enumerate(QUARTERS, start=2):
        c = ws.cell(row=3, column=i, value=q)
        style_header(c)
    total_col = N_PERIODS + 2
    c = ws.cell(row=3, column=total_col, value="5-Yr Total")
    style_header(c)

    assump_cols = {"Text": "B", "Voice": "C", "Video": "D"}

    for r, modality in enumerate(MODALITIES, start=4):
        ws.cell(row=r, column=1, value=modality).font = BOLD
        ws.cell(row=r, column=1).fill = SUBHEADER_FILL
        ws.cell(row=r, column=1).font = WHITE
        ws.cell(row=r, column=1).alignment = LEFT
        ws.cell(row=r, column=1).border = BORDER

        ac = assump_cols[modality]
        for i in range(N_PERIODS):
            col = i + 2
            users = (
                f"Assumptions!${ac}$5 * (1 + Assumptions!${ac}$6) ^ ({i}/4)"
            )
            tokens_per_quarter = (
                f"({users}) * 1000000 "
                f"* Assumptions!${ac}$7 "
                f"* Assumptions!${ac}$8 "
                f"* Assumptions!${ac}$9"
            )
            formula = f"=({tokens_per_quarter})/1e12"
            cell = ws.cell(row=r, column=col, value=formula)
            cell.number_format = "#,##0.0"
            cell.border = BORDER
            cell.alignment = RIGHT

        first = ws.cell(row=r, column=2).coordinate
        last = ws.cell(row=r, column=N_PERIODS + 1).coordinate
        tcell = ws.cell(row=r, column=total_col, value=f"=SUM({first}:{last})")
        tcell.number_format = "#,##0.0"
        style_total(tcell)

    total_row = 4 + len(MODALITIES)
    ws.cell(row=total_row, column=1, value="Total").font = BOLD
    style_total(ws.cell(row=total_row, column=1))
    ws.cell(row=total_row, column=1).alignment = LEFT
    for col in range(2, total_col + 1):
        first = ws.cell(row=4, column=col).coordinate
        last = ws.cell(row=4 + len(MODALITIES) - 1, column=col).coordinate
        c = ws.cell(row=total_row, column=col, value=f"=SUM({first}:{last})")
        c.number_format = "#,##0.0"
        style_total(c)

    ws.cell(row=total_row + 2, column=1, value="Annual Token Consumption (Trillions)").font = BOLD
    ws.merge_cells(start_row=total_row + 2, start_column=1, end_row=total_row + 2, end_column=6)

    header_r = total_row + 3
    ws.cell(row=header_r, column=1, value="Modality")
    style_header(ws.cell(row=header_r, column=1))
    for j, y in enumerate(YEARS, start=2):
        c = ws.cell(row=header_r, column=j, value=y)
        style_header(c)
    style_header(ws.cell(row=header_r, column=len(YEARS) + 2, value="5-Yr Total"))

    for k, modality in enumerate(MODALITIES, start=1):
        rr = header_r + k
        ws.cell(row=rr, column=1, value=modality).font = BOLD
        ws.cell(row=rr, column=1).fill = SUBHEADER_FILL
        ws.cell(row=rr, column=1).font = WHITE
        ws.cell(row=rr, column=1).border = BORDER
        forecast_row = 3 + k
        for j, _ in enumerate(YEARS, start=0):
            start_col = 2 + j * 4
            end_col = start_col + 3
            sc = ws.cell(row=forecast_row, column=start_col).coordinate
            ec = ws.cell(row=forecast_row, column=end_col).coordinate
            cell = ws.cell(
                row=rr,
                column=2 + j,
                value=f"=SUM({sc}:{ec})",
            )
            cell.number_format = "#,##0.0"
            cell.border = BORDER
            cell.alignment = RIGHT
        first = ws.cell(row=rr, column=2).coordinate
        last = ws.cell(row=rr, column=len(YEARS) + 1).coordinate
        c = ws.cell(row=rr, column=len(YEARS) + 2, value=f"=SUM({first}:{last})")
        c.number_format = "#,##0.0"
        style_total(c)

    total_annual_r = header_r + len(MODALITIES) + 1
    ws.cell(row=total_annual_r, column=1, value="Total").font = BOLD
    style_total(ws.cell(row=total_annual_r, column=1))
    ws.cell(row=total_annual_r, column=1).alignment = LEFT
    for col in range(2, len(YEARS) + 3):
        first = ws.cell(row=header_r + 1, column=col).coordinate
        last = ws.cell(row=header_r + len(MODALITIES), column=col).coordinate
        c = ws.cell(row=total_annual_r, column=col, value=f"=SUM({first}:{last})")
        c.number_format = "#,##0.0"
        style_total(c)

    widths = [22] + [11] * N_PERIODS + [13]
    widen(ws, widths)
    ws.freeze_panes = "B4"

    return total_row, header_r, total_annual_r


def build_charts(wb, forecast_ws, qrow_total, arow_header, arow_total):
    chart_ws = wb.create_sheet("Charts")
    chart_ws["A1"] = "Visualizations — Token Consumption by Modality"
    chart_ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    chart_ws.merge_cells("A1:M1")

    line = LineChart()
    line.title = "Quarterly Token Consumption by Modality (Trillions)"
    line.y_axis.title = "Tokens (Trillions)"
    line.x_axis.title = "Quarter"
    line.height = 11
    line.width = 22
    line.style = 12

    data = Reference(
        forecast_ws,
        min_col=1,
        max_col=N_PERIODS + 1,
        min_row=4,
        max_row=4 + len(MODALITIES) - 1,
    )
    line.add_data(data, titles_from_data=True, from_rows=True)
    cats = Reference(forecast_ws, min_col=2, max_col=N_PERIODS + 1, min_row=3, max_row=3)
    line.set_categories(cats)

    for s in line.series:
        s.smooth = True

    chart_ws.add_chart(line, "A3")

    bar = BarChart()
    bar.type = "col"
    bar.style = 11
    bar.grouping = "clustered"
    bar.title = "Annual Token Consumption by Modality (Trillions)"
    bar.y_axis.title = "Tokens (Trillions)"
    bar.x_axis.title = "Year"
    bar.height = 11
    bar.width = 22

    bar_data = Reference(
        forecast_ws,
        min_col=1,
        max_col=len(YEARS) + 1,
        min_row=arow_header + 1,
        max_row=arow_header + len(MODALITIES),
    )
    bar.add_data(bar_data, titles_from_data=True, from_rows=True)
    bar_cats = Reference(
        forecast_ws,
        min_col=2,
        max_col=len(YEARS) + 1,
        min_row=arow_header,
        max_row=arow_header,
    )
    bar.set_categories(bar_cats)
    bar.dataLabels = DataLabelList(showVal=False)

    chart_ws.add_chart(bar, "A26")

    stacked = BarChart()
    stacked.type = "col"
    stacked.style = 13
    stacked.grouping = "stacked"
    stacked.overlap = 100
    stacked.title = "Annual Token Mix — Stacked (Trillions)"
    stacked.y_axis.title = "Tokens (Trillions)"
    stacked.x_axis.title = "Year"
    stacked.height = 11
    stacked.width = 22
    stacked.add_data(bar_data, titles_from_data=True, from_rows=True)
    stacked.set_categories(bar_cats)
    chart_ws.add_chart(stacked, "A49")

    chart_ws.column_dimensions["A"].width = 16


def main() -> None:
    wb = Workbook()
    build_assumptions(wb.active)

    forecast_ws = wb.create_sheet("Forecast")
    qtotal_r, ahdr_r, atotal_r = build_forecast(forecast_ws)

    build_charts(wb, forecast_ws, qtotal_r, ahdr_r, atotal_r)

    order = ["Assumptions", "Forecast", "Charts"]
    wb._sheets = [wb[name] for name in order]

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
