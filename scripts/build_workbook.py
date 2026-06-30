#!/usr/bin/env python3
"""Build the GLM 5.2 vs leading frontier models token-consumption workbook.

Reads data/openrouter_token_consumption_june2026.csv and produces an .xlsx
workbook with an aggregated comparison table, a GLM 5.2 run-rate sheet, an
assumptions sheet (with live formulas) and a sources sheet.

Usage:
    python scripts/build_workbook.py
"""
from __future__ import annotations

import csv
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_PATH = os.path.join(ROOT, "data", "openrouter_token_consumption_june2026.csv")
OUT_DIR = os.path.join(ROOT, "output")
OUT_PATH = os.path.join(OUT_DIR, "GLM-5.2_vs_Frontier_Token_Consumption.xlsx")

AS_OF = "30 Jun 2026"

# ---- styling helpers --------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=15, color="1F3864")
SUB_FONT = Font(italic=True, size=10, color="595959")
GLM_FILL = PatternFill("solid", fgColor="FCE4D6")  # highlight GLM 5.2 row
ALT_FILL = PatternFill("solid", fgColor="F2F2F2")
BORDER = Border(*[Side(style="thin", color="D9D9D9")] * 4)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def to_float(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except ValueError:
        return None


def load_rows():
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def build():
    rows = load_rows()
    wb = Workbook()

    # ---------------- Assumptions sheet (named cells for live formulas) ------
    wsa = wb.active
    wsa.title = "Assumptions"
    wsa["A1"] = "Assumptions & Toggles"
    wsa["A1"].font = TITLE_FONT
    wsa["A3"] = "Weeks per month"
    wsa["B3"] = 4.345           # 365.25 / 7 / 12
    wsa["A4"] = "Input token share (blended price)"
    wsa["B4"] = 0.5
    wsa["A5"] = "Output token share (blended price)"
    wsa["B5"] = "=1-B4"
    wsa["A7"] = ("Implied spend = tokens (trillions) x blended price ($/Mtok), "
                 "which equals $ in millions.")
    wsa["A7"].font = SUB_FONT
    wsa["A8"] = ("Blended price = input_share x input$ + output_share x output$. "
                 "Edit B4 to re-flow every spend figure in the workbook.")
    wsa["A8"].font = SUB_FONT
    for cell in ("B3", "B4"):
        wsa[cell].fill = PatternFill("solid", fgColor="FFF2CC")
        wsa[cell].font = Font(bold=True)
    wsa.column_dimensions["A"].width = 38
    wsa.column_dimensions["B"].width = 12
    wb.defined_names.add  # no-op guard for older API; we reference by absolute addr

    WEEKS = "Assumptions!$B$3"
    INSHARE = "Assumptions!$B$4"
    OUTSHARE = "Assumptions!$B$5"

    # ---------------- Main comparison sheet ----------------------------------
    ws = wb.create_sheet("Monthly_Consumption")
    ws["A1"] = "Monthly Token Consumption — GLM 5.2 vs Leading Frontier Models"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"Source: OpenRouter public usage rankings & model pages. As of {AS_OF}. "
                "Token volumes in trillions (T). GLM 5.2 row highlighted.")
    ws["A2"].font = SUB_FONT

    headers = [
        "Rank\n(month)", "Model", "Developer", "Country", "License",
        "Monthly\ntokens (T)", "MoM\ngrowth %", "Latest weekly\ntokens (T)",
        "WoW\ngrowth %", "Monthly run-rate\n(T)", "Input\n$/Mtok", "Output\n$/Mtok",
        "Price\nbasis", "Blended\n$/Mtok", "Implied monthly spend\n(trailing, $M)",
        "Implied monthly spend\n(run-rate, $M)", "Notes",
    ]
    hrow = 4
    for c, h in enumerate(headers, start=1):
        ws.cell(row=hrow, column=c, value=h)
    style_header(ws, hrow, len(headers))

    r = hrow + 1
    first_data = r
    for row in rows:
        monthly = to_float(row["monthly_tokens_T"])
        weekly = to_float(row["latest_weekly_tokens_T"])
        in_list = to_float(row["input_price_list_usd_per_mtok"])
        out_list = to_float(row["output_price_list_usd_per_mtok"])
        in_real = to_float(row["input_price_realized_usd_per_mtok"])
        out_real = to_float(row["output_price_realized_usd_per_mtok"])

        # choose realized pricing if available, else list
        in_used = in_real if in_real is not None else in_list
        out_used = out_real if out_real is not None else out_list
        if in_real is not None or out_real is not None:
            basis = "realized"
        elif in_list is not None:
            basis = "list"
        else:
            basis = "n/a"

        ws.cell(row=r, column=1, value=to_float(row["rank_monthly"]))
        ws.cell(row=r, column=2, value=row["model"])
        ws.cell(row=r, column=3, value=row["developer"])
        ws.cell(row=r, column=4, value=row["country"])
        ws.cell(row=r, column=5, value=row["license_type"])
        ws.cell(row=r, column=6, value=monthly)
        mom = to_float(row["mom_growth_pct"])
        ws.cell(row=r, column=7, value=(mom / 100.0) if mom is not None else None)
        ws.cell(row=r, column=8, value=weekly)
        wow = to_float(row["weekly_growth_pct"])
        ws.cell(row=r, column=9, value=(wow / 100.0) if wow is not None else None)
        # monthly run-rate = weekly * weeks_per_month
        if weekly is not None:
            ws.cell(row=r, column=10, value=f"=H{r}*{WEEKS}")
        ws.cell(row=r, column=11, value=in_used)
        ws.cell(row=r, column=12, value=out_used)
        ws.cell(row=r, column=13, value=basis)
        # blended price formula (only when both prices present)
        if in_used is not None and out_used is not None:
            ws.cell(row=r, column=14, value=f"=K{r}*{INSHARE}+L{r}*{OUTSHARE}")
            # trailing spend
            if monthly is not None:
                ws.cell(row=r, column=15, value=f"=F{r}*N{r}")
            # run-rate spend
            ws.cell(row=r, column=16, value=f"=IF(J{r}=\"\",\"\",J{r}*N{r})")
        ws.cell(row=r, column=17, value=row["notes"])
        r += 1
    last_data = r - 1

    # number formats & per-row styling
    pct_cols = (7, 9)
    one_dp = (6, 8, 10)
    money = (11, 12, 14, 15, 16)
    for rr in range(first_data, last_data + 1):
        is_glm = ws.cell(row=rr, column=2).value == "GLM 5.2"
        for cc in range(1, len(headers) + 1):
            cell = ws.cell(row=rr, column=cc)
            cell.border = BORDER
            cell.alignment = LEFT if cc in (2, 3, 17) else CENTER
            if is_glm:
                cell.fill = GLM_FILL
            elif (rr - first_data) % 2 == 1:
                cell.fill = ALT_FILL
        for cc in pct_cols:
            ws.cell(row=rr, column=cc).number_format = "+0%;-0%"
        for cc in one_dp:
            ws.cell(row=rr, column=cc).number_format = "0.00"
        for cc in money:
            ws.cell(row=rr, column=cc).number_format = "$#,##0.00"
        ws.cell(row=rr, column=15).number_format = "$#,##0.0"
        ws.cell(row=rr, column=16).number_format = "$#,##0.0"
        ws.cell(row=rr, column=2).font = Font(bold=is_glm)

    # totals row
    tot = last_data + 2
    ws.cell(row=tot, column=2, value="Platform context")
    ws.cell(row=tot, column=2).font = Font(bold=True)
    ws.cell(row=tot, column=6, value=f"=SUM(F{first_data}:F{last_data})")
    ws.cell(row=tot, column=6).number_format = "0.00"
    ws.cell(row=tot, column=6).font = Font(bold=True)
    ws.cell(row=tot, column=17,
            value="Top-10 trailing-month subtotal (GLM 5.2 excluded; partial month). "
                  "OpenRouter platform-wide volume ~46T tokens/week (Jun 2026).")
    ws.cell(row=tot, column=17).alignment = LEFT

    widths = [8, 22, 14, 9, 13, 11, 9, 12, 9, 13, 9, 9, 9, 10, 16, 16, 52]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C5"

    # ---------------- GLM 5.2 run-rate sheet ---------------------------------
    wg = wb.create_sheet("GLM5.2_RunRate")
    wg["A1"] = "GLM 5.2 — Adoption Run-Rate"
    wg["A1"].font = TITLE_FONT
    wg["A2"] = ("Launched 13 Jun 2026 (GLM Coding Plan) and 16 Jun 2026 (MIT open weights + API). "
                "744B MoE, 40B active, 1M-token context. Excluded from the trailing-month top 10 "
                "because of its partial first month; run-rate is the cleaner read.")
    wg["A2"].font = SUB_FONT
    wg["A2"].alignment = LEFT
    wg.merge_cells("A2:E2")

    gh = ["Week ending", "Weekly tokens (T)", "WoW growth %", "Data quality", "Source / note"]
    for c, h in enumerate(gh, start=1):
        wg.cell(row=4, column=c, value=h)
    style_header(wg, 4, len(gh))
    glm_weeks = [
        ("21 Jun 2026", 1.27, None, "derived", "Implied from +66% WoW into week ending 28 Jun"),
        ("28 Jun 2026", 2.11, 0.66, "reported", "thestack.technology / OpenRouter ranking (#7)"),
    ]
    gr = 5
    for wk, tok, wow, q, src in glm_weeks:
        wg.cell(row=gr, column=1, value=wk)
        wg.cell(row=gr, column=2, value=tok).number_format = "0.00"
        if wow is not None:
            wg.cell(row=gr, column=3, value=wow).number_format = "+0%;-0%"
        wg.cell(row=gr, column=4, value=q)
        wg.cell(row=gr, column=5, value=src)
        for cc in range(1, 6):
            wg.cell(row=gr, column=cc).border = BORDER
            wg.cell(row=gr, column=cc).alignment = LEFT if cc == 5 else CENTER
        gr += 1

    wg.cell(row=gr + 1, column=1, value="Latest weekly (T)")
    wg.cell(row=gr + 1, column=2, value="=B6")
    wg.cell(row=gr + 1, column=2).number_format = "0.00"
    wg.cell(row=gr + 2, column=1, value="Monthly run-rate (T)")
    wg.cell(row=gr + 2, column=2, value=f"=B6*{WEEKS}")
    wg.cell(row=gr + 2, column=2).number_format = "0.00"
    wg.cell(row=gr + 3, column=1, value="Realized blended price ($/Mtok)")
    wg.cell(row=gr + 3, column=2, value=f"=0.447*{INSHARE}+3.31*{OUTSHARE}")
    wg.cell(row=gr + 3, column=2).number_format = "$#,##0.00"
    wg.cell(row=gr + 4, column=1, value="Implied monthly inference spend ($M)")
    wg.cell(row=gr + 4, column=2, value=f"=B{gr+2}*B{gr+3}")
    wg.cell(row=gr + 4, column=2).number_format = "$#,##0.0"
    for rr in range(gr + 1, gr + 5):
        wg.cell(row=rr, column=1).font = Font(bold=True)
        wg.cell(row=rr, column=2).fill = GLM_FILL
    wg.cell(row=gr + 6, column=1,
            value=("Interpretation: at the latest weekly run-rate GLM 5.2 annualizes to a "
                   "monthly-equivalent volume that would place it among the top 3-5 models, "
                   "despite not appearing in the trailing-month top 10."))
    wg.cell(row=gr + 6, column=1).font = SUB_FONT
    wg.cell(row=gr + 6, column=1).alignment = LEFT
    wg.merge_cells(start_row=gr + 6, start_column=1, end_row=gr + 6, end_column=5)
    for i, w in enumerate([20, 18, 14, 14, 52], start=1):
        wg.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Sources sheet ------------------------------------------
    wsrc = wb.create_sheet("Sources")
    wsrc["A1"] = "Sources & Methodology"
    wsrc["A1"].font = TITLE_FONT
    notes = [
        "Token-consumption deliverable focused on GLM 5.2 vs leading frontier models, as of 30 Jun 2026.",
        "",
        "Token volumes:",
        "  - OpenRouter LLM Rankings (live monthly/weekly token usage): https://openrouter.ai/rankings",
        "  - 'These Are The Most Popular AI Models On OpenRouter [June 2026]', officechai (monthly top 10).",
        "  - 'OpenRouter Monthly Token Usage Ranking 2026', aicost.org (cross-check).",
        "  - GLM 5.2 weekly figure: thestack.technology 'Opus-killer GLM-5.2... astonishing demand' (#7 ranking, 2.11T week ending 28 Jun, +66% WoW).",
        "  - OpenRouter blog 'The Open Weight Models that Matter: June 2026' (GLM 5.2 realized pricing $0.447/$3.31).",
        "",
        "Pricing (per 1M tokens, list and prompt-cache realized where shown):",
        "  - OpenRouter model pages: z-ai/glm-5.2, deepseek-v4-flash, deepseek-v4-pro, deepseek-v3.2,",
        "    anthropic/claude-opus-4.7, anthropic/claude-sonnet-4.6, moonshotai/kimi-k2.6.",
        "  - Gemini 3 Flash Preview list pricing $0.50/$3.00 (OpenRouter).",
        "  - Hy3 Preview input $0.063/M (output not disclosed); Owl Alpha & Nemotron (free) pricing not published.",
        "",
        "Methodology & assumptions (see Assumptions sheet, which drives all formulas):",
        "  - Monthly run-rate (T) = latest weekly tokens x weeks-per-month (4.345).",
        "  - Blended $/Mtok = input_share x input$ + output_share x output$ (default 50/50, editable).",
        "  - Implied monthly spend ($M) = tokens (trillions) x blended $/Mtok.  (1T tokens x $1/Mtok = $1M.)",
        "  - Realized (post prompt-caching) pricing used where published; otherwise provider list pricing.",
        "",
        "Caveats:",
        "  - OpenRouter reflects developer/API usage routed through that platform, not total ecosystem usage",
        "    (first-party apps such as ChatGPT, Gemini and Claude.ai consume far more tokens off-platform).",
        "  - GLM 5.2 launched mid-June, so its trailing-month cumulative understates its true run-rate.",
        "  - Implied spend is an estimate of inference $ on OpenRouter, not lab revenue or profit.",
        "  - Growth figures capped in source reporting (e.g. '+>999%') are entered at 999%.",
    ]
    for i, line in enumerate(notes, start=3):
        wsrc.cell(row=i, column=1, value=line)
    wsrc.column_dimensions["A"].width = 110

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    build()
