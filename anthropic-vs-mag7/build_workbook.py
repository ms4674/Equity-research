"""Build Anthropic_vs_Mag7_Spending_Comparison.xlsx.

Mag 7 income-statement figures are pulled from the SEC XBRL companyfacts API
(10-K / 10-Q as filed). Anthropic figures are transcribed from its Form S-1/A
(filed Nov 2, 2026; Note 8 "Supplemental Operating Expense Information" and
the MD&A quarterly tables). Amazon's "Technology and infrastructure" line uses
a company-specific XBRL tag that the companyfacts API does not expose, so it is
transcribed from Amazon's 10-K / 10-Q income statements.

Usage:  python build_workbook.py [--offline]
        --offline reuses cached companyfacts JSON in ./cache instead of hitting the SEC.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import urllib.request

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")
OUT = os.path.join(HERE, "Anthropic_vs_Mag7_Spending_Comparison.xlsx")
UA = "Equity-research workbook builder (contact: research@example.com)"

# --------------------------------------------------------------------------------------
# Anthropic (Form S-1/A, Nov 2, 2026). $ millions.
# Functional view: Note 8. By-nature view: consolidated statements of operations / MD&A.
# --------------------------------------------------------------------------------------
ANTHROPIC_FUNCTIONAL = {
    # period: (revenue, cost of revenue, R&D, S&M, G&A, Bartz settlement, operating income)
    "FY2023":  (41,     26,    422,    76,    68,    0,     -551),
    "FY2024":  (386,    752,   2275,   163,   176,   0,     -2980),
    "FY2025":  (4585,   2773,  6338,   911,   1123,  1500,  -8060),
    "H1 2025": (1163,   1019,  1850,   239,   190,   0,     -2135),
    "H1 2026": (16286,  2380,  9741,   3808,  1050,  0,     -693),
    "9M 2025": (2545,   1796,  3594,   530,   349,   1500,  -5224),
    "9M 2026": (33242,  4812,  19240,  6619,  3831,  0,     -1260),
}
# Derived: Q3 = 9M - H1 ; TTM Sep-26 = FY2025 + 9M 2026 - 9M 2025
def _derive(a, b, sign=1):
    return tuple(x + sign * y for x, y in zip(a, b))
ANTHROPIC_FUNCTIONAL["Q3 2025"] = _derive(ANTHROPIC_FUNCTIONAL["9M 2025"], ANTHROPIC_FUNCTIONAL["H1 2025"], -1)
ANTHROPIC_FUNCTIONAL["Q3 2026"] = _derive(ANTHROPIC_FUNCTIONAL["9M 2026"], ANTHROPIC_FUNCTIONAL["H1 2026"], -1)
ANTHROPIC_FUNCTIONAL["TTM Sep-2026"] = _derive(
    _derive(ANTHROPIC_FUNCTIONAL["FY2025"], ANTHROPIC_FUNCTIONAL["9M 2026"]), ANTHROPIC_FUNCTIONAL["9M 2025"], -1)

ANTHROPIC_BY_NATURE = {
    # period: (revenue, compute & infrastructure, personnel, sales/marketing/partnerships, other op & admin, Bartz, total opex, op income)
    "FY2023":  (41,    433,   95,   2,    62,   0,    592,   -551),
    "FY2024":  (386,   2522,  454,  43,   347,  0,    3366,  -2980),
    "FY2025":  (4585,  7326,  1839, 568,  1412, 1500, 12645, -8060),
    "9M 2025": (2545,  4381,  1027, 322,  539,  1500, 7769,  -5224),
    "9M 2026": (33242, 18808, 4863, 5529, 5302, 0,    34502, -1260),
}
ANTHROPIC_BY_NATURE["TTM Sep-2026"] = _derive(
    _derive(ANTHROPIC_BY_NATURE["FY2025"], ANTHROPIC_BY_NATURE["9M 2026"]), ANTHROPIC_BY_NATURE["9M 2025"], -1)

ANTHROPIC_QUARTERLY = [
    # quarter, revenue, compute & infra, personnel, S/M/P, other op & admin, total opex, op income, net loss
    ("Q1 2025", 376,   944,  218,  24,   125,  1311,  -935,  -6592),
    ("Q2 2025", 787,   1390, 320,  96,   181,  1987,  -1200, -6377),
    ("Q3 2025", 1382,  2047, 489,  202,  233,  4471,  -3089, -12138),  # opex incl. $1,500 Bartz settlement
    ("Q4 2025", 2040,  2945, 812,  246,  873,  4876,  -2836, -16858),
    ("Q1 2026", 4731,  3387, 1295, 1033, 683,  6398,  -1667, -14439),
    ("Q2 2026", 11555, 5558, 1857, 2135, 1031, 10581, 974,   -66309),
    ("Q3 2026", 16956, 9863, 1711, 2361, 3588, 17523, -567,  -23291),
]

# Amazon "Technology and infrastructure" (R&D-equivalent line) and "Fulfillment", $M, from 10-K/10-Q.
AMAZON_TECH_INFRA = {"FY2023": 85622, "FY2024": 88544, "FY2025": 108521,
                     "9M 2025": 79122, "9M 2026": 97539, "Q3 2025": 28962, "Q3 2026": 34814}

# --------------------------------------------------------------------------------------
# Mag 7 via SEC XBRL companyfacts
# --------------------------------------------------------------------------------------
COMPANIES = [  # display order
    ("Anthropic", None), ("Microsoft", 789019), ("Alphabet", 1652044), ("Amazon", 1018724),
    ("Meta", 1326801), ("Apple", 320193), ("Nvidia", 1045810), ("Tesla", 1318605),
]
TAGS = {
    "revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "cor":     ["CostOfRevenue", "CostOfGoodsAndServicesSold"],
    "rnd":     ["ResearchAndDevelopmentExpense"],
    "sm":      ["SellingAndMarketingExpense", "MarketingExpense"],
    "ga":      ["GeneralAndAdministrativeExpense"],
    "sga":     ["SellingGeneralAndAdministrativeExpense"],
    "opinc":   ["OperatingIncomeLoss"],
    "capex":   ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
}
CASHFLOW_METRICS = {"capex"}  # reported YTD only; quarters derived by differencing


def fetch_facts(cik: int, offline: bool) -> dict:
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, f"CIK{cik:010d}.json")
    if offline:
        return json.load(open(path))["facts"]["us-gaap"]
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        open(path, "wb").write(data)
    except Exception as e:  # fall back to cache if present
        if os.path.exists(path):
            print(f"warning: fetch failed for CIK {cik} ({e}); using cache", file=sys.stderr)
        else:
            raise
    return json.load(open(path))["facts"]["us-gaap"]


def _d(s): return dt.date.fromisoformat(s)
def _dur(k): return (_d(k[1]) - _d(k[0])).days


def series(facts, tags):
    """(start,end)->value for the tag with the most recent data; latest filing wins per period."""
    best = None
    for tg in tags:
        if tg not in facts:
            continue
        m = {}
        for v in facts[tg]["units"].get("USD", []):
            if "start" not in v:
                continue
            k = (v["start"], v["end"])
            if k not in m or v["filed"] > m[k][1]:
                m[k] = (v["val"], v["filed"])
        if not m:
            continue
        latest_end = max(e for _, e in m)
        if best is None or latest_end > best[0]:
            best = (latest_end, {k: x[0] for k, x in m.items()}, tg)
    return (best[1], best[2]) if best else ({}, None)


def shift_year(s, k):
    a, b = _d(k[0]), _d(k[1])
    for kk in s:
        if abs((_d(kk[0]) - a.replace(year=a.year - 1)).days) <= 7 and abs((_d(kk[1]) - b.replace(year=b.year - 1)).days) <= 7:
            return kk
    return None


def periods(s, fy_end_month, cashflow=False):
    """Annual (last 3 FYs), latest YTD + prior-year YTD, TTM, latest quarter + prior-year quarter."""
    annual = sorted([k for k in s if 355 <= _dur(k) <= 375 and _d(k[1]).month == fy_end_month], key=lambda k: k[1])
    annual = annual[-3:]
    out = {"annual": [(k, s[k]) for k in annual]}
    fy = annual[-1]
    later = [k for k in s if k[1] > fy[1]]
    latest_end = max((k[1] for k in later), default=fy[1])
    out["fy_end"], out["latest_end"] = fy[1], latest_end
    # sub-annual only: Amazon also tags trailing-12-month cash-flow periods, which must not be mistaken for YTD
    ytd = [k for k in later if k[1] == latest_end and 84 <= _dur(k) < 350]
    ytd = max(ytd, key=_dur) if ytd else None
    if ytd:
        py = shift_year(s, ytd)
        out["ytd"], out["ytd_prior"] = (ytd, s[ytd]), ((py, s[py]) if py else None)
        out["ttm"] = s[fy] + s[ytd] - s[py] if py else None
    else:
        out["ytd"] = out["ytd_prior"] = None
        out["ttm"] = s[fy]
    # latest quarter
    q = [k for k in s if k[1] == latest_end and 84 <= _dur(k) <= 98]
    if q and not cashflow:
        qk = q[0]; pq = shift_year(s, qk)
        out["q"], out["q_prior"] = (qk, s[qk]), ((pq, s[pq]) if pq else None)
    elif ytd and cashflow:
        # quarter = YTD - previous YTD within same fiscal year
        prev = [k for k in s if k[0] == ytd[0] and k[1] < ytd[1]]
        if prev:
            pk = max(prev, key=_dur)
            out["q"] = ((pk[1], ytd[1]), s[ytd] - s[pk])
            ppk = shift_year(s, pk); pyk = shift_year(s, ytd)
            out["q_prior"] = (None, s[pyk] - s[ppk]) if ppk and pyk else None
        else:
            out["q"] = ((ytd[0], ytd[1]), s[ytd]); pyk = shift_year(s, ytd)
            out["q_prior"] = (pyk, s[pyk]) if pyk else None
    else:
        # fiscal Q4 = FY - 9M
        nine = [k for k in s if k[0] == fy[0] and 260 <= _dur(k) <= 285]
        if nine:
            out["q"] = ((nine[0][1], fy[1]), s[fy] - s[nine[0]])
            pa = annual[-2] if len(annual) > 1 else None; pn = shift_year(s, nine[0])
            out["q_prior"] = (None, s[pa] - s[pn]) if pa and pn else None
        else:
            out["q"] = out["q_prior"] = None
    return out


def load_mag7(offline: bool):
    data = {}
    for name, cik in COMPANIES:
        if cik is None:
            continue
        facts = fetch_facts(cik, offline)
        rev_s, _ = series(facts, TAGS["revenue"])
        # fiscal year-end month = month of the most recent 10-K-length period
        fy_end_month = _d(max((k[1] for k in rev_s if 355 <= _dur(k) <= 375))).month
        data[name] = {}
        for metric, tags in TAGS.items():
            s, tg = series(facts, tags)
            if not s:
                continue
            p = periods(s, fy_end_month, cashflow=metric in CASHFLOW_METRICS)
            p["tag"] = tg
            data[name][metric] = p
    # Amazon: Technology & infrastructure stands in for R&D (not XBRL-exposed)
    am = data["Amazon"]
    rev = am["revenue"]
    am["rnd"] = {
        "annual": [((k[0], k[1]), AMAZON_TECH_INFRA[f"FY{_d(k[1]).year}"] * 1e6) for k, _ in rev["annual"]],
        "ytd": (rev["ytd"][0], AMAZON_TECH_INFRA["9M 2026"] * 1e6),
        "ytd_prior": (rev["ytd_prior"][0], AMAZON_TECH_INFRA["9M 2025"] * 1e6),
        "ttm": (AMAZON_TECH_INFRA["FY2025"] + AMAZON_TECH_INFRA["9M 2026"] - AMAZON_TECH_INFRA["9M 2025"]) * 1e6,
        "q": (rev["q"][0], AMAZON_TECH_INFRA["Q3 2026"] * 1e6),
        "q_prior": (rev["q_prior"][0], AMAZON_TECH_INFRA["Q3 2025"] * 1e6),
        "fy_end": rev["fy_end"], "latest_end": rev["latest_end"], "tag": "amzn:TechnologyAndInfrastructureExpense (10-K/10-Q)",
    }
    # Nvidia and Tesla disclose SG&A only; carry SG&A into the S&M slot and flag it
    for n in ("Nvidia", "Tesla"):
        data[n]["sm"] = dict(data[n]["sga"]); data[n]["sm"]["combined_sga"] = True
        data[n]["ga"] = None
    return data


# --------------------------------------------------------------------------------------
# Workbook helpers
# --------------------------------------------------------------------------------------
NAVY = "1F3864"; TEAL = "2E75B6"; LIGHT = "DDEBF7"; GREY = "F2F2F2"; AMBER = "FFF2CC"; GREEN = "E2EFDA"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=10, italic=True, color="595959")
F_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_SEC = Font(name="Calibri", size=12, bold=True, color=NAVY)
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_NORM = Font(name="Calibri", size=10)
F_NOTE = Font(name="Calibri", size=9, color="595959")
FILL_HDR = PatternFill("solid", fgColor=TEAL)
FILL_SEC = PatternFill("solid", fgColor=NAVY)
FILL_ANT = PatternFill("solid", fgColor=AMBER)
FILL_ALT = PatternFill("solid", fgColor=GREY)
FILL_DERIVED = PatternFill("solid", fgColor=GREEN)
NUM = '#,##0;(#,##0);"–"'
NUM1 = '#,##0.0;(#,##0.0);"–"'
PCT = '0.0%;(0.0%);"–"'
MULT = '0.00"x"'
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def M(v):  # $ -> $ millions
    return None if v is None else round(v / 1e6)


def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = F_TITLE
    if sub:
        ws["A2"] = sub; ws["A2"].font = F_SUB
    ws.sheet_view.showGridLines = False


def header_row(ws, r, c0, headers, fill=FILL_HDR, height=30):
    for i, h in enumerate(headers):
        cell = ws.cell(row=r, column=c0 + i, value=h)
        cell.font = F_HDR; cell.fill = fill; cell.alignment = CENTER; cell.border = BORDER
    ws.row_dimensions[r].height = height


def section(ws, r, text, span=12):
    ws.cell(row=r, column=1, value=text).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for c in range(1, span + 1):
        ws.cell(row=r, column=c).fill = FILL_SEC
    ws.row_dimensions[r].height = 20


def put(ws, r, c, v, fmt=None, bold=False, fill=None, align=None, font=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or (F_BOLD if bold else F_NORM)
    cell.border = BORDER
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    return cell


def note(ws, r, text, span=12):
    ws.cell(row=r, column=1, value=text).font = F_NOTE
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = max(15, 13 * (1 + len(text) // 150))


def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[get_column_letter(i)].width = x


def fy_label(end_iso):
    d = _d(end_iso)
    return f"FY{d.year} (to {d.strftime('%b %d, %Y')})"


def ttm_label(name, p):
    fe, le = _d(p["fy_end"]), _d(p["latest_end"])
    if fe == le:
        return f"FY{fe.year} (12 mo. to {fe.strftime('%b %d, %Y')})"
    months = round((le - fe).days / 30.4)
    return (f"TTM to {le.strftime('%b %d, %Y')} (FY{fe.year} + {months}M to {le.strftime('%b %Y')} "
            f"− {months}M to {le.replace(year=le.year - 1).strftime('%b %Y')})")


def q_label(p):
    if not p or not p.get("q"): return "n/a"
    s, e = p["q"][0]
    e = _d(e)
    return f"Qtr to {e.strftime('%b %d, %Y')}"


# --------------------------------------------------------------------------------------
# Sheets
# --------------------------------------------------------------------------------------
def row_values(name, data, which):
    """Return dict of $M values for a company for 'ttm' | 'fy' | 'q' | 'q_prior'."""
    if name == "Anthropic":
        key = {"ttm": "TTM Sep-2026", "fy": "FY2025", "q": "Q3 2026", "q_prior": "Q3 2025"}[which]
        rev, cor, rnd, sm, ga, bartz, opinc = ANTHROPIC_FUNCTIONAL[key]
        return dict(revenue=rev, cor=cor, rnd=rnd, sm=sm, ga=ga, opinc=opinc, sga=None)
    d = data[name]
    out = {}
    for m in ("revenue", "cor", "rnd", "sm", "ga", "opinc", "sga"):
        p = d.get(m)
        if not p:
            out[m] = None; continue
        if which == "ttm": v = p["ttm"]
        elif which == "fy": v = p["annual"][-1][1]
        elif which == "q": v = p["q"][1] if p.get("q") else None
        else: v = p["q_prior"][1] if p.get("q_prior") else None
        out[m] = M(v)
    return out


def period_text(name, data, which):
    if name == "Anthropic":
        return {"ttm": "TTM to Sep 30, 2026 (FY2025 + 9M 2026 − 9M 2025)", "fy": "FY2025 (to Dec 31, 2025)",
                "q": "Qtr to Sep 30, 2026", "q_prior": "Qtr to Sep 30, 2025"}[which]
    p = data[name]["revenue"]
    if which == "ttm": return ttm_label(name, p)
    if which == "fy": return fy_label(p["annual"][-1][0][1])
    return q_label(p)


def write_comparison_block(ws, r0, data, which, heading, with_growth=False):
    section(ws, r0, heading, 15)
    hdr = ["Company", "Period", "Revenue", "Cost of revenue", "Gross margin %", "R&D", "R&D % rev",
           "Sales & marketing", "S&M % rev", "G&A", "G&A % rev", "Operating income", "Op. margin %"]
    if with_growth:
        hdr += ["R&D YoY %", "Cost of rev. YoY %", "S&M YoY %"]
    header_row(ws, r0 + 1, 1, hdr)
    r = r0 + 2
    first = r
    for i, (name, _) in enumerate(COMPANIES):
        v = row_values(name, data, which)
        fill = FILL_ANT if name == "Anthropic" else (FILL_ALT if i % 2 == 0 else None)
        label = name
        if name == "Amazon": label = "Amazon ¹"
        if name in ("Nvidia", "Tesla"): label = f"{name} ²"
        if name == "Meta" and which in ("ttm", "q"): label = "Meta ³"
        put(ws, r, 1, label, bold=(name == "Anthropic"), fill=fill)
        put(ws, r, 2, period_text(name, data, which), fill=fill, font=F_NOTE)
        rev = v["revenue"]
        put(ws, r, 3, rev, NUM, fill=fill)
        put(ws, r, 4, v["cor"], NUM, fill=fill)
        put(ws, r, 5, (rev - v["cor"]) / rev if rev and v["cor"] is not None else None, PCT, fill=fill)
        put(ws, r, 6, v["rnd"], NUM, fill=fill)
        put(ws, r, 7, v["rnd"] / rev if rev and v["rnd"] is not None else None, PCT, fill=fill)
        put(ws, r, 8, v["sm"], NUM, fill=fill)
        put(ws, r, 9, v["sm"] / rev if rev and v["sm"] is not None else None, PCT, fill=fill)
        put(ws, r, 10, v["ga"], NUM, fill=fill)
        put(ws, r, 11, v["ga"] / rev if rev and v["ga"] is not None else None, PCT, fill=fill)
        put(ws, r, 12, v["opinc"], NUM, fill=fill)
        put(ws, r, 13, v["opinc"] / rev if rev and v["opinc"] is not None else None, PCT, fill=fill)
        if with_growth:
            pv = row_values(name, data, "q_prior")
            for j, m in enumerate(("rnd", "cor", "sm")):
                g = (v[m] / pv[m] - 1) if v.get(m) is not None and pv.get(m) else None
                put(ws, r, 14 + j, g, PCT, fill=fill)
        r += 1
    return first, r  # first data row, next free row


def sheet_summary(wb, data):
    ws = wb.active; ws.title = "Summary"
    title(ws, "Anthropic vs. Magnificent 7 — R&D, cost of revenue and sales & marketing",
          "US$ millions unless stated. Mag 7 from SEC 10-K/10-Q (XBRL companyfacts). Anthropic from Form S-1/A filed Nov 2, 2026 "
          "(Note 8 functional presentation). Built " + dt.date.today().strftime("%b %d, %Y") + ".")
    r = 4
    first_ttm, r = write_comparison_block(ws, r, data, "ttm", "A. Trailing twelve months — most recent reported period")
    r += 1
    first_fy, r = write_comparison_block(ws, r, data, "fy", "B. Latest completed fiscal year")
    r += 1
    first_q, r = write_comparison_block(ws, r, data, "q", "C. Most recent quarter (with year-over-year growth)", with_growth=True)
    r += 1
    notes = [
        "¹ Amazon does not report R&D. 'Technology and infrastructure' is shown in the R&D column; it also contains AWS infrastructure "
        "operating costs (incl. server depreciation), so it overstates R&D relative to peers. Amazon 'Cost of sales' excludes 'Fulfillment' "
        "(FY2025: $109,074M; 9M 2026: $86,919M), which most peers would put in cost of revenue.",
        "² Nvidia and Tesla report a single 'Selling, general and administrative' line. That combined SG&A figure is shown in the S&M column "
        "and G&A is left blank. Nvidia's latest period is its fiscal Q2 2027 (quarter to Jul 26, 2026); fiscal Q3 is reported Nov 2026.",
        "³ Meta's Q3 2026 G&A (and therefore TTM G&A and operating income) includes a $9.89B charge for the August 2026 settlement with "
        "52 US state attorneys general (teen mental-health litigation). Excluding it, TTM G&A ≈ $16.9B and TTM operating income ≈ $89.3B.",
        "Anthropic: the S-1 primary statements present expenses by nature (compute & infrastructure, personnel, etc.). The functional split used "
        "here is the supplemental Note 8 table. Cost of revenue is 'substantially all' compute for customer-serving workloads plus some personnel "
        "and overhead; R&D is mainly compute for model development, research and internal Claude use plus personnel; S&M is mainly channel-partner "
        "(AWS/Google Cloud/Azure marketplace) fees, marketing and compute for free-tier access. FY2025 and 9M 2025 operating income include the "
        "$1,500M Bartz copyright settlement, which is excluded from the G&A column. Q3 figures are derived as 9M less H1; TTM as FY2025 + 9M 2026 − 9M 2025.",
        "Anthropic Q3 2026 G&A of $2,781M includes a non-cash charge for the company's donation-matching program; Q3 2026 compute expense "
        "includes $5.0B of short-term compute that 'currently costs significantly more' than multi-year capacity (S-1/A MD&A).",
        "Gross margin for Anthropic is derived (revenue less functional cost of revenue); the prospectus does not present gross profit or gross margin.",
        "Apple began disclosing S&M and G&A separately in its FY2025 10-K (FY2024 restated); earlier years are SG&A only.",
        "Microsoft fiscal year ends Jun 30; Apple late September; Nvidia late January. Alphabet, Amazon, Meta and Tesla use calendar years.",
    ]
    section(ws, r, "Notes", 15); r += 1
    for n in notes:
        note(ws, r, n, 15); r += 1

    # Charts off the TTM block
    n = len(COMPANIES)
    cats = Reference(ws, min_col=1, min_row=first_ttm, max_row=first_ttm + n - 1)

    def bar(title_, col, fmt, anchor, y_title):
        ch = BarChart(); ch.type = "col"; ch.style = 10; ch.title = title_
        ch.y_axis.title = y_title; ch.y_axis.numFmt = fmt; ch.legend = None
        ch.add_data(Reference(ws, min_col=col, min_row=first_ttm, max_row=first_ttm + n - 1), titles_from_data=False)
        ch.set_categories(cats); ch.height = 7.5; ch.width = 16
        ws.add_chart(ch, anchor)

    bar("R&D as % of revenue (TTM)", 7, "0%", "Q4", "% of revenue")
    bar("Cost of revenue as % of revenue (TTM)", 4, "#,##0", "Q20", "US$ M")
    ws._charts[-1].title = "Cost of revenue, US$ M (TTM)"
    bar("Sales & marketing as % of revenue (TTM)", 9, "0%", "Q36", "% of revenue")
    bar("R&D, US$ M (TTM)", 6, "#,##0", "Q52", "US$ M")

    widths(ws, [14, 44, 11, 13, 11, 11, 10, 13, 10, 11, 10, 13, 11, 11, 13, 11])
    ws.freeze_panes = "C6"
    return ws


METRIC_LABELS = [("revenue", "Revenue"), ("cor", "Cost of revenue"), ("rnd", "Research & development"),
                 ("sm", "Sales & marketing"), ("ga", "General & administrative"), ("opinc", "Operating income")]


def sheet_detail(wb, data):
    ws = wb.create_sheet("Opex Detail")
    title(ws, "Line-item detail by company — fiscal years, year-to-date, trailing twelve months and latest quarter",
          "US$ millions. Period columns are company-specific (see 'Period key'). Shaded green = derived (TTM, quarter from YTD differencing, Anthropic Q3/TTM).")
    r = 4
    section(ws, r, "Period key (what each column means for each company)", 12); r += 1
    header_row(ws, r, 1, ["Company", "FY-2", "FY-1", "Latest FY", "Prior-yr YTD", "Latest YTD", "TTM", "Prior-yr Qtr", "Latest Qtr", "Source"]); r += 1
    for name, cik in COMPANIES:
        if name == "Anthropic":
            vals = ["Anthropic", "FY2023", "FY2024", "FY2025", "9M 2025", "9M 2026", "TTM to Sep 30, 2026", "Q3 2025 (derived)", "Q3 2026 (derived)",
                    "Form S-1/A (Nov 2, 2026), Note 8"]
            fill = FILL_ANT
        else:
            p = data[name]["revenue"]
            ann = [fy_label(k[1]) for k, _ in p["annual"]]
            while len(ann) < 3: ann.insert(0, "n/a")
            ytd = f"{p['ytd'][0][0]} → {p['ytd'][0][1]}" if p.get("ytd") else "n/a (FY = latest)"
            pytd = f"{p['ytd_prior'][0][0]} → {p['ytd_prior'][0][1]}" if p.get("ytd_prior") else "n/a"
            vals = [name] + ann + [pytd, ytd, ttm_label(name, p), "same qtr, prior yr", q_label(p), f"SEC XBRL companyfacts CIK {cik}"]
            fill = None
        for c, v in enumerate(vals, start=1):
            put(ws, r, c, v, fill=fill, font=F_NOTE if c > 1 else F_NORM)
        r += 1
    r += 1

    cols = ["Company", "FY-2", "FY-1", "Latest FY", "Prior-yr YTD", "Latest YTD", "TTM", "Prior-yr Qtr", "Latest Qtr", "Qtr YoY %", "TTM % of revenue", "XBRL tag / source"]

    def anth_vals(idx):
        A = ANTHROPIC_FUNCTIONAL
        return [A["FY2023"][idx], A["FY2024"][idx], A["FY2025"][idx], A["9M 2025"][idx], A["9M 2026"][idx],
                A["TTM Sep-2026"][idx], A["Q3 2025"][idx], A["Q3 2026"][idx]]

    anth_idx = {"revenue": 0, "cor": 1, "rnd": 2, "sm": 3, "ga": 4, "opinc": 6}
    for metric, label in METRIC_LABELS:
        section(ws, r, label, 12); r += 1
        header_row(ws, r, 1, cols); r += 1
        for i, (name, _) in enumerate(COMPANIES):
            fill = FILL_ANT if name == "Anthropic" else (FILL_ALT if i % 2 == 0 else None)
            if name == "Anthropic":
                vals = anth_vals(anth_idx[metric]); rev_ttm = ANTHROPIC_FUNCTIONAL["TTM Sep-2026"][0]; tag = "S-1/A Note 8 (functional)"
                derived = {6, 7, 8}
            else:
                p = data[name].get(metric)
                if not p:
                    put(ws, r, 1, name, fill=fill)
                    put(ws, r, 2, "not separately disclosed (see SG&A in S&M row)" if metric == "ga" else "n/a", fill=fill, font=F_NOTE)
                    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12); r += 1; continue
                ann = [M(v) for _, v in p["annual"]]
                while len(ann) < 3: ann.insert(0, None)
                vals = ann + [M(p["ytd_prior"][1]) if p.get("ytd_prior") else None, M(p["ytd"][1]) if p.get("ytd") else None,
                              M(p["ttm"]), M(p["q_prior"][1]) if p.get("q_prior") else None, M(p["q"][1]) if p.get("q") else None]
                rev_ttm = M(data[name]["revenue"]["ttm"]); tag = p["tag"]
                if p.get("combined_sga"): tag += "  (combined SG&A — S&M not split out)"
            # column 7 = TTM (always derived); Anthropic quarters (cols 8-9) are derived from 9M − H1
            derived_cols = {7} | ({8, 9} if name == "Anthropic" else set())
            put(ws, r, 1, name, bold=(name == "Anthropic"), fill=fill)
            for j, v in enumerate(vals, start=2):
                put(ws, r, j, v, NUM, fill=FILL_DERIVED if j in derived_cols else fill)
            q, pq = vals[7], vals[6]
            put(ws, r, 10, (q / pq - 1) if q is not None and pq else None, PCT, fill=fill)
            put(ws, r, 11, vals[5] / rev_ttm if vals[5] is not None and rev_ttm else None, PCT, fill=fill)
            put(ws, r, 12, tag, fill=fill, font=F_NOTE)
            r += 1
        r += 1

    # % of revenue matrix (TTM and latest FY)
    section(ws, r, "Cost structure as % of revenue", 12); r += 1
    header_row(ws, r, 1, ["Company", "Basis", "Cost of revenue", "R&D", "S&M", "G&A", "Operating margin", "Total opex (ex-CoR) % rev"]); r += 1
    for which, basis in (("ttm", "TTM"), ("fy", "Latest FY")):
        for i, (name, _) in enumerate(COMPANIES):
            v = row_values(name, data, which); rev = v["revenue"]
            fill = FILL_ANT if name == "Anthropic" else (FILL_ALT if i % 2 == 0 else None)
            put(ws, r, 1, name, fill=fill); put(ws, r, 2, basis, fill=fill)
            for c, m in zip(range(3, 7), ("cor", "rnd", "sm", "ga")):
                put(ws, r, c, v[m] / rev if v.get(m) is not None and rev else None, PCT, fill=fill)
            put(ws, r, 7, v["opinc"] / rev if v["opinc"] is not None and rev else None, PCT, fill=fill)
            opex = sum(x for x in (v["rnd"], v["sm"], v["ga"]) if x is not None)
            put(ws, r, 8, opex / rev if rev else None, PCT, fill=fill)
            r += 1
        r += 1
    note(ws, r, "Amazon 'Cost of revenue' = cost of sales only (fulfillment excluded); Amazon R&D column = technology & infrastructure. "
                "Nvidia/Tesla S&M column = combined SG&A. Anthropic quarter and TTM figures are derived from filed FY/9M/H1 functional tables.", 12)
    widths(ws, [14, 12, 12, 12, 13, 13, 13, 13, 13, 11, 14, 46])
    ws.freeze_panes = "B5"
    return ws


def sheet_inference_training(wb, data):
    ws = wb.create_sheet("Inference vs Training")
    title(ws, "Inference vs. training spend — Anthropic (filed proxy) vs. Mag 7 (AI infrastructure capex and disclosed positioning)",
          "US$ millions unless stated. No company in this set publishes a dollar split of compute between training and inference. "
          "Anthropic's functional expense allocation is the closest filed proxy; for the Mag 7 the comparable datapoint is total capex plus management commentary.")
    r = 4
    # ---- Section A: Anthropic filed proxy
    section(ws, r, "A. Anthropic — compute by workload, as allocated in the S-1/A functional presentation (Note 8)", 12); r += 1
    pers = ["FY2023", "FY2024", "FY2025", "9M 2025", "9M 2026", "TTM Sep-2026"]
    header_row(ws, r, 1, ["Line ($M)"] + pers + ["How to read", "", ""]); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10); r += 1
    BN, FN = ANTHROPIC_BY_NATURE, ANTHROPIC_FUNCTIONAL
    rows = [
        ("Revenue", [BN[p][0] for p in pers], NUM, None, "Consumption (API) ~80% / subscriptions ~20% of H1 2026 revenue."),
        ("Total compute & infrastructure expense (by nature)", [BN[p][1] for p in pers], NUM, None,
         "All workloads: customer serving, free tier, model training, research, internal Claude use. One fungible pool across AWS Trainium, Google TPU, Nvidia and AMD GPUs."),
        ("  Cost of revenue (functional) → customer-serving INFERENCE, upper bound", [FN[p][1] for p in pers], NUM, None,
         "'Substantially all' compute for customer-serving workloads, plus some personnel and allocated overhead. Treat as a ceiling on paid-inference compute."),
        ("  Compute outside cost of revenue → TRAINING / research / internal use / free tier, lower bound", [BN[p][1] - FN[p][1] for p in pers], NUM, FILL_DERIVED,
         "Derived: total compute less cost of revenue. Sits mostly in R&D (model development, research, internal usage) with a slice in S&M (free access)."),
        ("    Customer-serving share of compute (ceiling)", [FN[p][1] / BN[p][1] for p in pers], PCT, FILL_DERIVED, "Cost of revenue ÷ total compute."),
        ("    Training/research/internal/free share of compute (floor)", [1 - FN[p][1] / BN[p][1] for p in pers], PCT, FILL_DERIVED, "1 − the line above."),
        ("R&D expense (functional)", [FN[p][2] for p in pers], NUM, None,
         "Primarily compute for model development, research and internal usage, plus personnel, software and overhead. The main home of training spend."),
        ("S&M expense (functional)", [FN[p][3] for p in pers], NUM, None,
         "Channel-partner fees, marketing, personnel and compute for FREE access to Claude (an inference cost outside cost of revenue)."),
        ("Compute as % of revenue", [BN[p][1] / BN[p][0] for p in pers], PCT, FILL_DERIVED, "Fell from 1,056% (FY2023) to 57% (9M 2026); 48% in Q2 2026, 58% in Q3 2026."),
        ("Revenue per $1 of compute", [BN[p][0] / BN[p][1] for p in pers], MULT, FILL_DERIVED, "Company's preferred efficiency metric; peaked ~2.08x in Q2 2026, 1.72x in Q3 2026."),
        ("Derived gross margin (revenue − functional cost of revenue)", [(FN[p][0] - FN[p][1]) / FN[p][0] for p in pers], PCT, FILL_DERIVED,
         "Not a company-reported metric; excludes training, research, internal and free-tier compute."),
    ]
    for lbl, vals, fmt, fill, how in rows:
        put(ws, r, 1, lbl, bold=lbl.startswith(("Revenue", "Total", "R&D", "S&M")), fill=fill)
        for j, v in enumerate(vals, start=2):
            put(ws, r, j, v, fmt, fill=fill)
        put(ws, r, 8, how, fill=fill, font=F_NOTE, align=WRAP)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        ws.row_dimensions[r].height = 40
        r += 1
    r += 1
    # Quarterly compute
    section(ws, r, "A2. Anthropic — quarterly compute & infrastructure vs. revenue (S-1/A MD&A quarterly table)", 12); r += 1
    header_row(ws, r, 1, ["Quarter", "Revenue", "Compute & infra", "Compute % rev", "Rev per $ compute", "Personnel", "Sales, mktg & partnerships", "Other op & admin", "Total opex", "Operating income", "Net loss", "Note"]); r += 1
    qnotes = {"Q3 2025": "Opex incl. $1,500M Bartz settlement", "Q2 2026": "First operating profit", "Q3 2026": "Incl. $5.0B short-term compute; other op & admin incl. donation charge"}
    for q in ANTHROPIC_QUARTERLY:
        qn, rev, comp, pers_, smp, oth, opex, opi, nl = q
        put(ws, r, 1, qn, bold=True); put(ws, r, 2, rev, NUM); put(ws, r, 3, comp, NUM)
        put(ws, r, 4, comp / rev, PCT, fill=FILL_DERIVED); put(ws, r, 5, rev / comp, MULT, fill=FILL_DERIVED)
        put(ws, r, 6, pers_, NUM); put(ws, r, 7, smp, NUM); put(ws, r, 8, oth, NUM); put(ws, r, 9, opex, NUM); put(ws, r, 10, opi, NUM); put(ws, r, 11, nl, NUM)
        put(ws, r, 12, qnotes.get(qn, ""), font=F_NOTE)
        r += 1
    r += 1
    # Third-party estimates and commitments
    section(ws, r, "A3. Anthropic — outside estimates of the training / inference split, and forward compute commitments", 12); r += 1
    header_row(ws, r, 1, ["Item", "Period", "Training / R&D compute", "Inference compute", "Basis", "", "", "Source", ""])
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9); r += 1
    est = [
        ("Epoch AI estimate (from The Information reporting)", "FY2025", 4100, 2700, "$4.1B 'spent training models' taken as R&D compute; inference compute inferred from reported ~40% gross margin on ~$4.5B revenue. Filed FY2025 cost of revenue was $2,773M and total compute $7,326M.", "epoch.ai/data-insights/company-spending-breakdown"),
        ("Tom's Hardware reading of S-1", "H1 2026", 5930, 2380, "Training/research/internal ≥ $5.93B (at least); customer-serving = $2.38B cost of revenue; free-tier access ≤ ~$640M. Serving share of compute ≤ 27%.", "tomshardware.com (Oct 2026)"),
        ("Filed functional proxy (this workbook)", "9M 2026", 18808 - 4812, 4812, "Total compute $18,808M less cost of revenue $4,812M = $13,996M outside customer serving (includes free tier). Company does not disclose the split.", "S-1/A Note 8 and statements of operations"),
        ("Filed functional proxy (this workbook)", "Q3 2026", 9863 - 2432, 2432, "Q3 compute $9,863M less derived Q3 cost of revenue $2,432M. Q3 includes $5.0B of short-term compute bought at a premium to multi-year rates; $6.5B short-term for 9M 2026.", "S-1/A MD&A"),
    ]
    for item, per, tr, inf, basis, src in est:
        put(ws, r, 1, item, bold=True); put(ws, r, 2, per); put(ws, r, 3, tr, NUM); put(ws, r, 4, inf, NUM)
        put(ws, r, 5, basis, font=F_NOTE, align=WRAP); put(ws, r, 8, src, font=F_NOTE, align=WRAP)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 48; r += 1
    r += 1
    header_row(ws, r, 1, ["Compute commitment (as of / subsequent to Sep 30, 2026)", "Counterparty", "Minimum spend ($M)", "Window", "Terms", "", "", "Source", ""])
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9); r += 1
    commits = [
        ("Aggregate purchase commitments + leases (incl. not yet commenced)", "All (AWS, Google, Microsoft/Nvidia, Broadcom, data-center lessors)", 895400, "≈$447.5B expected before Dec 31, 2030", "'Substantially all' related to compute and infrastructure; take-or-pay / regardless of usage."),
        ("Google Cloud CSP hosting agreement (TPUs)", "Alphabet", 121300, "Apr 2026 – Jul 2033", "Non-cancelable; shortfall paid as true-up."),
        ("Amazon compute agreement (Trainium / AWS)", "Amazon", 110000, "May 2026 – Apr 2036", "Non-cancelable; shortfall paid as true-up."),
        ("Microsoft Azure agreement (Nvidia GPUs)", "Microsoft", 44900, "Nov 2026 – Jun 2033", "Non-cancelable except uncured breach / undelivered tranches."),
    ]
    for a, b, c, d_, e in commits:
        put(ws, r, 1, a, bold=True, align=WRAP); put(ws, r, 2, b, align=WRAP); put(ws, r, 3, c, NUM); put(ws, r, 4, d_, align=WRAP); put(ws, r, 5, e, font=F_NOTE, align=WRAP)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        put(ws, r, 8, "Form S-1/A (Nov 2, 2026), Commitments / Liquidity", font=F_NOTE, align=WRAP); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 36; r += 1
    r += 1

    # ---- Section B: Mag 7
    section(ws, r, "B. Mag 7 — AI infrastructure spend (total capex, as filed) and what each company says about training vs. inference", 12); r += 1
    hdr = ["Company", "Role in the stack", "Capex FY-2 ($M)", "Capex FY-1 ($M)", "Capex latest FY ($M)", "Capex TTM ($M)", "Capex TTM % rev",
           "2026 capex guidance", "Discloses training/inference $ split?", "Management positioning on training vs. inference", "Model-development spend proxy (R&D TTM, $M)", "Sources"]
    header_row(ws, r, 1, hdr, height=42); r += 1
    guidance = {
        "Microsoft": ("≈$170B calendar 2026 incl. finance leases (cut from ≈$175B on Oct 28, 2026)",
                      "Hosts OpenAI and MAI models on Azure; trains its own MAI models. Oct 2026: scaling Maia 200 custom silicon to be its 'primary inference platform' across Copilot, GitHub Copilot and Foundry. CFO: third-party GPU delays push some capacity into H1 2027; capex to accelerate in H1 2027. Anthropic committed ≥$44.9B to Azure (Nov 2026–Jun 2033).",
                      "Platform / inference host + model trainer"),
        "Alphabet": ("$195–205B for 2026 (held Oct 28, 2026); 'significant increase' expected in 2027",
                     "Trains Gemini; serves inference across Search, Workspace, Cloud and via Anthropic (≥$121.3B TPU commitment, 2026–2033). Split TPU 8 into TPU 8t (training) and TPU 8i (inference) in 2026. >3.2 quadrillion tokens/month processed (May 2026). No $ split disclosed.",
                     "Vertically integrated: trainer + inference host + chip designer"),
        "Amazon": ("≈$220B cash capex for 2026 (held Oct 29, 2026); 'meaningfully more' in 2027",
                   "Primarily an inference/training host via AWS (Trainium, Inferentia, Nvidia). Anthropic committed ≥$110B (2026–2036). CFO: Q3 capex 'primarily related to AWS and AI'. Capex also covers fulfillment, robotics and Kuiper. No $ split disclosed.",
                   "Cloud / inference host (plus Nova models)"),
        "Meta": ("$135–145B for 2026 incl. finance-lease principal (narrowed Oct 21, 2026); 'continue investing aggressively' in 2027",
                 "Only Mag 7 member with a published fleet split: ~10% experimentation / 20% training / 70% inference (power capacity, 2021 paper — pre-LLM era). Q3 2026: capacity gives 'flexibility to allocate resources across model training, core business growth, and ... new products'. MSL (superintelligence) training clusters Prometheus/Hyperion.",
                 "Model trainer + first-party inference (ads ranking, Meta AI)"),
        "Apple": ("None published", "Minimal frontier training; on-device inference plus Private Cloud Compute. Capex ($9.0B FY2026) is a fraction of hyperscaler levels. No $ split disclosed.",
                  "Device-side inference; buys models/compute"),
        "Nvidia": ("n/a (supplier)", "Supplier of training and inference accelerators. Said inference was 'more than half' of data-center compute revenue (Feb 2025 investor materials). Own capex is small; R&D is for chips/systems/software, not model training. FY2026 data-center revenue $193.7B; Q2 FY27 $89.0B.",
                   "Accelerator supplier"),
        "Tesla": ("≈$20B for 2026 (cut from ≈$25B on Q3 2026 call); AI compute, robotaxi fleet, Optimus, semiconductor fab",
                  "Trains FSD/Optimus models on Cortex 1 (>90 MW) and Cortex 2 (>115 MW) at Gigafactory Texas; onsite compute more than doubled in H1 2026. Inference runs on in-vehicle AI chips (AI5/AI6 with Samsung). Historically training-heavy by $; no $ split disclosed.",
                  "Model trainer (FSD) + edge inference"),
    }
    src_common = "SEC XBRL (PaymentsToAcquirePropertyPlantAndEquipment / PaymentsToAcquireProductiveAssets); Q3 2026 earnings releases & calls (CNBC Oct 28–29, 2026; Meta IR Oct 21, 2026); Anthropic S-1/A."
    for i, (name, _) in enumerate(COMPANIES):
        if name == "Anthropic":
            put(ws, r, 1, "Anthropic (for scale)", bold=True, fill=FILL_ANT); put(ws, r, 2, "Model trainer; rents compute (opex, not capex)", fill=FILL_ANT, align=WRAP)
            for c, v in zip(range(3, 7), (BN["FY2023"][1], BN["FY2024"][1], BN["FY2025"][1], BN["TTM Sep-2026"][1])):
                put(ws, r, c, v, NUM, fill=FILL_ANT)
            put(ws, r, 7, BN["TTM Sep-2026"][1] / BN["TTM Sep-2026"][0], PCT, fill=FILL_ANT)
            put(ws, r, 8, "Compute & infrastructure OPEX shown instead of capex. 9M 2026 capex only $925M + $1,955M lease prepayments; $895.4B of commitments ahead.", fill=FILL_ANT, font=F_NOTE, align=WRAP)
            put(ws, r, 9, "Partially — functional allocation (Section A)", fill=FILL_ANT, align=WRAP)
            put(ws, r, 10, "Compute is one fungible pool 'dynamically directed' between training, research, customer and internal workloads; 'most of our compute bill' is not paid inference. Agentic workloads 'consume significantly more compute per user interaction'.", fill=FILL_ANT, font=F_NOTE, align=WRAP)
            put(ws, r, 11, FN["TTM Sep-2026"][2], NUM, fill=FILL_ANT)
            put(ws, r, 12, "Form S-1/A (Nov 2, 2026)", fill=FILL_ANT, font=F_NOTE, align=WRAP)
            ws.row_dimensions[r].height = 96; r += 1; continue
        fill = FILL_ALT if i % 2 == 0 else None
        cap = data[name].get("capex"); rev = data[name]["revenue"]
        ann = [M(v) for _, v in cap["annual"]] if cap else [None] * 3
        while len(ann) < 3: ann.insert(0, None)
        g, pos, role = guidance[name]
        put(ws, r, 1, name, bold=True, fill=fill); put(ws, r, 2, role, fill=fill, align=WRAP)
        for c, v in zip(range(3, 6), ann): put(ws, r, c, v, NUM, fill=fill)
        put(ws, r, 6, M(cap["ttm"]) if cap else None, NUM, fill=fill)
        put(ws, r, 7, cap["ttm"] / rev["ttm"] if cap and rev["ttm"] else None, PCT, fill=fill)
        put(ws, r, 8, g, fill=fill, font=F_NOTE, align=WRAP)
        put(ws, r, 9, "Rough % split (2021, power capacity) — no $" if name == "Meta" else ("Directional only (inference >50% of DC compute revenue)" if name == "Nvidia" else "No"), fill=fill, align=WRAP)
        put(ws, r, 10, pos, fill=fill, font=F_NOTE, align=WRAP)
        put(ws, r, 11, M(data[name]["rnd"]["ttm"]), NUM, fill=fill)
        put(ws, r, 12, src_common, fill=fill, font=F_NOTE, align=WRAP)
        ws.row_dimensions[r].height = 96; r += 1
    r += 1
    notes_b = [
        "Capex here is cash purchases of property & equipment as tagged in XBRL. Microsoft and Meta guide on a basis that adds finance-lease principal/assets, so guidance is not directly comparable to the XBRL columns. "
        "Amazon's XBRL tag is gross purchases (before incentives/sale-leaseback proceeds). Fiscal years: Microsoft Jun, Apple Sep, Nvidia Jan; others Dec. Capex TTM for Alphabet, Meta, Tesla and Amazon runs to Sep 30, 2026; "
        "Microsoft to Sep 30, 2026 (FY26 + Q1 FY27 − Q1 FY26); Nvidia to Jul 26, 2026; Apple = FY2026.",
        "Hyperscaler capex is company-wide and includes non-AI spending (Amazon fulfillment/Kuiper, Alphabet/Microsoft office and network, etc.); none of the Mag 7 reports AI-only capex or a training-vs-inference dollar split. "
        "Anthropic's model (renting capacity under take-or-pay contracts) puts the same economic cost through operating expense rather than capex, so compare its compute OPEX to peers' capex only as an order-of-magnitude check.",
        "Amazon's 'Technology and infrastructure' (R&D proxy) is where AWS server depreciation lands, so for Amazon the R&D proxy already embeds a large inference/hosting cost; for Alphabet, Meta and Microsoft, depreciation on AI servers flows mainly through cost of revenue.",
    ]
    section(ws, r, "Notes to Section B", 12); r += 1
    for n in notes_b: note(ws, r, n, 12); r += 1
    r += 1

    # ---- Section C: industry reference points
    section(ws, r, "C. Industry reference points on the training ↔ inference mix (for triangulation)", 12); r += 1
    header_row(ws, r, 1, ["Reference", "Training share", "Inference share", "What it measures", "Source / date"]); r += 1
    refs = [
        ("Deloitte TMT Predictions 2026", "≈1/3", "≈2/3", "Share of all AI compute in 2026 (inference was ~1/3 in 2023, ~1/2 in 2025). Projection, not measurement.", "Deloitte, Nov 18, 2025"),
        ("Gartner", "≈45%", "≈55% (→ >65% by 2029)", "Share of AI-optimized IaaS spending supporting inference, 2026.", "Gartner via Introl (2026)"),
        ("Meta (Wu et al., 'Sustainable AI')", "≈20% training + 10% experimentation", "≈70%", "Rough split of Meta's AI power capacity; mostly ranking/recommendation models, pre-LLM.", "arXiv, Oct 2021"),
        ("Nvidia", "<50%", ">50%", "Share of Nvidia data-center compute revenue attributed to inference.", "Nvidia FY2025 Q4 investor materials, Feb 26, 2025"),
        ("OpenAI 2024 (Epoch AI estimate)", "≈$5.0B R&D compute", "≈$1.8B inference", "Lab-level compute spend; final training runs only ~9.6% of R&D compute. Hyperscaler-hosted inference (Azure) adds ~25% (up to 50%) on top of lab totals.", "Epoch AI, Oct 10, 2025 / Mar 23, 2026 / May 20, 2026"),
        ("Anthropic 2025 (Epoch AI estimate)", "≈$4.1B R&D compute (42% of spend)", "≈$2.7B inference (28% of spend)", "Lab-level; based on The Information reporting. Filed FY2025 total compute was $7,326M.", "Epoch AI, 2026"),
        ("Anthropic 9M 2026 (filed proxy, this workbook)", "≥ ~74% of compute (training/research/internal/free)", "≤ ~26% of compute (customer-serving)", "Functional allocation of $18,808M compute: $4,812M cost of revenue vs. $13,996M elsewhere.", "Form S-1/A Note 8"),
    ]
    for a, b, c, d_, e in refs:
        put(ws, r, 1, a, bold=True, align=WRAP); put(ws, r, 2, b, align=WRAP); put(ws, r, 3, c, align=WRAP); put(ws, r, 4, d_, font=F_NOTE, align=WRAP); put(ws, r, 5, e, font=F_NOTE, align=WRAP)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8); ws.row_dimensions[r].height = 42; r += 1
    r += 1
    section(ws, r, "Key takeaways", 12); r += 1
    take = [
        "Anthropic is the inverse of the hyperscalers: its filed numbers show at most ~26% of 9M 2026 compute went to paid customer serving, with ≥74% in training, research, internal use and free access — whereas industry estimates put inference at roughly two-thirds of total AI compute in 2026 and Meta's own (older) fleet split at ~70% inference.",
        "Because Anthropic rents capacity, its 'training + inference' bill is an operating expense ($18.8B in 9M 2026; $21.8B TTM) rather than capex. On a TTM basis that equals ~14% of Alphabet's capex, ~16% of Microsoft's and ~10% of Amazon's gross capex, and ~28–29% of Meta's ($76.5B) or Alphabet's ($74.5B) TTM R&D — the closest 'model-development' line at the peers.",
        "Anthropic's R&D intensity (62% of TTM revenue; 58% in 9M 2026) is 2–8x the Mag 7 (Meta ~32%, Alphabet ~16%, Microsoft ~11%, Apple ~10%, Tesla ~9%, Nvidia ~8%), while its derived gross margin (~84% TTM, 85.5% 9M 2026) is now at or above software-peer levels — the gross margin is high precisely because training compute is classified outside cost of revenue.",
        "Sales & marketing is the line where Anthropic diverges most from a software model: 20% of TTM revenue, dominated by cloud-marketplace channel fees (~48% of H1 2026 revenue came through AWS/Google Cloud/Azure) and free-tier compute, versus 5–8% at Microsoft, Alphabet and Meta.",
        "Forward commitments ($895.4B aggregate; ~$447.5B before end-2030) are 25x TTM revenue and are take-or-pay, so Anthropic's training/inference mix will be set by how fast paid inference demand fills capacity that arrives on a fixed schedule — the same utilisation risk the hyperscalers carry on their own capex.",
    ]
    for t in take: note(ws, r, "• " + t, 12); r += 1

    widths(ws, [44, 20, 16, 16, 18, 16, 14, 40, 26, 60, 18, 36])
    ws.freeze_panes = "B4"
    return ws


def sheet_anthropic(wb):
    ws = wb.create_sheet("Anthropic Detail")
    title(ws, "Anthropic, PBC — statements of operations (by nature) and functional supplemental table, Form S-1/A filed Nov 2, 2026",
          "US$ millions. Nine-month and quarterly figures unaudited. Q3 and TTM columns derived.")
    r = 4
    section(ws, r, "Operating expenses by nature (primary presentation)", 10); r += 1
    pers = ["FY2023", "FY2024", "FY2025", "9M 2025", "9M 2026", "TTM Sep-2026"]
    header_row(ws, r, 1, ["Line"] + pers); r += 1
    labels = ["Revenue", "Compute and infrastructure", "Personnel", "Sales, marketing, and partnerships", "Other operating and administrative", "Bartz settlement", "Total operating expenses", "Operating income (loss)"]
    for i, lbl in enumerate(labels):
        put(ws, r, 1, lbl, bold=lbl in ("Revenue", "Total operating expenses", "Operating income (loss)"))
        for j, p in enumerate(pers, start=2):
            put(ws, r, j, ANTHROPIC_BY_NATURE[p][i], NUM, fill=FILL_DERIVED if p.startswith("TTM") else None)
        r += 1
    r += 1
    section(ws, r, "Operating expenses by function (Note 8, Supplemental Operating Expense Information)", 10); r += 1
    pers2 = ["FY2023", "FY2024", "FY2025", "H1 2025", "H1 2026", "9M 2025", "9M 2026", "Q3 2025", "Q3 2026", "TTM Sep-2026"]
    header_row(ws, r, 1, ["Line"] + pers2); r += 1
    labels2 = ["Revenue", "Cost of revenue", "Research and development", "Sales and marketing", "General and administrative", "Bartz settlement", "Operating income (loss)"]
    for i, lbl in enumerate(labels2):
        put(ws, r, 1, lbl, bold=lbl in ("Revenue", "Operating income (loss)"))
        for j, p in enumerate(pers2, start=2):
            put(ws, r, j, ANTHROPIC_FUNCTIONAL[p][i], NUM, fill=FILL_DERIVED if p.startswith(("Q3", "TTM")) else None)
        r += 1
    # ratios
    ratio_rows = [("Gross margin (derived)", lambda t: (t[0] - t[1]) / t[0]), ("Cost of revenue % rev", lambda t: t[1] / t[0]),
                  ("R&D % rev", lambda t: t[2] / t[0]), ("S&M % rev", lambda t: t[3] / t[0]), ("G&A % rev", lambda t: t[4] / t[0]),
                  ("Operating margin", lambda t: t[6] / t[0])]
    for lbl, f in ratio_rows:
        put(ws, r, 1, lbl, font=F_NOTE)
        for j, p in enumerate(pers2, start=2):
            put(ws, r, j, f(ANTHROPIC_FUNCTIONAL[p]), PCT, fill=FILL_DERIVED)
        r += 1
    r += 1
    section(ws, r, "Quarterly (by nature)", 10); r += 1
    header_row(ws, r, 1, ["Quarter", "Revenue", "Compute & infra", "Personnel", "Sales, mktg & partnerships", "Other op & admin", "Total opex", "Operating income", "Net loss", "Rev QoQ %"]); r += 1
    prev = None
    for q in ANTHROPIC_QUARTERLY:
        put(ws, r, 1, q[0], bold=True)
        for j, v in enumerate(q[1:], start=2): put(ws, r, j, v, NUM)
        put(ws, r, 10, (q[1] / prev - 1) if prev else None, PCT, fill=FILL_DERIVED); prev = q[1]
        r += 1
    r += 1
    notes = [
        "Footnotes to Note 8: (1) Substantially all of cost of revenue consists of compute and infrastructure costs associated with customer-serving workloads, plus certain personnel-related and allocated overhead costs. "
        "(2) R&D primarily consists of compute and infrastructure costs for model development, research and internal usage workloads, personnel, software and allocated overhead. "
        "(3) S&M primarily consists of channel partner costs, marketing and advertising, personnel, compute for free access to the intelligence service, and allocated overhead. "
        "(4) G&A primarily consists of the donation matching obligation, personnel, facilities, allocated overhead, legal fees and non-income-based taxes.",
        "During 2025 the company changed its primary presentation from functional (cost of revenue / gross profit / R&D / S&M / G&A) to by-nature categories and recast prior periods; the functional view survives only in Note 8.",
        "Net loss is dominated by non-cash fair-value changes on convertible notes, derivative liabilities and SAFEs ($94.3B in 9M 2026; $135.1B of the $156.2B accumulated deficit since inception).",
        "Compute is expensed when capacity becomes available (typically procured 12–18 months ahead). 9M 2026 included $6.5B of short-term compute commitments ($5.0B in Q3 2026).",
    ]
    for n in notes: note(ws, r, n, 11); r += 1
    widths(ws, [40] + [13] * 10)
    ws.freeze_panes = "B6"
    return ws


def sheet_sources(wb, data):
    ws = wb.create_sheet("Sources & Method")
    title(ws, "Sources, definitions and method")
    r = 3
    section(ws, r, "Primary sources", 6); r += 1
    header_row(ws, r, 1, ["Company", "Document", "Period covered", "Where / how retrieved", "Retrieved"]); r += 1
    today = dt.date.today().isoformat()
    rows = [("Anthropic", "Form S-1/A (Amendment No. 1) — registration statement", "FY2023–FY2025 audited; 9M 2025 / 9M 2026 unaudited; quarterly Q1 2025–Q3 2026",
             "sec.gov/Archives/edgar/data/2130357/000162828026068676/anthropicpbc-sx1ax1.htm — Note 8 functional table; MD&A quarterly results; Commitments", today),
            ("Anthropic", "Form S-1 (initial, Oct 26, 2026)", "H1 2025 / H1 2026 functional table", "sec.gov/Archives/edgar/data/2130357/000162828026067168/anthropicpbc-sx1.htm", today)]
    for name, cik in COMPANIES[1:]:
        p = data[name]["revenue"]
        rows.append((name, "10-K and 10-Q financial statements via SEC XBRL companyfacts API", f"Annual {', '.join(fy_label(k[1]) for k,_ in p['annual'])}; latest period to {p['latest_end']}",
                     f"data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json (latest filed value used for each period; tags listed on 'Opex Detail')", today))
    rows.append(("Amazon", "10-K FY2025 (filed Feb 6, 2026), 10-Q Q3 2026 (Oct 30, 2026), 10-Q Q3 2025", "Technology & infrastructure; Fulfillment", "Transcribed from consolidated statements of operations (custom XBRL tag not exposed by companyfacts)", today))
    rows.append(("Meta", "Q3 2026 earnings release (Oct 21, 2026)", "Q3 2026", "investor.atmeta.com — $9.89B legal charge in G&A; 2026 capex guidance $135–145B", today))
    rows.append(("Microsoft / Alphabet / Amazon", "Q3 CY2026 earnings calls & releases (Oct 28–29, 2026)", "2026 capex guidance", "CNBC coverage Oct 29, 2026 (Microsoft ≈$170B, Alphabet $195–205B, Amazon ≈$220B)", today))
    rows.append(("Tesla", "FY2025 10-K; Q2/Q3 2026 updates", "Cortex training compute; 2026 capex ≈$20B", "sec.gov 10-K (tsla-20251231.htm); Q3 2026 call (CFO Dive)", today))
    rows.append(("Third-party", "Epoch AI (company spending breakdown; R&D vs training compute), Deloitte TMT Predictions 2026, Gartner, Tom's Hardware, PPC Land, Luminix", "Estimates", "See 'Inference vs Training' Section A3 and C", today))
    for row in rows:
        for c, v in enumerate(row, start=1):
            put(ws, r, c, v, align=WRAP, font=F_NOTE if c >= 3 else F_NORM)
        ws.row_dimensions[r].height = 42; r += 1
    r += 1
    section(ws, r, "Definitions and comparability", 6); r += 1
    defs = [
        ("Cost of revenue", "As reported. Amazon = 'Cost of sales' only (fulfillment shown separately by Amazon). Anthropic = functional cost of revenue (Note 8), substantially all customer-serving compute plus some personnel/overhead."),
        ("R&D", "As reported. Amazon = 'Technology and infrastructure' (includes AWS infrastructure costs; overstates R&D). Anthropic R&D is dominated by training/research/internal-use compute."),
        ("Sales & marketing", "As reported ('Sales and marketing' / 'Marketing and sales'). Nvidia and Tesla disclose only combined SG&A, which is placed in this column with G&A left blank. Anthropic S&M includes cloud-marketplace channel-partner fees and free-tier compute."),
        ("G&A", "As reported. Meta Q3 2026 includes a $9.89B litigation charge. Anthropic G&A includes donation-matching charges and excludes the separately presented $1,500M Bartz settlement (FY2025 / 9M 2025)."),
        ("TTM", "Latest fiscal year + latest year-to-date − prior-year year-to-date. Where the latest period is a fiscal year-end (Apple), TTM = FY."),
        ("Latest quarter", "Three-month figure as filed. Where only YTD is filed (cash-flow items such as capex), the quarter is derived by differencing."),
        ("Training vs inference", "No company in this set reports a dollar split. Anthropic's functional allocation is used as a filed proxy (cost of revenue ≈ customer-serving inference ceiling; compute outside cost of revenue ≈ training/research/internal/free-tier floor). Mag 7 are characterised by total capex, guidance and management statements, plus third-party estimates."),
        ("Currency / units", "US$ millions. XBRL values are stored in dollars and rounded to the nearest million here."),
        ("Reproduce", "python anthropic-vs-mag7/build_workbook.py  (add --offline to reuse ./cache). Anthropic and Amazon T&I figures are hard-coded constants at the top of the script."),
    ]
    for a, b in defs:
        put(ws, r, 1, a, bold=True, align=WRAP); put(ws, r, 2, b, align=WRAP, font=F_NOTE)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5); ws.row_dimensions[r].height = 40; r += 1
    widths(ws, [24, 46, 40, 70, 12])
    return ws


def main():
    offline = "--offline" in sys.argv
    data = load_mag7(offline)
    wb = Workbook()
    sheet_summary(wb, data)
    sheet_detail(wb, data)
    sheet_inference_training(wb, data)
    sheet_anthropic(wb)
    sheet_sources(wb, data)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
