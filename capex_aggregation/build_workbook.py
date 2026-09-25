"""Aggregate per-company capex into USD totals per fiscal year for the
S&P 500 and MSCI ACWI (MXWD proxy), and write an Excel workbook.

Reads: sp500_constituents.csv, acwi_constituents.csv, cache/*.json
Writes: ../Index_Capex_Aggregation.xlsx
"""

import datetime as dt
import json
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd
import xlsxwriter
import yfinance as yf

BASE = Path(__file__).parent
CACHE_DIR = BASE / "cache"
OUT_PATH = BASE.parent / "Index_Capex_Aggregation.xlsx"

FIRST_YEAR = 2022
LAST_YEAR = 2026
YEARS = list(range(FIRST_YEAR, LAST_YEAR + 1))
# Years with (near-)complete constituent coverage, shown in the chart.
CHART_YEARS = [2023, 2024, 2025]
ACWI_AS_OF = "Sep 23, 2026"


def load_cache() -> dict[str, dict]:
    out = {}
    for p in CACHE_DIR.glob("*.json"):
        rec = json.loads(p.read_text())
        out[rec["symbol"]] = rec
    return out


def capex_by_year(rec: dict) -> dict[int, float]:
    """Map fiscal-year label (calendar year of fiscal period end) -> |capex|,
    keeping the latest fiscal end if two fall in the same calendar year."""
    best_date: dict[int, str] = {}
    out: dict[int, float] = {}
    for date_str, value in rec.get("capex", {}).items():
        year = int(date_str[:4])
        if year not in YEARS or value is None:
            continue
        if year not in best_date or date_str > best_date[year]:
            best_date[year] = date_str
            out[year] = abs(float(value))
    return out


def fetch_fx_rates(currencies: set[str]) -> dict[str, dict[int, float]]:
    """Calendar-year average USD rate per unit of each currency."""
    rates: dict[str, dict[int, float]] = {"USD": {y: 1.0 for y in YEARS}}
    for cur in sorted(currencies - {"USD", None}):
        series = None
        for sym, invert in ((f"{cur}USD=X", False), (f"USD{cur}=X", True)):
            try:
                hist = yf.Ticker(sym).history(
                    start=f"{FIRST_YEAR}-01-01", interval="1mo",
                    auto_adjust=False, raise_errors=True)
                if hist is not None and not hist.empty:
                    series = hist["Close"]
                    if invert:
                        series = 1.0 / series
                    break
            except Exception:
                continue
        if series is None:
            print(f"WARNING: no FX rate found for {cur}")
            continue
        yearly = series.groupby(series.index.year).mean()
        rates[cur] = {int(y): float(v) for y, v in yearly.items() if int(y) in YEARS}
        # Pad any missing year with the closest available year.
        for y in YEARS:
            if y not in rates[cur] and rates[cur]:
                closest = min(rates[cur], key=lambda k: abs(k - y))
                rates[cur][y] = rates[cur][closest]
    return rates


def name_key(name: str | None) -> str:
    if not name:
        return ""
    return re.sub(r"[^A-Z0-9]", "", name.upper())[:3]


def build_companies(symbols: list[str], cache: dict, weights: dict[str, float],
                    ) -> list[dict]:
    """One record per company: dedupe share classes / cross-listings that
    share an identical cash-flow statement and name prefix."""
    groups: dict[tuple, list[str]] = defaultdict(list)
    for s in symbols:
        rec = cache.get(s)
        if not rec or not rec.get("capex"):
            continue
        fp = (
            rec.get("financial_currency"),
            tuple(sorted(rec["capex"].items())),
            name_key(rec.get("name")),
        )
        groups[fp].append(s)

    companies = []
    for fp, syms in groups.items():
        syms = sorted(syms, key=lambda s: -(weights.get(s) or 0.0))
        canonical = syms[0]
        rec = cache[canonical]
        companies.append({
            "symbol": canonical,
            "all_symbols": syms,
            "name": rec.get("name"),
            "currency": rec.get("financial_currency"),
            "capex_local": capex_by_year(rec),
            "weight": sum(weights.get(s) or 0.0 for s in syms),
        })
    return companies


def to_usd(companies: list[dict], fx: dict[str, dict[int, float]]) -> None:
    for c in companies:
        cur = c["currency"] or "USD"  # tiny minority missing currency
        rate = fx.get(cur)
        c["capex_usd"] = {}
        if rate is None:
            continue
        for y, v in c["capex_local"].items():
            c["capex_usd"][y] = v * rate.get(y, 1.0)


def aggregate(companies: list[dict]) -> pd.DataFrame:
    rows = []
    for y in YEARS:
        contributors = [c for c in companies if y in c.get("capex_usd", {})]
        rows.append({
            "fiscal_year": y,
            "capex_usd": sum(c["capex_usd"][y] for c in contributors),
            "n_companies": len(contributors),
            "weight_covered": sum(c["weight"] for c in contributors),
        })
    return pd.DataFrame(rows).set_index("fiscal_year")


def main() -> None:
    cache = load_cache()
    sp500 = pd.read_csv(BASE / "sp500_constituents.csv")
    acwi = pd.read_csv(BASE / "acwi_constituents.csv")

    acwi_weights = (acwi.dropna(subset=["yahoo_symbol"])
                    .groupby("yahoo_symbol")["weight_pct"].sum().to_dict())
    sp_symbols = list(sp500["yahoo_symbol"].dropna().unique())
    acwi_symbols = list(acwi["yahoo_symbol"].dropna().unique())

    sp_companies = build_companies(sp_symbols, cache, acwi_weights)
    acwi_companies = build_companies(acwi_symbols, cache, acwi_weights)

    currencies = {c["currency"] for c in acwi_companies + sp_companies
                  if c["currency"]}
    print(f"currencies to convert: {sorted(currencies)}")
    fx = fetch_fx_rates(currencies)
    for group in (sp_companies, acwi_companies):
        to_usd(group, fx)

    sp_agg = aggregate(sp_companies)
    acwi_agg = aggregate(acwi_companies)
    acwi_total_weight = float(acwi["weight_pct"].sum())
    print("\nS&P 500 aggregate capex (USD bn):")
    print((sp_agg["capex_usd"] / 1e9).round(1).to_string())
    print("\nMSCI ACWI aggregate capex (USD bn):")
    print((acwi_agg["capex_usd"] / 1e9).round(1).to_string())

    write_workbook(sp_companies, acwi_companies, sp_agg, acwi_agg,
                   sp_total=len(sp_symbols), acwi_total_weight=acwi_total_weight,
                   acwi_meta=acwi, sp_meta=sp500, fx=fx)
    print(f"\nwrote {OUT_PATH}")


# ---------------------------------------------------------------- Excel ----

def write_workbook(sp_companies, acwi_companies, sp_agg, acwi_agg,
                   sp_total, acwi_total_weight, acwi_meta, sp_meta, fx):
    wb = xlsxwriter.Workbook(str(OUT_PATH))
    f_title = wb.add_format({"bold": True, "font_size": 14})
    f_sub = wb.add_format({"font_color": "#666666", "italic": True})
    f_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                           "font_color": "white", "border": 1,
                           "text_wrap": True, "valign": "vcenter"})
    f_usd_bn = wb.add_format({"num_format": "#,##0.0", "border": 1})
    f_usd_mm = wb.add_format({"num_format": "#,##0"})
    f_int = wb.add_format({"num_format": "#,##0", "border": 1})
    f_pct = wb.add_format({"num_format": "0.0%", "border": 1})
    f_year = wb.add_format({"bold": True, "border": 1})
    f_note = wb.add_format({"font_color": "#666666"})
    f_wrap = wb.add_format({"text_wrap": True, "valign": "top"})
    f_bold = wb.add_format({"bold": True})
    f_weight = wb.add_format({"num_format": "0.000%"})
    f_fx = wb.add_format({"num_format": "0.000000"})

    # ------------------------------------------------------------ Summary
    ws = wb.add_worksheet("Summary")
    ws.write(0, 0, "Aggregate Capital Expenditure: S&P 500 vs MSCI ACWI (MXWD)",
             f_title)
    ws.write(1, 0, f"Bottom-up aggregation of constituent annual capex, "
                   f"converted to USD. Generated {dt.date.today():%b %d, %Y}.",
             f_sub)

    headers = ["Fiscal year", "S&P 500 capex (USD bn)", "S&P 500 companies",
               "S&P 500 coverage (# of 500)", "MSCI ACWI capex (USD bn)",
               "ACWI companies", "ACWI index weight covered"]
    for j, h in enumerate(headers):
        ws.write(3, j, h, f_hdr)
    r = 4
    for y in sp_agg.index:
        label = f"FY{y}" if y in CHART_YEARS else f"FY{y} (partial)"
        ws.write(r, 0, label, f_year)
        ws.write_number(r, 1, sp_agg.loc[y, "capex_usd"] / 1e9, f_usd_bn)
        ws.write_number(r, 2, int(sp_agg.loc[y, "n_companies"]), f_int)
        ws.write_number(r, 3, sp_agg.loc[y, "n_companies"] / sp_total, f_pct)
        ws.write_number(r, 4, acwi_agg.loc[y, "capex_usd"] / 1e9, f_usd_bn)
        ws.write_number(r, 5, int(acwi_agg.loc[y, "n_companies"]), f_int)
        ws.write_number(r, 6, acwi_agg.loc[y, "weight_covered"] / 100.0, f_pct)
        r += 1
    ws.set_column(0, 0, 16)
    ws.set_column(1, 6, 16)
    first_data_row = 4

    ws.write(r + 1, 0, "Fiscal year = calendar year in which the company's "
             "fiscal year ends. '(partial)' years are incomplete: Yahoo only "
             "provides ~4 recent fiscal years, so FY2022 misses many "
             "non-calendar fiscal-year companies, and FY2026 only includes "
             "fiscal years ending in 2026 that have been reported so far. "
             "FY2023-FY2025 are the comparable, near-complete years.", f_note)
    ws.write(r + 2, 0, "MSCI ACWI is proxied by iShares MSCI ACWI ETF "
             f"holdings as of {ACWI_AS_OF}. See Methodology sheet.", f_note)

    chart = wb.add_chart({"type": "column"})
    years_list = list(sp_agg.index)
    chart_first = first_data_row + years_list.index(CHART_YEARS[0])
    chart_last = first_data_row + years_list.index(CHART_YEARS[-1])
    chart.add_series({
        "name": "S&P 500",
        "categories": ["Summary", chart_first, 0, chart_last, 0],
        "values": ["Summary", chart_first, 1, chart_last, 1],
        "fill": {"color": "#1F4E78"},
    })
    chart.add_series({
        "name": "MSCI ACWI",
        "categories": ["Summary", chart_first, 0, chart_last, 0],
        "values": ["Summary", chart_first, 4, chart_last, 4],
        "fill": {"color": "#C55A11"},
    })
    chart.set_title({"name": "Aggregate capex, complete fiscal years (USD bn)"})
    chart.set_size({"width": 640, "height": 360})
    chart.set_legend({"position": "bottom"})
    ws.insert_chart(r + 4, 0, chart)

    # ------------------------------------------------------- Detail sheets
    sp_sector = dict(zip(sp_meta["yahoo_symbol"], sp_meta["sector"]))
    sp_name = dict(zip(sp_meta["yahoo_symbol"], sp_meta["name"]))
    acwi_first = acwi_meta.dropna(subset=["yahoo_symbol"]).drop_duplicates(
        "yahoo_symbol").set_index("yahoo_symbol")

    def detail_sheet(title, companies, extra_cols):
        ws = wb.add_worksheet(title)
        cols = (["Ticker(s)", "Company", *extra_cols, "Reporting currency"]
                + [f"FY{y} capex (USD mm)" for y in YEARS])
        for j, h in enumerate(cols):
            ws.write(0, j, h, f_hdr)
        companies = sorted(
            companies,
            key=lambda c: -max(c.get("capex_usd", {}).values(), default=0))
        for i, c in enumerate(companies, start=1):
            ws.write(i, 0, ", ".join(c["all_symbols"]))
            base_col = 2
            if title.startswith("S&P"):
                ws.write(i, 1, sp_name.get(c["symbol"]) or c["name"] or "")
                ws.write(i, 2, sp_sector.get(c["symbol"], ""))
                base_col = 3
            else:
                row_meta = (acwi_first.loc[c["symbol"]]
                            if c["symbol"] in acwi_first.index else None)
                ws.write(i, 1, c["name"] or (row_meta["name"] if row_meta is not None else ""))
                ws.write(i, 2, row_meta["sector"] if row_meta is not None else "")
                ws.write(i, 3, row_meta["location"] if row_meta is not None else "")
                ws.write_number(i, 4, c["weight"] / 100.0, f_weight)
                base_col = 5
            ws.write(i, base_col, c["currency"] or "n/a")
            for k, y in enumerate(YEARS):
                v = c.get("capex_usd", {}).get(y)
                if v is not None:
                    ws.write_number(i, base_col + 1 + k, v / 1e6, f_usd_mm)
        ws.set_column(0, 0, 16)
        ws.set_column(1, 1, 34)
        ws.set_column(2, base_col, 20)
        ws.set_column(base_col + 1, base_col + len(YEARS), 16)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(companies), len(cols) - 1)

    detail_sheet("S&P 500 Detail", sp_companies, ["GICS sector"])
    detail_sheet("MSCI ACWI Detail", acwi_companies,
                 ["Sector", "Country", "Index weight"])

    # ------------------------------------------------------- FX rates sheet
    ws = wb.add_worksheet("FX Rates")
    ws.write(0, 0, "Calendar-year average USD per unit of currency "
                   "(Yahoo Finance monthly closes)", f_bold)
    ws.write(2, 0, "Currency", f_hdr)
    for j, y in enumerate(YEARS):
        ws.write(2, 1 + j, y, f_hdr)
    for i, cur in enumerate(sorted(fx), start=3):
        ws.write(i, 0, cur)
        for j, y in enumerate(YEARS):
            if y in fx[cur]:
                ws.write_number(i, 1 + j, fx[cur][y], f_fx)
    ws.set_column(0, len(YEARS) + 1, 12)

    # ------------------------------------------------------- Methodology
    ws = wb.add_worksheet("Methodology")
    ws.set_column(0, 0, 110)
    notes = [
        ("Objective", f_bold),
        ("Aggregate annual capital expenditure of the constituents of the "
         "S&P 500 index and the MSCI ACWI index (Bloomberg ticker: MXWD), "
         "in US dollars, by fiscal year.", f_wrap),
        ("", None),
        ("Sources", f_bold),
        ("- S&P 500 constituents: Wikipedia 'List of S&P 500 companies' "
         "(503 share-class lines).", f_wrap),
        (f"- MSCI ACWI constituents: iShares MSCI ACWI ETF (ACWI) published "
         f"holdings as of {ACWI_AS_OF}, used as a proxy for the index. The "
         f"ETF holds ~2,200 of the index's ~2,500 names, covering nearly all "
         f"of its market capitalisation.", f_wrap),
        ("- Capital expenditure: Yahoo Finance annual cash-flow statements "
         "('Capital Expenditure' line), typically the last 4-5 fiscal years. "
         "Values are cash outflows for purchases of property, plant, "
         "equipment and similar; shown as positive numbers.", f_wrap),
        ("- FX: Yahoo Finance monthly FX closes, averaged per calendar year, "
         "applied to each company's reporting currency (see FX Rates sheet).",
         f_wrap),
        ("", None),
        ("Method", f_bold),
        ("- Fiscal year label = calendar year in which the company's fiscal "
         "year ends (e.g. Apple FY ending Sep 2025 counts in FY2025; Toyota "
         "FY ending Mar 2026 counts in FY2026).", f_wrap),
        ("- Multiple share classes / cross-listings of the same company "
         "(e.g. Alphabet A/C) are de-duplicated so each company counts once.",
         f_wrap),
        ("- Each company's capex is converted from its reporting currency to "
         "USD at that fiscal year's calendar-year average rate.", f_wrap),
        ("", None),
        ("Caveats", f_bold),
        ("- Constituent lists are as of the dates above; historical years use "
         "today's membership (no back-filling of historical index changes).",
         f_wrap),
        ("- Yahoo only provides ~4 recent fiscal years of statements, so "
         "FY2022 misses many companies with non-calendar fiscal years; "
         "FY2026 is incomplete because most fiscal years ending in 2026 have "
         "not yet been reported. FY2023-FY2025 are the comparable years.",
         f_wrap),
        ("- Banks, insurers and REITs frequently have no capital-expenditure "
         "line in Yahoo's normalised cash-flow statement (their spending is "
         "classified differently); such companies are excluded from the "
         "aggregate. This is the main reason S&P 500 company counts are "
         "~475 rather than 500.", f_wrap),
        ("- Financial companies often report little or no capex; companies "
         "with no reported capex line are excluded (see coverage columns).",
         f_wrap),
        ("- This is a bottom-up estimate from public data, not an official "
         "S&P / MSCI aggregate.", f_wrap),
    ]
    for i, (text, fmt) in enumerate(notes, start=0):
        if fmt:
            ws.write(i, 0, text, fmt)
        else:
            ws.write(i, 0, text)

    wb.close()


if __name__ == "__main__":
    main()
