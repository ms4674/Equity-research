"""Fetch annual capital expenditure for every index constituent from Yahoo
Finance, caching one JSON file per ticker so the run is resumable.

Usage: python3 fetch_capex.py [--retry-errors]
"""

import json
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import yfinance as yf

CACHE_DIR = Path(__file__).parent / "cache"
CAPEX_TAGS = ["CapitalExpenditure", "PurchaseOfPPE", "NetPPEPurchaseAndSale"]
N_THREADS = 2

print_lock = threading.Lock()
counter = {"done": 0, "ok": 0, "err": 0}


def cache_path(symbol: str) -> Path:
    return CACHE_DIR / (symbol.replace("/", "_") + ".json")


def fetch_one(symbol: str) -> dict:
    result = {"symbol": symbol}
    t = yf.Ticker(symbol)
    cf = t.get_cashflow(freq="yearly")
    capex = {}
    capex_tag = None
    if cf is not None and not cf.empty:
        for tag in CAPEX_TAGS:
            if tag in cf.index:
                row = cf.loc[tag]
                capex = {
                    str(k.date()): float(v)
                    for k, v in row.items()
                    if pd.notna(v)
                }
                capex_tag = tag
                break
    result["capex"] = capex
    result["capex_tag"] = capex_tag
    try:
        info = t.get_info()
        result["financial_currency"] = info.get("financialCurrency")
        result["name"] = info.get("shortName") or info.get("longName")
    except Exception as exc:  # capex may still be usable without info
        result["financial_currency"] = None
        result["info_error"] = str(exc)[:200]
    if not capex:
        result["error"] = "no capex data"
    return result


def worker(symbol: str) -> None:
    path = cache_path(symbol)
    try:
        result = fetch_one(symbol)
    except Exception as exc:
        result = {"symbol": symbol, "capex": {}, "error": str(exc)[:300]}
    # Never let a failed refetch clobber previously good data.
    if path.exists():
        old = json.loads(path.read_text())
        if not result.get("capex") and old.get("capex"):
            result["capex"] = old["capex"]
            result["capex_tag"] = old.get("capex_tag")
            result.pop("error", None)
        if not result.get("financial_currency") and old.get("financial_currency"):
            result["financial_currency"] = old["financial_currency"]
            result["name"] = result.get("name") or old.get("name")
    path.write_text(json.dumps(result))
    with print_lock:
        counter["done"] += 1
        counter["ok" if not result.get("error") else "err"] += 1
        if counter["done"] % 50 == 0:
            print(f"progress: {counter['done']} done "
                  f"({counter['ok']} ok, {counter['err']} err)", flush=True)
    time.sleep(0.8)  # courtesy delay per thread


def main() -> None:
    retry_errors = "--retry-errors" in sys.argv
    CACHE_DIR.mkdir(exist_ok=True)

    base = Path(__file__).parent
    sp500 = pd.read_csv(base / "sp500_constituents.csv")
    acwi = pd.read_csv(base / "acwi_constituents.csv")
    symbols = sorted(
        set(sp500["yahoo_symbol"].dropna()) | set(acwi["yahoo_symbol"].dropna())
    )

    todo = []
    for s in symbols:
        p = cache_path(s)
        if not p.exists():
            todo.append(s)
        elif retry_errors:
            rec = json.loads(p.read_text())
            if rec.get("error") or not rec.get("financial_currency"):
                todo.append(s)
    print(f"universe: {len(symbols)} symbols; to fetch: {len(todo)}", flush=True)

    with ThreadPoolExecutor(max_workers=N_THREADS) as ex:
        futures = [ex.submit(worker, s) for s in todo]
        for f in as_completed(futures):
            f.result()

    print(f"finished: {counter['done']} fetched "
          f"({counter['ok']} ok, {counter['err']} err)", flush=True)


if __name__ == "__main__":
    main()
