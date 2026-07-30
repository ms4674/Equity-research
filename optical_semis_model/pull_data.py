"""Pull live market/financial data for the optical-semiconductor universe via yfinance.

Writes market_data.json keyed by ticker. Run before build_model.py.
"""

import json
import time
from datetime import date

import yfinance as yf

TICKERS = [
    # Datacom optical transceivers & components
    "COHR", "LITE", "AAOI", "FN",
    "300308.SZ", "300502.SZ", "300394.SZ", "002281.SZ",
    # Optical semiconductors (DSP / PHY / TIA / driver / PIC)
    "MRVL", "AVGO", "CRDO", "MTSI", "SMTC", "ALAB", "POET",
    # Lithography & light-source optics
    "ASML", "7731.T", "7751.T", "6925.T", "7741.T",
    # Optical inspection & metrology
    "KLAC", "6920.T", "ONTO", "CAMT", "NVMI",
    # Substrates, materials & lasers
    "AXTI", "GLW", "IPGP",
    # Optical systems & EMS
    "CIEN", "CLS",
]

INFO_FIELDS = [
    "shortName", "longName", "currency", "financialCurrency", "exchange",
    "currentPrice", "marketCap", "enterpriseValue", "sharesOutstanding",
    "totalRevenue", "revenueGrowth", "grossMargins", "ebitdaMargins",
    "operatingMargins", "profitMargins", "ebitda", "freeCashflow",
    "totalCash", "totalDebt", "trailingPE", "forwardPE",
    "enterpriseToRevenue", "enterpriseToEbitda", "priceToSalesTrailing12Months",
    "trailingEps", "forwardEps", "beta", "fiftyTwoWeekHigh", "fiftyTwoWeekLow",
    "returnOnEquity", "earningsGrowth",
]


def one_year_return(ticker: yf.Ticker):
    try:
        closes = ticker.history(period="1y", auto_adjust=True)["Close"].dropna()
        if len(closes) < 20:
            return None
        return float(closes.iloc[-1] / closes.iloc[0] - 1.0)
    except Exception:
        return None


def fx_to_usd(currency: str) -> float:
    """Rough spot FX for converting non-USD caps to USD for comparability."""
    if currency in (None, "USD"):
        return 1.0
    pair = {"JPY": "JPY=X", "CNY": "CNY=X", "EUR": "EURUSD=X", "ILS": "ILS=X"}.get(currency)
    if pair is None:
        return 1.0
    try:
        h = yf.Ticker(pair).history(period="5d")
        rate = float(h["Close"].iloc[-1])
        # JPY=X / CNY=X quote USD->local, EURUSD=X quotes EUR->USD
        return rate if pair.endswith("USD=X") else 1.0 / rate
    except Exception:
        return 1.0


def main():
    out = {"as_of": str(date.today()), "fx": {}, "tickers": {}}
    fx_cache = {}
    for sym in TICKERS:
        for attempt in range(3):
            try:
                t = yf.Ticker(sym)
                info = t.info
                row = {f: info.get(f) for f in INFO_FIELDS}
                row["oneYearReturn"] = one_year_return(t)
                ccy = row.get("currency") or "USD"
                fin_ccy = row.get("financialCurrency") or ccy
                for c in (ccy, fin_ccy):
                    if c not in fx_cache:
                        fx_cache[c] = fx_to_usd(c)
                row["fxToUsd"] = fx_cache[ccy]
                row["finFxToUsd"] = fx_cache[fin_ccy]
                out["tickers"][sym] = row
                print(f"{sym:12s} {str(row.get('shortName')):35s} px={row.get('currentPrice')} mcap={row.get('marketCap')}")
                break
            except Exception as e:
                print(f"{sym}: attempt {attempt + 1} failed: {e}")
                time.sleep(2 * (attempt + 1))
        else:
            out["tickers"][sym] = {"error": "failed"}
        time.sleep(0.4)
    out["fx"] = fx_cache
    # NaN is invalid JSON and breaks openpyxl cells downstream
    for row in out["tickers"].values():
        for k, v in row.items():
            if isinstance(v, float) and v != v:
                row[k] = None
    with open("market_data.json", "w") as f:
        json.dump(out, f, indent=2)
    n_ok = sum(1 for v in out["tickers"].values() if "error" not in v)
    print(f"\nSaved market_data.json: {n_ok}/{len(TICKERS)} tickers OK")


if __name__ == "__main__":
    main()
