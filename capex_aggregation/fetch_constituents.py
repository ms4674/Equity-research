"""Fetch index constituents:
- S&P 500: Wikipedia list of S&P 500 companies.
- MSCI ACWI (Bloomberg: MXWD): proxied by iShares MSCI ACWI ETF (ACWI) holdings.

Outputs: sp500_constituents.csv, acwi_constituents.csv (with Yahoo Finance symbols).
"""

import io
import sys

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36"}

ISHARES_ACWI_URL = (
    "https://www.ishares.com/us/products/239600/ishares-msci-acwi-etf/"
    "latest-holdings.csv"
)

# Map iShares "Exchange" names to Yahoo Finance ticker suffixes.
EXCHANGE_SUFFIX = {
    "NASDAQ": "",
    "NYSE": "",
    "New York Stock Exchange Inc.": "",
    "Nyse Mkt Llc": "",
    "Cboe BZX": "",
    "Cboe BZX formerly known as BATS": "",
    "Toronto Stock Exchange": ".TO",
    "London Stock Exchange": ".L",
    "Euronext Amsterdam": ".AS",
    "Nyse Euronext - Euronext Paris": ".PA",
    "Nyse Euronext - Euronext Brussels": ".BR",
    "Nyse Euronext - Euronext Lisbon": ".LS",
    "Euronext Milan": ".MI",
    "Borsa Italiana": ".MI",
    "Xetra": ".DE",
    "Deutsche Boerse Xetra": ".DE",
    "Boerse Berlin": ".BE",
    "Deutsche Boerse Ag": ".DE",
    "SIX Swiss Exchange": ".SW",
    "Nasdaq Omx Nordic": ".ST",
    "Stockholmsborsen": ".ST",
    "Nasdaq Omx Helsinki Ltd.": ".HE",
    "Nasdaq Omx Copenhagen A/S": ".CO",
    "Omx Nordic Exchange Copenhagen A/S": ".CO",
    "Oslo Bors Asa": ".OL",
    "Wiener Boerse Ag": ".VI",
    "Irish Stock Exchange - All Market": ".IR",
    "Bolsa De Madrid": ".MC",
    "Tokyo Stock Exchange": ".T",
    "Hong Kong Exchanges And Clearing Ltd": ".HK",
    "Hong Kong Stock Exchange": ".HK",
    "Shanghai Stock Exchange": ".SS",
    "Shenzhen Stock Exchange": ".SZ",
    "Korea Exchange (Stock Market)": ".KS",
    "Korea Exchange (Kosdaq)": ".KQ",
    "Taiwan Stock Exchange": ".TW",
    "Gretai Securities Market": ".TWO",
    "Singapore Exchange": ".SI",
    "Asx - All Markets": ".AX",
    "New Zealand Exchange Ltd": ".NZ",
    "Bolsa Mexicana De Valores (Mexican Stock Exchange)": ".MX",
    "Bolsa Mexicana De Valores": ".MX",
    "Bolsa De Valores De Sao Paulo": ".SA",
    "XBSP": ".SA",
    "Bolsa De Comercio De Santiago": ".SN",
    "Santiago Stock Exchange": ".SN",
    "Bolsa De Valores De Colombia": ".CL",
    "Bolsa De Valores De Lima": ".LM",
    "Buenos Aires Stock Exchange": ".BA",
    "Johannesburg Stock Exchange": ".JO",
    "Tel Aviv Stock Exchange": ".TA",
    "National Stock Exchange Of India": ".NS",
    "Bombay Stock Exchange Ltd": ".BO",
    "Stock Exchange Of Thailand": ".BK",
    "Bursa Malaysia": ".KL",
    "Indonesia Stock Exchange": ".JK",
    "Philippine Stock Exchange Inc.": ".PS",
    "Warsaw Stock Exchange": ".WA",
    "Warsaw Stock Exchange/Equities/Main Market": ".WA",
    "Budapest Stock Exchange": ".BD",
    "Prague Stock Exchange": ".PR",
    "Borsa Istanbul": ".IS",
    "Istanbul Stock Exchange": ".IS",
    "Athens Exchange S.A. Cash Market": ".AT",
    "Abu Dhabi Securities Exchange": ".AD",
    "Dubai Financial Market": ".DU",
    "Saudi Stock Exchange": ".SR",
    "Qatar Exchange": ".QA",
    "Boursa Kuwait": ".KW",
    "Kuwait Stock Exchange": ".KW",
    "Egyptian Exchange": ".CA",
}


def yahoo_symbol(ticker: str, exchange: str, location: str) -> str | None:
    if not isinstance(ticker, str) or not ticker or ticker == "-":
        return None
    suffix = EXCHANGE_SUFFIX.get(exchange)
    if suffix is None:
        return None
    t = ticker.strip().replace(" ", "-")
    if suffix == "":
        # US listings: Yahoo uses '-' for share classes (BRK.B -> BRK-B).
        t = t.replace(".", "-")
    elif suffix in (".T", ".HK", ".SS", ".SZ", ".KS", ".KQ", ".TW", ".TWO"):
        # Numeric tickers: Yahoo pads HK to 4 digits.
        if suffix == ".HK":
            t = t.zfill(4)
    elif suffix == ".L":
        t = t.rstrip(".").replace(".", "-")  # e.g. BT.A -> BT-A.L
    else:
        t = t.rstrip(".")
    return t + suffix


def fetch_sp500() -> pd.DataFrame:
    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    tables = pd.read_html(io.StringIO(resp.text))
    df = tables[0]
    df = df.rename(
        columns={
            "Symbol": "ticker",
            "Security": "name",
            "GICS Sector": "sector",
        }
    )[["ticker", "name", "sector"]]
    df["yahoo_symbol"] = df["ticker"].str.replace(".", "-", regex=False)
    return df


def fetch_acwi() -> pd.DataFrame:
    resp = requests.get(ISHARES_ACWI_URL, headers=HEADERS, timeout=120)
    resp.raise_for_status()
    raw = resp.text
    # File begins with fund metadata; the holdings table starts at the
    # header line beginning with "Ticker,".
    lines = raw.splitlines()
    start = next(i for i, l in enumerate(lines) if l.startswith("Ticker,"))
    body = "\n".join(lines[start:])
    df = pd.read_csv(io.StringIO(body), thousands=",")
    df = df[df["Asset Class"] == "Equity"].copy()
    df = df.rename(
        columns={
            "Ticker": "ticker",
            "Name": "name",
            "Sector": "sector",
            "Weight (%)": "weight_pct",
            "Market Value": "market_value_usd",
            "Location": "location",
            "Exchange": "exchange",
        }
    )
    df["yahoo_symbol"] = [
        yahoo_symbol(t, e, loc)
        for t, e, loc in zip(df["ticker"], df["exchange"], df["location"])
    ]
    keep = [
        "ticker",
        "name",
        "sector",
        "location",
        "exchange",
        "weight_pct",
        "market_value_usd",
        "yahoo_symbol",
    ]
    return df[keep]


def main() -> None:
    sp500 = fetch_sp500()
    sp500.to_csv("sp500_constituents.csv", index=False)
    print(f"S&P 500: {len(sp500)} constituents")

    acwi = fetch_acwi()
    acwi.to_csv("acwi_constituents.csv", index=False)
    n_mapped = acwi["yahoo_symbol"].notna().sum()
    w_mapped = acwi.loc[acwi["yahoo_symbol"].notna(), "weight_pct"].sum()
    print(f"ACWI: {len(acwi)} equity holdings; {n_mapped} mapped to Yahoo "
          f"symbols covering {w_mapped:.2f}% of weight")
    unmapped = acwi[acwi["yahoo_symbol"].isna()]
    if len(unmapped):
        print("Unmapped exchanges:")
        print(unmapped.groupby("exchange")["weight_pct"].agg(["count", "sum"])
              .sort_values("sum", ascending=False).head(20))


if __name__ == "__main__":
    main()
