"""Quick chart helper for the revenue & token-consumption time series.

Usage:
    python plot_timeseries.py            # writes charts/*.png next to this script
    python plot_timeseries.py --show     # also opens the windows interactively

Requires: pandas, matplotlib (install via `pip install pandas matplotlib`).
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

HERE = Path(__file__).resolve().parent
DATA_DIR = HERE / "data"
CHART_DIR = HERE / "charts"


def _load_revenue() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "revenue_arr_timeseries.csv", parse_dates=["as_of_date"])
    df = df.sort_values(["company", "as_of_date"]).reset_index(drop=True)
    return df


def _load_tokens() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "monthly_token_consumption.csv", parse_dates=["as_of_date"])
    df = df.sort_values(["company", "as_of_date"]).reset_index(drop=True)
    return df


def plot_revenue(df: pd.DataFrame, out: Path) -> None:
    df = df[df["arr_usd_millions"] > 0]
    fig, ax = plt.subplots(figsize=(11, 6.5))
    for company, sub in df.groupby("company"):
        ax.plot(sub["as_of_date"], sub["arr_usd_millions"], marker="o", label=company)
    ax.set_yscale("log")
    ax.set_ylabel("Annualized run-rate revenue (USD millions, log scale)")
    ax.set_xlabel("Date")
    ax.set_title("AI revenue / ARR time series")
    ax.grid(True, which="both", alpha=0.3)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", fontsize=9)
    fig.tight_layout()
    fig.savefig(out, dpi=150)
    print(f"wrote {out}")


def plot_tokens(df: pd.DataFrame, out: Path) -> None:
    fig, ax = plt.subplots(figsize=(11, 6.5))
    plotted = df.dropna(subset=["tokens_per_month_trillions"])
    for company, sub in plotted.groupby("company"):
        ax.plot(
            sub["as_of_date"],
            sub["tokens_per_month_trillions"],
            marker="o",
            label=company,
        )
    ax.set_yscale("log")
    ax.set_ylabel("Tokens processed per month (trillions, log scale)")
    ax.set_xlabel("Date")
    ax.set_title("Monthly API token consumption")
    ax.grid(True, which="both", alpha=0.3)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    fig.tight_layout()
    fig.savefig(out, dpi=150)
    print(f"wrote {out}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--show", action="store_true")
    args = parser.parse_args()

    CHART_DIR.mkdir(exist_ok=True)
    rev = _load_revenue()
    tok = _load_tokens()
    plot_revenue(rev, CHART_DIR / "revenue_arr.png")
    plot_tokens(tok, CHART_DIR / "monthly_tokens.png")
    if args.show:
        plt.show()


if __name__ == "__main__":
    main()
