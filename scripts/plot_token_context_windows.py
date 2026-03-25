#!/usr/bin/env python3
"""
Token Count (Context Window) Time Series
Plots the evolution of maximum context window sizes across major AI companies.
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.dates as mdates
from matplotlib.lines import Line2D
from datetime import timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(SCRIPT_DIR, "..", "data", "token_context_windows.csv")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "charts")

COMPANY_STYLES = {
    "OpenAI": {"color": "#10A37F", "marker": "o"},
    "Anthropic": {"color": "#D97706", "marker": "s"},
    "Google": {"color": "#4285F4", "marker": "D"},
    "Meta": {"color": "#0668E1", "marker": "^"},
    "xAI": {"color": "#9333EA", "marker": "v"},
    "Cursor": {"color": "#F14E32", "marker": "P"},
}

OFFSETS_LOG = {
    "GPT-1": (8, 16),
    "GPT-2": (8, 16),
    "GPT-3": (8, -22),
    "GPT-3.5": (-40, -22),
    "GPT-4 (32K)": (35, 6),
    "GPT-3.5-turbo-16k": (-50, -22),
    "GPT-4 Turbo": (0, -24),
    "GPT-4o": (-45, 8),
    "GPT-4.1": (-40, -22),
    "GPT-5.4": (0, 14),
    "Claude 1": (0, 14),
    "Claude Instant 100K": (-20, -24),
    "Claude 2": (40, 8),
    "Claude 2.1": (0, 14),
    "Claude 3 Opus": (0, -24),
    "Claude 3.5 Sonnet": (0, 14),
    "Claude 3.5 Sonnet v2": (0, -24),
    "Claude 4 Sonnet": (-40, -24),
    "Claude Opus 4.6": (0, 14),
    "Gemini 1.0 Pro": (0, 14),
    "Gemini 1.5 Pro": (0, 14),
    "Gemini 1.5 Pro (2M)": (0, 14),
    "Gemini 2.0 Flash": (0, -24),
    "Gemini 2.5 Pro": (0, 14),
    "Gemini 3 Pro": (45, 6),
    "LLaMA 1": (0, 14),
    "LLaMA 2": (35, 6),
    "LLaMA 3": (0, 14),
    "LLaMA 3.1": (40, 6),
    "LLaMA 4 Maverick": (45, -6),
    "LLaMA 4 Scout": (0, 14),
    "Grok 1": (0, 14),
    "Grok 1.5": (40, 6),
    "Grok 2": (0, 14),
    "Grok 3": (40, -6),
    "Grok 4.1": (0, 14),
    "Grok 4.20": (40, 6),
    "cursor-small": (0, -24),
    "Composer 2": (0, -24),
}


def _token_label(tokens):
    if tokens >= 1_000_000:
        return f"{tokens / 1_000_000:.0f}M"
    if tokens >= 1_000:
        return f"{tokens / 1_000:.0f}K"
    return str(int(tokens))


def _token_fmt(x, _pos):
    return _token_label(x)


def _annotate(ax, model, date, tokens, color, offsets_map):
    text = f"{model}\n{_token_label(tokens)}"
    offset = offsets_map.get(model, (0, 14))
    ax.annotate(
        text,
        (date, tokens),
        textcoords="offset points",
        xytext=offset,
        fontsize=6,
        ha="center",
        color=color,
        fontweight="bold",
        alpha=0.85,
    )


def main():
    df = pd.read_csv(DATA_PATH, parse_dates=["release_date"])
    df = df.sort_values("release_date")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ---- Chart 1: Log-scale full history ----
    fig, ax = plt.subplots(figsize=(18, 10))
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")

    for company, style in COMPANY_STYLES.items():
        subset = df[df["company"] == company].copy()
        subset = subset.sort_values("release_date")
        subset = subset.groupby("release_date", as_index=False).agg(
            {"max_context_window_tokens": "max", "model": "last"}
        )
        ax.plot(
            subset["release_date"],
            subset["max_context_window_tokens"],
            color=style["color"],
            marker=style["marker"],
            markersize=7,
            linewidth=2.2,
            label=company,
            zorder=5,
        )
        for _, row in subset.iterrows():
            _annotate(
                ax,
                row["model"],
                row["release_date"],
                row["max_context_window_tokens"],
                style["color"],
                OFFSETS_LOG,
            )

    ax.set_yscale("log")
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(_token_fmt))
    ax.set_ylabel("Max Context Window (tokens, log scale)", fontsize=13, fontweight="bold")
    ax.set_xlabel("Release Date", fontsize=13, fontweight="bold")
    ax.set_title(
        "AI Model Context Window Size Over Time\nOpenAI  ·  Anthropic  ·  Google  ·  Meta  ·  xAI  ·  Cursor",
        fontsize=17,
        fontweight="bold",
        pad=18,
    )
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.grid(True, which="major", ls="-", alpha=0.25)
    ax.grid(True, which="minor", ls=":", alpha=0.12)

    handles = [
        Line2D([0], [0], color=s["color"], marker=s["marker"],
               markersize=8, linewidth=2.2, label=c)
        for c, s in COMPANY_STYLES.items()
    ]
    ax.legend(handles=handles, loc="upper left", fontsize=11,
              framealpha=0.9, edgecolor="#CCC")

    ax.annotate(
        "Source: Company announcements, API docs. Data as of Mar 2026.",
        xy=(0.99, 0.01), xycoords="axes fraction",
        fontsize=8, ha="right", va="bottom", color="#888",
    )

    plt.tight_layout()
    path_log = os.path.join(OUTPUT_DIR, "context_window_timeseries_log.png")
    fig.savefig(path_log, dpi=200, bbox_inches="tight")
    print(f"Saved: {path_log}")
    plt.close(fig)

    # ---- Chart 2: Linear scale, 2023–present (capped at 2M for readability) ----
    fig2, ax2 = plt.subplots(figsize=(18, 10))
    fig2.patch.set_facecolor("#FAFAFA")
    ax2.set_facecolor("#FAFAFA")

    cap = 2_500_000
    for company, style in COMPANY_STYLES.items():
        subset = df[df["company"] == company].copy()
        subset = subset[subset["release_date"] >= "2023-01-01"]
        subset = subset[subset["max_context_window_tokens"] <= cap]
        subset = subset.sort_values("release_date")
        subset = subset.groupby("release_date", as_index=False).agg(
            {"max_context_window_tokens": "max", "model": "last"}
        )
        if subset.empty:
            continue
        ax2.plot(
            subset["release_date"],
            subset["max_context_window_tokens"],
            color=style["color"],
            marker=style["marker"],
            markersize=7,
            linewidth=2.2,
            label=company,
            zorder=5,
        )
        for _, row in subset.iterrows():
            _annotate(
                ax2,
                row["model"],
                row["release_date"],
                row["max_context_window_tokens"],
                style["color"],
                OFFSETS_LOG,
            )

    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(_token_fmt))
    ax2.set_ylabel("Max Context Window (tokens)", fontsize=13, fontweight="bold")
    ax2.set_xlabel("Release Date", fontsize=13, fontweight="bold")
    ax2.set_title(
        "AI Model Context Window Size (2023–Present, Linear Scale)\n"
        "OpenAI  ·  Anthropic  ·  Google  ·  Meta  ·  xAI  ·  Cursor\n"
        "(Meta LLaMA 4 Scout 10M excluded for readability)",
        fontsize=15,
        fontweight="bold",
        pad=18,
    )
    ax2.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
    ax2.xaxis.set_minor_locator(mdates.MonthLocator())
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha="right")
    ax2.grid(True, which="major", ls="-", alpha=0.25)
    ax2.grid(True, which="minor", ls=":", alpha=0.12)

    handles2 = [
        Line2D([0], [0], color=s["color"], marker=s["marker"],
               markersize=8, linewidth=2.2, label=c)
        for c, s in COMPANY_STYLES.items()
    ]
    ax2.legend(handles=handles2, loc="upper left", fontsize=11,
               framealpha=0.9, edgecolor="#CCC")

    ax2.annotate(
        "Source: Company announcements, API docs. Data as of Mar 2026.",
        xy=(0.99, 0.01), xycoords="axes fraction",
        fontsize=8, ha="right", va="bottom", color="#888",
    )

    plt.tight_layout()
    path_lin = os.path.join(OUTPUT_DIR, "context_window_timeseries_linear.png")
    fig2.savefig(path_lin, dpi=200, bbox_inches="tight")
    print(f"Saved: {path_lin}")
    plt.close(fig2)


if __name__ == "__main__":
    main()
