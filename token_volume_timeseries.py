"""
Time series: Monthly API Token Volume (Trillions → Quadrillions)
Google vs OpenAI vs Anthropic — May 2025 to May 2026

Shows each provider's total tokens processed per month, separately.
Data from official announcements, earnings calls, and analyst estimates.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import pandas as pd
import numpy as np
from datetime import datetime

# ---------------------------------------------------------------------------
# Data: Monthly tokens processed (in trillions) for each provider
# Sources noted inline
# ---------------------------------------------------------------------------

# Google (Gemini ecosystem — all products including Search, API, consumer)
# Official: 480T (May '25), 980T (Jul '25), 1300T (Oct '25), 3200T (May '26)
google_data = [
    ("2025-05-01", 480),    # I/O 2025 keynote: Sundar Pichai
    ("2025-07-01", 980),    # Q2 update
    ("2025-10-01", 1300),   # The Decoder / analyst reports
    ("2026-01-01", 1800),   # Interpolated (growth decelerated per analysts)
    ("2026-03-01", 2400),   # Interpolated
    ("2026-05-19", 3200),   # I/O 2026 keynote: Sundar Pichai (3.2 quadrillion)
]

# OpenAI (API + ChatGPT combined platform)
# Derived: 6B/min (Oct '25) = ~259T/mo; 15B/min (Mar '26) = ~648T/mo
# Earlier: ~1.6T/day late 2024 = ~48T/month
openai_data = [
    ("2025-05-01", 100),    # Est. from growth curve (~2-3T/day)
    ("2025-08-01", 160),    # GPT-5 launch surge
    ("2025-10-01", 259),    # 6B tokens/min (Barclays/official)
    ("2025-12-01", 350),    # GPT-5.2 + enterprise growth
    ("2026-03-01", 648),    # 15B tokens/min (official)
    ("2026-05-20", 800),    # Est. continued growth trajectory
]

# Anthropic (Claude API + consumer)
# Estimated from: 87K orgs × 3.4B tokens/org/month (Q1 2026) = ~296T/month
# Q1 2025: ~50K orgs × 1.2B tokens/org = ~60T/month
# Revenue proxy: $9B→$30B ARR (3.3x in one quarter, Q1 2026)
anthropic_data = [
    ("2025-05-01", 60),     # Est. ~50K orgs × 1.2B tokens/org
    ("2025-08-01", 80),     # Steady growth pre-Claude 4
    ("2025-10-01", 100),    # Growing enterprise adoption
    ("2025-12-01", 130),    # End 2025 ($9B ARR)
    ("2026-03-01", 296),    # 87K orgs × 3.4B tokens/org (Q1 2026)
    ("2026-05-20", 400),    # Est. ($30B+ ARR, continued growth)
]

# ---------------------------------------------------------------------------
# Build DataFrames
# ---------------------------------------------------------------------------

def make_df(data, name):
    df = pd.DataFrame(data, columns=["date", "tokens_T"])
    df["date"] = pd.to_datetime(df["date"])
    df["provider"] = name
    return df

df_google = make_df(google_data, "Google")
df_openai = make_df(openai_data, "OpenAI")
df_anthropic = make_df(anthropic_data, "Anthropic")

# ---------------------------------------------------------------------------
# Plot — Separate panels (subplots) for each provider
# ---------------------------------------------------------------------------

fig, axes = plt.subplots(3, 1, figsize=(13, 11), sharex=True)

providers = [
    (df_google, "Google (Gemini)", "#4285F4", "480T → 3.2 Quadrillion (+567%)"),
    (df_openai, "OpenAI (GPT/o-series)", "#10A37F", "~100T → ~800T (+700%)"),
    (df_anthropic, "Anthropic (Claude)", "#D97706", "~60T → ~400T (+567%)"),
]

for ax, (df, label, color, subtitle) in zip(axes, providers):
    ax.fill_between(df["date"], 0, df["tokens_T"], alpha=0.15, color=color)
    ax.plot(df["date"], df["tokens_T"], color=color, linewidth=2.5, marker="o", markersize=7)
    ax.set_ylabel("Tokens / Month\n(Trillions)", fontsize=10)
    ax.set_title(f"{label}  —  {subtitle}", fontsize=12, fontweight="bold", color=color, loc="left")
    ax.grid(True, alpha=0.3, linestyle="--")
    ax.set_xlim(datetime(2025, 4, 15), datetime(2026, 6, 5))

    for _, row in df.iterrows():
        val = row["tokens_T"]
        if val >= 1000:
            label_text = f"{val/1000:.1f}Q"
        else:
            label_text = f"{int(val)}T"
        ax.annotate(
            label_text,
            xy=(row["date"], val),
            xytext=(0, 10),
            textcoords="offset points",
            fontsize=9,
            ha="center",
            fontweight="bold",
            color=color,
        )

# Add quadrillion threshold line to Google's chart
axes[0].axhline(y=1000, color="red", linestyle=":", alpha=0.5, linewidth=1.5)
axes[0].text(
    datetime(2025, 4, 20), 1050,
    "1 Quadrillion threshold",
    fontsize=9, color="red", alpha=0.7, fontstyle="italic"
)

axes[2].xaxis.set_major_locator(mdates.MonthLocator())
axes[2].xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
plt.xticks(rotation=45, ha="right")

fig.suptitle(
    "Monthly API Token Volume: From Trillions to Quadrillions\n(May 2025 – May 2026)",
    fontsize=14,
    fontweight="bold",
    y=0.98,
)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("token_volume_timeseries.png", dpi=150, bbox_inches="tight")
print("Chart saved to token_volume_timeseries.png")

# ---------------------------------------------------------------------------
# Also create a combined overlay chart for direct comparison
# ---------------------------------------------------------------------------

fig2, ax2 = plt.subplots(figsize=(13, 7))

for df, label, color, _ in providers:
    ax2.plot(df["date"], df["tokens_T"], color=color, linewidth=2.5, marker="o", markersize=8, label=label)
    ax2.fill_between(df["date"], 0, df["tokens_T"], alpha=0.08, color=color)

ax2.axhline(y=1000, color="red", linestyle=":", alpha=0.5, linewidth=1.5)
ax2.text(
    datetime(2026, 4, 1), 1050,
    "1 Quadrillion threshold",
    fontsize=10, color="red", alpha=0.7, fontstyle="italic"
)

ax2.set_xlabel("Date", fontsize=12)
ax2.set_ylabel("Tokens Processed Per Month (Trillions)", fontsize=12)
ax2.set_title(
    "API Token Volume: Google vs OpenAI vs Anthropic\n(Monthly, May 2025 – May 2026)",
    fontsize=14,
    fontweight="bold",
)
ax2.legend(fontsize=12, loc="upper left")
ax2.set_xlim(datetime(2025, 4, 15), datetime(2026, 6, 5))
ax2.set_ylim(bottom=0)
ax2.grid(True, alpha=0.3, linestyle="--")

ax2.xaxis.set_major_locator(mdates.MonthLocator())
ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
plt.xticks(rotation=45, ha="right")

def trillions_formatter(x, pos):
    if x >= 1000:
        return f"{x/1000:.1f}Q"
    return f"{int(x)}T"

ax2.yaxis.set_major_formatter(mticker.FuncFormatter(trillions_formatter))

plt.tight_layout()
plt.savefig("token_volume_combined.png", dpi=150, bbox_inches="tight")
print("Combined chart saved to token_volume_combined.png")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

print("\n" + "=" * 75)
print("API TOKEN VOLUME GROWTH — Monthly Tokens Processed (May 2025 → May 2026)")
print("=" * 75)
print(f"\n{'Provider':<25} {'May 2025':>12} {'May 2026':>12} {'Growth':>10} {'Note'}")
print("-" * 75)
print(f"{'Google (Gemini)':<25} {'480T':>12} {'3,200T':>12} {'6.7x':>10}  Crossed 1Q in Oct '25")
print(f"{'OpenAI (GPT/o-series)':<25} {'~100T':>12} {'~800T':>12} {'~8x':>10}  15B tokens/min by Mar '26")
print(f"{'Anthropic (Claude)':<25} {'~60T':>12} {'~400T':>12} {'~6.7x':>10}  87K orgs, 3.4B tokens/org")
print("-" * 75)
print(f"\n{'TOTAL (est.)':<25} {'~640T':>12} {'~4,400T':>12} {'~6.9x':>10}  Industry: 4.4 Quadrillion/mo")
print()
print("KEY MILESTONES:")
print("  • Google crossed 1 Quadrillion/month in Oct 2025 (first to do so)")
print("  • Google reached 3.2 Quadrillion/month by May 2026 (I/O keynote)")
print("  • OpenAI approaching 1 Quadrillion — likely crosses mid-2026")
print("  • Anthropic ~400T/month, on track for 1Q by late 2026/early 2027")
print()
print("CAVEATS:")
print("  • Google figure includes ALL token processing (Search, consumer, API)")
print("  • OpenAI figure from API rate (15B/min) includes ChatGPT + API")
print("  • Anthropic estimated from org count × avg tokens/org (official Q1 2026)")
print("  • Reasoning models inflate token counts ~10-17x vs prior gen models")
