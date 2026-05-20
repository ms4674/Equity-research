"""
Time series: Maximum context window (tokens) for Google, OpenAI, and Anthropic
over the past 12 months (May 2025 – May 2026).

Data sourced from official announcements and public documentation.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
from datetime import datetime

# ---------------------------------------------------------------------------
# Data: (date, max_tokens) for each provider's publicly available max context
# We track the maximum context window available via API at each point in time.
# ---------------------------------------------------------------------------

google_data = [
    # Gemini 1.5 Pro already offered 2M since May 2024; remained available
    ("2025-05-01", 2_000_000),
    ("2025-06-17", 2_000_000),  # Gemini 2.5 Pro GA (1M), but 1.5 Pro still 2M
    ("2025-09-01", 2_000_000),
    ("2025-12-01", 2_000_000),
    ("2026-02-19", 2_000_000),  # Gemini 3.1 Pro (1M); max still 2M
    ("2026-04-01", 2_000_000),  # Gemini 3.1 Ultra (2M) full rollout
    ("2026-05-20", 2_000_000),  # Gemini 3.5 Flash (1M); max still 2M
]

openai_data = [
    # o1 offered 200K; GPT-4o was 128K
    ("2025-05-01", 200_000),
    ("2025-08-01", 200_000),   # GPT-5 release (200K input)
    ("2025-12-01", 400_000),   # GPT-5.2 (~400K)
    ("2026-03-05", 1_000_000), # GPT-5.4 (1M context)
    ("2026-04-23", 1_048_576), # GPT-4.1 / GPT-5.5 (~1.05M)
    ("2026-05-20", 1_048_576),
]

anthropic_data = [
    # Claude 3.5 family: 200K context
    ("2025-05-01", 200_000),
    ("2025-08-18", 200_000),   # Context expansion confirmed at 200K
    ("2025-12-01", 200_000),   # Still 200K (beta 1M available with header)
    ("2026-03-13", 1_000_000), # Claude Opus 4.6 / Sonnet 4.6 — 1M GA
    ("2026-05-20", 1_000_000),
]

# ---------------------------------------------------------------------------
# Build DataFrames
# ---------------------------------------------------------------------------

def make_df(data, name):
    df = pd.DataFrame(data, columns=["date", "max_tokens"])
    df["date"] = pd.to_datetime(df["date"])
    df["provider"] = name
    return df

df_google = make_df(google_data, "Google (Gemini)")
df_openai = make_df(openai_data, "OpenAI (GPT/o-series)")
df_anthropic = make_df(anthropic_data, "Anthropic (Claude)")

# ---------------------------------------------------------------------------
# Plot
# ---------------------------------------------------------------------------

fig, ax = plt.subplots(figsize=(12, 6))

colors = {
    "Google (Gemini)": "#4285F4",
    "OpenAI (GPT/o-series)": "#10A37F",
    "Anthropic (Claude)": "#D97706",
}

for df in [df_google, df_openai, df_anthropic]:
    provider = df["provider"].iloc[0]
    ax.step(
        df["date"], df["max_tokens"] / 1e6,
        where="post",
        label=provider,
        color=colors[provider],
        linewidth=2.5,
    )
    ax.scatter(
        df["date"], df["max_tokens"] / 1e6,
        color=colors[provider],
        zorder=5,
        s=50,
    )

ax.set_xlabel("Date", fontsize=12)
ax.set_ylabel("Maximum Context Window (Million Tokens)", fontsize=12)
ax.set_title(
    "Context Window Growth: Google vs OpenAI vs Anthropic\n(Past 12 Months — May 2025 to May 2026)",
    fontsize=14,
    fontweight="bold",
)
ax.legend(fontsize=11, loc="center left")
ax.set_ylim(bottom=0, top=2.3)
ax.set_xlim(datetime(2025, 4, 15), datetime(2026, 6, 1))

ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
plt.xticks(rotation=45, ha="right")

ax.yaxis.set_major_locator(plt.MultipleLocator(0.25))
ax.grid(True, alpha=0.3, linestyle="--")

# Annotate key milestones
annotations = [
    ("2025-10-01", 2.0, "Google: 2M tokens (since May '24)", "Google (Gemini)", (0, 12)),
    ("2025-08-01", 0.2, "GPT-5\n200K", "OpenAI (GPT/o-series)", (0, -35)),
    ("2025-12-01", 0.4, "GPT-5.2\n400K", "OpenAI (GPT/o-series)", (15, 15)),
    ("2026-03-05", 1.0, "GPT-5.4  1M", "OpenAI (GPT/o-series)", (-30, 15)),
    ("2026-04-01", 1.0, "Claude 4.6  1M GA", "Anthropic (Claude)", (25, -30)),
]

for date_str, y_val, text, provider, offset in annotations:
    ax.annotate(
        text,
        xy=(pd.Timestamp(date_str), y_val),
        xytext=offset,
        textcoords="offset points",
        fontsize=8.5,
        ha="center",
        color=colors[provider],
        fontweight="bold",
        arrowprops=dict(arrowstyle="->", color=colors[provider], lw=0.8),
    )

plt.tight_layout()
plt.savefig("token_context_timeseries.png", dpi=150, bbox_inches="tight")
print("Chart saved to token_context_timeseries.png")

# ---------------------------------------------------------------------------
# Print summary table
# ---------------------------------------------------------------------------

print("\n" + "=" * 70)
print("SUMMARY: Maximum Context Window (Tokens) — May 2025 vs May 2026")
print("=" * 70)
print(f"{'Provider':<30} {'May 2025':>12} {'May 2026':>12} {'Increase':>12}")
print("-" * 70)

summary = [
    ("Google (Gemini)", 2_000_000, 2_000_000),
    ("OpenAI (GPT/o-series)", 200_000, 1_048_576),
    ("Anthropic (Claude)", 200_000, 1_000_000),
]

for name, start, end in summary:
    increase = f"{end/start:.1f}x"
    print(f"{name:<30} {start:>12,} {end:>12,} {increase:>12}")

print("-" * 70)
print("\nKey takeaway: Google held a massive lead at 2M tokens throughout.")
print("OpenAI grew 5.2x (200K → 1.05M). Anthropic grew 5x (200K → 1M).")
print("Both OpenAI and Anthropic are converging toward the 1M frontier,")
print("while Google has maintained 2M since May 2024 — a full year head start.")
