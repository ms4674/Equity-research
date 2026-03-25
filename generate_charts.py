"""
Generate time series visualizations for:
  1. MCP Server Count growth
  2. LLM Token Usage (daily tokens processed, in trillions)
  3. MCP SDK Monthly Downloads (npm + PyPI)

Outputs:
  - charts/mcp_dashboard.html  (interactive Plotly dashboard)
  - charts/mcp_server_count.png
  - charts/llm_token_usage.png
  - charts/mcp_sdk_downloads.png

Data sources:
  - MCP Server Count: Glama, DreamFactory, Pulse MCP, Bloomberry, Anthropic
  - Token Usage: OpenRouter/a16z State of AI 100T token study, NavyaAI
  - SDK Downloads: npm registry (@modelcontextprotocol/sdk), PyPI (mcp package)
"""

import os
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker

os.makedirs("charts", exist_ok=True)

COLORS = {
    "primary": "#6366f1",
    "secondary": "#f59e0b",
    "accent": "#10b981",
    "npm": "#cb3837",
    "pypi": "#3775a9",
    "total": "#6366f1",
    "bg": "#fafafa",
    "grid": "#e5e7eb",
    "text": "#1f2937",
}


def load_data():
    servers = pd.read_csv("data/mcp_server_count.csv", parse_dates=["date"])
    tokens = pd.read_csv("data/llm_token_usage.csv", parse_dates=["date"])
    sdk = pd.read_csv("data/mcp_sdk_downloads.csv")
    sdk["month"] = pd.to_datetime(sdk["month"] + "-01")
    return servers, tokens, sdk


def generate_plotly_dashboard(servers, tokens, sdk):
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=(
            "<b>MCP Server Count</b> — Ecosystem Growth Since Launch",
            "<b>Daily LLM Token Processing</b> — Industry-Wide (Trillions)",
            "<b>MCP SDK Monthly Downloads</b> — npm + PyPI",
        ),
        vertical_spacing=0.10,
        row_heights=[0.30, 0.30, 0.40],
    )

    # --- Panel 1: MCP Server Count ---
    fig.add_trace(
        go.Scatter(
            x=servers["date"], y=servers["server_count"],
            mode="lines+markers",
            name="MCP Servers",
            line=dict(color=COLORS["primary"], width=3),
            marker=dict(size=8, color=COLORS["primary"]),
            fill="tozeroy",
            fillcolor="rgba(99,102,241,0.10)",
            hovertemplate="<b>%{x|%b %Y}</b><br>Servers: %{y:,.0f}<extra></extra>",
        ),
        row=1, col=1,
    )
    fig.add_annotation(
        x=servers["date"].iloc[-1], y=servers["server_count"].iloc[-1],
        text=f"<b>{servers['server_count'].iloc[-1]:,}</b>",
        showarrow=True, arrowhead=2, ax=-40, ay=-30,
        font=dict(size=13, color=COLORS["primary"]),
        row=1, col=1,
    )

    # --- Panel 2: Token Usage ---
    fig.add_trace(
        go.Scatter(
            x=tokens["date"], y=tokens["daily_tokens_trillion"],
            mode="lines+markers",
            name="Daily Tokens (T)",
            line=dict(color=COLORS["secondary"], width=3),
            marker=dict(size=7, color=COLORS["secondary"]),
            fill="tozeroy",
            fillcolor="rgba(245,158,11,0.10)",
            hovertemplate="<b>%{x|%b %Y}</b><br>%{y:.2f}T tokens/day<extra></extra>",
        ),
        row=2, col=1,
    )
    fig.add_annotation(
        x=tokens["date"].iloc[-1], y=tokens["daily_tokens_trillion"].iloc[-1],
        text=f"<b>{tokens['daily_tokens_trillion'].iloc[-1]:.1f}T/day</b>",
        showarrow=True, arrowhead=2, ax=-50, ay=-25,
        font=dict(size=13, color=COLORS["secondary"]),
        row=2, col=1,
    )

    # --- Panel 3: SDK Downloads ---
    fig.add_trace(
        go.Bar(
            x=sdk["month"], y=sdk["npm_downloads"],
            name="npm (@modelcontextprotocol/sdk)",
            marker_color=COLORS["npm"],
            hovertemplate="<b>%{x|%b %Y}</b><br>npm: %{y:,.0f}<extra></extra>",
        ),
        row=3, col=1,
    )
    fig.add_trace(
        go.Bar(
            x=sdk["month"], y=sdk["pypi_downloads"],
            name="PyPI (mcp)",
            marker_color=COLORS["pypi"],
            hovertemplate="<b>%{x|%b %Y}</b><br>PyPI: %{y:,.0f}<extra></extra>",
        ),
        row=3, col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=sdk["month"], y=sdk["total_downloads"],
            mode="lines+markers",
            name="Total Downloads",
            line=dict(color=COLORS["total"], width=2.5, dash="dot"),
            marker=dict(size=6, color=COLORS["total"]),
            hovertemplate="<b>%{x|%b %Y}</b><br>Total: %{y:,.0f}<extra></extra>",
        ),
        row=3, col=1,
    )

    fig.update_layout(
        height=1200,
        title=dict(
            text=(
                "<b>MCP Ecosystem Dashboard</b>"
                "<br><sup>Server Count · LLM Token Processing · SDK Downloads</sup>"
            ),
            font=dict(size=22, color=COLORS["text"]),
            x=0.5,
        ),
        font=dict(family="Inter, system-ui, sans-serif", color=COLORS["text"]),
        plot_bgcolor="white",
        paper_bgcolor=COLORS["bg"],
        legend=dict(
            orientation="h", yanchor="bottom", y=-0.08, xanchor="center", x=0.5,
            font=dict(size=11),
        ),
        barmode="stack",
        margin=dict(t=100, b=80, l=70, r=30),
    )

    for i in range(1, 4):
        fig.update_xaxes(
            showgrid=True, gridcolor=COLORS["grid"], gridwidth=1,
            dtick="M2", tickformat="%b\n%Y",
            row=i, col=1,
        )
        fig.update_yaxes(
            showgrid=True, gridcolor=COLORS["grid"], gridwidth=1,
            row=i, col=1,
        )

    fig.update_yaxes(title_text="Servers", row=1, col=1)
    fig.update_yaxes(title_text="Trillion tokens/day", row=2, col=1)
    fig.update_yaxes(title_text="Downloads", row=3, col=1)

    fig.write_html("charts/mcp_dashboard.html", include_plotlyjs="cdn")
    print("  -> charts/mcp_dashboard.html")


def _style_ax(ax, title, ylabel):
    ax.set_title(title, fontsize=14, fontweight="bold", pad=12, color=COLORS["text"])
    ax.set_ylabel(ylabel, fontsize=11, color=COLORS["text"])
    ax.grid(True, alpha=0.3, color=COLORS["grid"])
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(colors=COLORS["text"])
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha="center", fontsize=9)


def generate_matplotlib_charts(servers, tokens, sdk):
    # --- Chart 1: Server Count ---
    fig1, ax1 = plt.subplots(figsize=(12, 5), facecolor=COLORS["bg"])
    ax1.set_facecolor("white")
    ax1.plot(servers["date"], servers["server_count"],
             color=COLORS["primary"], linewidth=2.5, marker="o", markersize=6, zorder=3)
    ax1.fill_between(servers["date"], servers["server_count"],
                     alpha=0.08, color=COLORS["primary"])
    ax1.annotate(
        f'{servers["server_count"].iloc[-1]:,}',
        xy=(servers["date"].iloc[-1], servers["server_count"].iloc[-1]),
        xytext=(-50, 15), textcoords="offset points",
        fontsize=12, fontweight="bold", color=COLORS["primary"],
        arrowprops=dict(arrowstyle="->", color=COLORS["primary"]),
    )
    _style_ax(ax1, "MCP Server Count — Ecosystem Growth Since Launch (Nov 2024)", "Number of Servers")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    fig1.tight_layout()
    fig1.savefig("charts/mcp_server_count.png", dpi=150, bbox_inches="tight")
    plt.close(fig1)
    print("  -> charts/mcp_server_count.png")

    # --- Chart 2: Token Usage ---
    fig2, ax2 = plt.subplots(figsize=(12, 5), facecolor=COLORS["bg"])
    ax2.set_facecolor("white")
    ax2.plot(tokens["date"], tokens["daily_tokens_trillion"],
             color=COLORS["secondary"], linewidth=2.5, marker="s", markersize=5, zorder=3)
    ax2.fill_between(tokens["date"], tokens["daily_tokens_trillion"],
                     alpha=0.08, color=COLORS["secondary"])
    ax2.annotate(
        f'{tokens["daily_tokens_trillion"].iloc[-1]:.1f}T/day',
        xy=(tokens["date"].iloc[-1], tokens["daily_tokens_trillion"].iloc[-1]),
        xytext=(-60, 15), textcoords="offset points",
        fontsize=12, fontweight="bold", color=COLORS["secondary"],
        arrowprops=dict(arrowstyle="->", color=COLORS["secondary"]),
    )
    _style_ax(ax2, "Daily LLM Token Processing — Industry-Wide (Trillions)", "Trillion tokens / day")
    fig2.tight_layout()
    fig2.savefig("charts/llm_token_usage.png", dpi=150, bbox_inches="tight")
    plt.close(fig2)
    print("  -> charts/llm_token_usage.png")

    # --- Chart 3: SDK Downloads ---
    fig3, ax3 = plt.subplots(figsize=(14, 6), facecolor=COLORS["bg"])
    ax3.set_facecolor("white")
    bar_width = 20
    npm_bars = ax3.bar(sdk["month"] - pd.Timedelta(days=bar_width / 2),
                       sdk["npm_downloads"], width=bar_width,
                       color=COLORS["npm"], alpha=0.85, label="npm (@modelcontextprotocol/sdk)")
    pypi_bars = ax3.bar(sdk["month"] - pd.Timedelta(days=bar_width / 2),
                        sdk["pypi_downloads"], width=bar_width,
                        bottom=sdk["npm_downloads"],
                        color=COLORS["pypi"], alpha=0.85, label="PyPI (mcp)")
    ax3.plot(sdk["month"], sdk["total_downloads"],
             color=COLORS["total"], linewidth=2, linestyle="--",
             marker="D", markersize=5, label="Total", zorder=3)
    _style_ax(ax3, "MCP SDK Monthly Downloads — npm + PyPI", "Downloads")
    ax3.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f"{x / 1e6:.0f}M" if x >= 1e6 else f"{x / 1e3:.0f}K" if x >= 1e3 else f"{x:.0f}"
    ))
    ax3.legend(loc="upper left", fontsize=10, framealpha=0.9)
    fig3.tight_layout()
    fig3.savefig("charts/mcp_sdk_downloads.png", dpi=150, bbox_inches="tight")
    plt.close(fig3)
    print("  -> charts/mcp_sdk_downloads.png")


def main():
    print("Loading data...")
    servers, tokens, sdk = load_data()

    print(f"  MCP Servers: {len(servers)} data points ({servers['date'].min():%b %Y} – {servers['date'].max():%b %Y})")
    print(f"  Token Usage: {len(tokens)} data points ({tokens['date'].min():%b %Y} – {tokens['date'].max():%b %Y})")
    print(f"  SDK Downloads: {len(sdk)} months ({sdk['month'].min():%b %Y} – {sdk['month'].max():%b %Y})")

    print("\nGenerating Plotly interactive dashboard...")
    generate_plotly_dashboard(servers, tokens, sdk)

    print("\nGenerating static PNG charts...")
    generate_matplotlib_charts(servers, tokens, sdk)

    print("\nDone! Files written to charts/")


if __name__ == "__main__":
    main()
