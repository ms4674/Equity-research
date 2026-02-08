"""
Rich console display for terminal output.

Provides colorful, formatted terminal output using the `rich` library
for quick at-a-glance analysis without needing to open the memo file.
"""

import logging
from typing import Optional

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from .stock_universe import MarketSnapshot, StockMover

logger = logging.getLogger(__name__)
console = Console()


def display_snapshot(snapshot: MarketSnapshot):
    """Display a market snapshot in the terminal with rich formatting."""

    # Title
    console.print()
    console.print(
        Panel(
            f"[bold white]Daily Software Stock Memo[/bold white]\n"
            f"[dim]{snapshot.date}[/dim]",
            style="bold blue",
            expand=False,
        )
    )
    console.print()

    # Breadth summary
    _display_breadth(snapshot)
    console.print()

    # Top gainers
    if snapshot.top_gainers:
        _display_movers_table("Top Gainers", snapshot.top_gainers, style="green")
        console.print()

    # Top losers
    if snapshot.top_losers:
        _display_movers_table("Top Losers", snapshot.top_losers, style="red")
        console.print()

    # Sector summary
    if snapshot.sector_summaries:
        _display_sector_summary(snapshot)
        console.print()


def _display_breadth(snapshot: MarketSnapshot):
    """Display market breadth summary."""
    total = snapshot.total_gainers + snapshot.total_losers + snapshot.total_unchanged

    # Color the average change
    avg_color = "green" if snapshot.universe_avg_change >= 0 else "red"

    breadth_text = (
        f"Stocks tracked: [bold]{total}[/bold]  |  "
        f"[green]Gainers: {snapshot.total_gainers}[/green]  |  "
        f"[red]Losers: {snapshot.total_losers}[/red]  |  "
        f"Unchanged: {snapshot.total_unchanged}  |  "
        f"Avg: [{avg_color}]{snapshot.universe_avg_change:+.2f}%[/{avg_color}]  |  "
        f"Breadth: {snapshot.breadth_ratio:.0%}"
    )

    console.print(Panel(breadth_text, title="Market Breadth", style="cyan"))


def _display_movers_table(title: str, movers: list, style: str = "white"):
    """Display a table of movers."""
    table = Table(title=title, style=style, show_lines=False, padding=(0, 1))

    table.add_column("#", style="dim", width=3)
    table.add_column("Ticker", style="bold", width=8)
    table.add_column("Name", width=25)
    table.add_column("Close", justify="right", width=10)
    table.add_column("Change %", justify="right", width=10)
    table.add_column("Volume", justify="right", width=14)
    table.add_column("Vol Ratio", justify="right", width=10)
    table.add_column("Mkt Cap", justify="right", width=10)

    for i, m in enumerate(movers, 1):
        chg_color = "green" if m.change_pct >= 0 else "red"
        vol_color = "yellow" if m.volume_ratio > 1.5 else "white"

        table.add_row(
            str(i),
            m.ticker,
            m.name[:24],
            f"${m.close:.2f}",
            f"[{chg_color}]{m.change_pct:+.2f}%[/{chg_color}]",
            f"{m.volume:,.0f}",
            f"[{vol_color}]{m.volume_ratio:.1f}x[/{vol_color}]",
            m.market_cap_str,
        )

    console.print(table)


def _display_sector_summary(snapshot: MarketSnapshot):
    """Display sub-sector performance summary."""
    table = Table(title="Sub-Sector Performance", style="blue", show_lines=False)

    table.add_column("Sub-Sector", width=25)
    table.add_column("Avg Change", justify="right", width=12)
    table.add_column("Best", width=20)
    table.add_column("Worst", width=20)
    table.add_column("# Stocks", justify="right", width=8)

    for name, summary in sorted(
        snapshot.sector_summaries.items(),
        key=lambda x: x[1].avg_change_pct,
        reverse=True,
    ):
        chg_color = "green" if summary.avg_change_pct >= 0 else "red"
        table.add_row(
            summary.name,
            f"[{chg_color}]{summary.avg_change_pct:+.2f}%[/{chg_color}]",
            summary.best_performer,
            summary.worst_performer,
            str(summary.total_stocks),
        )

    console.print(table)
