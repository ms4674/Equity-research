"""
Investment memo generator.

Produces formatted daily investment memos for software stock movers
in Markdown, HTML, or plain text format.
"""

import logging
import os
from datetime import datetime
from typing import Optional

from .stock_universe import MarketSnapshot, StockMover

logger = logging.getLogger(__name__)


def generate_memo(snapshot: MarketSnapshot, config: dict) -> str:
    """
    Generate a formatted investment memo from a market snapshot.

    Args:
        snapshot: MarketSnapshot with analysis results
        config: application configuration

    Returns:
        Formatted memo string
    """
    memo_config = config.get("memo", {})
    output_format = memo_config.get("output_format", "markdown").lower()

    if output_format == "markdown":
        return _generate_markdown_memo(snapshot, config)
    elif output_format == "html":
        md = _generate_markdown_memo(snapshot, config)
        return _markdown_to_html(md)
    else:
        return _generate_text_memo(snapshot, config)


def save_memo(memo_content: str, config: dict, format_ext: Optional[str] = None) -> str:
    """
    Save the memo to the output directory.

    Returns the path to the saved file.
    """
    memo_config = config.get("memo", {})
    output_dir = memo_config.get("output_dir", "output")
    output_format = format_ext or memo_config.get("output_format", "markdown")

    ext_map = {"markdown": "md", "html": "html", "text": "txt", "md": "md"}
    ext = ext_map.get(output_format, "md")

    os.makedirs(output_dir, exist_ok=True)

    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"software_stock_memo_{date_str}.{ext}"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(memo_content)

    logger.info("Memo saved to %s", filepath)
    return filepath


def _generate_markdown_memo(snapshot: MarketSnapshot, config: dict) -> str:
    """Generate a Markdown-formatted investment memo."""
    memo_config = config.get("memo", {})
    lines = []

    # ── Header ──────────────────────────────────────────────────────
    lines.append(f"# Daily Software Stock Memo")
    lines.append(f"**Date:** {snapshot.date}")
    lines.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ── Executive Summary ───────────────────────────────────────────
    lines.append("## Executive Summary")
    lines.append("")
    lines.append(_build_executive_summary(snapshot))
    lines.append("")

    # ── Market Breadth ──────────────────────────────────────────────
    lines.append("## Market Breadth")
    lines.append("")
    total = snapshot.total_gainers + snapshot.total_losers + snapshot.total_unchanged
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Total Stocks Tracked | {total} |")
    lines.append(f"| Gainers | {snapshot.total_gainers} |")
    lines.append(f"| Losers | {snapshot.total_losers} |")
    lines.append(f"| Unchanged | {snapshot.total_unchanged} |")
    lines.append(f"| Breadth Ratio (Gainers/Total) | {snapshot.breadth_ratio:.0%} |")
    lines.append(f"| Universe Avg Change | {snapshot.universe_avg_change:+.2f}% |")
    lines.append(f"| Universe Median Change | {snapshot.universe_median_change:+.2f}% |")
    lines.append("")

    # Breadth commentary
    if snapshot.breadth_ratio > 0.65:
        lines.append("> **Bullish breadth** — the majority of software stocks are participating in the rally.")
    elif snapshot.breadth_ratio < 0.35:
        lines.append("> **Bearish breadth** — widespread selling pressure across the software sector.")
    else:
        lines.append("> **Mixed breadth** — the software sector shows no clear directional consensus today.")
    lines.append("")

    # ── Top Gainers ─────────────────────────────────────────────────
    lines.append("---")
    lines.append("")
    lines.append("## Top Gainers")
    lines.append("")

    if snapshot.top_gainers:
        for i, mover in enumerate(snapshot.top_gainers, 1):
            lines.append(f"### {i}. {mover.name} ({mover.ticker}) — {mover.change_pct:+.2f}%")
            lines.append("")
            lines.append(_build_mover_detail(mover))
            lines.append("")
    else:
        lines.append("*No notable gainers today (above minimum threshold).*")
        lines.append("")

    # ── Top Losers ──────────────────────────────────────────────────
    lines.append("---")
    lines.append("")
    lines.append("## Top Losers")
    lines.append("")

    if snapshot.top_losers:
        for i, mover in enumerate(snapshot.top_losers, 1):
            lines.append(f"### {i}. {mover.name} ({mover.ticker}) — {mover.change_pct:+.2f}%")
            lines.append("")
            lines.append(_build_mover_detail(mover))
            lines.append("")
    else:
        lines.append("*No notable losers today (above minimum threshold).*")
        lines.append("")

    # ── Sector Performance ──────────────────────────────────────────
    if memo_config.get("include_sector_summary", True) and snapshot.sector_summaries:
        lines.append("---")
        lines.append("")
        lines.append("## Sub-Sector Performance")
        lines.append("")
        lines.append("| Sub-Sector | Avg Change | Median Change | Best | Worst | # Stocks |")
        lines.append("|------------|-----------|---------------|------|-------|----------|")

        for name, summary in sorted(
            snapshot.sector_summaries.items(),
            key=lambda x: x[1].avg_change_pct,
            reverse=True,
        ):
            lines.append(
                f"| {summary.name} | {summary.avg_change_pct:+.2f}% | "
                f"{summary.median_change_pct:+.2f}% | {summary.best_performer} | "
                f"{summary.worst_performer} | {summary.total_stocks} |"
            )
        lines.append("")

    # ── Volume Analysis ─────────────────────────────────────────────
    if memo_config.get("include_volume_analysis", True):
        lines.append("---")
        lines.append("")
        lines.append("## Volume Analysis")
        lines.append("")

        high_vol_movers = [
            m for m in snapshot.all_movers if m.is_high_volume and abs(m.change_pct) >= 1.0
        ]
        high_vol_movers.sort(key=lambda m: m.volume_ratio, reverse=True)

        if high_vol_movers:
            lines.append("Stocks with notable volume (>1.5x average) and significant price moves:")
            lines.append("")
            lines.append("| Ticker | Name | Change | Volume Ratio | Volume |")
            lines.append("|--------|------|--------|-------------|--------|")
            for m in high_vol_movers[:10]:
                lines.append(
                    f"| {m.ticker} | {m.name} | {m.change_pct:+.2f}% | "
                    f"{m.volume_ratio:.1f}x | {m.volume:,.0f} |"
                )
            lines.append("")
            lines.append(
                "> High volume moves are more likely to reflect institutional activity "
                "and carry greater significance for multi-day follow-through."
            )
        else:
            lines.append("*No stocks showed both high volume and significant price movement today.*")
        lines.append("")

    # ── Full Universe Table ─────────────────────────────────────────
    lines.append("---")
    lines.append("")
    lines.append("## Full Universe Performance")
    lines.append("")
    lines.append(
        "| Ticker | Name | Close | Change | Change % | Volume | Vol Ratio | Mkt Cap |"
    )
    lines.append(
        "|--------|------|-------|--------|----------|--------|-----------|---------|"
    )

    sorted_universe = sorted(snapshot.all_movers, key=lambda m: m.change_pct, reverse=True)
    for m in sorted_universe:
        lines.append(
            f"| {m.ticker} | {m.name} | ${m.close:.2f} | "
            f"{m.change:+.2f} | {m.change_pct:+.2f}% | "
            f"{m.volume:,.0f} | {m.volume_ratio:.1f}x | {m.market_cap_str} |"
        )
    lines.append("")

    # ── Disclaimer ──────────────────────────────────────────────────
    lines.append("---")
    lines.append("")
    lines.append("## Disclaimer")
    lines.append("")
    lines.append(
        "*This memo is generated automatically for informational purposes only. "
        "It does not constitute investment advice, a recommendation, or an offer "
        "to buy or sell any securities. Past performance is not indicative of "
        "future results. All data is sourced from public market feeds and may "
        "be subject to delays or inaccuracies. Always conduct your own research "
        "and consult a qualified financial advisor before making investment decisions.*"
    )
    lines.append("")

    return "\n".join(lines)


def _build_executive_summary(snapshot: MarketSnapshot) -> str:
    """Build the executive summary paragraph."""
    parts = []
    total = snapshot.total_gainers + snapshot.total_losers + snapshot.total_unchanged

    # Overall tone
    if snapshot.universe_avg_change > 1.0:
        tone = "strongly positive"
    elif snapshot.universe_avg_change > 0.25:
        tone = "modestly positive"
    elif snapshot.universe_avg_change > -0.25:
        tone = "mixed"
    elif snapshot.universe_avg_change > -1.0:
        tone = "modestly negative"
    else:
        tone = "strongly negative"

    parts.append(
        f"The software sector traded **{tone}** today across a universe of "
        f"**{total}** tracked stocks, with an average move of "
        f"**{snapshot.universe_avg_change:+.2f}%**."
    )

    # Top gainers mention
    if snapshot.top_gainers:
        top = snapshot.top_gainers[0]
        parts.append(
            f"**{top.name} ({top.ticker})** led the gainers, surging "
            f"**{top.change_pct:+.2f}%** to close at **${top.close:.2f}**."
        )

    # Top losers mention
    if snapshot.top_losers:
        bottom = snapshot.top_losers[0]
        parts.append(
            f"On the downside, **{bottom.name} ({bottom.ticker})** was the "
            f"biggest decliner, falling **{bottom.change_pct:+.2f}%** to "
            f"**${bottom.close:.2f}**."
        )

    # Breadth
    parts.append(
        f"Market breadth stood at **{snapshot.breadth_ratio:.0%}** "
        f"({snapshot.total_gainers} advancers vs. {snapshot.total_losers} decliners)."
    )

    # Sector highlights
    if snapshot.sector_summaries:
        best_sector = max(
            snapshot.sector_summaries.values(), key=lambda s: s.avg_change_pct
        )
        worst_sector = min(
            snapshot.sector_summaries.values(), key=lambda s: s.avg_change_pct
        )
        if best_sector.name != worst_sector.name:
            parts.append(
                f"Among sub-sectors, **{best_sector.name}** outperformed "
                f"({best_sector.avg_change_pct:+.2f}% avg) while "
                f"**{worst_sector.name}** lagged "
                f"({worst_sector.avg_change_pct:+.2f}% avg)."
            )

    return " ".join(parts)


def _build_mover_detail(mover: StockMover) -> str:
    """Build detailed analysis section for a single mover."""
    lines = []

    # Price table
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Close | ${mover.close:.2f} |")
    lines.append(f"| Change | {mover.change:+.2f} ({mover.change_pct:+.2f}%) |")

    if mover.open_price:
        lines.append(f"| Open | ${mover.open_price:.2f} |")
    if mover.high_price and mover.low_price:
        lines.append(f"| Day Range | ${mover.low_price:.2f} – ${mover.high_price:.2f} |")

    lines.append(f"| Volume | {mover.volume:,.0f} |")
    if mover.avg_volume_10d > 0:
        lines.append(f"| Avg Volume (10d) | {mover.avg_volume_10d:,.0f} |")
        lines.append(f"| Volume Ratio | {mover.volume_ratio:.1f}x |")

    lines.append(f"| Market Cap | {mover.market_cap_str} |")
    if mover.pe_ratio:
        lines.append(f"| P/E Ratio | {mover.pe_ratio:.1f}x |")
    if mover.fifty_two_week_high and mover.fifty_two_week_low:
        lines.append(
            f"| 52-Week Range | ${mover.fifty_two_week_low:.2f} – ${mover.fifty_two_week_high:.2f} |"
        )
    if mover.beta:
        lines.append(f"| Beta | {mover.beta:.2f} |")

    lines.append(f"| Sub-Sector | {mover.sub_sector} |")
    lines.append("")

    # Commentary
    commentary = _generate_mover_commentary(mover)
    if commentary:
        lines.append(f"> {commentary}")
        lines.append("")

    return "\n".join(lines)


def _generate_mover_commentary(mover: StockMover) -> str:
    """Generate brief analytical commentary for a mover."""
    parts = []

    # Volume signal
    if mover.volume_ratio > 3.0:
        parts.append(
            f"Volume is **{mover.volume_ratio:.1f}x** the 10-day average, "
            f"suggesting significant institutional interest."
        )
    elif mover.volume_ratio > 1.5:
        parts.append(
            f"Volume is **{mover.volume_ratio:.1f}x** above average, "
            f"indicating above-normal conviction in the move."
        )
    elif mover.volume_ratio < 0.5:
        parts.append(
            f"Volume is notably **below average** ({mover.volume_ratio:.1f}x), "
            f"suggesting the move may lack conviction."
        )

    # 52-week context
    if mover.fifty_two_week_high and mover.close:
        pct_from_high = (
            (mover.close - mover.fifty_two_week_high)
            / mover.fifty_two_week_high
            * 100
        )
        if abs(pct_from_high) < 3:
            parts.append(
                f"Trading near 52-week highs (**{pct_from_high:+.1f}%** from peak)."
            )
        elif pct_from_high < -30:
            parts.append(
                f"Still **{pct_from_high:.1f}%** below 52-week high, "
                f"indicating significant mean-reversion potential."
            )

    if mover.fifty_two_week_low and mover.close:
        pct_from_low = (
            (mover.close - mover.fifty_two_week_low) / mover.fifty_two_week_low * 100
        )
        if pct_from_low < 5:
            parts.append(
                f"Trading near 52-week lows (**{pct_from_low:+.1f}%** above trough), "
                f"which may attract value-oriented buyers."
            )

    # Valuation
    if mover.pe_ratio:
        if mover.pe_ratio > 80:
            parts.append(
                f"Valuation is elevated at **{mover.pe_ratio:.1f}x** earnings."
            )
        elif mover.pe_ratio < 15 and mover.pe_ratio > 0:
            parts.append(
                f"Relatively attractive valuation at **{mover.pe_ratio:.1f}x** earnings."
            )

    # Intraday range
    if mover.high_price and mover.low_price and mover.close:
        intraday_range = mover.high_price - mover.low_price
        range_pct = (intraday_range / mover.close) * 100
        if range_pct > 5:
            parts.append(
                f"Wide intraday range of **{range_pct:.1f}%**, reflecting elevated volatility."
            )

    return " ".join(parts)


def _generate_text_memo(snapshot: MarketSnapshot, config: dict) -> str:
    """Generate a plain text version of the memo (simplified)."""
    md = _generate_markdown_memo(snapshot, config)
    # Simple Markdown stripping for text output
    text = md.replace("**", "").replace("*", "").replace("#", "").replace(">", "")
    text = text.replace("---", "=" * 70)
    text = "\n".join(line for line in text.split("\n") if not line.strip().startswith("|--"))
    return text


def _markdown_to_html(md_content: str) -> str:
    """Convert Markdown memo to HTML with styling."""
    try:
        import markdown as md_lib
        html_body = md_lib.markdown(
            md_content, extensions=["tables", "fenced_code", "toc"]
        )
    except ImportError:
        # Fallback: wrap in pre tag
        html_body = f"<pre>{md_content}</pre>"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Software Stock Memo</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 960px;
            margin: 0 auto;
            padding: 2rem;
            background: #f8f9fa;
            color: #212529;
            line-height: 1.6;
        }}
        h1 {{ color: #1a1a2e; border-bottom: 3px solid #0066cc; padding-bottom: 0.5rem; }}
        h2 {{ color: #16213e; margin-top: 2rem; }}
        h3 {{ color: #0f3460; }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 1rem 0;
            font-size: 0.9rem;
        }}
        th, td {{
            border: 1px solid #dee2e6;
            padding: 0.5rem 0.75rem;
            text-align: left;
        }}
        th {{ background: #e9ecef; font-weight: 600; }}
        tr:nth-child(even) {{ background: #f8f9fa; }}
        blockquote {{
            border-left: 4px solid #0066cc;
            margin: 1rem 0;
            padding: 0.75rem 1rem;
            background: #e7f1ff;
            color: #004085;
        }}
        hr {{ border: none; border-top: 2px solid #dee2e6; margin: 2rem 0; }}
        strong {{ color: #0066cc; }}
    </style>
</head>
<body>
{html_body}
</body>
</html>"""
