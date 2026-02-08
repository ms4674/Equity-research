"""
Stock universe management and biggest movers detection.

Loads the software stock universe from configuration,
fetches market data, and identifies the biggest daily movers.
"""

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .data_provider import DataProvider

logger = logging.getLogger(__name__)


@dataclass
class StockMover:
    """Represents a notable stock mover for the day."""

    ticker: str
    name: str
    close: float
    change: float
    change_pct: float
    volume: int
    avg_volume_10d: int
    volume_ratio: float
    market_cap: Optional[float] = None
    pe_ratio: Optional[float] = None
    sector: str = "Technology"
    sub_sector: str = ""
    fifty_two_week_high: Optional[float] = None
    fifty_two_week_low: Optional[float] = None
    beta: Optional[float] = None
    open_price: Optional[float] = None
    high_price: Optional[float] = None
    low_price: Optional[float] = None

    @property
    def direction(self) -> str:
        """Return 'up' or 'down' based on the day's change."""
        return "up" if self.change_pct >= 0 else "down"

    @property
    def is_high_volume(self) -> bool:
        """Whether volume is notably above average (>1.5x)."""
        return self.volume_ratio > 1.5

    @property
    def market_cap_str(self) -> str:
        """Human-readable market cap string."""
        if self.market_cap is None:
            return "N/A"
        if self.market_cap >= 1e12:
            return f"${self.market_cap / 1e12:.1f}T"
        if self.market_cap >= 1e9:
            return f"${self.market_cap / 1e9:.1f}B"
        if self.market_cap >= 1e6:
            return f"${self.market_cap / 1e6:.1f}M"
        return f"${self.market_cap:,.0f}"


@dataclass
class SectorSummary:
    """Summary statistics for a sub-sector."""

    name: str
    avg_change_pct: float
    median_change_pct: float
    best_performer: str
    worst_performer: str
    total_stocks: int


@dataclass
class MarketSnapshot:
    """Complete market snapshot for the software stock universe."""

    date: str
    top_gainers: List[StockMover] = field(default_factory=list)
    top_losers: List[StockMover] = field(default_factory=list)
    all_movers: List[StockMover] = field(default_factory=list)
    sector_summaries: Dict[str, SectorSummary] = field(default_factory=dict)
    universe_avg_change: float = 0.0
    universe_median_change: float = 0.0
    total_gainers: int = 0
    total_losers: int = 0
    total_unchanged: int = 0
    breadth_ratio: float = 0.0  # gainers / (gainers + losers)


def load_stock_universe(config: dict) -> Tuple[List[str], Dict[str, str]]:
    """
    Load the stock universe from configuration.

    Returns:
        tickers: flat list of unique ticker symbols
        ticker_subsector_map: mapping of ticker -> sub-sector name
    """
    universe = config.get("stock_universe", {})
    tickers = []
    ticker_subsector = {}
    seen = set()

    for sub_sector, stocks in universe.items():
        for stock in stocks:
            ticker = stock["ticker"]
            if ticker not in seen:
                tickers.append(ticker)
                seen.add(ticker)
            # Map to sub-sector (last one wins for duplicates, which is fine)
            ticker_subsector[ticker] = sub_sector.replace("_", " ").title()

    logger.info(
        "Loaded %d unique tickers across %d sub-sectors",
        len(tickers),
        len(universe),
    )
    return tickers, ticker_subsector


def build_ticker_name_map(config: dict) -> Dict[str, str]:
    """Build a ticker -> company name map from config."""
    universe = config.get("stock_universe", {})
    name_map = {}
    for _sub_sector, stocks in universe.items():
        for stock in stocks:
            name_map[stock["ticker"]] = stock["name"]
    return name_map


def fetch_and_analyze(
    provider: DataProvider, config: dict
) -> MarketSnapshot:
    """
    Fetch market data and analyze to find the biggest movers.

    Args:
        provider: market data provider instance
        config: application configuration dict

    Returns:
        MarketSnapshot with top gainers, losers, and sector summaries
    """
    from datetime import datetime

    tickers, ticker_subsector = load_stock_universe(config)
    name_map = build_ticker_name_map(config)

    # Fetch data
    df = provider.get_daily_snapshot(tickers)

    if df.empty:
        logger.error("No data returned from provider")
        return MarketSnapshot(date=datetime.now().strftime("%Y-%m-%d"))

    logger.info("Retrieved data for %d tickers", len(df))

    # Build StockMover objects
    movers = []
    for _, row in df.iterrows():
        ticker = row["ticker"]
        mover = StockMover(
            ticker=ticker,
            name=row.get("name") or name_map.get(ticker, ticker),
            close=row.get("close", 0),
            change=row.get("change", 0),
            change_pct=row.get("change_pct", 0),
            volume=int(row.get("volume", 0)),
            avg_volume_10d=int(row.get("avg_volume_10d", 0)),
            volume_ratio=row.get("volume_ratio", 1.0),
            market_cap=row.get("market_cap"),
            pe_ratio=row.get("pe_ratio"),
            sector=row.get("sector", "Technology"),
            sub_sector=ticker_subsector.get(ticker, "Software"),
            fifty_two_week_high=row.get("fifty_two_week_high"),
            fifty_two_week_low=row.get("fifty_two_week_low"),
            beta=row.get("beta"),
            open_price=row.get("open"),
            high_price=row.get("high"),
            low_price=row.get("low"),
        )
        movers.append(mover)

    # Sort by absolute percent change
    movers.sort(key=lambda m: abs(m.change_pct), reverse=True)

    memo_config = config.get("memo", {})
    top_n = memo_config.get("top_movers_count", 5)
    min_threshold = memo_config.get("min_move_threshold", 1.0)

    # Split into gainers and losers
    gainers = sorted(
        [m for m in movers if m.change_pct > 0], key=lambda m: m.change_pct, reverse=True
    )
    losers = sorted(
        [m for m in movers if m.change_pct < 0], key=lambda m: m.change_pct
    )
    unchanged = [m for m in movers if m.change_pct == 0]

    # Filter by minimum threshold for "notable" movers
    notable_gainers = [m for m in gainers if abs(m.change_pct) >= min_threshold]
    notable_losers = [m for m in losers if abs(m.change_pct) >= min_threshold]

    # Sector summaries
    sector_summaries = _build_sector_summaries(movers)

    # Compute breadth
    n_gainers = len(gainers)
    n_losers = len(losers)
    breadth = n_gainers / (n_gainers + n_losers) if (n_gainers + n_losers) > 0 else 0.5

    # Universe averages
    all_changes = [m.change_pct for m in movers]
    avg_change = sum(all_changes) / len(all_changes) if all_changes else 0
    sorted_changes = sorted(all_changes)
    median_change = (
        sorted_changes[len(sorted_changes) // 2] if sorted_changes else 0
    )

    snapshot = MarketSnapshot(
        date=datetime.now().strftime("%Y-%m-%d"),
        top_gainers=notable_gainers[:top_n],
        top_losers=notable_losers[:top_n],
        all_movers=movers,
        sector_summaries=sector_summaries,
        universe_avg_change=round(avg_change, 2),
        universe_median_change=round(median_change, 2),
        total_gainers=n_gainers,
        total_losers=n_losers,
        total_unchanged=len(unchanged),
        breadth_ratio=round(breadth, 2),
    )

    logger.info(
        "Analysis complete: %d gainers, %d losers, %d notable movers",
        n_gainers,
        n_losers,
        len(notable_gainers) + len(notable_losers),
    )

    return snapshot


def _build_sector_summaries(movers: List[StockMover]) -> Dict[str, SectorSummary]:
    """Build per-sub-sector performance summaries."""
    from collections import defaultdict

    sector_data = defaultdict(list)
    for m in movers:
        sector_data[m.sub_sector].append(m)

    summaries = {}
    for sector_name, sector_movers in sector_data.items():
        changes = [m.change_pct for m in sector_movers]
        avg_chg = sum(changes) / len(changes) if changes else 0
        sorted_chg = sorted(changes)
        median_chg = sorted_chg[len(sorted_chg) // 2] if sorted_chg else 0

        best = max(sector_movers, key=lambda m: m.change_pct)
        worst = min(sector_movers, key=lambda m: m.change_pct)

        summaries[sector_name] = SectorSummary(
            name=sector_name,
            avg_change_pct=round(avg_chg, 2),
            median_change_pct=round(median_chg, 2),
            best_performer=f"{best.ticker} ({best.change_pct:+.1f}%)",
            worst_performer=f"{worst.ticker} ({worst.change_pct:+.1f}%)",
            total_stocks=len(sector_movers),
        )

    return summaries
