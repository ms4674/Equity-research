"""
Command-line interface for the Daily Software Stock Memo application.
"""

import logging
import sys

import click
import yaml

from .data_provider import create_provider
from .memo_generator import generate_memo, save_memo
from .stock_universe import fetch_and_analyze

logger = logging.getLogger(__name__)


def setup_logging(level: str = "INFO"):
    """Configure logging for the application."""
    log_level = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Suppress noisy third-party loggers
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("yfinance").setLevel(logging.WARNING)


def load_config(config_path: str) -> dict:
    """Load application configuration from YAML file."""
    try:
        with open(config_path, "r") as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        logger.warning("Config file not found: %s — using defaults", config_path)
        return {}
    except yaml.YAMLError as e:
        logger.error("Failed to parse config file: %s", e)
        sys.exit(1)


@click.group()
@click.option(
    "--config",
    "-c",
    default="config.yaml",
    help="Path to configuration file.",
    type=click.Path(),
)
@click.option(
    "--verbose", "-v", is_flag=True, default=False, help="Enable verbose (DEBUG) logging."
)
@click.pass_context
def cli(ctx, config, verbose):
    """Daily Software Stock Memo Generator.

    Connects to market data sources (Bloomberg or Yahoo Finance) and
    generates investment memos focused on the biggest movers in
    the software sector.
    """
    ctx.ensure_object(dict)
    cfg = load_config(config)
    log_level = "DEBUG" if verbose else cfg.get("logging", {}).get("level", "INFO")
    setup_logging(log_level)
    ctx.obj["config"] = cfg


@cli.command()
@click.option(
    "--format",
    "-f",
    "output_format",
    type=click.Choice(["markdown", "html", "text"]),
    default=None,
    help="Override output format from config.",
)
@click.option(
    "--provider",
    "-p",
    type=click.Choice(["bloomberg", "yahoo"]),
    default=None,
    help="Override data provider from config.",
)
@click.option(
    "--top-n",
    "-n",
    type=int,
    default=None,
    help="Override number of top movers to feature.",
)
@click.option(
    "--stdout",
    "print_stdout",
    is_flag=True,
    default=False,
    help="Print memo to stdout instead of (in addition to) saving to file.",
)
@click.pass_context
def generate(ctx, output_format, provider, top_n, print_stdout):
    """Generate today's daily software stock memo."""
    config = ctx.obj["config"]

    # Apply CLI overrides
    if output_format:
        config.setdefault("memo", {})["output_format"] = output_format
    if provider:
        config.setdefault("data_source", {})["provider"] = provider
    if top_n:
        config.setdefault("memo", {})["top_movers_count"] = top_n

    click.echo("=" * 60)
    click.echo("  Daily Software Stock Memo Generator")
    click.echo("=" * 60)
    click.echo()

    # Step 1: Connect to data provider
    provider_name = config.get("data_source", {}).get("provider", "yahoo")
    click.echo(f"[1/3] Connecting to data provider ({provider_name})...")
    try:
        data_provider = create_provider(config)
    except Exception as e:
        click.echo(f"ERROR: Failed to initialize data provider: {e}", err=True)
        sys.exit(1)

    if not data_provider.is_connected():
        click.echo("ERROR: Data provider is not connected.", err=True)
        sys.exit(1)
    click.echo(f"      Connected successfully.")
    click.echo()

    # Step 2: Fetch and analyze data
    click.echo("[2/3] Fetching market data and analyzing movers...")
    try:
        snapshot = fetch_and_analyze(data_provider, config)
    except Exception as e:
        logger.exception("Failed to fetch/analyze data")
        click.echo(f"ERROR: {e}", err=True)
        sys.exit(1)

    total = snapshot.total_gainers + snapshot.total_losers + snapshot.total_unchanged
    click.echo(f"      Retrieved data for {total} stocks.")
    click.echo(
        f"      Found {len(snapshot.top_gainers)} notable gainers, "
        f"{len(snapshot.top_losers)} notable losers."
    )
    click.echo()

    # Step 3: Generate memo
    click.echo("[3/3] Generating investment memo...")
    memo_content = generate_memo(snapshot, config)

    filepath = save_memo(memo_content, config)
    click.echo(f"      Memo saved to: {filepath}")
    click.echo()

    if print_stdout:
        click.echo("=" * 60)
        click.echo(memo_content)

    # Display rich console summary
    try:
        from .console_display import display_snapshot
        display_snapshot(snapshot)
    except ImportError:
        pass

    click.echo("Done.")


@cli.command()
@click.pass_context
def check(ctx):
    """Check data provider connectivity and stock universe configuration."""
    config = ctx.obj["config"]

    click.echo("Checking configuration...")
    click.echo()

    # Check stock universe
    from .stock_universe import load_stock_universe

    tickers, subsectors = load_stock_universe(config)
    click.echo(f"Stock universe: {len(tickers)} tickers across {len(set(subsectors.values()))} sub-sectors")
    click.echo(f"Sub-sectors: {', '.join(sorted(set(subsectors.values())))}")
    click.echo()

    # Check data provider
    provider_name = config.get("data_source", {}).get("provider", "yahoo")
    click.echo(f"Data provider: {provider_name}")
    try:
        provider = create_provider(config)
        status = "CONNECTED" if provider.is_connected() else "DISCONNECTED"
        click.echo(f"Status: {status}")
    except Exception as e:
        click.echo(f"Status: ERROR - {e}")

    click.echo()
    click.echo(f"Output format: {config.get('memo', {}).get('output_format', 'markdown')}")
    click.echo(f"Output directory: {config.get('memo', {}).get('output_dir', 'output')}")
    click.echo(f"Top movers count: {config.get('memo', {}).get('top_movers_count', 5)}")


@cli.command(name="list-stocks")
@click.pass_context
def list_stocks(ctx):
    """List all stocks in the configured universe."""
    config = ctx.obj["config"]
    universe = config.get("stock_universe", {})

    for sub_sector, stocks in universe.items():
        sector_name = sub_sector.replace("_", " ").title()
        click.echo(f"\n{sector_name}:")
        click.echo("-" * 40)
        for stock in stocks:
            click.echo(f"  {stock['ticker']:8s} {stock['name']}")

    from .stock_universe import load_stock_universe
    tickers, _ = load_stock_universe(config)
    click.echo(f"\nTotal unique tickers: {len(tickers)}")
