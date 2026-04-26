"""Command-line entry point for the archiver."""

from __future__ import annotations

import logging
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import click
from dateutil import parser as date_parser

from .archiver import Archiver
from .bloomberg_filter import BloombergScreenshotClassifier
from .fetcher import XClient

logger = logging.getLogger("modestproposal_bloomberg")


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.option(
    "--handle",
    default=lambda: os.environ.get("TARGET_HANDLE", "modestproposal1"),
    show_default="modestproposal1",
    help="X handle to archive (without the leading @).",
)
@click.option(
    "--output-dir",
    default=lambda: os.environ.get("OUTPUT_DIR", "data"),
    type=click.Path(file_okay=False, path_type=Path),
    show_default="data",
    help="Where to write screenshots and the metadata index.",
)
@click.option(
    "--max-pages",
    default=5,
    show_default=True,
    type=click.IntRange(min=1, max=32),
    help="How many pages of recent tweets to walk through.",
)
@click.option(
    "--since",
    default=None,
    help="Only consider tweets created at or after this ISO-8601 timestamp.",
)
@click.option(
    "--save-non-matches/--skip-non-matches",
    default=False,
    help="Persist images even when they are not classified as Bloomberg.",
)
@click.option(
    "--no-ocr",
    is_flag=True,
    default=False,
    help="Disable OCR even if pytesseract is installed.",
)
@click.option(
    "--verbose",
    "-v",
    count=True,
    help="Increase logging verbosity (-v for INFO, -vv for DEBUG).",
)
def main(
    handle: str,
    output_dir: Path,
    max_pages: int,
    since: Optional[str],
    save_non_matches: bool,
    no_ocr: bool,
    verbose: int,
) -> None:
    """Archive Bloomberg-style screenshots posted by an X user."""
    logging.basicConfig(
        level=logging.WARNING - 10 * min(verbose, 2),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    _load_dotenv(Path(".env"))
    bearer = os.environ.get("X_BEARER_TOKEN")
    if not bearer:
        click.echo(
            "error: X_BEARER_TOKEN is not set. Add it to your environment or to .env "
            "(see .env.example for a template).",
            err=True,
        )
        sys.exit(2)

    start_time = None
    if since:
        try:
            start_time = date_parser.parse(since)
        except (ValueError, TypeError) as exc:
            click.echo(f"error: could not parse --since={since!r}: {exc}", err=True)
            sys.exit(2)
        if start_time.tzinfo is None:
            start_time = start_time.replace(tzinfo=timezone.utc)

    client = XClient(bearer_token=bearer)
    classifier = BloombergScreenshotClassifier(ocr_enabled=False if no_ocr else None)
    archiver = Archiver(client=client, classifier=classifier, output_dir=output_dir)

    summary = archiver.run(
        handle=handle,
        max_pages=max_pages,
        start_time=start_time,
        save_non_matches=save_non_matches,
    )

    click.echo(
        f"Scanned {summary.total_media} photo attachments from @{handle}; "
        f"saved {summary.bloomberg_media} Bloomberg screenshot(s) into "
        f"{summary.output_dir}/screenshots."
    )


if __name__ == "__main__":  # pragma: no cover
    main()
