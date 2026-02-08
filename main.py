#!/usr/bin/env python3
"""
Daily Software Stock Memo Generator

Entry point for the application. Connects to Bloomberg (or Yahoo Finance
as a fallback) and generates investment memos on the biggest software
stock movers of the day.

Usage:
    python main.py generate              # Generate today's memo
    python main.py generate --stdout     # Also print to console
    python main.py generate -f html      # Output as HTML
    python main.py generate -p bloomberg # Use Bloomberg data
    python main.py check                 # Check connectivity
    python main.py list-stocks           # Show stock universe
"""

from src.cli import cli

if __name__ == "__main__":
    cli()
