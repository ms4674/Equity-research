"""Archiver for Bloomberg-style screenshots posted by @modestproposal1 on X.

The package exposes three small building blocks:

* :mod:`fetcher` - thin wrapper around the X (Twitter) API v2 to pull a user's
  recent tweets and any attached photos.
* :mod:`bloomberg_filter` - heuristics that decide whether a downloaded image
  looks like a Bloomberg Terminal / Bloomberg News screenshot.
* :mod:`archiver` - orchestration glue that ties the two together and writes
  the resulting images plus a metadata index to disk.

The CLI entry point lives in :mod:`modestproposal_bloomberg.cli`.
"""

from .archiver import Archiver, ArchiveResult
from .bloomberg_filter import BloombergScreenshotClassifier, ImageFeatures
from .fetcher import TweetMedia, XClient

__all__ = [
    "Archiver",
    "ArchiveResult",
    "BloombergScreenshotClassifier",
    "ImageFeatures",
    "TweetMedia",
    "XClient",
]
