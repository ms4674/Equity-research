"""High-level orchestration: fetch tweets, classify images, persist results."""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, List, Optional

from .bloomberg_filter import BloombergScreenshotClassifier, ImageFeatures
from .fetcher import TweetMedia, XClient, collect_media

logger = logging.getLogger(__name__)


@dataclass
class ArchiveResult:
    media: TweetMedia
    features: ImageFeatures
    saved_path: Optional[Path] = None

    def to_dict(self) -> dict:
        media = asdict(self.media)
        media["created_at"] = self.media.created_at.astimezone(timezone.utc).isoformat()
        return {
            "tweet": media,
            "features": self.features.to_dict(),
            "saved_path": str(self.saved_path) if self.saved_path else None,
        }


@dataclass
class ArchiveSummary:
    handle: str
    started_at: datetime
    finished_at: datetime
    total_media: int
    bloomberg_media: int
    output_dir: Path
    results: List[ArchiveResult] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "handle": self.handle,
            "started_at": self.started_at.isoformat(),
            "finished_at": self.finished_at.isoformat(),
            "total_media": self.total_media,
            "bloomberg_media": self.bloomberg_media,
            "output_dir": str(self.output_dir),
            "results": [r.to_dict() for r in self.results],
        }


class Archiver:
    def __init__(
        self,
        client: XClient,
        classifier: Optional[BloombergScreenshotClassifier] = None,
        output_dir: Path | str = "data",
    ) -> None:
        self.client = client
        self.classifier = classifier or BloombergScreenshotClassifier()
        self.output_dir = Path(output_dir)
        self.screenshots_dir = self.output_dir / "screenshots"
        self.metadata_path = self.output_dir / "index.jsonl"
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)

    def run(
        self,
        handle: str = "modestproposal1",
        max_pages: int = 5,
        start_time: Optional[datetime] = None,
        save_non_matches: bool = False,
    ) -> ArchiveSummary:
        started = datetime.now(tz=timezone.utc)
        seen_keys = self._load_seen_keys()
        results: List[ArchiveResult] = []
        bloomberg_count = 0
        total = 0

        media_iter: Iterable[TweetMedia] = collect_media(
            self.client, handle=handle, max_pages=max_pages, start_time=start_time
        )

        with self.metadata_path.open("a", encoding="utf-8") as index:
            for media in media_iter:
                total += 1
                if media.media_key in seen_keys:
                    logger.debug("Skipping already-seen media %s", media.media_key)
                    continue
                try:
                    image_bytes = self.client.download(media)
                except Exception as exc:
                    logger.warning("Failed to download %s: %s", media.image_url, exc)
                    continue

                features = self.classifier.classify(
                    image_bytes,
                    caption=media.tweet_text,
                    alt_text=media.alt_text or "",
                )
                saved_path: Optional[Path] = None
                if features.is_bloomberg or save_non_matches:
                    saved_path = self.screenshots_dir / media.safe_filename
                    saved_path.write_bytes(image_bytes)

                result = ArchiveResult(media=media, features=features, saved_path=saved_path)
                results.append(result)
                if features.is_bloomberg:
                    bloomberg_count += 1
                    logger.info(
                        "Archived %s screenshot from tweet %s (score=%.2f)",
                        features.classification,
                        media.tweet_id,
                        features.score,
                    )

                index.write(json.dumps(result.to_dict(), ensure_ascii=False) + "\n")
                index.flush()
                seen_keys.add(media.media_key)

        finished = datetime.now(tz=timezone.utc)
        summary = ArchiveSummary(
            handle=handle,
            started_at=started,
            finished_at=finished,
            total_media=total,
            bloomberg_media=bloomberg_count,
            output_dir=self.output_dir,
            results=results,
        )
        self._write_summary(summary)
        return summary

    def _load_seen_keys(self) -> set[str]:
        if not self.metadata_path.exists():
            return set()
        seen: set[str] = set()
        with self.metadata_path.open("r", encoding="utf-8") as index:
            for line in index:
                line = line.strip()
                if not line:
                    continue
                try:
                    record = json.loads(line)
                except json.JSONDecodeError:
                    continue
                media_key = (record.get("tweet") or {}).get("media_key")
                if media_key:
                    seen.add(media_key)
        return seen

    def _write_summary(self, summary: ArchiveSummary) -> None:
        path = self.output_dir / "last_run.json"
        path.write_text(
            json.dumps(summary.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
