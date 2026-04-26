"""Heuristics that decide whether an image is a Bloomberg screenshot.

We support two kinds of "Bloomberg screenshots" because both show up in
@modestproposal1's feed:

1. **Bloomberg Terminal** screens. These have a very distinctive look:
   black background, monospaced amber/orange text, often with cyan or
   white highlights. Color histogram analysis catches them well.
2. **Bloomberg.com / Bloomberg News** article screenshots. These are
   harder to detect from pixels alone, so we rely on OCR (when
   ``pytesseract`` is available) and look for the word ``Bloomberg`` in
   the image text or in an alt-text / tweet caption supplied by the
   caller.

The classifier returns a single :class:`ImageFeatures` object so callers
can both make a decision *and* persist the supporting evidence.
"""

from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass, field
from typing import Optional

from PIL import Image

logger = logging.getLogger(__name__)

try:  # pytesseract is optional; the classifier degrades gracefully without it.
    import pytesseract  # type: ignore

    _HAS_TESSERACT = True
except Exception:  # pragma: no cover - exercised only when pytesseract missing
    pytesseract = None  # type: ignore
    _HAS_TESSERACT = False


BLOOMBERG_TEXT_PATTERNS = [
    re.compile(r"\bbloomberg\b", re.IGNORECASE),
    re.compile(r"\bbbg\b", re.IGNORECASE),
    re.compile(r"<\s*GO\s*>", re.IGNORECASE),
]


@dataclass
class ImageFeatures:
    """The signals we collected from a single image."""

    width: int
    height: int
    dark_ratio: float
    amber_ratio: float
    cyan_ratio: float
    text_excerpt: str = ""
    matched_patterns: list[str] = field(default_factory=list)
    score: float = 0.0
    is_bloomberg: bool = False
    classification: str = "unknown"

    def to_dict(self) -> dict:
        return {
            "width": self.width,
            "height": self.height,
            "dark_ratio": round(self.dark_ratio, 4),
            "amber_ratio": round(self.amber_ratio, 4),
            "cyan_ratio": round(self.cyan_ratio, 4),
            "text_excerpt": self.text_excerpt,
            "matched_patterns": self.matched_patterns,
            "score": round(self.score, 3),
            "is_bloomberg": self.is_bloomberg,
            "classification": self.classification,
        }


class BloombergScreenshotClassifier:
    """Combine pixel statistics and OCR to score Bloomberg-likeness."""

    def __init__(
        self,
        terminal_threshold: float = 0.55,
        news_threshold: float = 0.5,
        ocr_enabled: Optional[bool] = None,
        sample_size: int = 256,
    ) -> None:
        self.terminal_threshold = terminal_threshold
        self.news_threshold = news_threshold
        self.ocr_enabled = _HAS_TESSERACT if ocr_enabled is None else ocr_enabled
        self.sample_size = sample_size

    def classify(
        self,
        image_bytes: bytes,
        caption: str = "",
        alt_text: str = "",
    ) -> ImageFeatures:
        with Image.open(io.BytesIO(image_bytes)) as img:
            features = self._image_features(img)

        ocr_text = ""
        if self.ocr_enabled and pytesseract is not None:
            try:
                with Image.open(io.BytesIO(image_bytes)) as img:
                    ocr_text = pytesseract.image_to_string(img)
            except Exception as exc:  # pragma: no cover - OCR failure is non-fatal
                logger.debug("pytesseract failed: %s", exc)

        haystack = "\n".join([ocr_text, caption or "", alt_text or ""]).strip()
        features.text_excerpt = ocr_text.strip()[:500]

        for pattern in BLOOMBERG_TEXT_PATTERNS:
            if pattern.search(haystack):
                features.matched_patterns.append(pattern.pattern)

        terminal_score = _terminal_score(features)
        news_score = _news_score(features, bool(features.matched_patterns))
        features.score = max(terminal_score, news_score)

        if terminal_score >= self.terminal_threshold:
            features.classification = "terminal"
            features.is_bloomberg = True
        elif news_score >= self.news_threshold and features.matched_patterns:
            features.classification = "news"
            features.is_bloomberg = True
        else:
            features.classification = "other"
            features.is_bloomberg = False

        return features

    def _image_features(self, img: Image.Image) -> ImageFeatures:
        rgb = img.convert("RGB")
        sample = rgb.copy()
        sample.thumbnail((self.sample_size, self.sample_size))
        pixels = list(sample.getdata())
        total = max(1, len(pixels))

        dark = 0
        amber = 0
        cyan = 0
        for r, g, b in pixels:
            brightness = (r + g + b) / 3
            if brightness < 50:
                dark += 1
            if r > 180 and 80 <= g <= 200 and b < 90:
                amber += 1
            if r < 120 and g > 150 and b > 150:
                cyan += 1

        return ImageFeatures(
            width=rgb.width,
            height=rgb.height,
            dark_ratio=dark / total,
            amber_ratio=amber / total,
            cyan_ratio=cyan / total,
        )


def _terminal_score(features: ImageFeatures) -> float:
    # Terminal screenshots have lots of dark pixels and a non-trivial chunk of
    # amber text. Cyan accents push the score higher.
    if features.dark_ratio < 0.4:
        return 0.0
    score = features.dark_ratio * 0.6 + features.amber_ratio * 6.0
    score += min(features.cyan_ratio * 4.0, 0.2)
    return min(score, 1.0)


def _news_score(features: ImageFeatures, has_text_match: bool) -> float:
    if not has_text_match:
        return 0.0
    aspect = features.width / max(features.height, 1)
    score = 0.6
    if 0.5 <= aspect <= 2.5:
        score += 0.2
    if features.dark_ratio < 0.6:
        score += 0.1
    return min(score, 1.0)
