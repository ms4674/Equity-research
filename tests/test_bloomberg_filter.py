"""Tests for the Bloomberg classifier heuristics."""

from __future__ import annotations

import io

from PIL import Image, ImageDraw

from modestproposal_bloomberg.bloomberg_filter import BloombergScreenshotClassifier


def _png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _terminal_screenshot(width: int = 600, height: int = 400) -> bytes:
    img = Image.new("RGB", (width, height), color=(0, 0, 0))
    draw = ImageDraw.Draw(img)
    for y in range(0, height, 12):
        draw.line([(10, y), (width - 10, y)], fill=(255, 153, 0), width=4)
    for y in range(20, height, 80):
        draw.line([(10, y), (width - 10, y)], fill=(0, 200, 220), width=3)
    return _png_bytes(img)


def _photo() -> bytes:
    img = Image.new("RGB", (400, 300), color=(120, 180, 220))
    draw = ImageDraw.Draw(img)
    draw.rectangle((50, 50, 350, 250), fill=(220, 200, 160))
    return _png_bytes(img)


def test_terminal_screenshot_is_classified_as_bloomberg() -> None:
    classifier = BloombergScreenshotClassifier(ocr_enabled=False)
    features = classifier.classify(_terminal_screenshot())
    assert features.is_bloomberg is True
    assert features.classification == "terminal"
    assert features.dark_ratio > 0.4
    assert features.amber_ratio > 0.05


def test_random_photo_is_not_bloomberg() -> None:
    classifier = BloombergScreenshotClassifier(ocr_enabled=False)
    features = classifier.classify(_photo())
    assert features.is_bloomberg is False
    assert features.classification == "other"


def test_news_screenshot_uses_caption_text() -> None:
    classifier = BloombergScreenshotClassifier(ocr_enabled=False)
    features = classifier.classify(
        _photo(),
        caption="From Bloomberg today: oil prices spike on Hormuz news",
    )
    assert features.is_bloomberg is True
    assert features.classification == "news"
    assert any("bloomberg" in p.lower() for p in features.matched_patterns)
