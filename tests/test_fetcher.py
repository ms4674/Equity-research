"""Tests for the X API client."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import pytest

from modestproposal_bloomberg.fetcher import XApiError, XClient, _full_size, _parse_datetime


class _FakeResponse:
    def __init__(self, status_code: int, payload: Any = None, content: bytes = b"") -> None:
        self.status_code = status_code
        self._payload = payload
        self.content = content
        self.text = "" if payload is None else str(payload)
        self.headers: dict[str, str] = {}

    def json(self) -> Any:
        return self._payload


class _FakeSession:
    def __init__(self, responses: list[_FakeResponse]) -> None:
        self._responses = list(responses)
        self.calls: list[tuple[str, str, dict]] = []
        self.headers: dict[str, str] = {}

    def request(self, method: str, url: str, **kwargs: Any) -> _FakeResponse:
        self.calls.append((method, url, kwargs))
        if not self._responses:
            raise AssertionError("Unexpected extra HTTP call: " + url)
        return self._responses.pop(0)


def test_full_size_url_appends_orig() -> None:
    assert _full_size("https://pbs.twimg.com/media/abc.jpg") == (
        "https://pbs.twimg.com/media/abc.jpg?name=orig"
    )
    assert _full_size("https://pbs.twimg.com/media/abc.jpg?name=small") == (
        "https://pbs.twimg.com/media/abc.jpg?name=orig"
    )


def test_parse_datetime_handles_z_suffix() -> None:
    parsed = _parse_datetime("2026-04-26T10:00:00.000Z")
    assert parsed == datetime(2026, 4, 26, 10, 0, tzinfo=timezone.utc)


def test_resolve_user_id_returns_id() -> None:
    session = _FakeSession([_FakeResponse(200, {"data": {"id": "12345", "username": "modestproposal1"}})])
    client = XClient(bearer_token="token", session=session)  # type: ignore[arg-type]
    assert client.resolve_user_id("modestproposal1") == "12345"
    assert session.calls[0][1].endswith("/users/by/username/modestproposal1")


def test_iter_media_yields_only_photos() -> None:
    payload = {
        "data": [
            {
                "id": "1",
                "text": "Bloomberg chart attached",
                "created_at": "2026-04-25T12:00:00.000Z",
                "attachments": {"media_keys": ["m1", "m2"]},
            }
        ],
        "includes": {
            "media": [
                {
                    "media_key": "m1",
                    "type": "photo",
                    "url": "https://pbs.twimg.com/media/abc.jpg",
                    "width": 1200,
                    "height": 800,
                },
                {
                    "media_key": "m2",
                    "type": "video",
                    "preview_image_url": "https://pbs.twimg.com/media/xyz.jpg",
                },
            ]
        },
        "meta": {},
    }
    session = _FakeSession([_FakeResponse(200, payload)])
    client = XClient(bearer_token="token", session=session)  # type: ignore[arg-type]
    media = list(client.iter_media(user_id="42", max_pages=1))
    assert len(media) == 1
    assert media[0].media_key == "m1"
    assert media[0].image_url.endswith("name=orig")
    assert media[0].tweet_url == "https://x.com/i/web/status/1"


def test_unrecoverable_error_raises() -> None:
    session = _FakeSession([_FakeResponse(401, {"errors": "nope"})])
    client = XClient(bearer_token="token", session=session)  # type: ignore[arg-type]
    with pytest.raises(XApiError):
        client.resolve_user_id("modestproposal1")
