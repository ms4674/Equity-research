"""Pull tweets and attached media from X's public API v2.

The module deliberately keeps the surface area small. We only need:

* the user id for a handle,
* a paginated stream of recent tweets with attached photos, and
* a helper to download the original-size photo bytes.

Authentication uses an app-only Bearer Token, which is sufficient for read
endpoints and avoids the OAuth 1.0a user-context dance.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Iterable, Iterator, Optional

import requests

logger = logging.getLogger(__name__)

X_API_BASE = "https://api.x.com/2"
DEFAULT_PAGE_SIZE = 100
MAX_RETRIES = 5


@dataclass
class TweetMedia:
    """A single image attached to a tweet."""

    tweet_id: str
    tweet_url: str
    tweet_text: str
    created_at: datetime
    media_key: str
    image_url: str
    width: Optional[int] = None
    height: Optional[int] = None
    alt_text: Optional[str] = None
    extras: dict = field(default_factory=dict)

    @property
    def safe_filename(self) -> str:
        ts = self.created_at.strftime("%Y%m%dT%H%M%SZ")
        suffix = self.image_url.rsplit(".", 1)[-1].split("?", 1)[0].lower()
        if suffix not in {"jpg", "jpeg", "png", "webp"}:
            suffix = "jpg"
        return f"{ts}_{self.tweet_id}_{self.media_key}.{suffix}"


class XApiError(RuntimeError):
    """Raised when the X API returns an unrecoverable error."""


class XClient:
    """Minimal client for the X API v2 read endpoints we care about."""

    def __init__(
        self,
        bearer_token: str,
        session: Optional[requests.Session] = None,
        timeout: float = 30.0,
    ) -> None:
        if not bearer_token:
            raise ValueError("bearer_token is required")
        self._session = session or requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {bearer_token}",
                "User-Agent": "modestproposal-bloomberg-archiver/0.1",
            }
        )
        self._timeout = timeout

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        for attempt in range(MAX_RETRIES):
            response = self._session.request(method, url, timeout=self._timeout, **kwargs)
            if response.status_code == 429:
                # Honor the rate-limit reset header when present, otherwise
                # fall back to exponential backoff (2, 4, 8, 16, 32 seconds).
                reset = response.headers.get("x-rate-limit-reset")
                if reset and reset.isdigit():
                    sleep_for = max(1, int(reset) - int(time.time()))
                else:
                    sleep_for = 2 ** (attempt + 1)
                logger.warning("Rate limited by X; sleeping %ss", sleep_for)
                time.sleep(min(sleep_for, 120))
                continue
            if 500 <= response.status_code < 600:
                sleep_for = 2 ** (attempt + 1)
                logger.warning(
                    "X API %s returned %s; retrying in %ss",
                    url,
                    response.status_code,
                    sleep_for,
                )
                time.sleep(sleep_for)
                continue
            return response
        raise XApiError(f"Exhausted retries calling {url}")

    def resolve_user_id(self, handle: str) -> str:
        """Return the numeric user id for a handle (without the leading @)."""
        handle = handle.lstrip("@")
        url = f"{X_API_BASE}/users/by/username/{handle}"
        response = self._request("GET", url)
        if response.status_code != 200:
            raise XApiError(
                f"Failed to resolve handle @{handle}: {response.status_code} {response.text}"
            )
        payload = response.json().get("data") or {}
        if "id" not in payload:
            raise XApiError(f"X API response did not include an id: {payload}")
        return payload["id"]

    def iter_media(
        self,
        user_id: str,
        max_pages: int = 5,
        start_time: Optional[datetime] = None,
        exclude_replies: bool = True,
        exclude_retweets: bool = True,
    ) -> Iterator[TweetMedia]:
        """Yield :class:`TweetMedia` objects for photos attached to recent tweets.

        Only photo media is yielded; videos and animated GIFs are skipped because
        the archiver is purpose-built for static Bloomberg screenshots.
        """
        url = f"{X_API_BASE}/users/{user_id}/tweets"
        params = {
            "max_results": str(DEFAULT_PAGE_SIZE),
            "tweet.fields": "created_at,attachments,entities,text",
            "expansions": "attachments.media_keys",
            "media.fields": "media_key,type,url,preview_image_url,width,height,alt_text",
        }
        excludes = []
        if exclude_replies:
            excludes.append("replies")
        if exclude_retweets:
            excludes.append("retweets")
        if excludes:
            params["exclude"] = ",".join(excludes)
        if start_time is not None:
            params["start_time"] = start_time.astimezone(timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%SZ"
            )

        pages_seen = 0
        while True:
            response = self._request("GET", url, params=params)
            if response.status_code != 200:
                raise XApiError(
                    f"Failed to fetch tweets for {user_id}: "
                    f"{response.status_code} {response.text}"
                )
            payload = response.json()
            tweets = payload.get("data") or []
            media_index = {
                m["media_key"]: m
                for m in (payload.get("includes") or {}).get("media") or []
            }

            for tweet in tweets:
                attachments = (tweet.get("attachments") or {}).get("media_keys") or []
                if not attachments:
                    continue
                created_at = _parse_datetime(tweet.get("created_at"))
                tweet_id = tweet["id"]
                tweet_url = f"https://x.com/i/web/status/{tweet_id}"
                for media_key in attachments:
                    media = media_index.get(media_key)
                    if not media or media.get("type") != "photo":
                        continue
                    image_url = media.get("url") or media.get("preview_image_url")
                    if not image_url:
                        continue
                    yield TweetMedia(
                        tweet_id=tweet_id,
                        tweet_url=tweet_url,
                        tweet_text=tweet.get("text", ""),
                        created_at=created_at,
                        media_key=media_key,
                        image_url=_full_size(image_url),
                        width=media.get("width"),
                        height=media.get("height"),
                        alt_text=media.get("alt_text"),
                    )

            pages_seen += 1
            next_token = (payload.get("meta") or {}).get("next_token")
            if not next_token or pages_seen >= max_pages:
                break
            params["pagination_token"] = next_token

    def download(self, media: TweetMedia) -> bytes:
        response = self._request("GET", media.image_url)
        if response.status_code != 200:
            raise XApiError(
                f"Failed to download {media.image_url}: {response.status_code}"
            )
        return response.content


def _parse_datetime(value: Optional[str]) -> datetime:
    if not value:
        return datetime.now(tz=timezone.utc)
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    return datetime.fromisoformat(value).astimezone(timezone.utc)


def _full_size(image_url: str) -> str:
    """Twitter's CDN serves originals when ``name=orig`` is requested."""
    if "?" in image_url:
        base, query = image_url.split("?", 1)
        params = [p for p in query.split("&") if not p.startswith("name=")]
        params.append("name=orig")
        return f"{base}?{'&'.join(params)}"
    return f"{image_url}?name=orig"


def collect_media(
    client: XClient,
    handle: str,
    max_pages: int = 5,
    start_time: Optional[datetime] = None,
) -> Iterable[TweetMedia]:
    """Convenience wrapper used by the archiver and the CLI."""
    user_id = client.resolve_user_id(handle)
    return client.iter_media(user_id=user_id, max_pages=max_pages, start_time=start_time)
