from __future__ import annotations

from typing import Any

import requests


def fetchYoutubeMetadata(url: str) -> dict[str, Any]:
    response = requests.get(
        "https://www.youtube.com/oembed",
        params={"url": url, "format": "json"},
        timeout=10,
    )
    response.raise_for_status()
    return response.json()
