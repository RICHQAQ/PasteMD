"""Shared, stdlib-only update feed format used by the client and release scripts."""

import re
from urllib.parse import urlparse

DEFAULT_MANIFEST_URL = "https://download.richqaq.cn/pastemd/latest.json"
DEFAULT_UPDATE_CHANNEL = "stable"
VERSION_PATTERN = r"\d+(?:\.\d+){1,3}(?:(?:[-.]?)(?:dev|alpha|beta|rc)\.?\d*)?"


def valid_version(version: str) -> bool:
    return isinstance(version, str) and re.fullmatch(VERSION_PATTERN, version) is not None


def trusted_download_url(url: str, manifest_url: str = DEFAULT_MANIFEST_URL) -> bool:
    if not isinstance(url, str):
        return False
    parsed, feed = urlparse(url), urlparse(manifest_url)
    if parsed.scheme != "https" or parsed.username or parsed.password or parsed.fragment:
        return False
    if parsed.netloc == "github.com":
        return parsed.path.startswith("/RICHQAQ/PasteMD/releases/download/")
    # A configured HTTPS feed may serve packages only on its own origin.
    return feed.scheme == "https" and parsed.netloc == feed.netloc


def parse_manifest(data: dict, manifest_url: str, channel: str = "stable") -> dict:
    """Normalize R2 metadata to the existing release structure, rejecting bad feeds."""
    if (channel not in ("stable", "preview") or not isinstance(data, dict)
            or data.get("schema_version") != 1 or data.get("channel") != channel
            or not valid_version(data.get("version"))
            or (channel == "stable" and re.search(r"dev|alpha|beta|rc", data["version"]))
            or (channel == "preview" and not re.search(r"dev|alpha|beta|rc", data["version"]))):
        raise ValueError("Invalid update manifest or channel mismatch")
    assets = data.get("assets")
    if not isinstance(assets, list) or not assets:
        raise ValueError("Missing update assets")
    normalized = []
    for asset in assets:
        if not isinstance(asset, dict):
            raise ValueError("Invalid asset")
        name, size, sha = asset.get("name"), asset.get("size"), asset.get("sha256")
        urls = asset.get("urls")
        if (not isinstance(name, str) or not re.fullmatch(r"[A-Za-z0-9_.-]+", name)
                or not isinstance(size, int) or isinstance(size, bool) or size <= 0
                or not isinstance(sha, str) or not re.fullmatch(r"[a-fA-F0-9]{64}", sha)
                or not isinstance(urls, list) or not urls
                or not all(trusted_download_url(url, manifest_url) for url in urls)
                or asset.get("platform") not in ("windows", "macos")
                or asset.get("arch") not in ("x86_64", "arm64", "universal")):
            raise ValueError(f"Invalid update asset: {name}")
        normalized.append({**asset, "state": "uploaded", "digest": "sha256:" + sha.lower(),
                           "browser_download_url": urls[0]})
    return {"tag_name": "v" + data["version"], "body": data.get("release_notes") or "",
            "html_url": data.get("release_url") or "", "assets": normalized}
