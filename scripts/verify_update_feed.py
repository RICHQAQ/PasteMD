#!/usr/bin/env python3
"""Check that the published manifest and installer URLs are publicly readable."""

import argparse
import json
from pathlib import Path
import sys
import time
import urllib.request

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pastemd.utils.update_manifest import parse_manifest


def verify(expected: dict, feed_url: str) -> bool:
    request = urllib.request.Request(feed_url, headers={"User-Agent": "PasteMD-CI-Verification"})
    with urllib.request.urlopen(request, timeout=30) as response:
        if response.geturl().split(":", 1)[0] != "https":
            raise ValueError("Public feed redirected outside HTTPS")
        body = response.read(1024 * 1024 + 1)
        if len(body) > 1024 * 1024:
            raise ValueError("Public feed exceeds 1 MiB")
        actual = json.loads(body)
    parse_manifest(actual, feed_url, expected["channel"])
    if actual != expected:
        return False  # CDN may retain the previous valid pointer for up to 300 s.
    for asset in actual["assets"]:
        request = urllib.request.Request(asset["urls"][0], method="HEAD")
        with urllib.request.urlopen(request, timeout=30) as response:
            if (response.geturl().split(":", 1)[0] != "https"
                    or int(response.headers.get("Content-Length", "-1")) != asset["size"]):
                raise ValueError(f"Public installer size/URL mismatch: {asset['name']}")
        print(f"Public installer verified: {asset['name']} ({asset['size']} bytes)")
    return True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--feed-url", required=True)
    args = parser.parse_args()
    expected = json.loads(args.manifest.read_text(encoding="utf-8"))
    # Validate the origin before making network requests.
    parse_manifest(expected, args.feed_url, expected["channel"])
    for attempt in range(23):
        if verify(expected, args.feed_url):
            print(f"Public feed verified: {args.feed_url} ({expected['version']})")
            return
        if attempt < 22:
            print("Waiting 15 s for the cached update pointer to expire...", flush=True)
            time.sleep(15)
    raise ValueError("Public update pointer did not become visible within 330 s")


if __name__ == "__main__":
    main()
