#!/usr/bin/env python3
"""Publish immutable packages to R2, then advance the stable update pointer.

Credentials are read from environment variables only. --dry-run generates the
same validated manifest without importing boto3 or touching Cloudflare.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from urllib.parse import quote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pastemd import __version__
from pastemd.utils.update_manifest import parse_manifest, valid_version


DEFAULT_PREFIX = "pastemd"


def validate_tag(tag: str) -> str:
    version = tag.removeprefix("v")
    if not tag.startswith("v") or not valid_version(version):
        raise ValueError(f"Unsupported release tag: {tag}")
    if version != __version__:
        raise ValueError(f"Tag {tag} does not match pastemd.__version__ ({__version__})")
    return version


def stable_version(version: str) -> bool:
    return valid_version(version) and not re.search(r"dev|alpha|beta|rc", version)


def version_key(version: str) -> tuple:
    match = re.fullmatch(r"(\d+(?:\.\d+){1,3})(?:[-.]?(dev|alpha|beta|rc)\.?(\d*))?", version)
    if not match:
        raise ValueError("Invalid version ordering")
    numbers = tuple(map(int, match[1].split(".")))
    rank = {"dev": 0, "alpha": 1, "beta": 2, "rc": 3, None: 4}[match[2]]
    return numbers + (0,) * (4 - len(numbers)), rank, int(match[3] or 0)


def build_manifest(tag: str, artifact_dir: Path, public_base: str, prefix: str,
                   notes: str = "", github_fallback: bool = True) -> tuple[dict, list[tuple[Path, str]]]:
    version = validate_tag(tag)
    parsed = urlparse(public_base)
    if (parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password
            or parsed.query or parsed.fragment):
        raise ValueError("R2_PUBLIC_BASE_URL must be an HTTPS base URL")
    if not re.fullmatch(r"[A-Za-z0-9_/-]+", prefix) or ".." in prefix:
        raise ValueError("Unsafe R2_KEY_PREFIX")
    base = public_base.rstrip("/") + "/" + prefix.strip("/")
    assets, uploads = [], []
    found = set()
    for path in sorted(artifact_dir.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in (".exe", ".dmg"):
            continue
        name = path.name
        if name == f"PasteMD_pandoc-Setup_v{version}.exe":
            platform_name, arch = "windows", "x86_64"
        else:
            match = re.fullmatch(rf"PasteMD-{re.escape(version)}-(arm64|x86_64|universal)\.dmg", name)
            if not match:
                raise ValueError(f"Unexpected installer filename: {name}")
            platform_name, arch = "macos", match[1]
        if (platform_name, arch) in found:
            raise ValueError("Duplicate platform/architecture package")
        found.add((platform_name, arch))
        size = path.stat().st_size
        if not size:
            raise ValueError(f"Empty installer: {name}")
        with path.open("rb") as file:
            sha = hashlib.file_digest(file, "sha256").hexdigest()
        key = f"{prefix.strip('/')}/releases/{tag}/{name}"
        uploads.append((path, key))
        urls = [f"{base}/releases/{tag}/{name}"]
        if github_fallback:
            urls.append(f"https://github.com/RICHQAQ/PasteMD/releases/download/{quote(tag, safe='')}/{quote(name, safe='')}")
        assets.append({"name": name, "platform": platform_name, "arch": arch,
                       "size": size, "sha256": sha,
                       "urls": urls})
    if not any(p == "windows" for p, _ in found) or not any(p == "macos" for p, _ in found):
        raise ValueError("A release needs both Windows and macOS installers")
    manifest = {"schema_version": 1, "version": version,
                "channel": "stable" if stable_version(version) else "preview",
                "published_at": datetime.now(timezone.utc).isoformat(),
                "release_url": f"https://github.com/RICHQAQ/PasteMD/releases/tag/{tag}" if github_fallback else "",
                "release_notes": notes, "assets": assets}
    parse_manifest(manifest, f"{base}/latest.json", manifest["channel"])
    return manifest, uploads


def publish(s3, bucket: str, manifest: dict, uploads: list[tuple[Path, str]], prefix: str,
            publish_preview: bool = False) -> None:
    if publish_preview and manifest["channel"] != "preview":
        raise ValueError("Preview publishing requires a prerelease version")
    def head(key):
        try:
            return s3.head_object(Bucket=bucket, Key=key)
        except s3.exceptions.ClientError as exc:
            if exc.response["Error"]["Code"] in ("404", "NoSuchKey", "NotFound"):
                return None
            raise

    # Check all pre-existing objects before writing anything. A tag is immutable;
    # rebuilt packages under the same tag must not poison cached client downloads.
    for (path, key), asset in zip(uploads, manifest["assets"]):
        old = head(key)
        if old and (old["ContentLength"] != asset["size"] or old.get("Metadata", {}).get("sha256") != asset["sha256"]):
            raise ValueError(f"R2 object {key} has different bytes; publish a new version tag")
    for (path, key), asset in zip(uploads, manifest["assets"]):
        if not head(key):
            content_type = "application/x-apple-diskimage" if path.suffix == ".dmg" else "application/vnd.microsoft.portable-executable"
            s3.upload_file(str(path), bucket, key, ExtraArgs={
                "ContentType": content_type, "CacheControl": "public, max-age=31536000, immutable",
                "Metadata": {"sha256": asset["sha256"]},
            })
        uploaded = head(key)
        if not uploaded or uploaded["ContentLength"] != asset["size"] or uploaded.get("Metadata", {}).get("sha256") != asset["sha256"]:
            raise ValueError(f"Post-upload verification failed: {key}")
        print(f"Verified R2 object: {key} ({asset['size']} bytes)")
    payload = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
    s3.put_object(Bucket=bucket, Key=f"{prefix}/releases/v{manifest['version']}/manifest.json",
                  Body=payload, ContentType="application/json; charset=utf-8", CacheControl="public, max-age=300")
    if manifest["channel"] != "stable" and not publish_preview:
        print("Preview release uploaded; stable latest.json is unchanged")
        return
    latest_key = f"{prefix}/latest-preview.json" if publish_preview else f"{prefix}/latest.json"
    old_head = head(latest_key)
    if old_head:
        previous = json.loads(s3.get_object(Bucket=bucket, Key=latest_key)["Body"].read())
        old_version = previous.get("version", "")
        if (not valid_version(old_version) or previous.get("channel") != manifest["channel"]
                or stable_version(old_version) != stable_version(manifest["version"])):
            raise ValueError("Existing update pointer is invalid; inspect it before replacing")
        if version_key(old_version) > version_key(manifest["version"]):
            print(f"Newer version {old_version} already exists; {latest_key} is unchanged")
            return
    s3.put_object(Bucket=bucket, Key=latest_key, Body=payload,
                  ContentType="application/json; charset=utf-8", CacheControl="public, max-age=300")
    print(f"Published {manifest['channel']} feed: {latest_key}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", default=os.environ.get("TAG_NAME", ""))
    parser.add_argument("--check-version", action="store_true")
    parser.add_argument("--artifacts", type=Path, default=Path("artifacts"))
    parser.add_argument("--notes", type=Path)
    parser.add_argument("--output", type=Path, default=Path("update-manifest.json"))
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--publish-preview", action="store_true")
    parser.add_argument("--no-github-fallback", action="store_true")
    args = parser.parse_args()
    validate_tag(args.tag)
    if args.check_version:
        print(f"Release tag matches source: {args.tag}")
        return
    prefix = os.environ.get("R2_KEY_PREFIX", DEFAULT_PREFIX).strip("/")
    manifest, uploads = build_manifest(args.tag, args.artifacts,
        os.environ.get("R2_PUBLIC_BASE_URL", "https://download.richqaq.cn"), prefix,
        args.notes.read_text(encoding="utf-8") if args.notes else "", not args.no_github_fallback)
    if args.publish_preview and manifest["channel"] != "preview":
        raise ValueError("--publish-preview requires a prerelease version")
    args.output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.dry_run:
        print(f"Validated {len(uploads)} packages; manifest written to {args.output}")
        return
    required = ("R2_ACCOUNT_ID", "R2_BUCKET", "R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY", "R2_PUBLIC_BASE_URL")
    missing = [key for key in required if not os.environ.get(key)]
    if missing:
        raise ValueError("Missing R2 settings: " + ", ".join(missing))
    import boto3
    from botocore.config import Config
    s3 = boto3.client("s3", endpoint_url=f"https://{os.environ['R2_ACCOUNT_ID']}.r2.cloudflarestorage.com",
                      region_name="auto", aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
                      aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
                      config=Config(request_checksum_calculation="when_required", response_checksum_validation="when_required"))
    publish(s3, os.environ["R2_BUCKET"], manifest, uploads, prefix, args.publish_preview)


if __name__ == "__main__":
    main()
