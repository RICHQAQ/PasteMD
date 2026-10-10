#!/usr/bin/env python3
"""Bake a run-specific preview version and feed into a disposable CI checkout."""

import argparse
import json
import os
from pathlib import Path
import re
from urllib.parse import urlparse


def prepare(root: Path, run_number: int, public_base: str) -> str:
    parsed = urlparse(public_base)
    if (run_number < 1 or parsed.scheme != "https" or not parsed.netloc
            or parsed.path not in ("", "/") or parsed.username or parsed.password
            or parsed.query or parsed.fragment):
        raise ValueError("A positive run number and an HTTPS domain without a path are required")
    init_path = root / "pastemd/__init__.py"
    init = init_path.read_text(encoding="utf-8")
    match = re.search(r'^__version__ = "(\d+(?:\.\d+){1,3})"$', init, re.MULTILINE)
    if not match:
        raise ValueError("The source version must be stable before preparing a test checkout")
    numbers = match[1].split(".")
    numbers[-1] = str(int(numbers[-1]) + 1)
    version = ".".join(numbers) + f"dev{run_number}"
    feed = public_base.rstrip("/") + "/pastemd-test/latest-preview.json"
    manifest_path = root / "pastemd/utils/update_manifest.py"
    source = manifest_path.read_text(encoding="utf-8")
    for name, value in (("DEFAULT_MANIFEST_URL", feed), ("DEFAULT_UPDATE_CHANNEL", "preview")):
        source, count = re.subn(rf'^{name} = "[^"\n]*"$',
                               lambda _: f"{name} = {json.dumps(value)}", source, flags=re.MULTILINE)
        if count != 1:
            raise ValueError(f"Missing or duplicate build constant: {name}")
    # Validate all inputs before changing either file.
    init_path.write_text(init[:match.start(1)] + version + init[match.end(1):], encoding="utf-8")
    manifest_path.write_text(source, encoding="utf-8")
    return version


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run-number", type=int, required=True)
    parser.add_argument("--public-base", required=True)
    args = parser.parse_args()
    version = prepare(Path(__file__).resolve().parents[1], args.run_number, args.public_base)
    if os.environ.get("GITHUB_ENV"):
        with open(os.environ["GITHUB_ENV"], "a", encoding="utf-8") as file:
            file.write(f"TAG_NAME=v{version}\n")
    print(f"Test checkout prepared: v{version} (preview channel)")


if __name__ == "__main__":
    main()
