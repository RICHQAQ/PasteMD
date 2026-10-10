"""Verified downloads and platform installers, independent of the tray/Tk UI."""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
import platform
import plistlib
import re
import shutil
import subprocess
import struct
import sys
import tempfile
import threading
import time
import urllib.request
from urllib.parse import urlparse
from dataclasses import dataclass
from typing import Callable

from ..config.paths import get_log_dir
from .version_checker import VersionChecker
from .update_manifest import trusted_download_url
from .logging import log
from .https import build_https_opener


# Keep the existing install identity: installer.iss escapes the opening brace
# but currently has TWO literal closing braces. Also recognize the conventional
# form for older/manual builds. Changing AppId would create a second installation.
WINDOWS_UNINSTALL_IDS = (
    "{4f3f2b18-55a3-4f40-98f6-d01a3e3e0220}}_is1",
    "{4f3f2b18-55a3-4f40-98f6-d01a3e3e0220}_is1",
)


class UpdateError(Exception):
    """An update failure that can be translated by the presentation layer."""

    def __init__(self, key: str, **details):
        self.key, self.details = key, details
        super().__init__(key)


class UpdateCancelled(Exception):
    pass


@dataclass(frozen=True)
class UpdateAsset:
    name: str
    url: str
    size: int
    sha256: str
    fallback_urls: tuple[str, ...] = ()


def select_asset(release: dict, system: str | None = None,
                 machine: str | None = None) -> UpdateAsset:
    """Only accept our installer for the selected version/platform, never source ZIPs."""
    system = system or platform.system()
    machine = (machine or platform.machine()).lower()
    version = release["latest_version"]
    escaped = re.escape(version)
    if system == "Windows":
        if machine not in ("amd64", "x86_64", "arm64", "aarch64"):
            raise UpdateError("unsupported_platform")
        pattern = rf"PasteMD_pandoc-Setup_v{escaped}\.exe"
    elif system == "Darwin":
        arch = {"aarch64": "arm64", "amd64": "x86_64"}.get(machine, machine)
        if arch not in ("arm64", "x86_64"):
            raise UpdateError("unsupported_platform")
        pattern = rf"PasteMD-{escaped}(?:-(?:{arch}|universal))?\.dmg"
    else:
        raise UpdateError("unsupported_platform")
    matches = [a for a in release.get("assets", [])
               if isinstance(a, dict) and a.get("state") == "uploaded"
               and re.fullmatch(pattern, a.get("name", ""), re.IGNORECASE)
               and (not a.get("platform") or a["platform"] == ("macos" if system == "Darwin" else "windows"))
               and (system != "Darwin" or not a.get("arch") or a["arch"] in (arch, "universal"))]
    # Prefer a labelled native build to the historical unlabelled DMG.
    if system == "Darwin":
        native = [a for a in matches if a["name"].lower().endswith(f"-{arch}.dmg")]
        matches = native or matches
    if len(matches) != 1:
        raise UpdateError("no_asset")
    asset = matches[0]
    url = asset.get("browser_download_url", "")
    urls = asset.get("urls") or [url]
    if not all(trusted_download_url(u, release.get("manifest_url", "")) for u in urls):
        raise UpdateError("invalid_asset")
    digest = asset.get("digest") or ""
    if not re.fullmatch(r"sha256:[a-fA-F0-9]{64}", digest):
        raise UpdateError("missing_checksum")
    size = asset.get("size")
    if not isinstance(size, int) or isinstance(size, bool) or size <= 0:
        raise UpdateError("invalid_asset")
    return UpdateAsset(asset["name"], urls[0], size, digest[7:].lower(), tuple(urls[1:]))


def download_asset(asset: UpdateAsset, destination: Path, cancel: threading.Event,
                   progress: Callable[[int, int], None]) -> Path:
    """Retry network failures through proxy/direct and then alternate mirrors."""
    partial = destination.with_suffix(destination.suffix + ".part")
    last_error = "download_failed"
    try:
        for url in (asset.url, *asset.fallback_urls):
            for use_proxy in (True, False):
                if cancel.is_set():
                    raise UpdateCancelled()
                try:
                    mode = "system proxy" if use_proxy else "direct"
                    log(f"[update] Downloading {url}, mode={mode}, expected={asset.size}")
                    opener = build_https_opener(use_proxy)
                    request = urllib.request.Request(url, headers={"User-Agent": "PasteMD-Updater"})
                    digest = hashlib.sha256()
                    received = 0
                    progress(0, asset.size)
                    with opener.open(request, timeout=15) as response, partial.open("wb") as output:
                        if response.status != 200 or urlparse(response.geturl()).scheme != "https":
                            raise OSError("Invalid download response or insecure redirect")
                        while True:
                            if cancel.is_set():
                                raise UpdateCancelled()
                            block = response.read(256 * 1024)
                            if not block:
                                break
                            received += len(block)
                            if received > asset.size:
                                raise UpdateError("checksum_failed")
                            output.write(block)
                            digest.update(block)
                            progress(received, asset.size)
                    if cancel.is_set():
                        raise UpdateCancelled()
                    actual = digest.hexdigest()
                    if received != asset.size or actual != asset.sha256:
                        log(f"[update] Verification failed: size={received}/{asset.size}, "
                            f"sha256={actual}, expected={asset.sha256}")
                        raise UpdateError("checksum_failed")
                    partial.replace(destination)
                    log(f"[update] Download verified: {destination.name}, sha256={actual}")
                    return destination
                except UpdateError as exc:
                    last_error = exc.key
                    # Never install a corrupt result; try the next source.
                    log(f"[update] Corrupt package from {url}; trying another source")
                    break
                except (OSError, TimeoutError) as exc:
                    last_error = "download_failed"
                    log(f"[update] Download failed ({mode}): {type(exc).__name__}: {exc}")
        raise UpdateError(last_error)
    finally:
        partial.unlink(missing_ok=True)


@dataclass(frozen=True)
class InstallTarget:
    path: Path
    system: str
    all_users: bool = False


def get_install_target() -> InstallTarget:
    """Never overwrite a source checkout, portable directory or mounted DMG."""
    if platform.system() == "Darwin":
        for raw in (sys.executable, sys.argv[0]):
            for parent in Path(raw).resolve().parents:
                if parent.suffix == ".app":
                    info = _read_plist(parent)
                    if info.get("CFBundleIdentifier") != "com.richqaq.pastemd":
                        continue
                    if str(parent).startswith("/Volumes/"):
                        raise UpdateError("move_to_applications")
                    if not os.access(parent.parent, os.W_OK | os.X_OK):
                        raise UpdateError("no_permission")
                    return InstallTarget(parent, "Darwin")
        raise UpdateError("installed_only")
    if platform.system() == "Windows":
        import winreg
        exe = Path(sys.executable).resolve()
        if exe.name.lower() != "pastemd.exe":
            raise UpdateError("installed_only")
        root_key = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        for hive, all_users in ((winreg.HKEY_CURRENT_USER, False), (winreg.HKEY_LOCAL_MACHINE, True)):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                for app_id in WINDOWS_UNINSTALL_IDS:
                    try:
                        with winreg.OpenKey(hive, root_key + "\\" + app_id, 0, winreg.KEY_READ | view) as handle:
                            location = winreg.QueryValueEx(handle, "InstallLocation")[0]
                        if location and Path(location).resolve() == exe.parent:
                            return InstallTarget(exe.parent, "Windows", all_users)
                    except OSError:
                        continue
        raise UpdateError("installed_only")
    raise UpdateError("unsupported_platform")


def _read_plist(app: Path) -> dict:
    try:
        with (app / "Contents/Info.plist").open("rb") as source:
            return plistlib.load(source)
    except (OSError, plistlib.InvalidFileException):
        raise UpdateError("invalid_bundle") from None


def _run(args: list[str], timeout: int = 120) -> subprocess.CompletedProcess:
    try:
        log(f"[update] Running: {args}")
        return subprocess.run(args, check=True, capture_output=True, text=True, timeout=timeout)
    except (OSError, subprocess.SubprocessError) as exc:
        log(f"[update] Command failed: {exc}; stderr={getattr(exc, 'stderr', '')}")
        raise UpdateError("prepare_failed") from None


def _verify_mac_bundle(candidate: Path, installed: Path, version: str) -> None:
    info = _read_plist(candidate)
    if (info.get("CFBundleIdentifier") != "com.richqaq.pastemd"
            or str(info.get("CFBundleShortVersionString")) != version):
        raise UpdateError("invalid_bundle")
    # Preserve the installed designated requirement, including the signing team.
    requirement = _run(["/usr/bin/codesign", "-d", "-r-", str(installed)])
    match = re.search(r"designated => (.+)", requirement.stdout + requirement.stderr)
    if not match or "anchor apple" not in match.group(1):
        raise UpdateError("invalid_signature")
    try:
        _run(["/usr/bin/codesign", "--verify", "--deep", "--strict", "-R",
              match.group(1), str(candidate)])
        _run(["/usr/sbin/spctl", "--assess", "--type", "execute", str(candidate)])
    except UpdateError:
        raise UpdateError("invalid_signature") from None
    executable = info.get("CFBundleExecutable", "")
    if not executable or Path(executable).name != executable:
        raise UpdateError("invalid_bundle")
    if platform.machine() not in _binary_architectures(candidate / "Contents/MacOS" / executable):
        raise UpdateError("wrong_architecture")


def _binary_architectures(executable: Path) -> set[str]:
    """Read Mach-O headers without requiring lipo/Xcode on the user's machine."""
    cpu_names = {0x01000007: "x86_64", 0x0100000C: "arm64"}
    try:
        with executable.open("rb") as source:
            header = source.read(8 + 16 * 32)
        magic = header[:4]
        if magic in (b"\xcf\xfa\xed\xfe", b"\xfe\xed\xfa\xcf"):
            endian = "<" if magic[0] == 0xCF else ">"
            return {cpu_names.get(struct.unpack_from(endian + "I", header, 4)[0], "unknown")}
        fat_formats = {b"\xca\xfe\xba\xbe": (">", 20), b"\xbe\xba\xfe\xca": ("<", 20),
                       b"\xca\xfe\xba\xbf": (">", 32), b"\xbf\xba\xfe\xca": ("<", 32)}
        if magic not in fat_formats:
            raise ValueError("Not a supported Mach-O executable")
        endian, entry_size = fat_formats[magic]
        count = struct.unpack_from(endian + "I", header, 4)[0]
        if not 1 <= count <= 16:
            raise ValueError("Invalid Mach-O architecture count")
        return {cpu_names.get(struct.unpack_from(endian + "I", header, 8 + i * entry_size)[0], "unknown")
                for i in range(count)}
    except (OSError, ValueError, struct.error) as exc:
        log(f"[update] Invalid executable header: {exc}")
        raise UpdateError("invalid_bundle") from None



# Positional arguments keep paths out of shell source; the old app is renamed only
# after its process exits. The backup survives if recovery itself fails.
MAC_HELPER = r'''#!/bin/bash
set -u
pid="$1"; target="$2"; staged="$3"; stage_dir="$4"; work_dir="$5"; ready="$6"
backup="$stage_dir/previous.app"
moved=0
finish() {
    code=$?
    echo "[$(/bin/date)] Update helper finished, exit=$code"
    if [ "$code" -ne 0 ]; then
        if [ "$moved" -eq 1 ] && [ -d "$backup" ]; then
            /bin/rm -rf "$target"
            /bin/mv "$backup" "$target" || exit 1
            /usr/bin/open "$target" || true
        fi
        /usr/bin/osascript -e 'display alert "PasteMD update failed" message "The previous version was restored when possible. See the PasteMD update log."' || true
    else
        /bin/rm -rf "$stage_dir"
    fi
    /bin/rm -rf "$work_dir"
}
trap finish EXIT
echo "[$(/bin/date)] Waiting for PasteMD process $pid to exit"
/usr/bin/touch "$ready" || exit 1
for ((i=0; i<120; i++)); do
    /bin/kill -0 "$pid" 2>/dev/null || break
    /bin/sleep 1
done
/bin/kill -0 "$pid" 2>/dev/null && exit 1
echo "[$(/bin/date)] Replacing $target; backup=$backup"
/bin/mv "$target" "$backup" || exit 1
moved=1
/bin/mv "$staged" "$target" || exit 1
echo "[$(/bin/date)] Starting $target"
/usr/bin/open "$target" || exit 1
'''


# Inno's silent install skips [Run]; this helper waits for it and explicitly
# restarts PasteMD. UAC belongs to the installer, never to a downloaded script.
WINDOWS_HELPER = r'''param([string]$ConfigPath)
$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
try {
    Write-Output "Waiting for PasteMD process $($config.pid) to exit"
    $old = Get-Process -Id $config.pid -ErrorAction SilentlyContinue
    New-Item -ItemType File -Path $config.ready -Force | Out-Null
    if ($old -and -not $old.WaitForExit(120000)) { throw 'PasteMD did not exit' }
    $parameters = @('/SILENT', '/NORESTART', '/SUPPRESSMSGBOXES',
        ('/DIR="' + $config.target + '"'), ('/LOG="' + $config.installer_log + '"'))
    Write-Output "Installing into $($config.target)"
    if ($config.all_users) {
        $parameters += '/ALLUSERS'
        $installer = Start-Process -FilePath $config.installer -ArgumentList $parameters -Verb RunAs -PassThru -Wait
    } else {
        $parameters += '/CURRENTUSER'
        $installer = Start-Process -FilePath $config.installer -ArgumentList $parameters -PassThru -Wait
    }
    if ($installer.ExitCode -ne 0) { throw ('Installer exit code: ' + $installer.ExitCode) }
    Start-Process -FilePath (Join-Path $config.target 'PasteMD.exe')
} catch {
    Write-Output ("[$(Get-Date -Format o)] Update failed: " + $_.Exception.Message)
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(('Update failed. See ' + $config.log), 'PasteMD') | Out-Null
    $exe = Join-Path $config.target 'PasteMD.exe'
    if (Test-Path -LiteralPath $exe) { Start-Process -FilePath $exe }
} finally {
    Remove-Item -LiteralPath $config.work_dir -Recurse -Force -ErrorAction SilentlyContinue
}
'''


class PreparedUpdate:
    """Own staged files until a detached helper takes responsibility for cleanup."""

    def __init__(self, target: InstallTarget, package: Path, work_dir: Path,
                 version: str, cancel: threading.Event):
        self.target, self.package, self.work_dir = target, package, work_dir
        self.stage_dir: Path | None = None
        self.handed_off = False
        if target.system == "Darwin":
            try:
                self._prepare_mac(version, cancel)
            except BaseException:
                self.cleanup()
                raise

    def _prepare_mac(self, version: str, cancel: threading.Event) -> None:
        mount = self.work_dir / "mount"
        mount.mkdir()
        attached = False
        try:
            _run(["/usr/bin/hdiutil", "attach", str(self.package), "-readonly", "-nobrowse",
                  "-mountpoint", str(mount)])
            attached = True
            candidate = mount / "PasteMD.app"
            _verify_mac_bundle(candidate, self.target.path, version)
            if cancel.is_set():
                raise UpdateCancelled()
            self.stage_dir = Path(tempfile.mkdtemp(prefix=".pastemd-update-", dir=self.target.path.parent))
            _run(["/usr/bin/ditto", str(candidate), str(self.stage_dir / "PasteMD.app")])
            _verify_mac_bundle(self.stage_dir / "PasteMD.app", self.target.path, version)
        finally:
            if attached:
                try:
                    _run(["/usr/bin/hdiutil", "detach", str(mount)])
                except UpdateError:
                    _run(["/usr/bin/hdiutil", "detach", "-force", str(mount)])
        if cancel.is_set():
            raise UpdateCancelled()

    def launch(self) -> None:
        import json
        log_dir = Path(get_log_dir())
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "update-install.log"
        log(f"[update] Launching installer helper for {self.target.path}; log={log_path}")
        ready = self.work_dir / "helper.ready"
        if self.target.system == "Darwin":
            helper = self.work_dir / "install.sh"
            helper.write_text(MAC_HELPER, encoding="utf-8")
            command = ["/bin/bash", str(helper), str(os.getpid()), str(self.target.path),
                       str(self.stage_dir / "PasteMD.app"), str(self.stage_dir),
                       str(self.work_dir), str(ready)]
            kwargs = {"start_new_session": True}
        else:
            helper = self.work_dir / "install.ps1"
            helper.write_text(WINDOWS_HELPER, encoding="utf-8-sig")
            config_path = self.work_dir / "install.json"
            config_path.write_text(json.dumps({
                "pid": os.getpid(), "target": str(self.target.path),
                "all_users": self.target.all_users, "installer": str(self.package),
                "log": str(log_path), "installer_log": str(log_dir / "update-installer.log"), "work_dir": str(self.work_dir), "ready": str(ready),
            }), encoding="utf-8")
            powershell = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
            command = [str(powershell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(helper), str(config_path)]
            kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS}
        with log_path.open("ab") as output:
            process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=output,
                                       stderr=output, **kwargs)
        for _ in range(100):
            if process.poll() is not None:
                raise UpdateError("prepare_failed")
            if ready.exists():
                log("[update] Installer helper ready; exiting PasteMD for installation")
                self.handed_off = True
                return
            time.sleep(0.05)
        process.terminate()
        process.wait(timeout=5)
        raise UpdateError("prepare_failed")

    def cleanup(self) -> None:
        if not self.handed_off:
            if self.stage_dir:
                shutil.rmtree(self.stage_dir, ignore_errors=True)
            shutil.rmtree(self.work_dir, ignore_errors=True)


def refresh_release_assets(release: dict) -> dict:
    """Old API mirrors may omit digests; retrieve the same tag from GitHub."""
    if release.get("manifest_url") and any(a.get("urls") for a in release.get("assets", [])):
        return release
    if release.get("assets") and all(a.get("digest") for a in release["assets"] if isinstance(a, dict)):
        return release
    checker = VersionChecker(release["current_version"])
    checker._prepare_ssl_environment()
    from urllib.parse import quote
    url = "https://api.github.com/repos/RICHQAQ/PasteMD/releases/tags/" + quote("v" + release["latest_version"], safe="")
    data = checker._fetch_release_url(url)
    if data and data.get("tag_name", "").lstrip("vV") == release["latest_version"] and not data.get("draft") and not data.get("prerelease"):
        return {**release, "assets": data.get("assets") or []}
    return release
