"""Update trust, recovery, threading and publication invariants."""

import copy
import hashlib
import io
import json
from pathlib import Path
import queue
import subprocess
import threading
from types import SimpleNamespace

import pytest

from pastemd import __version__
from pastemd.service.update import UpdateSession
from pastemd.utils import updater
from pastemd.utils.update_manifest import DEFAULT_MANIFEST_URL, parse_manifest
from pastemd.utils.version_checker import VersionChecker
from scripts import publish_update


def manifest():
    name = f"PasteMD_pandoc-Setup_v{__version__}.exe"
    return {"schema_version": 1, "channel": "stable", "version": __version__,
            "release_notes": "Some notes", "assets": [{"name": name, "platform": "windows",
            "arch": "x86_64", "size": 7, "sha256": hashlib.sha256(b"package").hexdigest(),
            "urls": [f"https://download.richqaq.cn/pastemd/releases/v{__version__}/{name}",
                     f"https://github.com/RICHQAQ/PasteMD/releases/download/v{__version__}/{name}"]}]}


def release():
    data = parse_manifest(manifest(), DEFAULT_MANIFEST_URL)
    return {"latest_version": __version__, "current_version": "0.0.1",
            "manifest_url": DEFAULT_MANIFEST_URL, "has_update": True, "assets": data["assets"]}


def test_r2_feed_is_used_without_github(monkeypatch):
    checker = VersionChecker("0.0.1")
    calls = []
    monkeypatch.setattr(checker, "_fetch_release_url", lambda url: calls.append(url) or manifest())
    result = checker.check_update()
    assert calls == [DEFAULT_MANIFEST_URL]
    assert result["has_update"] and result["assets"][0]["urls"][0].startswith("https://download.")
    assert result["release_notes"] == "Some notes"


def test_bad_r2_feed_falls_back_to_release_api(monkeypatch):
    checker = VersionChecker("0.0.1")
    calls = []
    github = {"tag_name": "v" + __version__, "assets": []}
    def fetch(url):
        calls.append(url)
        return {"schema_version": 9} if url == DEFAULT_MANIFEST_URL else github
    monkeypatch.setattr(checker, "_fetch_release_url", fetch)
    assert checker.check_update()["has_update"]
    assert calls == [DEFAULT_MANIFEST_URL, checker.GITHUB_API_URL]


@pytest.mark.parametrize("latest,current,expected", [
    ("0.1.7.10", "0.1.7.9", True), ("0.1.7.6", "0.1.7.6", False),
    ("0.1.7.6", "0.1.7.6rc1", True), ("0.1.7.6rc2", "0.1.7.6rc1", True),
    ("0.1.7.6beta1", "0.1.7.6alpha3", True), ("0.1.7.6dev1", "0.1.7.6", False),
    ("0.1.7", "0.1.7.0", False),
])
def test_version_order(latest, current, expected):
    assert VersionChecker(current)._compare_versions(latest, current) is expected


@pytest.mark.parametrize("mutation", [
    lambda m: m.update(schema_version=2),
    lambda m: m.update(channel="preview"),
    lambda m: m.update(version="../1.0"),
    lambda m: m["assets"][0].update(name="../../installer.exe"),
    lambda m: m["assets"][0].update(sha256=""),
    lambda m: m["assets"][0].update(size=-1),
    lambda m: m["assets"][0].update(size=True),
    lambda m: m["assets"][0].update(urls=["https://evil.test/install.exe"]),
    lambda m: m["assets"][0].update(urls=["http://download.richqaq.cn/install.exe"]),
    lambda m: m["assets"][0].update(urls=["https://github.com/other/project/install.exe"]),
])
def test_invalid_manifests_rejected(mutation):
    data = manifest()
    mutation(data)
    with pytest.raises(ValueError):
        parse_manifest(data, DEFAULT_MANIFEST_URL)


def test_select_native_arch_and_reject_missing_package():
    data = release()
    assert updater.select_asset(data, "Windows", "AMD64").name.endswith(".exe")
    with pytest.raises(updater.UpdateError, match="no_asset"):
        updater.select_asset(data, "Darwin", "arm64")
    base = data["assets"][0]
    data["assets"] = [{**base, "name": f"PasteMD-{__version__}-{arch}.dmg", "platform": "macos", "arch": arch}
                      for arch in ("arm64", "x86_64")]
    assert "-arm64.dmg" in updater.select_asset(data, "Darwin", "arm64").name
    assert "-x86_64.dmg" in updater.select_asset(data, "Darwin", "x86_64").name


class Response(io.BytesIO):
    status = 200
    def geturl(self):
        return "https://cdn.test/object"


def mock_download(monkeypatch, responses):
    calls = []
    def open_request(request, timeout):
        calls.append(request.full_url)
        item = responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return Response(item)
    monkeypatch.setattr(updater.urllib.request, "build_opener", lambda *args: SimpleNamespace(open=open_request))
    return calls


def test_download_verifies_then_promotes_partial(tmp_path, monkeypatch):
    asset = updater.select_asset(release(), "Windows", "AMD64")
    mock_download(monkeypatch, [b"package"])
    destination = tmp_path / asset.name
    progress = []
    assert updater.download_asset(asset, destination, threading.Event(), lambda n, total: progress.append(n)) == destination
    assert destination.read_bytes() == b"package"
    assert progress[0] == 0 and progress[-1] == asset.size
    assert not destination.with_suffix(".exe.part").exists()


def test_network_fallback_and_clean_retry(tmp_path, monkeypatch, caplog):
    caplog.set_level("INFO", logger="pastemd")
    asset = updater.select_asset(release(), "Windows", "AMD64")
    calls = mock_download(monkeypatch, [TimeoutError("slow"), OSError("offline"), b"package"])
    updater.download_asset(asset, tmp_path / asset.name, threading.Event(), lambda *args: None)
    assert calls == [asset.url, asset.url, asset.fallback_urls[0]]
    assert "TimeoutError: slow" in caplog.text
    assert "Download verified" in caplog.text


@pytest.mark.parametrize("bad", [b"pack", b"garbage", b"package-extra"])
def test_corrupt_download_never_installed(tmp_path, monkeypatch, bad):
    asset = updater.select_asset(release(), "Windows", "AMD64")
    mock_download(monkeypatch, [bad, bad])
    destination = tmp_path / asset.name
    with pytest.raises(updater.UpdateError, match="checksum_failed"):
        updater.download_asset(asset, destination, threading.Event(), lambda *args: None)
    assert list(tmp_path.iterdir()) == []


def test_cancel_during_download_removes_partial(tmp_path, monkeypatch):
    asset = updater.select_asset(release(), "Windows", "AMD64")
    mock_download(monkeypatch, [b"package"])
    cancel = threading.Event()
    def progress(n, total):
        if n:
            cancel.set()
    with pytest.raises(updater.UpdateCancelled):
        updater.download_asset(asset, tmp_path / asset.name, cancel, progress)
    assert list(tmp_path.iterdir()) == []


def test_source_checkout_cannot_be_replaced(monkeypatch):
    monkeypatch.setattr(updater.platform, "system", lambda: "Darwin")
    monkeypatch.setattr(updater.sys, "executable", "/usr/bin/python3")
    monkeypatch.setattr(updater.sys, "argv", ["/tmp/main.py"])
    with pytest.raises(updater.UpdateError, match="installed_only"):
        updater.get_install_target()


def test_session_callbacks_only_run_on_polling_thread(monkeypatch):
    from pastemd.service import update
    monkeypatch.setattr(update.VersionChecker, "check_update", lambda self: release())
    callback_threads = []
    session = UpdateSession({}, lambda: callback_threads.append(threading.get_ident()), lambda: None)
    session.check()
    session.worker.join(timeout=2)
    assert session.state == "checking"
    session.poll()
    assert session.state == "available"
    assert set(callback_threads) == {threading.get_ident()}


def test_cancel_race_cleans_prepared_result():
    session = UpdateSession({}, lambda: None, lambda: None)
    cleaned = []
    session.state = "downloading"
    session.events.put(("ready", SimpleNamespace(cleanup=lambda: cleaned.append(True))))
    session.cancel()
    session.poll()
    assert session.state == "cancelled" and session.prepared is None and cleaned == [True]


def test_install_launch_failure_is_visible_and_cleans_staging():
    session = UpdateSession({}, lambda: None, lambda: pytest.fail("must not quit"))
    cleaned = []
    def fail():
        raise updater.UpdateError("prepare_failed")
    session.prepared = SimpleNamespace(launch=fail, cleanup=lambda: cleaned.append(True))
    session.state = "ready"
    session.install()
    session.worker.join(timeout=2)
    session.poll()
    assert session.state == "failed" and session.error_key == "prepare_failed"
    assert cleaned == [True] and session.prepared is None


def test_mac_signature_mismatch_stops_before_replacement(monkeypatch, tmp_path):
    candidate = tmp_path / "new.app"
    candidate.mkdir()
    monkeypatch.setattr(updater, "_read_plist", lambda _: {"CFBundleIdentifier": "com.richqaq.pastemd", "CFBundleShortVersionString": __version__})
    calls = []
    def command(args):
        calls.append(args)
        if "--verify" in args:
            raise updater.UpdateError("prepare_failed")
        return SimpleNamespace(stdout='designated => anchor apple generic and identifier "com.richqaq.pastemd"', stderr="")
    monkeypatch.setattr(updater, "_run", command)
    with pytest.raises(updater.UpdateError, match="invalid_signature"):
        updater._verify_mac_bundle(candidate, tmp_path / "old.app", __version__)
    assert not any("ditto" in arg for args in calls for arg in args)


@pytest.mark.parametrize("launch_success", [True, False])
def test_mac_detached_helper_replacement_and_rollback(tmp_path, launch_success):
    target = tmp_path / "PasteMD with $ and ' spaces.app"
    stage = tmp_path / ".stage"
    work = tmp_path / "work"
    staged = stage / "PasteMD.app"
    target.mkdir(); stage.mkdir(); work.mkdir(); staged.mkdir()
    (target / "version").write_text("old")
    (staged / "version").write_text("new")
    # Real rename/rollback on throwaway dirs; replace only GUI command paths.
    launch = tmp_path / "launch"
    alert = tmp_path / "alert"
    launch.write_text("#!/bin/bash\nexit " + ("0" if launch_success else "1") + "\n")
    alert.write_text("#!/bin/bash\nexit 0\n")
    launch.chmod(0o700); alert.chmod(0o700)
    script = tmp_path / "install.sh"
    script.write_text(updater.MAC_HELPER.replace("/usr/bin/open", '"' + str(launch) + '"').replace("/usr/bin/osascript", '"' + str(alert) + '"'))
    old = subprocess.Popen(["/usr/bin/true"])
    old.wait()
    result = subprocess.run(["/bin/bash", str(script), str(old.pid), str(target), str(staged), str(stage), str(work), str(work / "ready")], capture_output=True, text=True, timeout=5)
    assert result.returncode == (0 if launch_success else 1)
    assert (target / "version").read_text() == ("new" if launch_success else "old")
    assert not work.exists()


def create_artifacts(tmp_path):
    for name in (f"PasteMD_pandoc-Setup_v{__version__}.exe", f"PasteMD-{__version__}-arm64.dmg"):
        (tmp_path / name).write_bytes(b"package")
    return publish_update.build_manifest("v" + __version__, tmp_path, "https://download.richqaq.cn", "pastemd", "Release notes")


def test_release_manifest_is_usable_by_client(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    assert len(uploads) == 2
    normalized = parse_manifest(data, DEFAULT_MANIFEST_URL)
    assert len(normalized["assets"]) == 2
    assert data["assets"][0]["sha256"] == hashlib.sha256(b"package").hexdigest()


def test_mismatched_tag_and_missing_platform_rejected(tmp_path):
    with pytest.raises(ValueError, match="does not match"):
        publish_update.validate_tag("v999.0")
    (tmp_path / f"PasteMD-{__version__}-arm64.dmg").write_bytes(b"package")
    with pytest.raises(ValueError, match="both Windows and macOS"):
        publish_update.build_manifest("v" + __version__, tmp_path, "https://download.richqaq.cn", "pastemd")


class FakeS3:
    class ClientError(Exception):
        response = {"Error": {"Code": "404"}}
    exceptions = SimpleNamespace(ClientError=ClientError)
    def __init__(self):
        self.objects = {}
        self.writes = []
        self.fail_upload = False
    def head_object(self, Bucket, Key):
        if Key not in self.objects:
            raise self.ClientError()
        obj = self.objects[Key]
        return {"ContentLength": len(obj["Body"]), "Metadata": obj.get("Metadata", {})}
    def upload_file(self, path, bucket, key, ExtraArgs):
        if self.fail_upload:
            raise OSError("upload failed")
        self.put_object(Bucket=bucket, Key=key, Body=Path(path).read_bytes(), **ExtraArgs)
    def put_object(self, Bucket, Key, **kwargs):
        self.writes.append(Key)
        self.objects[Key] = kwargs
    def get_object(self, Bucket, Key):
        return {"Body": io.BytesIO(self.objects[Key]["Body"])}


def test_latest_is_last_and_failed_upload_cannot_advance_it(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    s3 = FakeS3()
    s3.fail_upload = True
    with pytest.raises(OSError):
        publish_update.publish(s3, "bucket", data, uploads, "pastemd")
    assert "pastemd/latest.json" not in s3.writes
    s3.fail_upload = False
    publish_update.publish(s3, "bucket", data, uploads, "pastemd")
    assert s3.writes[-1] == "pastemd/latest.json"
    assert s3.objects["pastemd/latest.json"]["CacheControl"] == "public, max-age=300"
    assert all(s3.objects[key]["CacheControl"].endswith("immutable") for _, key in uploads)


def test_same_tag_different_bytes_cannot_overwrite_cache(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    s3 = FakeS3()
    publish_update.publish(s3, "bucket", data, uploads, "pastemd")
    s3.writes.clear()
    modified = copy.deepcopy(data)
    modified["assets"][0]["sha256"] = "0" * 64
    with pytest.raises(ValueError, match="different bytes"):
        publish_update.publish(s3, "bucket", modified, uploads, "pastemd")
    assert s3.writes == []


def test_older_release_cannot_roll_back_stable_feed(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    s3 = FakeS3()
    s3.objects["pastemd/latest.json"] = {"Body": json.dumps({"version": "999.0", "channel": "stable"}).encode()}
    publish_update.publish(s3, "bucket", data, uploads, "pastemd")
    assert "pastemd/latest.json" not in s3.writes
    assert json.loads(s3.objects["pastemd/latest.json"]["Body"])["version"] == "999.0"


def test_preview_release_does_not_change_stable_feed(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    data["channel"] = "preview"
    data["version"] += "dev1"
    s3 = FakeS3()
    publish_update.publish(s3, "bucket", data, uploads, "pastemd")
    assert "pastemd/latest.json" not in s3.writes


def test_preview_is_opt_in_and_cannot_fall_back_to_production(monkeypatch):
    preview = manifest()
    preview.update(channel="preview", version=__version__ + "dev2")
    feed = "https://download.richqaq.cn/pastemd-test/latest-preview.json"
    with pytest.raises(ValueError):
        parse_manifest(preview, feed)
    assert parse_manifest(preview, feed, "preview")["tag_name"].endswith("dev2")
    checker = VersionChecker(__version__ + "dev1", feed, "preview")
    monkeypatch.setattr(checker, "_fetch_release_url", lambda _: preview)
    assert checker.check_update()["has_update"]
    calls = []
    monkeypatch.setattr(checker, "_fetch_release_url", lambda url: calls.append(url))
    assert checker.check_update() is None
    assert calls == [feed]


def test_preview_pointer_is_last_and_cannot_change_stable_or_downgrade(tmp_path, monkeypatch):
    monkeypatch.setattr(publish_update, "__version__", __version__ + "dev2")
    version = publish_update.__version__
    for name in (f"PasteMD_pandoc-Setup_v{version}.exe", f"PasteMD-{version}-arm64.dmg"):
        (tmp_path / name).write_bytes(b"package")
    data, uploads = publish_update.build_manifest("v" + version, tmp_path,
        "https://pub-test.r2.dev", "pastemd-test", github_fallback=False)
    assert all(len(asset["urls"]) == 1 for asset in data["assets"])
    assert data["release_url"] == ""
    s3 = FakeS3()
    publish_update.publish(s3, "test-bucket", data, uploads, "pastemd-test", publish_preview=True)
    assert s3.writes[-1] == "pastemd-test/latest-preview.json"
    assert not any(key.endswith("/latest.json") for key in s3.writes)
    newer = {**data, "version": __version__ + "dev10"}
    s3.objects["pastemd-test/latest-preview.json"]["Body"] = json.dumps(newer).encode()
    s3.writes.clear()
    publish_update.publish(s3, "test-bucket", data, uploads, "pastemd-test", publish_preview=True)
    assert "pastemd-test/latest-preview.json" not in s3.writes


def test_stable_version_cannot_publish_preview_pointer(tmp_path):
    data, uploads = create_artifacts(tmp_path)
    s3 = FakeS3()
    with pytest.raises(ValueError, match="prerelease"):
        publish_update.publish(s3, "bucket", data, uploads, "pastemd-test", publish_preview=True)
    assert not s3.writes


@pytest.mark.parametrize("suffix", ["dev1", "-alpha.2", ".beta3", "rc1", "-rc.2"])
def test_preview_publisher_accepts_all_supported_suffixes(tmp_path, monkeypatch, suffix):
    version = __version__ + suffix
    monkeypatch.setattr(publish_update, "__version__", version)
    for name in (f"PasteMD_pandoc-Setup_v{version}.exe", f"PasteMD-{version}-arm64.dmg"):
        (tmp_path / name).write_bytes(b"package")
    data, _ = publish_update.build_manifest("v" + version, tmp_path,
        "https://pub-test.r2.dev", "pastemd-test", github_fallback=False)
    assert parse_manifest(data, "https://pub-test.r2.dev/pastemd-test/latest-preview.json", "preview")


def test_ci_preview_preparation_does_not_edit_production_checkout(tmp_path):
    from scripts.prepare_test_build import prepare
    init = tmp_path / "pastemd/__init__.py"
    constants = tmp_path / "pastemd/utils/update_manifest.py"
    constants.parent.mkdir(parents=True)
    init.write_text('__version__ = "0.1.7.6"\n')
    constants.write_text('DEFAULT_MANIFEST_URL = "https://production.test/latest.json"\nDEFAULT_UPDATE_CHANNEL = "stable"\n')
    original = Path("pastemd/__init__.py").read_text()
    assert prepare(tmp_path, 123, "https://pub-test.r2.dev") == "0.1.7.7dev123"
    assert 'DEFAULT_UPDATE_CHANNEL = "preview"' in constants.read_text()
    assert "https://pub-test.r2.dev/pastemd-test/latest-preview.json" in constants.read_text()
    assert Path("pastemd/__init__.py").read_text() == original


def test_public_feed_verification_checks_package_lengths(monkeypatch):
    from scripts import verify_update_feed
    expected = manifest()
    feed = DEFAULT_MANIFEST_URL
    requests = []
    def open_request(request, timeout):
        requests.append(request.get_method())
        response = io.BytesIO(json.dumps(expected).encode())
        response.geturl = lambda: request.full_url
        response.headers = {"Content-Length": "7"}
        return response
    monkeypatch.setattr(verify_update_feed.urllib.request, "urlopen", open_request)
    assert verify_update_feed.verify(expected, feed)
    assert requests == ["GET", "HEAD"]
    expected["assets"][0]["size"] = 8
    with pytest.raises(ValueError, match="mismatch"):
        verify_update_feed.verify(expected, feed)


def test_window_close_keeps_download_and_ready_requires_user_click():
    import tkinter as tk
    from pastemd.presentation.update.dialog import UpdateDialog
    try:
        root = tk.Tk()
        root.withdraw()
    except tk.TclError as exc:
        pytest.skip(str(exc))
    closed, installs = [], []
    session = UpdateSession({}, lambda: None, lambda: None)
    session.state = "downloading"
    session.release = release()
    session.received, session.total, session.speed = 5, 10, 1024
    window = UpdateDialog(root, session, lambda: closed.append(True), lambda: installs.append(True))
    root.update_idletasks()
    assert float(window.progress["value"]) == 50
    assert str(window.primary["state"]) == "disabled"
    window.close()
    assert session.state == "downloading" and not session.cancel_event.is_set()
    assert closed == [True]
    session.state = "ready"
    window = UpdateDialog(root, session, lambda: None, lambda: installs.append(True))
    root.update_idletasks()
    assert str(window.primary["state"]) == "normal"
    assert installs == []
    window._primary()
    assert installs == [True]
    window.close()
    root.destroy()


def test_tray_progress_and_cancel_control_without_browser(monkeypatch):
    from pastemd.core.state import app_state
    from pastemd.config.defaults import DEFAULT_CONFIG
    from pastemd.presentation.tray.menu import TrayMenuManager
    monkeypatch.setattr(app_state, "config", copy.deepcopy(DEFAULT_CONFIG))
    monkeypatch.setattr(app_state, "ui_queue", queue.Queue())
    manager = TrayMenuManager(None, SimpleNamespace(notify=lambda *args, **kwargs: None))
    manager.update_session.state = "downloading"
    manager.update_session.received, manager.update_session.total = 5, 10
    assert "50%" in manager._update_menu_text()
    menu = manager.build_menu()
    from pastemd.i18n import t
    assert any(item.text == t("update.cancel") for item in menu.items)
    manager._on_cancel_update(None, None)
    app_state.ui_queue.get_nowait()()
    assert manager.update_session.cancel_event.is_set()
    assert manager.update_session.state == "cancelling"


def test_windows_helper_is_detached_and_config_keeps_paths_out_of_code(monkeypatch, tmp_path):
    target = updater.InstallTarget(tmp_path / "installed folder", "Windows", True)
    work = tmp_path / "work"
    work.mkdir()
    package = work / "installer.exe"
    prepared = updater.PreparedUpdate(target, package, work, __version__, threading.Event())
    monkeypatch.setattr(updater, "get_log_dir", lambda: str(tmp_path / "logs"))
    monkeypatch.setattr(subprocess, "CREATE_NO_WINDOW", 0x08000000, raising=False)
    monkeypatch.setattr(subprocess, "DETACHED_PROCESS", 0x8, raising=False)
    calls = []
    def spawn(command, **kwargs):
        calls.append((command, kwargs))
        config = json.loads((work / "install.json").read_text())
        Path(config["ready"]).touch()
        return SimpleNamespace(poll=lambda: None)
    monkeypatch.setattr(subprocess, "Popen", spawn)
    prepared.launch()
    config = json.loads((work / "install.json").read_text())
    assert config["target"] == str(target.path) and config["all_users"] is True
    assert config["installer_log"] != config["log"]
    assert calls[0][1]["creationflags"] & subprocess.DETACHED_PROCESS
    assert str(target.path) not in (work / "install.ps1").read_text(encoding="utf-8-sig")
    prepared.cleanup()
    assert work.exists()  # helper owns cleanup after handoff


def test_macho_architecture_check_needs_no_developer_tools(tmp_path):
    import struct
    binary = tmp_path / "PasteMD"
    binary.write_bytes(struct.pack("<II", 0xFEEDFACF, 0x0100000C))
    assert updater._binary_architectures(binary) == {"arm64"}
    binary.write_bytes(struct.pack("<II", 0xFEEDFACF, 0x01000007))
    assert updater._binary_architectures(binary) == {"x86_64"}
    binary.write_bytes(struct.pack(">II", 0xCAFEBABE, 2) +
        struct.pack(">IIIII", 0x01000007, 0, 0, 0, 0) +
        struct.pack(">IIIII", 0x0100000C, 0, 0, 0, 0))
    assert updater._binary_architectures(binary) == {"arm64", "x86_64"}
    binary.write_bytes(b"bad header")
    with pytest.raises(updater.UpdateError, match="invalid_bundle"):
        updater._binary_architectures(binary)


def test_installer_registry_id_matches_current_script():
    setup = Path("installer.iss").read_text()
    app_id = setup.split("AppId=", 1)[1].splitlines()[0].replace("{{", "{")
    assert app_id + "_is1" in updater.WINDOWS_UNINSTALL_IDS


@pytest.mark.parametrize("all_users", [False, True])
def test_windows_install_detection_preserves_original_privilege_mode(monkeypatch, tmp_path, all_users):
    import sys
    installed = tmp_path / "PasteMD"
    monkeypatch.setattr(updater.platform, "system", lambda: "Windows")
    monkeypatch.setattr(sys, "executable", str(installed / "PasteMD.exe"))
    class RegistryHandle:
        def __enter__(self):
            return self
        def __exit__(self, *args):
            pass
    def open_key(hive, key, _, access):
        if hive != (2 if all_users else 1) or not key.endswith(updater.WINDOWS_UNINSTALL_IDS[0]):
            raise OSError("not found")
        return RegistryHandle()
    monkeypatch.setitem(sys.modules, "winreg", SimpleNamespace(
        HKEY_CURRENT_USER=1, HKEY_LOCAL_MACHINE=2, KEY_WOW64_64KEY=4, KEY_WOW64_32KEY=8,
        KEY_READ=16, OpenKey=open_key, QueryValueEx=lambda *args: (str(installed), None)))
    target = updater.get_install_target()
    assert target.path == installed and target.all_users is all_users
