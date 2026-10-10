"""Development version overrides must stay separate from installer identity."""

import hashlib
import shutil
from types import SimpleNamespace

import pytest

from pastemd.service import update
from pastemd.utils.updater import InstallTarget, select_asset
from pastemd.utils.version_checker import VersionChecker


ACTUAL = "0.1.7.7dev29"
FEED = "https://download-test.example.com/pastemd-test/latest-preview.json"


def preview():
    name = f"PasteMD_pandoc-Setup_v{ACTUAL}.exe"
    return {"schema_version": 1, "channel": "preview", "version": ACTUAL,
            "release_notes": "Test update", "assets": [{
                "name": name, "platform": "windows", "arch": "x86_64", "size": 7,
                "sha256": hashlib.sha256(b"package").hexdigest(),
                "urls": [f"https://download-test.example.com/pastemd-test/releases/v{ACTUAL}/{name}"],
            }]}


def session(monkeypatch, dev):
    monkeypatch.setattr(update, "__version__", ACTUAL)
    monkeypatch.setattr(VersionChecker, "_fetch_release_url", lambda self, url: preview())
    return update.UpdateSession({"dev": dev, "update_channel": "preview", "update_manifest_url": FEED},
                                lambda: None, lambda: None)


def check(current):
    current.check()
    current.worker.join(timeout=2)
    assert not current.worker.is_alive()
    current.poll()


@pytest.mark.parametrize("dev,simulated", [
    ({"enabled": True, "version": "0.1.7.6"}, True),
    ({"enabled": True, "version": "0.1.7.7dev28"}, True),
    ({"enabled": False, "version": "0.1.7.6"}, False),
    ({"enabled": "false", "version": "0.1.7.6"}, False),
    ({"enabled": 1, "version": "0.1.7.6"}, False),
    ({"enabled": True, "version": ""}, False),
    ({"enabled": True, "version": "../../bad"}, False),
    ({"enabled": True, "version": 123}, False),
    (True, False),
    (None, False),
])
def test_dev_only_changes_current_version_when_explicitly_enabled(monkeypatch, dev, simulated):
    current = session(monkeypatch, dev)
    check(current)
    assert current.state == ("available" if simulated else "latest")
    assert current.debug_version_active is simulated
    assert update.__version__ == ACTUAL
    if simulated:
        assert current.release["current_version"] == dev["version"]
        assert current.release["latest_version"] == ACTUAL
    else:
        assert current.comparison_version == ACTUAL


def test_disabling_dev_restores_normal_checks_without_restart(monkeypatch, caplog):
    caplog.set_level("INFO", logger="pastemd")
    current = session(monkeypatch, {"enabled": True, "version": "0.1.7.6"})
    check(current)
    assert current.state == "available"
    assert f"actual={ACTUAL}, comparison=0.1.7.6" in caplog.text
    current.config = {**current.config, "dev": {"enabled": False, "version": "0.1.7.6"}}
    check(current)
    assert current.state == "latest" and current.release is None
    assert not current.debug_version_active


def test_install_preparation_keeps_real_manifest_version(monkeypatch, tmp_path):
    current = session(monkeypatch, {"enabled": True, "version": "0.1.7.6"})
    check(current)
    received = []
    monkeypatch.setattr(update, "get_install_target", lambda: InstallTarget(tmp_path, "Windows"))
    monkeypatch.setattr(update, "select_asset", lambda data: select_asset(data, "Windows", "AMD64"))

    def download(asset, destination, cancel, progress):
        destination.write_bytes(b"package")
        progress(7, 7)
        return destination

    class Prepared:
        def __init__(self, target, package, work_dir, version, cancel):
            received.append((version, package.name))
            self.work_dir = work_dir

        def cleanup(self):
            shutil.rmtree(self.work_dir)

    monkeypatch.setattr(update, "download_asset", download)
    monkeypatch.setattr(update, "PreparedUpdate", Prepared)
    current.download()
    current.worker.join(timeout=2)
    assert not current.worker.is_alive()
    current.poll()
    assert current.state == "ready"
    assert received == [(ACTUAL, f"PasteMD_pandoc-Setup_v{ACTUAL}.exe")]
    current.shutdown()


def test_advanced_settings_preserve_other_dev_flags():
    from pastemd.presentation.settings.dialog import SettingsDialog

    settings = object.__new__(SettingsDialog)
    value = lambda x: SimpleNamespace(get=lambda: x)
    settings.excel_enable_var = value(True)
    settings.excel_format_var = value(True)
    settings.paste_delay_var = value("0.3")
    settings.dev_enabled_var = value(True)
    settings.dev_version_var = value(" 0.1.7.6 ")
    config = {"dev": {"enabled": False, "future_debug_flag": True}}
    settings._collect_advanced(config)
    assert config["dev"] == {"enabled": True, "version": "0.1.7.6", "future_debug_flag": True}


def test_invalid_debug_version_cannot_be_saved(monkeypatch):
    from pastemd.presentation.settings.dialog import SettingsDialog

    settings = object.__new__(SettingsDialog)
    settings.current_config = {"dev": {"enabled": True, "version": "bad"}}
    settings._confirm_keep_formula_enable = lambda: True
    settings._tab_specs, settings._tab_created = {}, set()
    saved, errors = [], []
    settings.config_loader = SimpleNamespace(save=lambda config: saved.append(config))
    settings._show_topmost_message = lambda title, message, kind: errors.append((message, kind))
    settings._on_save()
    assert saved == []
    assert len(errors) == 1 and errors[0][1] == "error"
