"""One update session shared by the tray and update window.

Workers only enqueue results; poll() applies state and calls the UI on the main
thread. Closing a window leaves the download running and visible in the tray.
"""

from __future__ import annotations

from pathlib import Path
import queue
import shutil
import tempfile
import threading
import time
import traceback
from typing import Callable

from .. import __version__
from ..utils.logging import log
from ..utils.update_manifest import DEFAULT_MANIFEST_URL, DEFAULT_UPDATE_CHANNEL
from ..utils.version_checker import VersionChecker
from ..utils.updater import (PreparedUpdate, UpdateCancelled, UpdateError,
                             download_asset, get_install_target, refresh_release_assets,
                             select_asset)


class UpdateSession:
    BUSY = {"checking", "downloading", "preparing", "installing", "cancelling"}

    def __init__(self, config: dict, changed: Callable[[], None], installed: Callable[[], None]):
        self.config, self.changed, self.installed = config, changed, installed
        self.state = "idle"
        self.release: dict | None = None
        self.received = self.total = 0
        self.speed = 0.0
        self.error_key = ""
        self.prepared: PreparedUpdate | None = None
        self.cancel_event = threading.Event()
        self.events: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.manual_check = False

    @property
    def busy(self) -> bool:
        return self.state in self.BUSY

    def _set_state(self, state: str) -> None:
        self.state = state
        log(f"[update] State: {state}")
        self.changed()

    def _start(self, state: str, operation: Callable[[], None]) -> None:
        self.error_key = ""
        self.cancel_event.clear()
        self._set_state(state)

        def run():
            try:
                operation()
            except UpdateCancelled:
                log("[update] Download/preparation cancelled")
                self.events.put(("cancelled", None))
            except UpdateError as exc:
                log(f"[update] {exc.key}\n{traceback.format_exc()}")
                self.events.put(("failed", exc.key))
            except Exception:
                log(f"[update] Unexpected failure\n{traceback.format_exc()}")
                self.events.put(("failed", "unexpected_error"))

        self.worker = threading.Thread(target=run, name="PasteMD-Update", daemon=True)
        self.worker.start()

    def check(self, manual: bool = True) -> None:
        if self.busy or self.prepared:
            return
        self.manual_check = manual

        def operation():
            checker = VersionChecker(__version__, self.config.get("update_manifest_url", DEFAULT_MANIFEST_URL),
                                     self.config.get("update_channel", DEFAULT_UPDATE_CHANNEL))
            result = checker.check_update()
            if result is None:
                raise UpdateError("check_failed")
            self.events.put(("checked", result))

        self._start("checking", operation)

    def download(self) -> None:
        if self.busy or not self.release or self.prepared:
            return
        self.manual_check = True
        release = self.release.copy()
        self.received = self.total = 0
        self.speed = 0.0

        def operation():
            target = get_install_target()
            asset = select_asset(refresh_release_assets(release))
            work_dir = Path(tempfile.mkdtemp(prefix="pastemd-update-"))
            prepared = None
            handed_to_ui = False
            try:
                if shutil.disk_usage(work_dir).free < asset.size + 64 * 1024 * 1024:
                    raise UpdateError("no_space")
                last_emit = 0.0
                start = time.monotonic()

                def progress(received: int, total: int):
                    nonlocal last_emit, start
                    now = time.monotonic()
                    if received == 0:
                        start = now
                    if received == 0 or received == total or now - last_emit >= 0.3:
                        speed = received / max(now - start, 0.1)
                        self.events.put(("progress", (received, total, speed)))
                        last_emit = now

                package = download_asset(asset, work_dir / asset.name, self.cancel_event, progress)
                self.events.put(("preparing", None))
                prepared = PreparedUpdate(target, package, work_dir, release["latest_version"], self.cancel_event)
                if self.cancel_event.is_set():
                    raise UpdateCancelled()
                self.events.put(("ready", prepared))
                handed_to_ui = True
            finally:
                if not handed_to_ui:
                    if prepared:
                        prepared.cleanup()
                    else:
                        shutil.rmtree(work_dir, ignore_errors=True)

        self._start("downloading", operation)

    def cancel(self) -> None:
        if self.state in {"downloading", "preparing"}:
            self.cancel_event.set()
            self._set_state("cancelling")
        elif self.prepared and self.state != "installing":
            self.prepared.cleanup()
            self.prepared = None
            self._set_state("cancelled")

    def install(self) -> None:
        if self.prepared is None or self.busy:
            return
        prepared = self.prepared

        def operation():
            prepared.launch()
            self.events.put(("installed", None))

        self._start("installing", operation)

    def poll(self) -> None:
        """Call from the Tk/main thread, also for headless session tests."""
        while True:
            try:
                event, value = self.events.get_nowait()
            except queue.Empty:
                break
            if event == "progress":
                self.received, self.total, self.speed = value
                self.changed()
            elif event == "checked":
                if value.get("has_update"):
                    self.release = value
                    self._set_state("available")
                else:
                    self.release = None
                    self._set_state("latest")
            elif event == "ready":
                # A cancel can arrive after the worker's final check, before poll.
                if self.cancel_event.is_set():
                    value.cleanup()
                    self._set_state("cancelled")
                else:
                    self.prepared = value
                    self._set_state("ready")
            elif event == "failed":
                if self.prepared:
                    self.prepared.cleanup()
                    self.prepared = None
                self.error_key = value
                self._set_state("failed")
            elif event == "preparing":
                if not self.cancel_event.is_set():
                    self._set_state("preparing")
            elif event == "installed":
                self.installed()
            else:
                self._set_state(event)

    def shutdown(self) -> None:
        """Wait before application exit so mounted/staged resources can be cleaned."""
        self.cancel_event.set()
        if self.worker and self.worker.is_alive():
            self.worker.join()
        self.poll()
        if self.prepared:
            self.prepared.cleanup()
            self.prepared = None
