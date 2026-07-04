"""macOS Dock visibility and activation helpers.

This app is primarily a menu-bar (tray) utility. On macOS, we keep the Dock icon
hidden most of the time, and only show it while UI windows are open.
"""

from __future__ import annotations

import sys
from typing import Optional

from ...core.state import app_state
from ...config.paths import get_app_icon_path
from ..logging import log


def _get_refcount() -> int:
    return int(getattr(app_state, "macos_ui_refcount", 0) or 0)


def _set_refcount(value: int) -> None:
    setattr(app_state, "macos_ui_refcount", max(0, int(value)))


def _ns_app():
    try:
        from AppKit import NSApplication

        return NSApplication.sharedApplication()
    except Exception:
        return None


def set_dock_visible(visible: bool) -> None:
    """
    Toggle Dock icon visibility on macOS.

    Uses activation policy:
      - Regular: shows Dock icon
      - Accessory: hides Dock icon (still allows status bar item)
    """
    if sys.platform != "darwin":
        return

    app = _ns_app()
    if app is None:
        return

    try:
        from AppKit import (
            NSApplicationActivationPolicyAccessory,
            NSApplicationActivationPolicyRegular,
        )

        policy = (
            NSApplicationActivationPolicyRegular
            if visible
            else NSApplicationActivationPolicyAccessory
        )
        if visible:
            _set_application_icon(app)
        app.setActivationPolicy_(policy)
    except Exception as exc:
        log(f"Failed to set Dock visibility: {exc}")


def _set_application_icon(app) -> None:
    """Set the Dock/app-switcher icon to the full-color app icon."""
    try:
        if not hasattr(app, "setApplicationIconImage_"):
            return
        from AppKit import NSImage

        icon = NSImage.alloc().initWithContentsOfFile_(get_app_icon_path())
        if icon is not None:
            app.setApplicationIconImage_(icon)
    except Exception as exc:
        log(f"Failed to set Dock icon: {exc}")


def activate_app() -> None:
    """Bring app to foreground (best-effort)."""
    if sys.platform != "darwin":
        return

    app = _ns_app()
    if app is None:
        return

    try:
        app.activateIgnoringOtherApps_(True)
    except Exception as exc:
        log(f"Failed to activate app: {exc}")


def begin_ui_session() -> None:
    """
    Mark that a UI window is being shown.

    Increments a refcount and ensures Dock icon is visible while UI is open.
    """
    if sys.platform != "darwin":
        return

    refcount = _get_refcount() + 1
    _set_refcount(refcount)

    if refcount == 1:
        set_dock_visible(True)
        activate_app()


def end_ui_session() -> None:
    """
    Mark that a UI window was closed.

    Decrements refcount and hides Dock icon again when no UI windows remain.
    """
    if sys.platform != "darwin":
        return

    refcount = _get_refcount() - 1
    _set_refcount(refcount)

    if refcount <= 0:
        set_dock_visible(False)
