"""macOS permissions tab for settings dialog."""

from __future__ import annotations

import datetime as _dt
import subprocess
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional, Tuple

from ...i18n import t
from ...utils.logging import log


_Status = Optional[bool]


class MacOSPermissionsTab:
    """Build and manage the macOS permissions page."""

    def __init__(self, notebook: ttk.Notebook, root: tk.Tk):
        self.notebook = notebook
        self.root = root
        self.frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.frame, text=t("settings.tab.permissions"))

        self._items = []
        self._refresh_job = None
        self._refresh_interval_ms = 2000
        self._last_checked_var = tk.StringVar(value=t("settings.permissions.last_checked", time="--:--:--"))

        self._build_ui()
        self.refresh()
        self._schedule_refresh()

    def select(self) -> None:
        """Select this tab in the notebook."""
        try:
            self.notebook.select(self.frame)
        except Exception as exc:
            log(f"Failed to select permissions tab: {exc}")

    def refresh(self) -> None:
        """Refresh permissions status."""
        for item in self._items:
            status = self._safe_check(item["checker"])
            status_text, status_color = self._format_status(status)
            item["status_var"].set(status_text)
            try:
                item["status_label"].configure(foreground=status_color)
            except Exception:
                pass
            self._update_request_button(item, status)

        timestamp = _dt.datetime.now().strftime("%H:%M:%S")
        self._last_checked_var.set(t("settings.permissions.last_checked", time=timestamp))

    def _schedule_refresh(self) -> None:
        if not self._is_alive():
            return

        def _tick():
            if not self._is_alive():
                return
            try:
                self.refresh()
            except Exception as exc:
                log(f"Failed to refresh permissions status: {exc}")
            self._schedule_refresh()

        try:
            self._refresh_job = self.root.after(self._refresh_interval_ms, _tick)
        except Exception as exc:
            log(f"Failed to schedule permissions refresh: {exc}")

    def _build_ui(self) -> None:
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)

        canvas = tk.Canvas(self.frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        canvas.grid(row=0, column=0, sticky=tk.NSEW)

        content = ttk.Frame(canvas)
        self._content_frame = content

        window_id = canvas.create_window((0, 0), window=content, anchor="nw")

        def _on_content_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        content.bind("<Configure>", _on_content_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            self._scroll_canvas_if_alive(canvas, event)

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        content.columnconfigure(1, weight=1)

        intro = ttk.Label(
            content,
            text=t("settings.permissions.intro"),
            wraplength=520,
            justify=tk.LEFT,
        )
        intro.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        hint = ttk.Label(
            content,
            text=t("settings.permissions.add_hint"),
            wraplength=520,
            justify=tk.LEFT,
            foreground="gray",
        )
        hint.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        self._items = [
            self._make_item(
                key="accessibility",
                title=t("settings.permissions.accessibility.title"),
                desc=t("settings.permissions.accessibility.desc"),
                checker=self._check_accessibility,
                open_settings=self._open_accessibility_settings,
                request_access=self._request_accessibility,
            ),
            self._make_item(
                key="screen_recording",
                title=t("settings.permissions.screen_recording.title"),
                desc=t("settings.permissions.screen_recording.desc"),
                checker=self._check_screen_recording,
                open_settings=self._open_screen_recording_settings,
                request_access=self._request_screen_recording,
            ),
            self._make_item(
                key="input_monitoring",
                title=t("settings.permissions.input_monitoring.title"),
                desc=t("settings.permissions.input_monitoring.desc"),
                checker=self._check_input_monitoring,
                open_settings=self._open_input_monitoring_settings,
                request_access=self._request_input_monitoring,
            ),
            self._make_item(
                key="automation",
                title=t("settings.permissions.automation.title"),
                desc=t("settings.permissions.automation.desc"),
                checker=self._check_automation,
                open_settings=self._open_automation_settings,
                request_access=self._request_automation,
            ),
        ]

        for index, item in enumerate(self._items, start=1):
            row = index * 3
            title_label = ttk.Label(content, text=item["title"], font=("", 10, "bold"))
            title_label.grid(row=row, column=0, sticky=tk.W, pady=(8, 2))

            status_label = ttk.Label(content, textvariable=item["status_var"])
            status_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=(8, 2))
            item["status_label"] = status_label

            action_frame = ttk.Frame(content)
            action_frame.grid(row=row, column=2, sticky=tk.E, pady=(8, 2))
            action_btn = ttk.Button(
                action_frame,
                text=t("settings.permissions.open_settings"),
                command=item["open_settings"],
                width=12,
            )
            action_btn.pack(side=tk.RIGHT)
            if item.get("request_access"):
                request_btn = ttk.Button(
                    action_frame,
                    text=t("settings.permissions.request_access"),
                    command=item["request_access"],
                    width=10,
                )
                item["request_btn"] = request_btn

            desc_label = ttk.Label(
                content,
                text=item["desc"],
                wraplength=520,
                justify=tk.LEFT,
                foreground="gray",
            )
            desc_label.grid(row=row + 1, column=0, columnspan=3, sticky=tk.W, pady=(0, 6))

        refresh_btn = ttk.Button(
            content,
            text=t("settings.permissions.refresh"),
            command=self.refresh,
            width=12,
        )
        refresh_btn.grid(row=20, column=0, sticky=tk.W, pady=(10, 2))

        last_checked_label = ttk.Label(
            content,
            textvariable=self._last_checked_var,
            foreground="gray",
        )
        last_checked_label.grid(row=20, column=1, columnspan=2, sticky=tk.W, padx=(10, 0), pady=(10, 2))

    @staticmethod
    def _scroll_canvas_if_alive(canvas: tk.Canvas, event) -> None:
        """Ignore stale global mousewheel callbacks after the tab is destroyed."""
        try:
            if not canvas.winfo_exists():
                return
            delta = getattr(event, "delta", 0)
            if delta:
                canvas.yview_scroll(int(-delta / 120), "units")
        except tk.TclError:
            return

    def _make_item(
        self,
        *,
        key: str,
        title: str,
        desc: str,
        checker: Callable[[], _Status],
        open_settings: Callable[[], None],
        request_access: Optional[Callable[[], None]] = None,
    ) -> dict:
        return {
            "key": key,
            "title": title,
            "desc": desc,
            "checker": checker,
            "open_settings": open_settings,
            "request_access": request_access,
            "status_var": tk.StringVar(value=t("settings.permissions.status.checking")),
            "status_label": None,
            "request_btn": None,
        }

    def _format_status(self, status: _Status) -> Tuple[str, str]:
        if status is True:
            return t("settings.permissions.status.granted"), "#2a7b2e"
        if status is False:
            return t("settings.permissions.status.missing"), "#b32323"
        return t("settings.permissions.status.unknown"), "#666666"

    def _safe_check(self, checker: Callable[[], _Status]) -> _Status:
        try:
            return checker()
        except Exception as exc:
            log(f"Permission check failed: {exc}")
            return None

    def _update_request_button(self, item: dict, status: _Status) -> None:
        btn = item.get("request_btn")
        if not btn:
            return
        try:
            if status is False:
                if not btn.winfo_ismapped():
                    btn.pack(side=tk.RIGHT, padx=(0, 6))
            else:
                if btn.winfo_ismapped():
                    btn.pack_forget()
        except Exception as exc:
            log(f"Failed to update request button: {exc}")

    def _check_accessibility(self) -> _Status:
        # 1) Quartz 路线（最简单）
        try:
            import Quartz
            if hasattr(Quartz, "AXIsProcessTrustedWithOptions"):
                try:
                    # 常量通常就是这个
                    prompt_key = getattr(Quartz, "kAXTrustedCheckOptionPrompt", None)
                    options = {prompt_key: False} if prompt_key is not None else {}
                    return bool(Quartz.AXIsProcessTrustedWithOptions(options))
                except Exception as exc:
                    log(f"AXIsProcessTrustedWithOptions failed (Quartz): {exc}")

            if hasattr(Quartz, "AXIsProcessTrusted"):
                try:
                    return bool(Quartz.AXIsProcessTrusted())
                except Exception as exc:
                    log(f"AXIsProcessTrusted failed (Quartz): {exc}")

        except Exception as exc:
            # Quartz 不可用，不要当成“缺权限”
            log(f"Quartz not available for accessibility check: {exc}")

        # 2) ctypes fallback：直接调 ApplicationServices
        try:
            import ctypes
            app_services = ctypes.CDLL(
                "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices"
            )

            # Boolean AXIsProcessTrusted(void);
            app_services.AXIsProcessTrusted.restype = ctypes.c_bool
            app_services.AXIsProcessTrusted.argtypes = []
            return bool(app_services.AXIsProcessTrusted())
        except Exception as exc:
            log(f"AXIsProcessTrusted failed (ctypes): {exc}")
            return None

    def _check_automation(self) -> _Status:
        script = 'tell application "System Events" to get name of processes'
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=2,
            )
        except subprocess.TimeoutExpired:
            return False
        except Exception as exc:
            log(f"Automation check failed: {exc}")
            return None

        if result.returncode == 0:
            return True

        combined = " ".join([result.stdout or "", result.stderr or ""]).lower()
        if "not authorized" in combined or "not authorised" in combined or "not permitted" in combined:
            return False
        return None

    def _check_screen_recording(self) -> _Status:
        try:
            import Quartz
        except Exception as exc:
            log(f"Quartz not available for screen recording check: {exc}")
            return None

        if hasattr(Quartz, "CGPreflightScreenCaptureAccess"):
            try:
                return bool(Quartz.CGPreflightScreenCaptureAccess())
            except Exception as exc:
                log(f"CGPreflightScreenCaptureAccess failed: {exc}")
                return None
        return None

    def _check_input_monitoring(self) -> _Status:
        try:
            import Quartz
        except Exception as exc:
            log(f"Quartz not available for input monitoring check: {exc}")
            return None

        if hasattr(Quartz, "CGPreflightListenEventAccess"):
            try:
                return bool(Quartz.CGPreflightListenEventAccess())
            except Exception as exc:
                log(f"CGPreflightListenEventAccess failed: {exc}")
                return None
        return None

    def _open_accessibility_settings(self) -> None:
        self._open_system_settings("x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility")

    def _request_accessibility(self) -> None:
        try:
            import Quartz
        except Exception as exc:
            log(f"Quartz not available for accessibility request: {exc}")
            self._open_accessibility_settings()
            return

        try:
            prompt_key = getattr(Quartz, "kAXTrustedCheckOptionPrompt", None)
            if prompt_key is None and hasattr(Quartz, "kAXTrustedCheckOptionPrompt"):
                prompt_key = Quartz.kAXTrustedCheckOptionPrompt
            if hasattr(Quartz, "AXIsProcessTrustedWithOptions"):
                Quartz.AXIsProcessTrustedWithOptions({prompt_key: True})
            else:
                Quartz.AXIsProcessTrusted()
        except Exception as exc:
            log(f"Accessibility request failed: {exc}")
            self._open_accessibility_settings()
        else:
            self.root.after(1200, self.refresh)

    def _open_screen_recording_settings(self) -> None:
        self._open_system_settings("x-apple.systempreferences:com.apple.preference.security?Privacy_ScreenCapture")

    def _request_screen_recording(self) -> None:
        try:
            import Quartz
        except Exception as exc:
            log(f"Quartz not available for screen recording request: {exc}")
            self._open_screen_recording_settings()
            return

        try:
            if hasattr(Quartz, "CGRequestScreenCaptureAccess"):
                Quartz.CGRequestScreenCaptureAccess()
            else:
                self._open_screen_recording_settings()
        except Exception as exc:
            log(f"Screen recording request failed: {exc}")
            self._open_screen_recording_settings()
        else:
            self.root.after(1200, self.refresh)

    def _open_input_monitoring_settings(self) -> None:
        self._open_system_settings("x-apple.systempreferences:com.apple.preference.security?Privacy_ListenEvent")

    def _request_input_monitoring(self) -> None:
        try:
            import Quartz
        except Exception as exc:
            log(f"Quartz not available for input monitoring request: {exc}")
            self._open_input_monitoring_settings()
            return

        try:
            if hasattr(Quartz, "CGRequestListenEventAccess"):
                Quartz.CGRequestListenEventAccess()
            else:
                self._open_input_monitoring_settings()
        except Exception as exc:
            log(f"Input monitoring request failed: {exc}")
            self._open_input_monitoring_settings()
        else:
            self.root.after(1200, self.refresh)

    def _open_automation_settings(self) -> None:
        self._open_system_settings("x-apple.systempreferences:com.apple.preference.security?Privacy_Automation")

    def _request_automation(self) -> None:
        script = 'tell application "System Events" to get name of processes'
        try:
            subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=2,
            )
        except subprocess.TimeoutExpired:
            self._open_automation_settings()
        except Exception as exc:
            log(f"Automation request failed: {exc}")
            self._open_automation_settings()
        else:
            self.root.after(1200, self.refresh)

    def _open_system_settings(self, url: str) -> None:
        try:
            subprocess.run(["open", url], check=False)
        except Exception as exc:
            log(f"Failed to open System Settings: {exc}")

    def _is_alive(self) -> bool:
        try:
            return bool(self.root.winfo_exists())
        except Exception:
            return False
