"""A non-modal window backed by the persistent tray update session."""

import tkinter as tk
from tkinter import ttk

from ... import __version__
from ...i18n import t
from ...utils.fs import open_dir
from ...config.paths import get_log_dir, get_log_path
from ...service.update import UpdateSession


def format_bytes(value: float) -> str:
    if value >= 1024 * 1024:
        return f"{value / (1024 * 1024):.1f} MB"
    return f"{value / 1024:.0f} KB"


def session_status(session: UpdateSession) -> str:
    if session.state == "downloading":
        percent = int(session.received / session.total * 100) if session.total else 0
        return t("update.status.downloading", percent=percent,
                 received=format_bytes(session.received), total=format_bytes(session.total),
                 speed=format_bytes(session.speed))
    if session.state == "failed":
        return t("update.error." + session.error_key)
    return t("update.status." + session.state)


class UpdateDialog:
    def __init__(self, parent, session: UpdateSession, on_close, on_install):
        self.session, self.on_close, self.on_install = session, on_close, on_install
        self.root = tk.Toplevel(parent)
        self.root.title(t("update.title"))
        self.root.geometry("580x440")
        self.root.minsize(460, 360)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<Escape>", lambda _event: self.close())
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)
        self.heading = ttk.Label(frame, font=("TkDefaultFont", 15, "bold"))
        self.heading.pack(anchor="w")
        self.current_version_label = ttk.Label(frame, text=t("update.current_version", version=__version__))
        self.current_version_label.pack(anchor="w", pady=(5, 10))
        self.debug_version_label = ttk.Label(frame, foreground="gray", wraplength=520)
        notes_frame = ttk.Frame(frame)
        notes_frame.pack(fill="both", expand=True)
        self.notes = tk.Text(notes_frame, wrap="word", height=8, relief="flat", padx=10, pady=10)
        scrollbar = ttk.Scrollbar(notes_frame, command=self.notes.yview)
        self.notes.configure(yscrollcommand=scrollbar.set)
        self.notes.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._notes_version = None
        self.status = ttk.Label(frame, wraplength=520)
        self.status.pack(anchor="w", pady=(12, 6))
        self.progress = ttk.Progressbar(frame, maximum=100)
        self.progress.pack(fill="x")
        ttk.Label(frame, text=t("update.install_hint"), wraplength=520).pack(anchor="w", pady=8)
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(4, 0))
        ttk.Button(buttons, text=t("update.open_logs"), command=self._open_logs).pack(side="left")
        self.primary = ttk.Button(buttons, command=self._primary)
        self.primary.pack(side="right")
        self.secondary = ttk.Button(buttons, command=self._secondary)
        self.secondary.pack(side="right", padx=8)
        self.render()
        self.focus()

    def _open_logs(self):
        get_log_path()  # ensure the directory exists
        open_dir(get_log_dir())

    def focus(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def render(self):
        session = self.session
        if session.debug_version_active:
            self.debug_version_label.configure(text=t("update.debug_version", version=session.comparison_version))
            self.debug_version_label.pack(after=self.current_version_label, anchor="w", pady=(0, 8))
        else:
            self.debug_version_label.pack_forget()
        version = session.release["latest_version"] if session.release else None
        self.heading.configure(text=t("update.new_version", version=version) if version else t("update.title"))
        if self._notes_version != version or self._notes_version is None:
            self.notes.configure(state="normal")
            self.notes.delete("1.0", "end")
            self.notes.insert("1.0", (session.release or {}).get("release_notes") or t("update.no_notes"))
            self.notes.configure(state="disabled")
            self._notes_version = version
        self.status.configure(text=session_status(session))
        self.progress["value"] = session.received / session.total * 100 if session.total else 0
        installing = session.state == "installing"
        ready = session.state == "ready"
        self.primary.configure(text=t("update.install" if ready else "update.download" if version else "tray.menu.check_update"),
                               state="disabled" if session.busy else "normal")
        cancellable = session.state in {"downloading", "preparing", "ready"}
        self.secondary.configure(text=t("update.cancel" if cancellable else "update.close"),
                                 state="disabled" if installing else "normal")

    def _primary(self):
        if self.session.state == "ready":
            self.on_install()
        elif self.session.release:
            self.session.download()
        else:
            self.session.check()

    def _secondary(self):
        if self.session.state in {"downloading", "preparing", "ready"}:
            self.session.cancel()
        else:
            self.close()

    def close(self):
        if self.session.state == "installing":
            return
        self.root.destroy()
        self.on_close()
