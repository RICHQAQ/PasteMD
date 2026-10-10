"""Tray menu construction and callbacks."""

import os
import subprocess
import pystray
import threading
import time
import webbrowser
from typing import Optional

from ... import __version__
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.paths import get_log_path, get_config_path
from ...service.notification.manager import NotificationManager
from ...utils.fs import ensure_dir, open_dir, open_file
from ...utils.logging import log
from ...service.update import UpdateSession
from ..update.dialog import UpdateDialog, session_status
from ...utils.system_detect import is_windows, is_macos
from ...i18n import t, iter_languages, get_language, set_language, get_language_label, get_no_app_action_map
from .icon import create_status_icon
from ..hotkey.dialog import HotkeyDialog
from ..settings.dialog import SettingsDialog

try:
    if is_macos():
        from ...utils.macos.dock import begin_ui_session, end_ui_session, activate_app
    else:  # pragma: no cover
        begin_ui_session = end_ui_session = activate_app = lambda *args, **kwargs: None
except Exception:  # pragma: no cover
    begin_ui_session = end_ui_session = activate_app = lambda *args, **kwargs: None


class TrayMenuManager:
    """托盘菜单管理器"""
    
    def __init__(self, config_loader: ConfigLoader, notification_manager: NotificationManager):
        self.config_loader = config_loader
        self.notification_manager = notification_manager
        self.restart_hotkey_callback = None  # 将由外部设置
        self.pause_hotkey_callback = None  # 暂停热键监听
        self.resume_hotkey_callback = None  # 恢复热键监听
        self.update_dialog = None
        self.update_session = UpdateSession(app_state.config, self._refresh_update_ui, self._finish_update)
        self._updates_started = False
        self._quitting = False
        self._last_update_state = "idle"
        self._last_tray_refresh = 0.0
        self.hotkey_dialog = None
        self.settings_dialog = None
    
    def set_restart_hotkey_callback(self, callback):
        """设置重启热键的回调函数"""
        self.restart_hotkey_callback = callback
    
    def set_pause_hotkey_callback(self, callback):
        """设置暂停热键的回调函数"""
        self.pause_hotkey_callback = callback
    
    def set_resume_hotkey_callback(self, callback):
        """设置恢复热键的回调函数"""
        self.resume_hotkey_callback = callback
    
    def build_menu(self) -> pystray.Menu:
        """构建托盘菜单"""
        config = app_state.config

        # 构建常规菜单项
        normal_menu_items = [
            pystray.MenuItem(
                t("tray.menu.hotkey_display", hotkey=app_state.config['hotkey']),
                lambda icon, item: None,
                enabled=False
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                t("tray.menu.enable_hotkey"),
                self._on_toggle_enabled,
                checked=lambda item: app_state.enabled
            ),
            pystray.MenuItem(
                t("tray.menu.show_notifications"),
                self._on_toggle_notify,
                checked=lambda item: config.get("notify", True)
            )
        ]

        if is_windows():
            normal_menu_items.append(
                pystray.MenuItem(
                    t("tray.menu.move_cursor"),
                    self._on_toggle_move_cursor,
                    checked=lambda item: config.get("move_cursor_to_end", True)
                )
            )

        # 构建版本菜单项
        version_menu_items = [
            pystray.MenuItem(
                t("tray.menu.current_version", version=__version__),
                lambda icon, item: None,
                enabled=False
            ),
        ]
        session = self.update_session
        if session.state != "idle":
            version_menu_items.append(pystray.MenuItem(
                lambda item: self._update_menu_text(), self._on_show_update,
            ))
        version_menu_items.append(pystray.MenuItem(
            t("tray.menu.check_update"), self._on_check_update,
            enabled=not session.busy and session.prepared is None,
        ))
        if session.state in {"downloading", "preparing", "ready"}:
            version_menu_items.append(pystray.MenuItem(t("update.cancel"), self._on_cancel_update))

        return pystray.Menu(
            *normal_menu_items,
            self._build_no_app_action_menu(),
            self._build_html_formatting_menu(),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("tray.menu.set_hotkey"), self._on_set_hotkey),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                t("tray.menu.keep_file"),
                self._on_toggle_keep,
                checked=lambda item: config.get("keep_file", False)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("tray.menu.open_save_dir"), self._on_open_save_dir),
            pystray.MenuItem(t("tray.menu.open_log"), self._on_open_log),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("settings.dialog.title"), self._on_open_settings),
            pystray.Menu.SEPARATOR,
            *version_menu_items,
            pystray.MenuItem(
                t("tray.menu.about"),
                self._on_open_about_page
            ),
            pystray.MenuItem(t("tray.menu.quit"), self._on_quit)
        )

    # 菜单回调函数
    def _on_toggle_enabled(self, icon, item):
        """切换热键启用状态"""
        app_state.enabled = not app_state.enabled
        icon.icon = create_status_icon(ok=app_state.enabled)
        
        status = t("tray.status.hotkey_enabled") if app_state.enabled else t("tray.status.hotkey_paused")
        icon.menu = self.build_menu()
        self.notification_manager.notify("PasteMD", status, ok=app_state.enabled)
    
    def _on_set_hotkey(self, icon, item):
        """设置热键，确保 Tk 对话框在主线程 UI 队列中运行"""
        def save_hotkey(new_hotkey: str):
            """保存新热键并重启热键绑定"""
            try:
                # 更新配置
                app_state.config["hotkey"] = new_hotkey
                app_state.hotkey_str = new_hotkey
                self._save_config()
                
                # 重启热键绑定
                def _restart():
                    if self.restart_hotkey_callback:
                        self.restart_hotkey_callback()
                try:
                    if app_state.root and app_state.root.winfo_exists():
                        app_state.root.after(200, _restart)
                    else:
                        _restart()
                except Exception as e:
                    log(f"Failed to schedule hotkey restart: {e}")
                    _restart()
                
                # 刷新菜单
                try:
                    icon.menu = self.build_menu()
                    if hasattr(icon, "update_menu"):
                        icon.update_menu()
                except Exception as e:
                    log(f"Failed to refresh tray menu after hotkey save: {e}")

                # 如果设置窗口正在打开，实时刷新里面的热键展示
                try:
                    if self.settings_dialog and self.settings_dialog.is_alive():
                        self.settings_dialog.refresh_hotkey_display()
                except Exception as e:
                    log(f"Failed to refresh settings hotkey display: {e}")
                
                log(f"Hotkey changed to: {new_hotkey}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.status.hotkey_saved", hotkey=new_hotkey),
                    ok=True)
            except Exception as e:
                log(f"Failed to save hotkey: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.error.hotkey_save_failed", error=str(e)),
                    ok=False)
                raise

        def show_dialog_on_main():
            """在主线程打开/销毁 Tk 对话框"""
            dialog_shown = False
            previous_block = bool(getattr(app_state, "ui_block_hotkeys", False))
            try:
                if self.hotkey_dialog and self.hotkey_dialog.is_alive():
                    activate_app()
                    self.hotkey_dialog.restore_and_focus()
                    return
                
                # 打开弹窗期间屏蔽热键触发（macOS 下不做 stop/start，避免崩溃）
                app_state.ui_block_hotkeys = True
                if is_windows() and self.pause_hotkey_callback:
                    self.pause_hotkey_callback()

                begin_ui_session()

                def _on_dialog_close():
                    end_ui_session()
                    if is_windows() and self.resume_hotkey_callback:
                        self.resume_hotkey_callback()

                dialog = HotkeyDialog(
                    current_hotkey=app_state.hotkey_str,
                    on_save=save_hotkey,
                    on_close=_on_dialog_close,
                )
                self.hotkey_dialog = dialog

                def _clear_dialog_ref(event=None):
                    if getattr(event, "widget", None) is dialog.root or event is None:
                        self.hotkey_dialog = None
                        app_state.ui_block_hotkeys = previous_block

                dialog.root.bind("<Destroy>", _clear_dialog_ref)
                dialog.show()
                dialog_shown = True
            except Exception as e:
                log(f"Failed to show hotkey dialog: {e}")
                self.notification_manager.notify("PasteMD", t("tray.error.open_hotkey_dialog", error=str(e)), ok=False)
            finally:
                if not dialog_shown:
                    app_state.ui_block_hotkeys = previous_block
                    end_ui_session()
                    if is_windows() and self.resume_hotkey_callback:
                        self.resume_hotkey_callback()

        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(show_dialog_on_main)
        else:
            # 兜底：未获取到 UI 队列时，仍在当前线程执行
            show_dialog_on_main()
    
    def _on_toggle_notify(self, icon, item):
        """切换通知状态"""
        current = app_state.config.get("notify", True)
        app_state.config["notify"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        if app_state.config["notify"]:
            self.notification_manager.notify("PasteMD", t("tray.status.notifications_enabled"), ok=True)
        else:
            log("Notifications disabled via tray toggle")
    
    def _on_toggle_move_cursor(self, icon, item):
        """切换插入后光标移动到末尾状态"""
        current = app_state.config.get("move_cursor_to_end", True)
        app_state.config["move_cursor_to_end"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.move_cursor_on") if app_state.config["move_cursor_to_end"] else t("tray.status.move_cursor_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
        
    def _on_toggle_excel(self, icon, item):
        """切换启用 Excel 插入"""
        current = app_state.config.get("enable_excel", True)
        app_state.config["enable_excel"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.excel_insert_on") if app_state.config["enable_excel"] else t("tray.status.excel_insert_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
        
    def _on_toggle_excel_format(self, icon, item):
        """切换 Excel 粘贴时是否保留格式"""
        current = app_state.config.get("excel_keep_format", True)
        app_state.config["excel_keep_format"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.excel_format_on") if app_state.config["excel_keep_format"] else t("tray.status.excel_format_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_toggle_keep(self, icon, item):
        """切换保留文件状态"""
        current = app_state.config.get("keep_file", False)
        app_state.config["keep_file"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.keep_file_on") if app_state.config["keep_file"] else t("tray.status.keep_file_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_open_save_dir(self, icon, item):
        """打开保存目录"""
        save_dir = app_state.config.get("save_dir", "")
        save_dir = os.path.expandvars(save_dir)
        ensure_dir(save_dir)
        open_dir(save_dir)
    
    def _on_open_log(self, icon, item):
        """打开日志文件"""
        log_path = get_log_path()
        if not os.path.exists(log_path):
            # 创建空日志文件
            open(log_path, "w", encoding="utf-8").close()
        open_file(log_path)
    
    def _on_open_settings(self, icon, item):
        """打开设置界面"""
        self._open_settings(icon, item, None)

    def open_settings_tab(self, tab_key: Optional[str]) -> None:
        """Open settings and select a specific tab."""
        self._open_settings(getattr(app_state, "icon", None), None, tab_key)

    def _open_settings(self, icon, item, select_tab: Optional[str]) -> None:
        """打开设置界面"""
        def on_settings_save():
            """设置保存后的回调"""
            # 刷新菜单以反映可能的配置更改（如语言）
            set_language(app_state.config.get("language", "en-US"))
            try:
                tray_icon = icon or getattr(app_state, "icon", None)
                if tray_icon is not None:
                    tray_icon.menu = self.build_menu()
                    if hasattr(tray_icon, "update_menu"):
                        tray_icon.update_menu()
            except Exception as e:
                log(f"Failed to refresh tray menu after settings save: {e}")

        def show_dialog_on_main():
            """在主线程显示设置对话框"""
            try:
                if self.settings_dialog and self.settings_dialog.is_alive():
                    activate_app()
                    if select_tab:
                        try:
                            self.settings_dialog.select_tab(select_tab)
                        except Exception as e:
                            log(f"Failed to select settings tab: {e}")
                    self.settings_dialog.restore_and_focus()
                    return

                begin_ui_session()

                def _on_dialog_close():
                    end_ui_session()
                    if self.resume_hotkey_callback:
                        self.resume_hotkey_callback()

                dialog = SettingsDialog(
                    on_save=on_settings_save,
                    on_close=_on_dialog_close,
                    initial_tab=select_tab,
                )
                self.settings_dialog = dialog

                try:
                    dialog.set_open_hotkey_dialog(
                        lambda: self._on_set_hotkey(icon, None)
                    )
                except Exception:
                    pass

                def _clear_settings_dialog(event=None):
                    if getattr(event, "widget", None) is dialog.root or event is None:
                        self.settings_dialog = None

                dialog.root.bind("<Destroy>", _clear_settings_dialog)
                dialog.show()
            except Exception as e:
                log(f"Failed to show settings dialog: {e}")
                self.notification_manager.notify("PasteMD", f"Error opening settings: {e}", ok=False)
                end_ui_session()

        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(show_dialog_on_main)
        else:
            show_dialog_on_main()

    def _build_html_formatting_menu(self) -> pystray.MenuItem:
        """构建 HTML 格式化子菜单"""
        return pystray.MenuItem(
            t("tray.menu.html_formatting"),
            pystray.Menu(
                pystray.MenuItem(
                    t("tray.menu.strikethrough_to_del"),
                    self._on_toggle_html_strikethrough,
                    checked=lambda item: self._get_html_formatting_option("strikethrough_to_del", True),
                ),
            ),
        )

    def _build_no_app_action_menu(self) -> pystray.MenuItem:
        """构建无应用时动作子菜单"""
        return pystray.MenuItem(
            t("tray.menu.no_app_action"),
            pystray.Menu(
                pystray.MenuItem(
                    t("action.open"),
                    self._on_set_no_app_action,
                    checked=lambda item: self._get_no_app_action() == "open",
                ),
                pystray.MenuItem(
                    t("action.save"),
                    self._on_set_no_app_action,
                    checked=lambda item: self._get_no_app_action() == "save",
                ),
                pystray.MenuItem(
                    t("action.clipboard"),
                    self._on_set_no_app_action,
                    checked=lambda item: self._get_no_app_action() == "clipboard",
                ),
                pystray.MenuItem(
                    t("action.none"),
                    self._on_set_no_app_action,
                    checked=lambda item: self._get_no_app_action() == "none",
                ),
            ),
        )

    def _get_no_app_action(self) -> str:
        """获取当前无应用动作设置"""
        return app_state.config.get("no_app_action", "open")

    def _on_set_no_app_action(self, icon, item):
        """设置无应用动作"""
        # 从标签文本映射到配置值（反转 action_map）
        action_map = get_no_app_action_map()
        reverse_action_map = {v: k for k, v in action_map.items()}
        
        # 获取点击的菜单项的文本
        clicked_text = getattr(item, 'text', '')
        
        # 设置新的动作
        new_action = reverse_action_map.get(clicked_text, "open")
        app_state.config["no_app_action"] = new_action
        self._save_config()
        icon.menu = self.build_menu()
        
        # 显示状态通知
        status_map = {
            "open": t("tray.status.no_app_action_open"),
            "save": t("tray.status.no_app_action_save"),
            "clipboard": t("tray.status.no_app_action_clipboard"),
            "none": t("tray.status.no_app_action_none")
        }
        
        status = status_map.get(new_action, "")
        self.notification_manager.notify("PasteMD", status, ok=True)

    def _get_html_formatting_option(self, key: str, default: bool) -> bool:
        options = app_state.config.get("html_formatting", {})
        if isinstance(options, dict):
            return bool(options.get(key, default))
        return default

    def _on_toggle_html_strikethrough(self, icon, item):
        """切换删除线转 <del> 的 HTML 格式化配置"""
        current = self._get_html_formatting_option("strikethrough_to_del", True)
        if not isinstance(app_state.config.get("html_formatting"), dict):
            app_state.config["html_formatting"] = {}
        app_state.config["html_formatting"]["strikethrough_to_del"] = not current
        self._save_config()
        icon.menu = self.build_menu()

        status = (
            t("tray.status.html_strike_on")
            if app_state.config["html_formatting"].get("strikethrough_to_del", True)
            else t("tray.status.html_strike_off")
        )
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _queue_update_ui(self, callback):
        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(callback)

    def start_updates(self):
        """Called once on the main thread after Tk has been initialized."""
        if self._updates_started:
            return
        self._updates_started = True

        def poll():
            if self._quitting:
                return
            self.update_session.poll()
            app_state.root.after(250, poll)

        app_state.root.after(250, poll)
        self.update_session.check(manual=False)

    def _update_menu_text(self):
        session = self.update_session
        if session.state == "available" and session.release:
            return t("tray.menu.new_version", version=session.release["latest_version"])
        if session.state == "failed":
            return t("update.tray.failed")
        return session_status(session)

    def _refresh_update_ui(self):
        session = self.update_session
        changed_state = self._last_update_state != session.state
        now = time.monotonic()
        icon = getattr(app_state, "icon", None)
        if icon and (changed_state or now - self._last_tray_refresh >= 1.0):
            self._last_tray_refresh = now
            try:
                icon.menu = self.build_menu()
                icon.update_menu()
            except Exception as exc:
                log(f"[update] Tray refresh failed: {exc}")
        if self.update_dialog:
            self.update_dialog.render()
        self._last_update_state = session.state
        if changed_state and session.state == "failed":
            app_state.ui_block_hotkeys = False
        if changed_state and session.state in {"available", "latest", "failed", "ready"}:
            if session.manual_check and session.state in {"available", "latest", "failed"}:
                self._show_update_dialog()
            if session.state == "ready":
                self.notification_manager.notify("PasteMD", t("update.status.ready"), ok=True)
            elif session.state == "failed" and session.manual_check:
                self.notification_manager.notify("PasteMD", t("update.error." + session.error_key), ok=False)

    def _on_check_update(self, icon, item):
        self._queue_update_ui(lambda: self.update_session.check(manual=True))

    def _on_show_update(self, icon, item):
        def show():
            self.update_session.manual_check = True
            self._show_update_dialog()
        self._queue_update_ui(show)

    def _on_cancel_update(self, icon, item):
        self._queue_update_ui(self.update_session.cancel)

    def _show_update_dialog(self):
        if self._quitting:
            return
        if self.update_dialog:
            activate_app()
            self.update_dialog.focus()
            return
        begin_ui_session()

        def closed():
            self.update_dialog = None
            end_ui_session()

        try:
            self.update_dialog = UpdateDialog(app_state.root, self.update_session, closed, self._install_update)
            activate_app()
        except Exception as exc:
            end_ui_session()
            log(f"[update] Failed to open update window: {exc}")

    def _install_update(self):
        if app_state.is_running():
            from tkinter import messagebox
            messagebox.showinfo("PasteMD", t("update.wait_for_task"), parent=self.update_dialog.root)
            return
        app_state.ui_block_hotkeys = True
        self.update_session.install()

    def _finish_update(self):
        self._on_quit(getattr(app_state, "icon", None), None)

    def _on_open_about_page(self, icon, item):
        """打开关于页面"""
        # macOS 使用专门的介绍页面
        if is_macos():
            about_url = "https://pastemd.richqaq.cn/macos"
        else:
            about_url = "http://pastemd.richqaq.cn"
        try:
            webbrowser.open(about_url)
            log(f"Opening about page: {about_url}")
        except Exception as e:
            log(f"Failed to open browser: {e}")
            self.notification_manager.notify(
                "PasteMD",
                t("tray.error.open_about_page", url=about_url),
                ok=False
            )
    
    def _on_quit(self, icon, item):
        """Keep the UI alive while cancellation releases a download/mounted DMG."""
        self._queue_update_ui(lambda: self._quit_when_update_stops(icon))

    def _quit_when_update_stops(self, icon):
        self._quitting = True
        session = self.update_session
        session.cancel_event.set()
        if session.worker and session.worker.is_alive():
            app_state.root.after(100, lambda: self._quit_when_update_stops(icon))
            return
        session.shutdown()
        if icon:
            icon.stop()
        if app_state.quit_event is None:
            app_state.quit_event = threading.Event()
        app_state.quit_event.set()

    def _save_config(self):
        """保存配置"""
        try:
            self.config_loader.save(app_state.config)
        except Exception as e:
            log(f"Failed to save config: {e}")
