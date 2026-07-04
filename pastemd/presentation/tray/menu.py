"""Tray menu construction and callbacks."""

import os
import subprocess
import pystray
import threading
import webbrowser
from typing import Optional

from ... import __version__
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.paths import get_log_path, get_config_path
from ...service.notification.manager import NotificationManager
from ...utils.fs import ensure_dir, open_dir, open_file
from ...utils.logging import log
from ...utils.version_checker import VersionChecker
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

    def __init__(self, config_loader: ConfigLoader, notification_manager: NotificationManager, history_manager=None):
        self.config_loader = config_loader
        self.notification_manager = notification_manager
        self.history_manager = history_manager
        self.restart_hotkey_callback = None  # 将由外部设置
        self.pause_hotkey_callback = None  # 暂停热键监听
        self.resume_hotkey_callback = None  # 恢复热键监听
        self.version_checker = None  # 将由外部设置或按需创建
        self.latest_version = None  # 存储最新版本号
        self.latest_release_url = None  # 存储最新版本的下载链接
        self.hotkey_dialog = None
        self.settings_dialog = None
        self.history_dialog = None
    
    def set_restart_hotkey_callback(self, callback):
        """设置重启热键的回调函数"""
        self.restart_hotkey_callback = callback
    
    def set_pause_hotkey_callback(self, callback):
        """设置暂停热键的回调函数"""
        self.pause_hotkey_callback = callback
    
    def set_resume_hotkey_callback(self, callback):
        """设置恢复热键的回调函数"""
        self.resume_hotkey_callback = callback

    def handle_app_reopen(self, icon=None) -> None:
        """Handle macOS app reopen/second-launch events."""
        if self._focus_existing_dialog():
            return
        self._on_open_settings(icon or getattr(app_state, "icon", None), None)

    def _focus_existing_dialog(self) -> bool:
        for attr_name in ("hotkey_dialog", "history_dialog", "settings_dialog"):
            dialog = getattr(self, attr_name, None)
            if dialog is None:
                continue
            if not self._dialog_is_alive(dialog):
                setattr(self, attr_name, None)
                continue
            if self._focus_dialog(dialog):
                return True
        return False

    def _dialog_is_alive(self, dialog) -> bool:
        is_alive = getattr(dialog, "is_alive", None)
        if callable(is_alive):
            try:
                return bool(is_alive())
            except Exception:
                return False
        return True

    def _focus_dialog(self, dialog) -> bool:
        try:
            activate_app()
            restore_and_focus = getattr(dialog, "restore_and_focus", None)
            if callable(restore_and_focus):
                restore_and_focus()
                return True

            lift = getattr(dialog, "lift", None)
            if callable(lift):
                lift()
            focus_force = getattr(dialog, "focus_force", None)
            if callable(focus_force):
                focus_force()
            return callable(lift) or callable(focus_force)
        except Exception:
            return False
    
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
        if self.latest_version:
            version_menu_items.append(
                pystray.MenuItem(
                    t("tray.menu.new_version", version=self.latest_version),
                    self._on_open_release_page,
                    enabled=True
                )
            )
        else:
            version_menu_items.append(
                pystray.MenuItem(
                    t("tray.menu.check_update"),
                    self._on_check_update
                )
            )

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
            pystray.MenuItem(t("tray.menu.paste_history"), self._on_open_history),
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

    def _on_open_history(self, icon, item):
        """打开历史记录界面"""
        if self.history_dialog is not None:
            # Already open - bring to front
            try:
                self.history_dialog.lift()
                self.history_dialog.focus_force()
            except Exception:
                self.history_dialog = None
            return
        self._open_history_dialog()

    def _open_history_dialog(self):
        """在主线程创建并显示历史记录对话框"""
        from ...core.state import app_state
        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(self._do_show_history_dialog)
            # macOS: 弹窗期间需要显示 Dock 图标
            if is_macos():
                try:
                    from ...utils.macos.dock import begin_ui_session
                    begin_ui_session()
                except Exception:
                    pass

    def _do_show_history_dialog(self):
        """在主线程中显示历史记录对话框"""
        try:
            from ..history.dialog import HistoryDialog
            self.history_dialog = HistoryDialog(
                self.history_manager,
                self.notification_manager,
                on_close=self._on_history_dialog_closed,
            )
            self.history_dialog.show()
        except Exception as e:
            log(f"Failed to open history dialog: {e}")
            if is_macos():
                try:
                    from ...utils.macos.dock import end_ui_session
                    end_ui_session()
                except Exception:
                    pass

    def _on_history_dialog_closed(self):
        """历史对话框关闭回调"""
        self.history_dialog = None
        if is_macos():
            try:
                from ...utils.macos.dock import end_ui_session
                end_ui_session()
            except Exception:
                pass

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
    
    def _on_check_update(self, icon, item):
        """检查更新"""
        # 在后台线程中检查更新，避免阻塞 UI
        def check_in_background():
            try:
                # 导入版本号
                from ... import __version__
                
                checker = VersionChecker(__version__)
                result = checker.check_update()
                
                if result is None:
                    # 网络错误或检查失败
                    log("Version check failed - network error")
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_failure')}",
                        t("tray.update.network_error"),
                        ok=False
                    )
                elif result.get("has_update"):
                    latest_version = result.get("latest_version")
                    release_url = result.get("release_url")
                    
                    # 使用 update_version_info 方法更新版本信息并重新绘制菜单
                    self.update_version_info(icon, latest_version, release_url)
                    
                    # 通知用户有新版本，并自动打开下载页面
                    message = t("tray.update.opening_release", version=latest_version)
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_new_version')}",
                        message,
                        ok=True
                    )
                    
                    # 自动打开下载页面
                    try:
                        webbrowser.open(release_url)
                    except Exception as e:
                        log(f"Failed to open browser: {e}")
                    
                    log(f"New version available: {latest_version}")
                    log(f"Download URL: {release_url}")
                else:
                    # 无需更新，通知用户已是最新版本
                    current_version = result.get("current_version")
                    log(f"Already on latest version: {current_version}")
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_latest')}",
                        t("tray.update.latest_version", version=current_version),
                        ok=True
                    )
            except Exception as e:
                error_text = str(e)
                short_error = error_text if len(error_text) <= 15 else error_text[:12] + "..."
                self.notification_manager.notify(
                    f"PasteMD - {t('tray.update.title_unexpected_error')}",
                    t("tray.update.error_with_message", error=short_error),
                    ok=False
                )
                log(f"Error checking update: {e}")
        
        # 启动后台线程
        thread = threading.Thread(target=check_in_background, daemon=True)
        thread.start()
    
    def _on_open_release_page(self, icon, item):
        """打开发布页面"""
        if self.latest_release_url:
            try:
                webbrowser.open(self.latest_release_url)
                log(f"Opening release page: {self.latest_release_url}")
            except Exception as e:
                log(f"Failed to open browser: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.error.open_release_page"),
                    ok=False
                )

    def update_version_info(self, icon, latest_version: str, release_url: str):
        """更新最新版本信息"""
        self.latest_version = latest_version
        self.latest_release_url = release_url
        icon.menu = self.build_menu()
    
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
        """退出应用程序"""
        icon.stop()
        
        # 设置退出事件（AppState 中已经声明了这个属性）
        if app_state.quit_event is None:
            import threading
            app_state.quit_event = threading.Event()
        
        app_state.quit_event.set()
        
        # 发送退出信号到主程序
        if getattr(app_state, "ui_queue", None):
            try:
                app_state.ui_queue.put(None)
            except Exception as e:
                log(f"Failed to send quit signal: {e}")
    
    def _save_config(self):
        """保存配置"""
        try:
            self.config_loader.save(app_state.config)
        except Exception as e:
            log(f"Failed to save config: {e}")
