"""Tray menu construction and callbacks."""

import os
import pystray
import threading
import webbrowser

from ... import __version__
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.paths import get_log_path, get_config_path
from ...domains.notification.manager import NotificationManager
from ...utils.fs import ensure_dir
from ...utils.logging import log
from ...utils.version_checker import VersionChecker
from .icon import create_status_icon
from ..hotkey.dialog import HotkeyDialog


class TrayMenuManager:
    """托盘菜单管理器"""
    
    def __init__(self, config_loader: ConfigLoader, notification_manager: NotificationManager):
        self.config_loader = config_loader
        self.notification_manager = notification_manager
        self.restart_hotkey_callback = None  # 将由外部设置
        self.version_checker = None  # 将由外部设置或按需创建
        self.latest_version = None  # 存储最新版本号
        self.latest_release_url = None  # 存储最新版本的下载链接
    
    def set_restart_hotkey_callback(self, callback):
        """设置重启热键的回调函数"""
        self.restart_hotkey_callback = callback
    
    def build_menu(self) -> pystray.Menu:
        """构建托盘菜单"""
        
        config = app_state.config
        
        # 构建版本菜单项
        version_menu_items = [
            pystray.MenuItem(
                f"当前版本: {__version__}",
                lambda icon, item: None,
                enabled=False
            ),
        ]
        
        # 如果有新版本，显示新版本号
        if self.latest_version:
            version_menu_items.append(
                pystray.MenuItem(
                    f"✨ 新版本: {self.latest_version}",
                    self._on_open_release_page,
                    enabled=True
                )
            )
        else:
            version_menu_items.append(
                pystray.MenuItem(
                    "检查更新",
                    self._on_check_update
                )
            )
        
        return pystray.Menu(
            # 快捷显示
            pystray.MenuItem(
                f"快捷键: {app_state.config['hotkey']}",
                lambda icon, item: None,
                enabled=False
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "启用热键",
                self._on_toggle_enabled,
                checked=lambda item: app_state.enabled
            ),
            pystray.MenuItem(
                "弹窗通知",
                self._on_toggle_notify,
                checked=lambda item: config.get("notify", True)
            ),
            pystray.MenuItem(
                "无应用时自动打开",
                self._on_toggle_auto_open,
                checked=lambda item: config.get("auto_open_on_no_app", True)
            ),
            pystray.MenuItem(
                "插入后光标移动到末尾",
                self._on_toggle_move_cursor,
                checked=lambda item: config.get("move_cursor_to_end", True)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("设置热键", self._on_set_hotkey),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "保留生成文件",
                self._on_toggle_keep,
                checked=lambda item: config.get("keep_file", False)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "启动插入excel",
                self._on_toggle_excel,
                checked=lambda item: config.get("enable_excel", True)
            ),
            pystray.MenuItem(
                "启动excel解析特殊格式",
                self._on_toggle_excel_format,
                checked=lambda item: config.get("excel_keep_format", True)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("打开保存目录", self._on_open_save_dir),
            pystray.MenuItem("查看日志", self._on_open_log),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("编辑配置", self._on_edit_config),
            pystray.MenuItem("重载配置/热键", self._on_reload),
            pystray.Menu.SEPARATOR,
            *version_menu_items,
            pystray.MenuItem(
                "关于",
                self._on_open_about_page
            ),
            pystray.MenuItem("退出", self._on_quit)
        )
    
    # 菜单回调函数
    def _on_toggle_enabled(self, icon, item):
        """切换热键启用状态"""
        app_state.enabled = not app_state.enabled
        icon.icon = create_status_icon(ok=app_state.enabled)
        
        status = "已启用热键" if app_state.enabled else "已暂停热键"
        icon.menu = self.build_menu()
        self.notification_manager.notify("PasteMD", status, ok=app_state.enabled)
    
    def _on_set_hotkey(self, icon, item):
        """设置热键"""
        def save_hotkey(new_hotkey: str):
            """保存新热键并重启热键绑定"""
            try:
                # 更新配置
                app_state.config["hotkey"] = new_hotkey
                app_state.hotkey_str = new_hotkey
                self._save_config()
                
                # 重启热键绑定
                if self.restart_hotkey_callback:
                    self.restart_hotkey_callback()
                
                # 刷新菜单
                icon.menu = self.build_menu()
                
                log(f"Hotkey changed to: {new_hotkey}")
                self.notification_manager.notify(
                    "PasteMD",
                    f"热键已更新为：{new_hotkey}",
                    ok=True)
            except Exception as e:
                log(f"Failed to save hotkey: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    f"保存热键失败：{str(e)}",
                    ok=False)
                raise
        
        # 直接在主线程中显示对话框
        # tkinter 必须在主线程中运行，不能使用后台线程
        try:
            dialog = HotkeyDialog(
                current_hotkey=app_state.hotkey_str,
                on_save=save_hotkey
            )
            dialog.show()
        except Exception as e:
            log(f"Failed to show hotkey dialog: {e}")
            self.notification_manager.notify("PasteMD", f"打开热键设置失败：{str(e)}", ok=False)
    
    def _on_toggle_notify(self, icon, item):
        """切换通知状态"""
        current = app_state.config.get("notify", True)
        app_state.config["notify"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        if app_state.config["notify"]:
            self.notification_manager.notify("PasteMD", "已开启通知", ok=True)
        else:
            log("Notifications disabled via tray toggle")
    
    def _on_toggle_auto_open(self, icon, item):
        """切换无应用时自动打开状态"""
        current = app_state.config.get("auto_open_on_no_app", True)
        app_state.config["auto_open_on_no_app"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = "已开启无应用时自动打开" if app_state.config["auto_open_on_no_app"] else "已关闭无应用时自动打开"
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_toggle_move_cursor(self, icon, item):
        """切换插入后光标移动到末尾状态"""
        current = app_state.config.get("move_cursor_to_end", True)
        app_state.config["move_cursor_to_end"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = "已开启插入后光标移动到末尾" if app_state.config["move_cursor_to_end"] else "已关闭插入后光标移动到末尾"
        self.notification_manager.notify("PasteMD", status, ok=True)
        
    def _on_toggle_excel(self, icon, item):
        """切换启用 Excel 插入"""
        current = app_state.config.get("enable_excel", True)
        app_state.config["enable_excel"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_manager.notify("PasteMD", f"Excel 插入功能：{'开启' if not current else '关闭'}", ok=True)
        
    def _on_toggle_excel_format(self, icon, item):
        """切换 Excel 粘贴时是否保留格式"""
        current = app_state.config.get("excel_keep_format", True)
        app_state.config["excel_keep_format"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        self.notification_manager.notify("PasteMD", f"Excel 格式保留：{'开启' if not current else '关闭'}", ok=True)
    
    def _on_toggle_keep(self, icon, item):
        """切换保留文件状态"""
        current = app_state.config.get("keep_file", False)
        app_state.config["keep_file"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = "保留文件：开启" if app_state.config["keep_file"] else "保留文件：关闭"
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_open_save_dir(self, icon, item):
        """打开保存目录"""
        save_dir = app_state.config.get("save_dir", "")
        save_dir = os.path.expandvars(save_dir)
        ensure_dir(save_dir)
        os.startfile(save_dir)
    
    def _on_open_log(self, icon, item):
        """打开日志文件"""
        log_path = get_log_path()
        if not os.path.exists(log_path):
            # 创建空日志文件
            open(log_path, "w", encoding="utf-8").close()
        os.startfile(log_path)
    
    def _on_edit_config(self, icon, item):
        """编辑配置文件"""
        config_path = get_config_path()
        if not os.path.exists(config_path):
            self._save_config()  # 创建默认配置文件
        os.startfile(config_path)
    
    def _on_reload(self, icon, item):
        """重载配置和热键"""
        try:
            app_state.config = self.config_loader.load()
            app_state.hotkey_str = app_state.config.get("hotkey", "<ctrl>+b")
            
            if self.restart_hotkey_callback:
                self.restart_hotkey_callback()
            icon.menu = self.build_menu()
            self.notification_manager.notify("PasteMD", "配置已重载", ok=True)
        except Exception as e:
            log(f"Failed to reload config: {e}")
            self.notification_manager.notify("PasteMD", "配置重载失败", ok=False)
    
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
                        "PasteMD - 检查更新失败",
                        "网络连接失败，请稍后再试",
                        ok=False
                    )
                elif result.get("has_update"):
                    latest_version = result.get("latest_version")
                    release_url = result.get("release_url")
                    
                    # 使用 update_version_info 方法更新版本信息并重新绘制菜单
                    self.update_version_info(icon, latest_version, release_url)
                    
                    # 通知用户有新版本，并自动打开下载页面
                    message = f"发现新版本 {latest_version}，正在为您打开下载页面..."
                    self.notification_manager.notify(
                        "PasteMD - 有新版本",
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
                        "PasteMD - 已是最新版本",
                        f"当前版本 {current_version} 已是最新版本",
                        ok=True
                    )
            except Exception as e:
                self.notification_manager.notify(
                    "PasteMD - 更新检查失败",
                    f"检查更新时出错：{str(e) if len(str(e)) <= 15 else str(e)[:12] + '...'}",
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
                    "无法打开浏览器，请手动访问 GitHub Releases 页面",
                    ok=False
                )

    def update_version_info(self, icon, latest_version: str, release_url: str):
        """更新最新版本信息"""
        self.latest_version = latest_version
        self.latest_release_url = release_url
        icon.menu = self.build_menu()
    
    def _on_open_about_page(self, icon, item):
        """打开关于页面"""
        about_url = "http://pastemd.richqaq.cn"
        try:
            webbrowser.open(about_url)
            log(f"Opening about page: {about_url}")
        except Exception as e:
            log(f"Failed to open browser: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "无法打开浏览器，请手动访问关于页面\nhttp://pastemd.richqaq.cn",
                ok=False
            )
    
    def _on_quit(self, icon, item):
        """退出应用程序"""
        icon.stop()
    
    def _save_config(self):
        """保存配置"""
        try:
            self.config_loader.save(app_state.config)
        except Exception as e:
            log(f"Failed to save config: {e}")
