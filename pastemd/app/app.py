"""Application entry point and initialization."""

import threading
import sys

try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RichQAQ.PasteMD")
except Exception:
    pass

from .. import __version__
from ..core.state import app_state
from ..core.singleton import check_single_instance
from ..config.loader import ConfigLoader
from ..config.paths import get_app_icon_path
from ..utils.logging import log
from ..utils.version_checker import VersionChecker
from ..domains.notification.manager import NotificationManager
from .wiring import Container


def initialize_application() -> Container:
    """初始化应用程序"""
    # 1. 加载配置
    config_loader = ConfigLoader()
    config = config_loader.load()
    app_state.config = config
    app_state.hotkey_str = config.get("hotkey", "<ctrl>+b")
    
    # 2. 创建依赖注入容器
    container = Container()
    
    log("Application initialized successfully")
    return container


def show_startup_notification(notification_manager: NotificationManager) -> None:
    """显示启动通知"""
    try:
        # 确保图标路径存在（仅用于验证）
        get_app_icon_path()
        notification_manager.notify(
            "PasteMD",
            "启动成功，已经运行在后台。",
            ok=True
        )
    except Exception as e:
        log(f"Failed to show startup notification: {e}")


def check_update_in_background(notification_manager: NotificationManager, tray_menu_manager=None) -> None:
    """在后台检查版本更新"""
    def _check():
        try:

            checker = VersionChecker(__version__)
            result = checker.check_update()
            
            if result and result.get("has_update"):
                latest_version = result.get("latest_version")
                release_url = result.get("release_url")
                
                # 使用菜单管理器的方法更新版本信息并重新绘制菜单
                if tray_menu_manager and app_state.icon:
                    tray_menu_manager.update_version_info(app_state.icon, latest_version, release_url)
                
                log(f"New version available: {latest_version}")
                log(f"Download URL: {release_url}")
        except Exception as e:
            log(f"Background version check failed: {e}")
    
    # 启动后台线程，不阻塞主程序
    thread = threading.Thread(target=_check, daemon=True)
    thread.start()


def main() -> None:
    """应用程序主入口点"""
    try:
        # 检查单实例运行
        if not check_single_instance():
            log("Application is already running")
            sys.exit(1)
        
        # 初始化应用程序
        container = initialize_application()
        
        # 启动热键监听
        hotkey_runner = container.get_hotkey_runner()
        hotkey_runner.start()
        
        # 获取通知管理器和菜单管理器
        notification_manager = container.get_notification_manager()
        tray_menu_manager = container.tray_menu_manager
        
        # 显示启动通知
        show_startup_notification(notification_manager)
        
        # 启动后台版本检查（无需显示通知）
        check_update_in_background(notification_manager, tray_menu_manager)
        
        # 启动托盘（阻塞运行）
        tray_runner = container.get_tray_runner()
        tray_runner.run()
        
    except KeyboardInterrupt:
        log("Application interrupted by user")
    except Exception as e:
        log(f"Fatal error: {e}")
        raise
    finally:
        # 释放锁
        if app_state.instance_checker:
            app_state.instance_checker.release_lock()
        log("Application shutting down")


if __name__ == "__main__":
    main()
