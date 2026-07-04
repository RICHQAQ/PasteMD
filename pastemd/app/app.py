"""Application entry point and initialization."""

import sys
import threading
import queue
import tkinter as tk
from ..config.paths import get_app_icon_path, is_first_launch
from ..utils.system_detect import is_macos

# macOS: 首次启动时打开使用说明页面
IS_FIRST_LAUNCH = is_first_launch()

if is_macos() and IS_FIRST_LAUNCH:
    try:
        import webbrowser
        webbrowser.open("https://pastemd.richqaq.cn/macos")
    except Exception:
        pass

# 设置 Windows 应用程序 ID (仅在 Windows 上)
try:
    import ctypes
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RichQAQ.PasteMD")
except Exception:
    pass

from ..utils.dpi import set_dpi_awareness

from .. import __version__
from ..core.state import app_state
from ..core.singleton import check_single_instance
from ..config.loader import ConfigLoader
from ..utils.logging import log
from ..utils.version_checker import VersionChecker
from ..service.notification.manager import NotificationManager
from ..i18n import FALLBACK_LANGUAGE, detect_system_language, set_language, t
from .wiring import Container


def initialize_application() -> tuple[Container, dict]:
    """初始化应用程序
    
    Returns:
        tuple: (container, workflow_conflicts)
            - container: 依赖注入容器
            - workflow_conflicts: 预留字段，当前返回空字典
    """
    # 1. 加载配置
    config_loader = ConfigLoader()
    config = config_loader.load()
    app_state.config = config
    app_state.hotkey_str = config.get("hotkey", "<ctrl>+<shift>+b")

    language_value = config.get("language")
    if not language_value:
        detected_language = detect_system_language()
        if detected_language:
            language = detected_language
        else:
            language = FALLBACK_LANGUAGE
        app_state.config["language"] = language
        try:
            config_loader.save(app_state.config)
        except Exception as exc:
            log(f"Failed to persist auto-detected language: {exc}")
        log(f"First launch: detected system language '{language}'")
    else:
        language = str(language_value)
    set_language(language)
    
    # 2. 创建依赖注入容器
    container = Container()

    # 3. 启动历史记录服务
    try:
        container.history_manager.start()
        # Wire history manager into the workflow router
        from .workflows.router import router as workflow_router
        workflow_router.set_history_manager(container.history_manager)
    except Exception as exc:
        log(f"Failed to start history manager: {exc}")

    log("Application initialized successfully")
    return container, {}


def show_startup_notification(notification_manager: NotificationManager) -> None:
    """显示启动通知"""
    try:
        # 检查是否启用开机通知
        if app_state.config.get("startup_notify", True) is False:
            return
        
        # 确保图标路径存在（仅用于验证）
        get_app_icon_path()
        notification_manager.notify(
            "PasteMD",
            t("app.startup.success"),
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
        # 设置 DPI 感知（尽早调用）
        set_dpi_awareness()

        # 检查单实例运行
        if not check_single_instance():
            # 已有实例：macOS 上尝试通知已运行实例打开设置页（类似“再次点击应用图标”）
            if is_macos():
                try:
                    from ..utils.macos.ipc import send_command

                    if send_command("open_settings"):
                        sys.exit(0)
                except Exception as exc:
                    log(f"Failed to send reopen command: {exc}")

            log("Application is already running")
            sys.exit(1)
        
        # 初始化应用程序
        container, workflow_conflicts = initialize_application()

        # 初始化 UI 队列，确保 Tk 等 UI 操作始终在主线程
        ui_queue: queue.Queue = queue.Queue()
        app_state.ui_queue = ui_queue
        
        # 初始化退出事件
        import threading
        app_state.quit_event = threading.Event()
        
        # 初始化全局 Tk 实例 (解决 Tcl_AsyncDelete 问题)
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        app_state.root = root

        # macOS: 默认隐藏 Dock 图标，仅在弹窗打开时临时显示
        if is_macos():
            try:
                from ..utils.macos.dock import set_dock_visible

                set_dock_visible(False)
            except Exception as exc:
                log(f"Failed to hide Dock icon: {exc}")

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
        
        # 获取托盘运行器
        tray_runner = container.get_tray_runner()
        
        # macOS: 托盘必须在主线程初始化（NSWindow 限制），但可以在后台运行
        # Windows: 托盘可以在后台线程运行
        if is_macos():
            # 在主线程初始化托盘，然后分离运行
            tray_runner.setup()
            # 使用 setup_detached 后，托盘会在后台线程运行
            if IS_FIRST_LAUNCH:
                def _open_permissions_settings():
                    try:
                        tray_menu_manager.open_settings_tab("permissions")
                    except Exception as exc:
                        log(f"Failed to open permissions settings on first launch: {exc}")

                if ui_queue is not None:
                    ui_queue.put(_open_permissions_settings)
                else:
                    _open_permissions_settings()
        else:
            # Windows: 直接在后台线程运行
            threading.Thread(target=tray_runner.run, daemon=True).start()

        # macOS: 启动本地 IPC，支持“再次启动/点击应用图标”时唤起设置页
        if is_macos():
            try:
                from ..utils.macos.ipc import start_server

                def _handle_command(cmd: str) -> None:
                    if cmd == "open_settings":
                        # TrayMenuManager 内部会投递到 UI 队列，确保 Tk 在主线程
                        try:
                            tray_menu_manager.handle_app_reopen(app_state.icon)
                        except Exception as exc:
                            log(f"Failed to open settings from IPC: {exc}")

                start_server(_handle_command)
            except Exception as exc:
                log(f"Failed to start IPC server: {exc}")

            # macOS: Finder/Launchpad 二次打开已运行的 .app 通常不会启动新进程，
            # 而是发送 Reopen 事件；这里显式捕获并唤起设置页。
            try:
                from ..utils.macos.reopen import install_reopen_handler

                install_reopen_handler(
                    lambda: tray_menu_manager.handle_app_reopen(app_state.icon)
                )
            except Exception as exc:
                log(f"Failed to install macOS reopen handler: {exc}")

        # UI 队列处理函数
        def process_ui_queue():
            try:
                # 检查退出事件
                quit_event = getattr(app_state, 'quit_event', None)
                if quit_event and quit_event.is_set():
                    # 退出事件被触发 - 清理所有窗口
                    _cleanup_all_windows()
                    root.quit()
                    return
                
                while True:
                    # 非阻塞获取任务
                    task = ui_queue.get_nowait()
                    if task is None:
                        # 退出信号 - 清理所有窗口
                        _cleanup_all_windows()
                        root.quit()
                        return
                    try:
                        task()
                    except Exception as e:
                        log(f"UI task error: {e}")
            except queue.Empty:
                pass
            finally:
                # 继续轮询 (100ms)
                root.after(100, process_ui_queue)
        
        def _cleanup_all_windows():
            """清理所有窗口"""
            try:
                # 销毁所有子窗口（Toplevel窗口）
                for widget in root.winfo_children():
                    try:
                        if hasattr(widget, 'destroy'):
                            widget.destroy()
                    except Exception as e:
                        log(f"Failed to destroy widget: {e}")
            except Exception as e:
                log(f"Error during window cleanup: {e}")

        # 启动队列处理
        root.after(100, process_ui_queue)
        
        # 进入主事件循环
        root.mainloop()
        
    except KeyboardInterrupt:
        log("Application interrupted by user")
    except Exception as e:
        log(f"Fatal error: {e}")
        raise
    finally:
        # 停止历史记录服务
        try:
            if 'container' in dir():
                container.history_manager.stop()
        except Exception:
            pass
        # 释放锁
        if app_state.instance_checker:
            app_state.instance_checker.release_lock()
        log("Application shutting down")


if __name__ == "__main__":
    main()
