"""Dependency injection and object wiring."""

from ..core.state import app_state
from ..config.loader import ConfigLoader
from ..domains.notification.manager import NotificationManager
from ..app.workflows.paste_workflow import PasteWorkflow
from ..presentation.tray.menu import TrayMenuManager
from ..presentation.tray.run import TrayRunner
from ..presentation.hotkey.run import HotkeyRunner


class Container:
    """依赖注入容器"""
    
    def __init__(self):
        # 基础服务
        self.config_loader = ConfigLoader()
        self.notification_manager = NotificationManager()
        
        # 业务工作流
        self.paste_workflow = PasteWorkflow()
        
        # UI 组件
        self.tray_menu_manager = TrayMenuManager(
            self.config_loader,
            self.notification_manager
        )
        self.tray_runner = TrayRunner(self.tray_menu_manager)
        self.hotkey_runner = HotkeyRunner(
            self.paste_workflow.execute,
            self.notification_manager,
            self.config_loader
        )
        
        # 设置热键重启回调
        self.tray_menu_manager.set_restart_hotkey_callback(
            self.hotkey_runner.restart
        )
        
        # 设置热键暂停/恢复回调（用于录制时暂停）
        self.tray_menu_manager.set_pause_hotkey_callback(
            self.hotkey_runner.get_hotkey_manager().pause
        )
        
        # 恢复热键回调
        def on_hotkey_resumed():
            """热键恢复后的回调函数"""
            if app_state.enabled:
                self.hotkey_runner.debounce_manager.trigger_async(
                    self.paste_workflow.execute
                )
        
        def resume_hotkey():
            """恢复热键监听"""
            self.hotkey_runner.get_hotkey_manager().resume(on_hotkey_resumed)
        
        self.tray_menu_manager.set_resume_hotkey_callback(resume_hotkey)
    
    def get_paste_workflow(self) -> PasteWorkflow:
        return self.paste_workflow
    
    def get_hotkey_runner(self) -> HotkeyRunner:
        return self.hotkey_runner
    
    def get_tray_runner(self) -> TrayRunner:
        return self.tray_runner
    
    def get_notification_manager(self) -> NotificationManager:
        return self.notification_manager
