"""Tray icon runner."""

import pystray

from ...core.state import app_state
from .icon import create_status_icon
from .menu import TrayMenuManager


class TrayRunner:
    """托盘运行器"""
    
    def __init__(self, menu_manager: TrayMenuManager):
        self.menu_manager = menu_manager
    
    def run(self, app_name: str = "PasteMD") -> None:
        """启动托盘图标"""
        # 创建初始图标
        tray_icon = create_status_icon(ok=True)
        
        # 创建托盘实例
        icon = pystray.Icon(
            app_name,
            tray_icon,
            app_name,
            self.menu_manager.build_menu()
        )
        
        # 保存图标实例到全局状态
        app_state.icon = icon
        
        # 启动托盘（阻塞运行）
        icon.run()
