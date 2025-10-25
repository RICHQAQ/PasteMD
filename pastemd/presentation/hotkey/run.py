"""Hotkey UI entry point."""

from ...domains.hotkey.manager import HotkeyManager
from ...domains.hotkey.debounce import DebounceManager
from ...core.state import app_state


class HotkeyRunner:
    """热键运行器"""
    
    def __init__(self, controller_callback):
        self.hotkey_manager = HotkeyManager()
        self.debounce_manager = DebounceManager()
        self.controller_callback = controller_callback
    
    def start(self) -> None:
        """启动热键监听"""
        hotkey = app_state.hotkey_str
        
        def on_hotkey():
            if app_state.enabled:
                self.debounce_manager.trigger_async(self.controller_callback)
        
        self.hotkey_manager.bind(hotkey, on_hotkey)
    
    def stop(self) -> None:
        """停止热键监听"""
        self.hotkey_manager.unbind()
    
    def restart(self) -> None:
        """重启热键监听"""
        self.stop()
        self.start()
