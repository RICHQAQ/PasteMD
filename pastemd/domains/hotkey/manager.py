"""Hotkey binding manager."""

from typing import Optional, Callable
from pynput import keyboard

from ...utils.logging import log


class HotkeyManager:
    """热键管理器"""
    
    def __init__(self):
        self.listener: Optional[keyboard.GlobalHotKeys] = None
        self.current_hotkey: Optional[str] = None
    
    def bind(self, hotkey: str, callback: Callable[[], None]) -> None:
        """
        绑定全局热键
        
        Args:
            hotkey: 热键字符串 (例如: "<ctrl>+b")
            callback: 热键触发时的回调函数
        """
        # 停止现有监听器
        self.unbind()
        
        try:
            mapping = {hotkey: callback}
            self.listener = keyboard.GlobalHotKeys(mapping)
            self.listener.daemon = True
            self.listener.start()
            self.current_hotkey = hotkey
            log(f"Hotkey bound: {hotkey}")
            
        except Exception as e:
            log(f"Failed to bind hotkey {hotkey}: {e}")
            raise
    
    def unbind(self) -> None:
        """解绑当前热键"""
        if self.listener:
            try:
                self.listener.stop()
                log(f"Hotkey unbound: {self.current_hotkey}")
            except Exception as e:
                log(f"Error stopping hotkey listener: {e}")
            finally:
                self.listener = None
                self.current_hotkey = None
    
    def restart(self, hotkey: str, callback: Callable[[], None]) -> None:
        """重启热键绑定"""
        self.unbind()
        self.bind(hotkey, callback)
    
    def is_bound(self) -> bool:
        """检查是否有热键绑定"""
        return self.listener is not None and self.current_hotkey is not None
