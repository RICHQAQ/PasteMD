"""Hotkey binding manager."""

from typing import Optional, Callable
from pynput import keyboard

from ...utils.logging import log


class HotkeyManager:
    """热键管理器 - 负责全局热键监听和触发"""
    
    def __init__(self):
        self.global_listener: Optional[keyboard.GlobalHotKeys] = None  # 全局热键监听器
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
            self.global_listener = keyboard.GlobalHotKeys(mapping)
            self.global_listener.daemon = True
            self.global_listener.start()
            self.current_hotkey = hotkey
            log(f"Hotkey bound: {hotkey}")
            
        except Exception as e:
            log(f"Failed to bind hotkey {hotkey}: {e}")
            raise
    
    def unbind(self) -> None:
        """解绑当前热键"""
        if self.global_listener:
            try:
                self.global_listener.stop()
                log(f"Hotkey unbound: {self.current_hotkey}")
            except Exception as e:
                log(f"Error stopping hotkey listener: {e}")
            finally:
                self.global_listener = None
                self.current_hotkey = None
    
    def restart(self, hotkey: str, callback: Callable[[], None]) -> None:
        """重启热键绑定"""
        self.unbind()
        self.bind(hotkey, callback)
    
    def is_bound(self) -> bool:
        """检查是否有热键绑定"""
        return self.global_listener is not None and self.current_hotkey is not None
    
    def pause(self) -> None:
        """暂停热键监听（用于录制时避免触发）"""
        if self.global_listener:
            try:
                self.global_listener.stop()
                log(f"Hotkey paused: {self.current_hotkey}")
            except Exception as e:
                log(f"Error pausing hotkey listener: {e}")
    
    def resume(self, callback: Callable[[], None]) -> None:
        """恢复热键监听"""
        if self.current_hotkey and not self.global_listener:
            try:
                mapping = {self.current_hotkey: callback}
                self.global_listener = keyboard.GlobalHotKeys(mapping)
                self.global_listener.daemon = True
                self.global_listener.start()
                log(f"Hotkey resumed: {self.current_hotkey}")
            except Exception as e:
                log(f"Error resuming hotkey listener: {e}")
