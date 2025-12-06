"""Global runtime state management."""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
import threading


@dataclass
class AppState:
    """全局应用状态"""
    enabled: bool = True
    running: bool = False
    last_fire: float = 0.0
    last_ok: bool = True
    hotkey_str: str = "<ctrl>+b"
    config: Dict[str, Any] = field(default_factory=dict)

    # UI组件引用
    listener: Optional[Any] = None  # pynput.keyboard.GlobalHotKeys
    icon: Optional[Any] = None      # pystray.Icon

    # 单实例检查器
    instance_checker: Optional[Any] = None  # SingleInstanceChecker

    # UI 任务队列，确保 Tk 等 UI 操作仅在主线程运行
    ui_queue: Optional[Any] = None

    # 线程锁
    _lock: threading.Lock = field(default_factory=threading.Lock)
    
    def with_lock(self, func):
        """线程安全执行函数"""
        with self._lock:
            return func()
    
    def set_running(self, running: bool):
        """线程安全设置运行状态"""
        with self._lock:
            self.running = running
    
    def is_running(self) -> bool:
        """线程安全检查运行状态"""
        with self._lock:
            return self.running


# 全局状态实例
app_state = AppState()
