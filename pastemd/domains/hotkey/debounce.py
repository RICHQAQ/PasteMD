"""Hotkey trigger debouncing and mutual exclusion."""

import time
import threading
from typing import Callable

from ...core.constants import FIRE_DEBOUNCE_SEC
from ...core.state import app_state
from ...utils.logging import log


class DebounceManager:
    """热键触发防抖管理器"""
    
    def __init__(self):
        pass
    
    def trigger_async(self, callback: Callable[[], None]) -> None:
        """
        异步触发回调，带防抖和互斥处理
        
        Args:
            callback: 要执行的回调函数
        """
        now = time.time()
        
        # 防抖：短时间内重复触发直接忽略
        if now - app_state.last_fire < FIRE_DEBOUNCE_SEC:
            return
        
        app_state.last_fire = now
        
        # 互斥：如果已有任务在运行，直接返回
        if app_state.is_running():
            return
        
        # 启动后台线程执行实际工作
        def worker():
            app_state.set_running(True)
            try:
                callback()
            except Exception as e:
                log(f"Callback execution failed: {e}")
            finally:
                app_state.set_running(False)
        
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
