"""Single instance check to prevent multiple application instances."""

import ctypes

from .state import app_state
from ..utils.logging import log


# Windows API 函数定义
kernel32 = ctypes.windll.kernel32

# 常量
ERROR_ALREADY_EXISTS = 183


class SingleInstanceChecker:
    """检查和管理应用的单实例运行 - 使用 Windows Mutex"""
    
    def __init__(self, app_name: str = "Global\\PasteMD-Mutex"):
        self.app_name = app_name
        self.mutex_handle = None
    
    def is_already_running(self) -> bool:
        """
        检查是否已有实例在运行（使用 Windows Mutex）
        
        Returns:
            bool: 如果已有实例运行返回 True，否则返回 False
        """
        try:
            # 创建或打开一个命名互斥体
            # 如果互斥体已经存在，GetLastError 会返回 ERROR_ALREADY_EXISTS
            self.mutex_handle = kernel32.CreateMutexW(
                None,  # 默认安全属性
                True,  # 初始拥有者
                self.app_name  # 互斥体名称
            )
            
            if self.mutex_handle:
                # 检查是否是因为已存在而返回的句柄
                last_error = kernel32.GetLastError()
                if last_error == ERROR_ALREADY_EXISTS:
                    log("Mutex already exists, another instance is running")
                    return True
                else:
                    log("Mutex created successfully")
                    return False
            else:
                log("Failed to create mutex")
                return False
                
        except Exception as e:
            log(f"Error checking single instance: {e}")
            return False
    
    def acquire_lock(self) -> bool:
        """
        获取应用锁（Mutex 已经在 is_already_running 中创建）
        
        Returns:
            bool: 成功获取锁返回 True
        """
        # Mutex 已经在 is_already_running 中创建和获取
        if self.mutex_handle:
            log("Mutex lock acquired")
            return True
        return False
    
    def release_lock(self) -> None:
        """释放应用锁"""
        try:
            if self.mutex_handle:
                # 释放互斥体
                kernel32.ReleaseMutex(self.mutex_handle)
                # 关闭句柄
                kernel32.CloseHandle(self.mutex_handle)
                self.mutex_handle = None
                log("Mutex released")
        except Exception as e:
            log(f"Error releasing mutex: {e}")


def check_single_instance() -> bool:
    """
    检查并确保应用只有一个实例运行
    
    Returns:
        bool: 如果这是第一个实例返回 True，否则返回 False 并退出程序
    """
    checker = SingleInstanceChecker()
    
    # 检查是否已有实例在运行
    if checker.is_already_running():
        log("Another instance of the application is already running")
        return False
    
    # 尝试获取锁
    if not checker.acquire_lock():
        log("Failed to acquire application lock")
        return False
    
    # 保存检查器实例以便后续释放锁
    app_state.instance_checker = checker
    
    return True
