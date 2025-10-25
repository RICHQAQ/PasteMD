"""COM interop utilities."""

import pythoncom
from functools import wraps


def ensure_com(func):
    """
    装饰器：确保在 COM 环境中执行函数
    
    自动初始化和清理 COM 环境，避免线程问题
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                # 静默处理清理异常
                pass
    return wrapper
