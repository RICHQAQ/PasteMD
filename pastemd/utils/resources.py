"""PyInstaller resource path utilities."""

import os
import sys


def resource_path(relative_path: str) -> str:
    """
    获取资源文件的绝对路径，支持 PyInstaller 打包后的环境
    
    Args:
        relative_path: 相对于基础目录的路径
        
    Returns:
        资源文件的绝对路径
    """
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller 打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    
    # 开发环境：相对于项目根目录
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_dir, relative_path)
