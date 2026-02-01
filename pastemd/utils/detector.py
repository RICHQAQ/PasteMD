# -*- coding: utf-8 -*-
"""
跨平台应用检测模块
自动根据操作系统选择对应的实现
"""

from __future__ import annotations
import sys
from typing import Literal

AppType = Literal["word", "wps", "excel", "wps_excel", "onenote", "powerpoint", ""]

# 根据平台导入对应的实现
if sys.platform == "darwin":
    from .macos.detector import (
        detect_active_app as _detect_active_app,
        detect_wps_type as _detect_wps_type,
        get_frontmost_window_title as _get_frontmost_window_title,
    )
elif sys.platform == "win32":
    from .win32.detector import (
        detect_active_app as _detect_active_app,
        detect_wps_type as _detect_wps_type,
    )
    from .win32.window import (
        get_foreground_window_title as _get_frontmost_window_title,
    )
else:
    # 不支持的平台，提供空实现
    def _detect_active_app() -> str:
        return ""
    
    def _detect_wps_type() -> str:
        return ""
    
    def _get_frontmost_window_title() -> str:
        return ""


def detect_active_app() -> AppType:
    """
    检测当前活跃的插入目标应用
    
    跨平台统一接口，自动调用对应平台的实现
    
    Returns:
        "word": Microsoft Word
        "excel": Microsoft Excel
        "wps": WPS 文字
        "wps_excel": WPS 表格
        "": 未识别或不支持的应用
        
    Examples:
        >>> app = detect_active_app()
        >>> if app == "word":
        ...     print("检测到 Word")
    """
    return _detect_active_app()


def detect_wps_type() -> AppType:
    """
    检测 WPS 应用的具体类型 (文字/表格)
    
    跨平台统一接口，自动调用对应平台的实现
    
    Returns:
        "wps": WPS 文字
        "wps_excel": WPS 表格
        "": 无法识别
        
    Note:
        此函数应在确认当前应用是 WPS 后调用
        不同平台的实现方式有所不同：
        - Windows: 优先使用 COM 对象判断，其次使用窗口标题
        - macOS: 使用窗口标题判断（无 COM 接口）
    """
    return _detect_wps_type()


def is_office_app(app_type: str) -> bool:
    """
    判断是否为支持的 Office 应用
    
    Args:
        app_type: detect_active_app() 返回的应用类型
        
    Returns:
        True 如果是支持的 Office 应用
    """
    return app_type in ("word", "excel", "wps", "wps_excel")


def is_word_like(app_type: str) -> bool:
    """
    判断是否为文字处理应用（Word 或 WPS 文字）
    
    Args:
        app_type: detect_active_app() 返回的应用类型
        
    Returns:
        True 如果是文字处理应用
    """
    return app_type in ("word", "wps")


def is_excel_like(app_type: str) -> bool:
    """
    判断是否为表格应用（Excel 或 WPS 表格）
    
    Args:
        app_type: detect_active_app() 返回的应用类型
        
    Returns:
        True 如果是表格应用
    """
    return app_type in ("excel", "wps_excel")


def get_app_display_name(app_type: str) -> str:
    """
    获取应用的显示名称
    
    Args:
        app_type: detect_active_app() 返回的应用类型
        
    Returns:
        应用的友好显示名称
    """
    display_names = {
        "word": "Microsoft Word",
        "excel": "Microsoft Excel",
        "wps": "WPS 文字",
        "wps_excel": "WPS 表格",
        "": "未知应用",
    }
    return display_names.get(app_type, app_type)


def get_frontmost_window_title() -> str:
    """
    获取当前前台窗口的标题
    
    跨平台统一接口，自动调用对应平台的实现
    
    Returns:
        窗口标题字符串，失败时返回空字符串
    """
    return _get_frontmost_window_title()


__all__ = [
    "AppType",
    "detect_active_app",
    "detect_wps_type",
    "is_office_app",
    "is_word_like",
    "is_excel_like",
    "get_app_display_name",
    "get_frontmost_window_title",
]
