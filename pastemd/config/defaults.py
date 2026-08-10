"""Default configuration values."""

import os
import sys
from typing import Dict, Any
from .paths import resource_path
from ..utils.system_detect import is_macos, is_windows


def find_pandoc() -> str:
    """
    查找 pandoc 路径，兼容：
    - PyInstaller 单文件（exe 同级 pandoc）
    - PyInstaller 非单文件
    - Nuitka 单文件 / 非单文件
    - Inno 安装
    - 源码运行（系统 pandoc）
    - macOS 打包和源码运行
    """
    # 根据操作系统确定可执行文件名
    pandoc_binary = "pandoc.exe" if is_windows() else "pandoc"

    if is_macos():
        base_dir = os.path.dirname(sys.executable)
        candidate = os.path.join(base_dir, "pandoc", "bin", pandoc_binary)
        if os.path.exists(candidate):
            return candidate
    
    # exe/可执行文件 同级 pandoc
    exe_dir = os.path.dirname(sys.executable)
    candidate = os.path.join(exe_dir, "pandoc", pandoc_binary)
    if os.path.exists(candidate):
        return candidate

    # 打包资源路径（Nuitka / PyInstaller onedir / 新方案）
    candidate = resource_path(f"pandoc/{pandoc_binary}")
    if os.path.exists(candidate):
        return candidate

    # 兜底：系统 pandoc
    return "pandoc"


def get_default_save_dir() -> str:
    """获取默认保存目录,跨平台兼容"""
    if is_windows():
        return os.path.expandvars(r"%USERPROFILE%\Documents\pastemd")
    else:
        # macOS 和 Linux
        return os.path.expanduser("~/Documents/pastemd")


# 保留应用列表（不允许添加到可扩展工作流）
RESERVED_APPS = {"word", "wps", "excel", "wps_excel"}


DEFAULT_CONFIG: Dict[str, Any] = {
    "hotkey": "<ctrl>+<shift>+b",
    "pandoc_path": find_pandoc(),
    "reference_docx": None,
    "save_dir": get_default_save_dir(),
    "keep_file": False,
    "notify": True,
    "startup_notify": True,
    "enable_excel": True,
    "excel_keep_format": True,
    "paste_delay_s": 0.3,
    "no_app_action": "open",  # 无应用检测时的动作：open=自动打开, save=仅保存, clipboard=复制到剪贴板, none=无操作
    "md_disable_first_para_indent": True,
    "html_disable_first_para_indent": True,
    "markdown_hard_line_breaks": False,
    "horizontal_rule_style": "default",  # default=Pandoc 原生横线, paragraph_border=Office/WPS 段落边框线
    "docx_auto_table_layout": False,  # 实验性：根据内容自动调整 Word/WPS 表格列宽
    "html_formatting": {
        "strikethrough_to_del": True,
    },
    "move_cursor_to_end": True,
    "Keep_original_formula": False,
    "enable_latex_replacements": True,
    "fix_single_dollar_block": True,
    # Pandoc 下载远程资源（如图片）时附加的请求头（映射为 --request-header）。
    "pandoc_request_headers": [
        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ],
    "pandoc_filters": [],
    "pandoc_filters_by_conversion": {
        "md_to_docx": [],
        "html_to_docx": [],
        "html_to_md": [],
        "md_to_html": [],
        "md_to_rtf": [],
        "md_to_latex": [],
    },
    # 可扩展工作流配置
    # apps 格式（win）: [{"name": "Notion", "id": "/path/to/app", "window_patterns": [".*Notion.*"]}, ...]
    # apps 格式（macOS）: [{"name": "Notion", "id": "com.notionlabs.Notion", "window_patterns": [".*Notion.*"]}, ...]
    # window_patterns: 可选的正则表达式数组，用于匹配窗口标题（如浏览器中的网页）
    "extensible_workflows": {
        "html": {
            "enabled": True,  # 默认开启
            "apps": [],
            "keep_formula_latex": True,  # True = $...$, False = MathML
        },
        "md": {
            "enabled": True,  # 默认开启
            "apps": [],
            "html_formatting": {
                "css_font_to_semantic": True,
                "bold_first_row_to_header": True,
            },
        },
        "latex": {
            "enabled": True,  # 默认开启
            "apps": [],
        },
        "file": {
            "enabled": True,  # 默认开启
            "apps": [],
        },
    },
}
