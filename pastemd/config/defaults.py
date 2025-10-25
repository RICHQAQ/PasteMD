"""Default configuration values."""

import os
import sys
from typing import Dict, Any


DEFAULT_CONFIG: Dict[str, Any] = {}
if os.path.exists(os.path.join(os.path.dirname(sys.executable), "pandoc", "pandoc.exe")):
    DEFAULT_CONFIG = {
        "hotkey": "<ctrl>+b",
        "pandoc_path": os.path.join(os.path.dirname(sys.executable), "pandoc", "pandoc.exe"),
        "reference_docx": None,  # 可选：Pandoc 参考模板；不需要就设为 None
        "save_dir": r"%USERPROFILE%\Documents\pastemd",
        "keep_file": False,
        "notify": True,
        "enable_excel": True,  # 是否启用智能识别 Markdown 表格并粘贴到 Excel
        "excel_keep_format": True,  # Excel 粘贴时是否保留格式（粗体、斜体等）
        "auto_open_on_no_app": True  # 当未检测到应用时，自动创建文件并用默认应用打开
    }
else:
    DEFAULT_CONFIG = {
        "hotkey": "<ctrl>+b",
        "pandoc_path": "pandoc",
        "reference_docx": None,  # 可选：Pandoc 参考模板；不需要就设为 None
        "save_dir": r"%USERPROFILE%\Documents\pastemd",
        "keep_file": False,
        "notify": True,
        "enable_excel": True,  # 是否启用智能识别 Markdown 表格并粘贴到 Excel
        "excel_keep_format": True,  # Excel 粘贴时是否保留格式（粗体、斜体等）
        "auto_open_on_no_app": True  # 当未检测到应用时，自动创建文件并用默认应用打开
    }
