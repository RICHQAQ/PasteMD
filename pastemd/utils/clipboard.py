"""Cross-platform clipboard operations.

This module provides a unified interface for clipboard operations across different platforms.
It automatically detects the operating system and imports the appropriate implementation.
"""

import os
import sys
from ..core.errors import ClipboardError

try:
    from bs4 import BeautifulSoup
except Exception:  # pragma: no cover
    BeautifulSoup = None


# 根据操作系统导入对应的实现
if sys.platform == "darwin":
    from .macos.clipboard import (
        get_clipboard_text,
        set_clipboard_text,
        is_clipboard_empty,
        is_clipboard_html,
        get_clipboard_html,
        set_clipboard_rich_text,
        copy_files_to_clipboard,
        is_clipboard_files,
        get_clipboard_files,
        get_markdown_files_from_clipboard,
        read_markdown_files_from_clipboard,
        preserve_clipboard,
    )
    from .macos.keystroke import simulate_paste
    # read_file_with_encoding 从共享模块导入
    from .clipboard_file_utils import read_file_with_encoding
elif sys.platform == "win32":
    from .win32.clipboard import (
        get_clipboard_text,
        set_clipboard_text,
        is_clipboard_empty,
        is_clipboard_html,
        get_clipboard_html,
        set_clipboard_rich_text,
        copy_files_to_clipboard,
        is_clipboard_files,
        get_clipboard_files,
        get_markdown_files_from_clipboard,
        read_markdown_files_from_clipboard,
        preserve_clipboard,
    )
    from .win32.keystroke import simulate_paste
    # read_file_with_encoding 从共享模块导入
    from .clipboard_file_utils import read_file_with_encoding
else:
    # 其他平台的后备实现（仅支持基本文本功能）
    import pyperclip

    def get_clipboard_text() -> str:
        """
        获取剪贴板文本内容

        Returns:
            剪贴板文本内容

        Raises:
            ClipboardError: 剪贴板操作失败时
        """
        try:
            text = pyperclip.paste()
            if text is None:
                return ""
            return text
        except Exception as e:
            raise ClipboardError(f"Failed to read clipboard: {e}")

    def is_clipboard_empty() -> bool:
        """
        检查剪贴板是否为空

        Returns:
            True 如果剪贴板为空或只包含空白字符
        """
        try:
            text = get_clipboard_text()
            return not text or not text.strip()
        except ClipboardError:
            return True

    def is_clipboard_html() -> bool:
        """
        检查剪切板内容是否为 HTML 富文本

        Note:
            在不支持的平台上始终返回 False

        Returns:
            False (不支持的平台)
        """
        return False

    def get_clipboard_html(config: dict | None = None) -> str:
        """
        获取剪贴板 HTML 富文本内容

        Note:
            在不支持的平台上会抛出异常

        Raises:
            ClipboardError: 不支持的平台
        """
        raise ClipboardError(f"HTML clipboard operations not supported on {sys.platform}")

    def set_clipboard_rich_text(
        *,
        html: str | None = None,
        rtf_bytes: bytes | None = None,
        docx_bytes: bytes | None = None,
        text: str | None = None,
    ) -> None:
        raise ClipboardError(
            f"Rich-text clipboard operations not supported on {sys.platform}"
        )

    def simulate_paste(*, timeout_s: float = 5.0) -> None:
        raise ClipboardError(f"Paste keystroke not supported on {sys.platform}")


# 导出公共接口
__all__ = [
    "get_clipboard_text",
    "set_clipboard_text",
    "is_clipboard_empty",
    "is_clipboard_html",
    "get_clipboard_html",
    "ClipboardError",
]

# 条件导出文件操作/富文本/粘贴快捷键 (Windows 和 macOS)
if sys.platform in ("win32", "darwin"):
    __all__.extend([
        "set_clipboard_rich_text",
        "simulate_paste",
        "copy_files_to_clipboard",
        "is_clipboard_files",
        "get_clipboard_files",
        "get_markdown_files_from_clipboard",
        "read_markdown_files_from_clipboard",
        "read_file_with_encoding",
        "preserve_clipboard",
        "capture_clipboard_content",
    ])


def capture_clipboard_content() -> tuple[str, str, str, bool, int]:
    """捕获剪贴板内容，返回 (preview, full_text, original_html, from_md_file, md_file_count)。

    必须在 workflow.execute() 之前调用，以便 workflow 复用预捕获的数据。

    from_md_file / md_file_count: 指示内容是否来自剪贴板中的 .md 文件列表。
    """
    try:
        text = get_clipboard_text() or ""
        found = False
        files_data: list[tuple[str, str]] = []

        # 始终检测 .md 文件（CF_HDROP / file URL，与 CF_TEXT / CF_HTML 独立互不影响）
        try:
            found, files_data, _ = read_markdown_files_from_clipboard()
        except Exception:
            pass

        if found and files_data:
            parts = [f"[{fn}]\n{content.strip()}" for fn, content in files_data]
            text = "\n\n".join(parts)
        elif not text.strip():
            # 无 md 文件也无文本 → 尝试检测普通文件
            try:
                if is_clipboard_files():
                    files = get_clipboard_files()
                    if files:
                        names = [os.path.basename(f) for f in files[:10]]
                        text = "Files: " + ", ".join(names)
                        if len(files) > 10:
                            text += f" (+{len(files) - 10} more)"
            except Exception:
                pass
        html = ""
        try:
            html = get_clipboard_html() or ""
        except Exception:
            pass

        html_text = _html_to_plain_text(html)
        if html_text and not found:
            text = html_text
        elif not text.strip() and not is_clipboard_empty():
            text = "[non-text clipboard content]"

        preview = text.strip()[:200] if text else ""

        return preview, text, html, found, len(files_data) if found else 0
    except Exception:
        return "", "", "", False, 0


def _html_to_plain_text(html: str) -> str:
    if not html or BeautifulSoup is None:
        return ""
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        try:
            soup = BeautifulSoup(html, "html.parser")
        except Exception:
            return ""

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return " ".join(soup.get_text(" ", strip=True).split())
