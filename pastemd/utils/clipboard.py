"""Clipboard operations."""

import pyperclip
import win32clipboard as wc
import time
from ..core.errors import ClipboardError


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
    检查剪切板内容是否为 HTML 富文本 (CF_HTML / "HTML Format")

    Returns:
        True 如果剪贴板中存在 HTML 富文本格式；否则 False
    """
    # 优先使用 pywin32；若不可用则退回 ctypes
    try:
        fmt = wc.RegisterClipboardFormat("HTML Format")

        # 某些应用会暂时占用剪贴板，这里做几次轻量重试
        for _ in range(3):
            try:
                wc.OpenClipboard()
                try:
                    return bool(wc.IsClipboardFormatAvailable(fmt))
                finally:
                    wc.CloseClipboard()
            except Exception:
                time.sleep(0.03)
        return False
    except Exception:
        # 无 pywin32 或异常，使用 ctypes 直连 Win32 API
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        RegisterClipboardFormatW = user32.RegisterClipboardFormatW
        RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
        RegisterClipboardFormatW.restype = wintypes.UINT

        OpenClipboard = user32.OpenClipboard
        OpenClipboard.argtypes = [wintypes.HWND]
        OpenClipboard.restype = wintypes.BOOL

        CloseClipboard = user32.CloseClipboard
        CloseClipboard.argtypes = []
        CloseClipboard.restype = wintypes.BOOL

        IsClipboardFormatAvailable = user32.IsClipboardFormatAvailable
        IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
        IsClipboardFormatAvailable.restype = wintypes.BOOL

        fmt = RegisterClipboardFormatW("HTML Format")
        if not fmt:
            return False

        for _ in range(3):
            if OpenClipboard(None):
                try:
                    return bool(IsClipboardFormatAvailable(fmt))
                finally:
                    CloseClipboard()
            time.sleep(0.03)
        return False


def get_clipboard_html() -> str:
    """
    获取剪贴板 HTML 富文本内容，并清理 SVG 等不可用内容
    
    返回 CF_HTML 格式中的 Fragment 部分（实际网页复制的内容），
    并自动移除 <svg> 标签和 .svg 图片引用。

    Returns:
        清理后的 HTML 富文本内容

    Raises:
        ClipboardError: 剪贴板操作失败时
    """
    try:
        fmt = wc.RegisterClipboardFormat("HTML Format")
        cf_html = None
        
        # 重试机制，避免剪贴板被占用
        for _ in range(3):
            try:
                wc.OpenClipboard()
                try:
                    if wc.IsClipboardFormatAvailable(fmt):
                        data = wc.GetClipboardData(fmt)
                        # data 可能是 bytes 或 str
                        if isinstance(data, bytes):
                            cf_html = data.decode("utf-8", errors="ignore")
                        else:
                            cf_html = data
                        break
                finally:
                    wc.CloseClipboard()
            except Exception:
                time.sleep(0.03)
        
        if not cf_html:
            raise ClipboardError("No HTML format data in clipboard")
        
        # 解析 CF_HTML 格式，提取 Fragment
        fragment = _extract_html_fragment(cf_html)
        
        # 清理 SVG 等不可用内容
        cleaned = _clean_html_content(fragment)
        
        return cleaned
        
    except Exception as e:
        raise ClipboardError(f"Failed to read HTML from clipboard: {e}")


def _extract_html_fragment(cf_html: str) -> str:
    """
    从 CF_HTML 格式中提取 Fragment 部分
    
    Args:
        cf_html: CF_HTML 格式的完整文本
        
    Returns:
        Fragment HTML 内容
    """
    import re
    
    # 提取元数据
    meta = {}
    for line in cf_html.splitlines():
        if line.strip().startswith("<!--"):
            break
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip()
    
    # 尝试使用偏移量提取 Fragment
    sf = meta.get("StartFragment")
    ef = meta.get("EndFragment")
    if sf and ef and sf.isdigit() and ef.isdigit():
        try:
            start_fragment = int(sf)
            end_fragment = int(ef)
            return cf_html[start_fragment:end_fragment]
        except Exception:
            pass
    
    # 兜底：使用注释锚点提取
    m = re.search(r"<!--StartFragment-->(.*)<!--EndFragment-->", cf_html, flags=re.S)
    if m:
        return m.group(1)
    
    # 再兜底：提取完整 HTML
    start_html = int(meta.get("StartHTML", "0"))
    end_html = int(meta.get("EndHTML", str(len(cf_html))))
    try:
        return cf_html[start_html:end_html]
    except Exception:
        return cf_html


def _clean_html_content(html: str) -> str:
    """
    清理 HTML 内容，移除 SVG 等不可用元素
    
    Args:
        html: 原始 HTML 内容
        
    Returns:
        清理后的 HTML 内容
    """
    try:
        from bs4 import BeautifulSoup
        
        soup = BeautifulSoup(html, "lxml")
        
        # 删除所有 <svg> 标签
        for svg in soup.find_all("svg"):
            svg.decompose()
        
        # 删除 src 指向 .svg 的 <img> 标签
        for img in soup.find_all("img", src=True):
            if img["src"].lower().endswith(".svg"):
                img.decompose()
        
        # 返回清理后的 HTML（包含最小壳）
        return f"<!DOCTYPE html>\n<meta charset='utf-8'>\n{str(soup)}"
        
    except ImportError:
        # 如果没有 BeautifulSoup，使用简单的正则清理
        import re
        html = re.sub(r"<svg[^>]*>.*?</svg>", "", html, flags=re.DOTALL | re.IGNORECASE)
        html = re.sub(r'<img[^>]*src=["\'][^"\']*\.svg["\'][^>]*>', "", html, flags=re.IGNORECASE)
        return html


