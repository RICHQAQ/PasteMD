"""File system utilities."""

import os
import pathlib
import subprocess
import tempfile
import re
from datetime import datetime
from typing import Optional, List
from bs4 import BeautifulSoup
from .system_detect import is_windows, is_macos


def ensure_dir(path: str) -> None:
    """确保目录存在，如不存在则创建"""
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)


def open_dir(path: str) -> None:
    """在文件管理器中打开目录"""
    path = os.path.abspath(path)
    if os.path.isdir(path):
        if is_windows():
            os.startfile(path)
        elif is_macos():
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])


def open_file(path: str) -> None:
    """使用默认应用打开文件"""
    path = os.path.abspath(path)
    if os.path.isfile(path):
        if is_windows():
            os.startfile(path)
        elif is_macos():
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])


def extract_title_from_markdown(md_text: str, max_chars: int = 30) -> Optional[str]:
    """
    从 Markdown 文本中提取标题，递减查找
    优先级：H1 → H2 → H3 → H4 → H5 → H6 → 第一句话
    
    Args:
        md_text: Markdown 文本
        max_chars: 最大字符数
        
    Returns:
        标题文本，如果没有则返回第一句话，都没有则返回 None
    """
    lines = md_text.strip().split('\n')
    
    # 递减查找标题（从 H1 到 H6）
    for heading_level in range(1, 7):
        heading_marker = '#' * heading_level
        for line in lines:
            line = line.strip()
            # 匹配标题格式（#...后跟空格）
            match = re.match(rf'^{re.escape(heading_marker)}\s+(.+?)$', line)
            if match:
                title = match.group(1).strip()
                # 清理标题中的特殊字符
                cleaned = sanitize_filename(title, max_length=max_chars)
                if cleaned:
                    return cleaned
    
    # 如果没有找到任何标题，尝试使用第一句话
    for line in lines:
        line = line.strip()
        # 跳过空行和特殊 Markdown 标记
        if not line or line.startswith('|') or line.startswith('-') or \
           line.startswith('*') or line.startswith('`') or line.startswith('>'):
            continue
        
        # 移除 Markdown 格式标记（粗体、斜体等）
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', line)  # **bold** -> bold
        text = re.sub(r'\*(.+?)\*', r'\1', text)      # *italic* -> italic
        text = re.sub(r'__(.+?)__', r'\1', text)      # __bold__ -> bold
        text = re.sub(r'_(.+?)_', r'\1', text)        # _italic_ -> italic
        text = re.sub(r'`(.+?)`', r'\1', text)        # `code` -> code
        text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)  # [link](url) -> link
        
        text = text.strip()
        if text:
            # 清理特殊字符并限制长度
            cleaned = sanitize_filename(text, max_length=max_chars)
            if cleaned:
                return cleaned
    
    return None


def extract_title_from_html(html_text: str, max_chars: int = 30) -> Optional[str]:
    """从 HTML 文本中提取合适的标题"""
    if not html_text:
        return None

    soup = None
    try:
        soup = BeautifulSoup(html_text, "lxml")
    except Exception:
        try:
            soup = BeautifulSoup(html_text, "html.parser")
        except Exception:
            return None

    if soup.title and soup.title.string:
        candidate = sanitize_filename(soup.title.string.strip(), max_length=max_chars)
        if candidate:
            return candidate

    for heading_level in range(1, 7):
        for tag in soup.find_all(f"h{heading_level}"):
            text = tag.get_text(strip=True)
            if text:
                candidate = sanitize_filename(text, max_length=max_chars)
                if candidate:
                    return candidate

    for text in soup.stripped_strings:
        candidate = sanitize_filename(text, max_length=max_chars)
        if candidate:
            return candidate

    return None


def extract_table_name_from_data(table_data: List[List[str]], max_chars: int = 30) -> Optional[str]:
    """
    从表格数据中提取表名
    使用表头（第一行）的内容组成表名
    
    Args:
        table_data: 二维数组表格数据
        max_chars: 最大字符数
        
    Returns:
        表名（使用第一行数据拼接），如果表格为空则返回 None
    """
    if not table_data or len(table_data) == 0:
        return None
    
    # 使用第一行（通常是表头）的前 6 列作为表名
    first_row = table_data[0]
    if first_row:
        # 取前 2 列
        cells = first_row[:min(6, len(first_row))]
        # 移除空单元格并拼接
        cells = [cell.strip() for cell in cells if cell.strip()]
        if cells:
            table_name = "_".join(cells)
            table_name = sanitize_filename(table_name, max_length=max_chars)
            return table_name if table_name else None
    
    return None


def sanitize_filename(filename: str, max_length: int = 100, fallback: str = "document") -> str:
    """
    清理文件名，移除不允许的字符
    
    Args:
        filename: 原始文件名
        max_length: 最大长度
        fallback: 清理后为空时的兜底名称（传空字符串表示返回空）
        
    Returns:
        清理后的文件名
    """
    # 移除或替换不允许的字符
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '_', filename)
    
    # 移除多个连续的下划线
    cleaned = re.sub(r'_+', '_', cleaned)
    
    # 移除前后的下划线，以及 Windows 不允许的尾随点和空格
    cleaned = cleaned.strip('_').rstrip(' .')
    
    # 限制长度
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length].rstrip('_').rstrip(' .')

    reserved_names = {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }
    stem, ext = os.path.splitext(cleaned)
    if stem.upper() in reserved_names:
        cleaned = f"{stem}_{ext}"
    
    return cleaned or fallback


def extract_name_from_source_files(source_filenames: Optional[List[str]]) -> Optional[str]:
    """
    从来源 MD 文件名生成本次输出的文件名（不含扩展名）

    多个来源文件时用下划线连接；与内容命名不同，不按 30 字截断，仅受整体长度上限约束。

    Args:
        source_filenames: 来源文件名列表（可含扩展名）

    Returns:
        清理后的文件名；来源为空或清理后为空则返回 None（由调用方回退到内容命名）
    """
    if not source_filenames:
        return None

    stems = []
    for filename in source_filenames:
        if not filename:
            continue
        stem = os.path.splitext(os.path.basename(str(filename)))[0].strip()
        if stem:
            stems.append(stem)

    if not stems:
        return None

    cleaned = sanitize_filename("_".join(stems), max_length=100, fallback="")
    return cleaned or None


def generate_unique_path(base_path: str) -> str:
    """
    如果文件已存在，生成唯一的路径（添加时间戳）
    
    Args:
        base_path: 基础文件路径
        
    Returns:
        唯一的文件路径
    """
    if not os.path.exists(base_path):
        return base_path
    
    # 文件已存在，添加时间戳
    dir_path = os.path.dirname(base_path)
    filename = os.path.basename(base_path)
    name, ext = os.path.splitext(filename)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_filename = f"{name}_{timestamp}{ext}"
    
    return os.path.join(dir_path, new_filename)


def generate_output_path(keep_file: bool, save_dir: str, md_text: str = "",
                         table_data: Optional[List[List[str]]] = None,
                         html_text: str = "",
                         source_filenames: Optional[List[str]] = None,
                         md_name_mode: str = "content") -> str:
    """
    生成输出文件路径，优先使用内容中提取的名称
    
    Args:
        keep_file: 是否保留文件
        save_dir: 保存目录
    md_text: Markdown 文本（用于提取标题）
        table_data: 表格数据（用于提取表名）
    html_text: HTML 富文本（用于提取标题）
        source_filenames: 来源 MD 文件名列表（按原文件名命名时使用）
        md_name_mode: 命名方式，content=根据文档内容, original=使用原文件名
        
    Returns:
        输出文件的完整路径
    """
    filename = None
    file_ext = "xlsx" if table_data is not None else "docx"
    
    # 优先级 1: 如果是表格，使用表名
    if table_data is not None:
        table_name = extract_table_name_from_data(table_data)
        if table_name:
            filename = f"{table_name}.{file_ext}"
    
    # 优先级 2: 来源为 MD 文件且设置了按原文件名命名
    if filename is None and md_name_mode == "original":
        source_name = extract_name_from_source_files(source_filenames)
        if source_name:
            filename = f"{source_name}.{file_ext}"

    # 优先级 3: 如果是 HTML，使用 HTML 标题
    if filename is None and html_text:
        html_title = extract_title_from_html(html_text)
        if html_title:
            filename = f"{html_title}.{file_ext}"

    # 优先级 4: 如果是文档，使用标题
    if filename is None and md_text:
        title = extract_title_from_markdown(md_text)
        if title:
            filename = f"{title}.{file_ext}"
    
    # 优先级 5: 使用时间戳
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"md_paste_{timestamp}.{file_ext}"
    
    if keep_file:
        ensure_dir(save_dir)
        base_path = os.path.join(save_dir, filename)
        return generate_unique_path(base_path)
    else:
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        return generate_unique_path(temp_path)
