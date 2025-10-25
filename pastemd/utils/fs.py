"""File system utilities."""

import os
import pathlib
import tempfile
import re
from datetime import datetime
from typing import Optional, List


def ensure_dir(path: str) -> None:
    """确保目录存在，如不存在则创建"""
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)


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


def sanitize_filename(filename: str, max_length: int = 100) -> str:
    """
    清理文件名，移除不允许的字符
    
    Args:
        filename: 原始文件名
        max_length: 最大长度
        
    Returns:
        清理后的文件名
    """
    # 移除或替换不允许的字符
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '_', filename)
    
    # 移除多个连续的下划线
    cleaned = re.sub(r'_+', '_', cleaned)
    
    # 移除前后的下划线
    cleaned = cleaned.strip('_')
    
    # 限制长度
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length].rstrip('_')
    
    return cleaned or "document"


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
                         table_data: Optional[List[List[str]]] = None) -> str:
    """
    生成输出文件路径，优先使用内容中提取的名称
    
    Args:
        keep_file: 是否保留文件
        save_dir: 保存目录
        md_text: Markdown 文本（用于提取标题）
        table_data: 表格数据（用于提取表名）
        
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
    
    # 优先级 2: 如果是文档，使用标题
    if filename is None and md_text:
        title = extract_title_from_markdown(md_text)
        if title:
            filename = f"{title}.{file_ext}"
    
    # 优先级 3: 使用时间戳
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
