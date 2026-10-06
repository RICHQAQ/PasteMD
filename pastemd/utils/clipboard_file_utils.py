"""Platform-independent clipboard file utilities.

This module contains shared logic for file operations related to clipboard,
such as reading files with encoding detection, filtering Markdown files, etc.
These utilities don't depend on platform-specific APIs and can be reused by
both Windows and macOS implementations.
"""

import os
from .md_image_paths import resolve_markdown_image_paths
from ..utils.logging import log
from ..core.errors import ClipboardError


def read_file_with_encoding(file_path: str) -> str:
    """
    读取文件内容，自动检测编码
    
    尝试顺序：utf-8 -> gbk -> gb2312 -> utf-8-sig
    
    Args:
        file_path: 文件路径
        
    Returns:
        文件内容字符串
        
    Raises:
        ClipboardError: 读取失败时
    """
    encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig"]
    
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                content = f.read()
                log(f"Successfully read file '{file_path}' with encoding: {encoding}")
                return content
        except UnicodeDecodeError:
            log(f"Failed to decode '{file_path}' with encoding: {encoding}")
            continue
        except Exception as e:
            log(f"Error reading file '{file_path}' with encoding {encoding}: {e}")
            continue
    
    # 所有编码都失败
    raise ClipboardError(
        f"Failed to read file '{file_path}' with any supported encoding: {encodings}"
    )


def filter_markdown_files(file_paths: list[str]) -> list[str]:
    """
    从文件路径列表中过滤出 Markdown 文件
    
    只返回扩展名为 .md 或 .markdown 的文件
    
    Args:
        file_paths: 文件路径列表
        
    Returns:
        Markdown 文件的绝对路径列表（按文件名排序）
    """
    md_extensions = (".md", ".markdown")
    md_files = [
        f for f in file_paths
        if os.path.isfile(f) and f.lower().endswith(md_extensions)
    ]
    
    # 按文件名排序
    md_files.sort(key=lambda x: os.path.basename(x).lower())
    
    log(f"Found {len(md_files)} Markdown files from {len(file_paths)} total files")
    return md_files


def read_markdown_files(file_paths: list[str]) -> tuple[bool, list[tuple[str, str]], list[tuple[str, str]]]:
    """
    读取 Markdown 文件内容
    
    读取失败的文件会被跳过，继续处理其它文件。
    
    Args:
        file_paths: Markdown 文件路径列表
        
    Returns:
        (found, files_data, errors) 元组：
        - found: 是否成功读取至少一个文件
        - files_data: [(filename, content), ...] 成功读取的文件名和内容列表
        - errors: [(filename, error_message), ...] 读取失败的文件和错误信息
    """
    if not file_paths:
        return False, [], []
    
    files_data: list[tuple[str, str]] = []
    errors: list[tuple[str, str]] = []
    
    for file_path in file_paths:
        filename = os.path.basename(file_path)
        try:
            content = read_file_with_encoding(file_path)
            try:
                # 相对路径图片锚定到源文件目录，否则 Pandoc 以 save_dir 为 cwd 找不到
                content = resolve_markdown_image_paths(content, os.path.dirname(file_path))
            except Exception as e:
                # 图片路径解析失败不应影响文件内容本身的使用
                log(f"Failed to resolve image paths in '{filename}': {e}")
            files_data.append((filename, content))
            log(f"Successfully read MD file: {filename}")
        except Exception as e:
            # 记录失败信息，但继续处理其他文件
            error_msg = str(e)
            log(f"Failed to read MD file '{filename}': {error_msg}")
            errors.append((filename, error_msg))
    
    # found 为 True 当且仅当至少成功读取了一个文件
    return len(files_data) > 0, files_data, errors
