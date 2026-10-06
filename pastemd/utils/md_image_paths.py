# -*- coding: utf-8 -*-
"""Markdown 相对路径图片引用锚定到指定目录。

复制 .md 文件粘贴时，Pandoc 的工作目录与源文件目录无关（如 save_dir），
相对路径的图片引用会因找不到文件而降级为占位文字。把相对路径改写为
绝对路径后，图片即可像远程 URL 一样被 Pandoc 嵌入 DOCX。
"""

import os
import re
from urllib.parse import unquote

from .markdown_utils import find_inline_code_spans, iter_lines_with_code_state
from .md_patterns import IMAGE_REF_RE, NOT_RELATIVE_PATH_RE


def resolve_markdown_image_paths(md_text: str, base_dir: str) -> str:
    """
    把相对路径且文件确实存在的图片引用改写为基于 base_dir 的绝对路径。

    远程 URL、绝对路径与定位不到的文件保持原样，代码块与行内代码内不处理。

    Args:
        md_text: Markdown 文本
        base_dir: 相对路径的基准目录（通常为源文件所在目录）

    Returns:
        图片引用为绝对路径的 Markdown 文本
    """
    if not md_text or not base_dir or "![" not in md_text:
        return md_text

    base_dir = os.path.abspath(base_dir)
    out = []
    for line, in_code in iter_lines_with_code_state(md_text.split("\n")):
        if in_code or "![" not in line:
            out.append(line)
        else:
            out.append(_rewrite_line(line, base_dir))
    return "\n".join(out)


def _rewrite_line(line: str, base_dir: str) -> str:
    code_spans = find_inline_code_spans(line)

    def replace(match: re.Match[str]) -> str:
        if any(start <= match.start() < end for start, end in code_spans):
            return match.group(0)
        return _rewrite(match, base_dir)

    return IMAGE_REF_RE.sub(replace, line)


def _rewrite(match: re.Match[str], base_dir: str) -> str:
    prefix, lt, path, gt, title = match.groups()
    absolute = _to_absolute(path.strip(), base_dir)
    if absolute is None or bool(lt) != bool(gt):
        return match.group(0)
    if " " in absolute and not lt:
        lt, gt = "<", ">"
    return f"{prefix}{lt}{absolute}{gt}{title or ''})"


def _to_absolute(path: str, base_dir: str) -> str | None:
    """path 为相对路径且能定位到文件时返回绝对路径（正斜杠），否则 None。"""
    if not path or NOT_RELATIVE_PATH_RE.match(path):
        return None

    for candidate in (path, unquote(path)):
        full = os.path.normpath(os.path.join(base_dir, candidate))
        if os.path.isfile(full):
            return full.replace("\\", "/")
    return None
