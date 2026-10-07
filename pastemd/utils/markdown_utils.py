"""Markdown processing utilities - pure functions without workflow dependencies."""


import re
from collections.abc import Iterator

from .md_patterns import FENCE_MARKER_RE, INLINE_CODE_SPAN_RE, LIST_MARKER_RE


def merge_markdown_contents(files_data: list[tuple[str, str]]) -> str:
    """
    合并多个 MD 文件内容
    
    Args:
        files_data: [(filename, content), ...] 列表
        
    Returns:
        合并后的 Markdown 内容
        
    Notes:
        - 单文件：直接返回内容
        - 多文件：按原顺序拼接 `<!-- Source: filename -->` 注释 + content.strip() + 空行分隔
    """
    if len(files_data) == 1:
        # 单个文件直接返回内容
        return files_data[0][1]
    
    # 多个文件用 HTML 注释标记来源
    merged_parts = []
    for filename, content in files_data:
        merged_parts.append(f"<!-- Source: {filename} -->")
        merged_parts.append(content.strip())
        merged_parts.append("")  # 空行分隔
    
    return "\n".join(merged_parts)

def has_backtick_fenced_code_block(text: str) -> bool:
    """
    检测 ``` 这种 fenced code block，并要求起始/结束围栏成对出现。
    """
    if not text:
        return False

    pattern = re.compile(
        r'^\s{0,3}(`{3,})[^\n]*\n'   # 开始：``` 或更多反引号，允许 ```python
        r'[\s\S]*?\n'               # 内容（非贪婪）
        r'^\s{0,3}\1\s*$',          # 结束：同样数量的反引号
        re.MULTILINE
    )
    return bool(pattern.search(text))


def has_latex_math(text: str) -> bool:
    """
    检测常见 LaTeX 数学公式：
    行内：$...$ 或 \\( ... \\)
    块级：$$...$$ 或 \\[ ... \\]
    这里对 $...$ 不做内容限制（更宽松，误判风险也更高，比如 $100）。
    """
    if not text:
        return False

    # 块级：$$...$$（允许跨行）
    if re.search(r'\$\$[\s\S]*?\$\$', text):
        return True

    # 块级：\[...\]（允许跨行）
    if re.search(r'\\\[[\s\S]*?\\\]', text):
        return True

    # 行内：\(...\)（不跨行）
    if re.search(r'\\\([^\n]*?\\\)', text):
        return True

    # 行内：$...$（不跨行；排除 $$...$$）
    if re.search(r'(?<!\$)\$(?!\$)[^\n$]+(?<!\$)\$(?!\$)', text):
        return True

    return False

def is_markdown(text: str) -> bool:
    if not text or not isinstance(text, str):
        return False

    if has_backtick_fenced_code_block(text):
        return True

    if has_latex_math(text):
        return True

    md_patterns = [
        r'^\s{0,3}#{1,6}\s+',        # 标题
        r'\[.+?\]\(.+?\)',           # 链接
        r'^\s*[-*+]\s+',             # 无序列表
        r'^\s*\d+\.\s+',             # 有序列表
        r'^>\s+',                    # 引用
        r'`[^`]+`',                  # 行内代码
        r'!\[.*?\]\(.+?\)',          # 图片
        r'(\*\*|__).+?(\*\*|__)',    # 粗体
        r'(\*|_).+?(\*|_)',          # 斜体（可能误判）
    ]

    for p in md_patterns:
        if re.search(p, text, re.MULTILINE):
            return True
    return False


def iter_lines_with_code_state(lines: list[str]) -> Iterator[tuple[str, bool]]:
    """
    逐行产出 (行内容, 是否位于 ``` / ~~~ 围栏代码块内)，围栏行本身按代码行处理。

    围栏语义对齐 CommonMark/pandoc：
    - 围栏缩进最多 3 列（tab 按 4 列展开），缩进 4 列的 ``` 是缩进代码块内容；
    - 列表项内允许以内容列为基准的围栏（内容列 <= 缩进 <= 内容列+3）；
    - 闭合围栏须同字符且长度不少于开启围栏，其后只能有空白。

    供各类"跳过代码块"的 Markdown 文本变换共用。
    """
    in_code = False
    fence = ""
    list_content_col = -1  # 当前列表项的内容列，-1 表示不在列表项内
    for line in lines:
        expanded = line.expandtabs(4)
        indent = len(expanded) - len(expanded.lstrip(" "))
        stripped = line.strip()
        marker_match = FENCE_MARKER_RE.match(stripped)
        fence_eligible = marker_match is not None and (
            indent <= 3
            or (0 <= list_content_col <= indent <= list_content_col + 3)
        )

        if in_code:
            if fence_eligible:
                marker = marker_match.group(1)
                if (
                    marker[0] == fence[0]
                    and len(marker) >= len(fence)
                    and not stripped[len(marker):].strip()
                ):
                    in_code, fence = False, ""
                yield line, True
                continue
            # 代码内容行不更新列表上下文
            yield line, True
            continue

        if not stripped:
            yield line, False
            continue

        # 列表标记行刷新内容列；缩进小于内容列的非空行结束列表项
        marker_m = LIST_MARKER_RE.match(stripped)
        if marker_m:
            list_content_col = indent + marker_m.end()
            yield line, False
            continue
        if list_content_col >= 0 and indent < list_content_col:
            list_content_col = -1

        if fence_eligible:
            in_code, fence = True, marker_match.group(1)
            yield line, True
            continue
        yield line, False


def find_inline_code_spans(line: str) -> list[tuple[int, int]]:
    """返回行内代码 span 的 (start, end) 字符区间列表（含定界反引号）。"""
    return [(m.start(), m.end()) for m in INLINE_CODE_SPAN_RE.finditer(line)]
