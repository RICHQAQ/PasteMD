# -*- coding: utf-8 -*-
"""Markdown 中内嵌 HTML 表格 → Markdown 表格转换。

Pandoc 的 markdown reader 不会把 raw HTML <table> 解析为原生表格：
DOCX writer 会丢弃 RawBlock html 标签，只留下单元格文本，导致表格被
"拍平"成零散段落。因此在预处理阶段把内嵌 HTML 表格改写成管道表格。
"""

import re
from typing import List, Optional

try:
    from bs4 import BeautifulSoup, Tag  # type: ignore
except Exception:  # pragma: no cover - BeautifulSoup 缺失时跳过转换
    BeautifulSoup = None  # type: ignore
    Tag = None  # type: ignore

_TABLE_BLOCK_RE = re.compile(r"<table\b[^>]*>.*?</table\s*>", re.IGNORECASE | re.DOTALL)


def convert_embedded_html_tables(md_text: str) -> str:
    """把 Markdown 文本中的内嵌 HTML <table> 块转换为 Markdown 管道表格。

    代码块（``` / ~~~）内的内容不做处理；单个表格解析失败时保持原样。

    Args:
        md_text: 原始 Markdown 文本

    Returns:
        转换后的 Markdown 文本
    """
    if not md_text or "<table" not in md_text.lower():
        return md_text
    if BeautifulSoup is None:
        return md_text

    lines = md_text.split("\n")
    out: List[str] = []
    pending: List[str] = []
    in_code = False
    fence = ""

    def flush_pending() -> None:
        if not pending:
            return
        out.append(_convert_tables_in_chunk("\n".join(pending)))
        pending.clear()

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            flush_pending()
            if not in_code:
                in_code, fence = True, stripped[:3]
            elif stripped.startswith(fence):
                in_code, fence = False, ""
            out.append(line)
            continue
        if in_code:
            out.append(line)
        else:
            pending.append(line)

    flush_pending()
    return "\n".join(out)


def _convert_tables_in_chunk(chunk: str) -> str:
    """对一段非代码 Markdown 文本执行 HTML 表格替换。"""
    if "<table" not in chunk.lower():
        return chunk

    # 注释内的表格不转换（避免把注释掉的内容激活为真实表格）
    comment_spans = [(m.start(), m.end()) for m in re.finditer(r"<!--.*?-->", chunk, re.DOTALL)]
    line_starts = [m.start() for m in re.finditer(r"(?m)^", chunk)]

    def _in_comment(pos: int) -> bool:
        return any(s <= pos < e for s, e in comment_spans)

    def _line_indent(pos: int) -> int:
        line_start = 0
        for s in line_starts:
            if s <= pos:
                line_start = s
            else:
                break
        prefix = chunk[line_start:pos]
        return len(prefix) - len(prefix.lstrip(" \t"))

    def _replace(match: "re.Match[str]") -> str:
        if _in_comment(match.start()):
            return match.group(0)
        # 行首缩进 >= 4 空格是缩进代码块，pandoc 视为代码，不应转换
        if _line_indent(match.start()) >= 4:
            return match.group(0)
        table_md = _html_table_to_markdown(match.group(0))
        if not table_md:
            return match.group(0)
        # 前后补空行，确保表格独立成块（多余空行由 normalize_markdown 压缩）
        return f"\n\n{table_md}\n\n"

    return _TABLE_BLOCK_RE.sub(_replace, chunk)


def _html_table_to_markdown(html: str) -> Optional[str]:
    """把一个 HTML <table> 片段解析为 Markdown 管道表格。

    Returns:
        管道表格文本；解析失败或表格为空时返回 None（保持原样）。
    """
    try:
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find("table")
        if table is None:
            return None

        rows: List[List[str]] = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"], recursive=False)
            if not cells:
                cells = tr.find_all(["td", "th"])
            row = [_cell_text(cell) for cell in cells]
            if row:
                rows.append(row)

        # 去掉整行为空的间隔行（HTML 表格常用 colspan 空行做视觉分隔）
        rows = [row for row in rows if any(cell for cell in row)]
        if not rows:
            return None

        # GFM 管道表格必须有表头行；HTML 无 th 时以首行充当表头
        col_count = max(len(row) for row in rows)
        padded = [row + [""] * (col_count - len(row)) for row in rows]

        lines = []
        lines.append("| " + " | ".join(padded[0]) + " |")
        lines.append("| " + " | ".join(["---"] * col_count) + " |")
        for row in padded[1:]:
            lines.append("| " + " | ".join(row) + " |")
        return "\n".join(lines)
    except Exception:
        return None


def _cell_text(cell: "Tag") -> str:
    """提取单元格文本：换行折叠为空格，竖线转义以保持管道表格结构。"""
    for br in cell.find_all("br"):
        br.replace_with(" ")
    text = cell.get_text(" ", strip=True)
    text = " ".join(text.split())
    return text.replace("|", "\\|")
