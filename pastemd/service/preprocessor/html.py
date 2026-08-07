"""HTML content preprocessor."""

import re

from bs4 import BeautifulSoup, NavigableString, Tag
from .base import BasePreprocessor
from ...utils.html_formatter import (
    clean_html_content,
    convert_css_font_to_semantic,
    convert_strikethrough_to_del,
    promote_bold_first_row_to_header,
)
from ...utils.logging import log


OBSIDIAN_CLIPBOARD_MARKER = "<!-- obsidian -->"
PRESERVE_NEWLINE_WHITE_SPACE_RE = re.compile(
    r"(?:^|;)\s*white-space\s*:\s*(?:pre-wrap|pre-line|break-spaces)\b",
    re.IGNORECASE,
)
NEWLINE_EXCLUDED_TAGS = {"script", "style", "textarea", "pre", "code"}


def _restore_data_math_source_latex(soup: BeautifulSoup) -> None:
    """Restore LaTeX from clipboard math nodes that expose their source as data."""
    for tag in soup.select('[role="math"][data-math-source]'):
        source = tag.get("data-math-source")
        if not isinstance(source, str) or not source.strip():
            continue

        source = source.strip()
        is_display = tag.select_one(".katex-display") is not None
        if is_display:
            latex = (
                source
                if source.startswith("$$") and source.endswith("$$")
                else f"$$\n{source}\n$$"
            )
        else:
            latex = source if source.startswith("$") and source.endswith("$") else f"${source}$"
        tag.replace_with(NavigableString(latex))


def _wrap_obsidian_math_latex(soup: BeautifulSoup, html: str) -> None:
    """Restore LaTeX delimiters for Obsidian clipboard math nodes."""
    if OBSIDIAN_CLIPBOARD_MARKER not in html:
        return

    for tag in soup.select("span.math.math-inline"):
        text = tag.get_text().strip()
        if not text:
            continue
        latex = text if text.startswith("$") else f"${text}$"
        tag.replace_with(latex)

    for tag in soup.select(".math.math-block"):
        text = tag.get_text().strip()
        if not text:
            continue
        latex = text if text.startswith("$$") else f"$${text}$$"
        tag.replace_with(latex)


def _convert_preserved_newlines_to_br(soup: BeautifulSoup) -> None:
    """Make CSS-preserved text newlines explicit before Pandoc consumes HTML."""
    for tag in soup.find_all(style=PRESERVE_NEWLINE_WHITE_SPACE_RE):
        if not isinstance(tag, Tag):
            continue
        for text_node in list(tag.find_all(string=True)):
            parent = text_node.parent
            if not parent or any(
                ancestor.name and ancestor.name.lower() in NEWLINE_EXCLUDED_TAGS
                for ancestor in text_node.parents
            ):
                continue

            text = str(text_node).replace("\r\n", "\n").replace("\r", "\n")
            if "\n" not in text:
                continue

            replacement = []
            parts = text.split("\n")
            for index, part in enumerate(parts):
                if part:
                    replacement.append(NavigableString(part))
                if index < len(parts) - 1:
                    replacement.append(soup.new_tag("br"))

            text_node.replace_with(*replacement)


class HtmlPreprocessor(BasePreprocessor):
    """HTML 内容预处理器（无状态）"""

    def process(self, html: str, config: dict) -> str:
        """
        预处理 HTML 内容

        处理步骤:
        1. 清理无效元素（SVG等）
        2. 转换删除线标记
        3. 清理 LaTeX 公式块中的 br 标签
        4. 其他自定义处理...

        Args:
            html: 原始 HTML 内容
            config: 配置字典

        Returns:
            预处理后的 HTML 内容
        """
        log("Preprocessing HTML content")

        # 使用 html_formatter 进行清理
        soup = BeautifulSoup(html, "html.parser")
        _restore_data_math_source_latex(soup)
        _wrap_obsidian_math_latex(soup, html)
        clean_html_content(soup, config)
        _convert_preserved_newlines_to_br(soup)

        html_formatting = config.get("html_formatting") or config.get("Html_formatting") or {}
        if not isinstance(html_formatting, dict):
            html_formatting = {}
        if html_formatting.get("strikethrough_to_del", True):
            convert_strikethrough_to_del(soup)
        if html_formatting.get("css_font_to_semantic", True):
            convert_css_font_to_semantic(soup)
        if html_formatting.get("bold_first_row_to_header", False):
            promote_bold_first_row_to_header(soup)

        # unwrap_li_paragraphs(soup)
        # remove_empty_paragraphs(soup)

        html_output = str(soup)
        
        # 仅在 HTML 不包含 DOCTYPE 时才添加
        if "<!DOCTYPE" not in html_output.upper():
            html_output = f"<!DOCTYPE html>\n<meta charset='utf-8'>\n{html_output}"
        
        return html_output
