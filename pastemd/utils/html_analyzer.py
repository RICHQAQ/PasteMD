"""Helpers for inspecting clipboard HTML fragments before conversion."""

from __future__ import annotations

from typing import Iterable, Set

try:
    from bs4 import BeautifulSoup, FeatureNotFound  # type: ignore
except Exception:  # pragma: no cover - BeautifulSoup is in requirements
    BeautifulSoup = None  # type: ignore
    FeatureNotFound = None  # type: ignore


# HTML 标签中能提供语义结构的元素集合
SEMANTIC_TAGS: Set[str] = {
    "p",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "ul",
    "ol",
    "li",
    "dl",
    "dt",
    "dd",
    "table",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "th",
    "td",
    "col",
    "colgroup",
    "pre",
    "code",
    "blockquote",
    "figure",
    "figcaption",
    "math",
    "section",
    "article",
    "header",
    "footer",
    "aside",
    "nav",
    "hr",
}

# 复制按钮常见的包裹标签（通常不包含真实结构）
INLINE_WRAPPER_TAGS: Set[str] = {
    "span",
    "font",
    "strong",
    "em",
    "b",
    "i",
    "u",
    "sub",
    "sup",
    "s",
    "del",
    "mark",
    "a",
}

# Markdown 语法特征，用于辅助判断 HTML 是否只是 Markdown 文本
MARKDOWN_HINTS: Iterable[str] = (
    "\n#",
    "\n##",
    "\n- ",
    "\n* ",
    "\n1.",
    "```",
    "**",
    "__",
    "~~",
    "> ",
    "$$",
    "\\(",
    "\\)",
    "|",
    "\n---",
    "\n***",
    "`",
)


def _count_semantic_tags(html_soup) -> int:
    """统计 HTML 中带有语义结构的标签数量。"""
    body = html_soup.body or html_soup
    count = 0
    for tag in body.find_all(True):
        name = tag.name.lower()
        if name in SEMANTIC_TAGS:
            count += 1
    return count


def _only_contains_inline_wrappers(html_soup) -> bool:
    """判断 HTML 是否只包含 wrapper / inline 标签。"""
    body = html_soup.body or html_soup
    for tag in body.find_all(True):
        name = tag.name.lower()
        if name in ("html", "head", "body", "meta", "style"):
            continue
        if name not in INLINE_WRAPPER_TAGS:
            return False
    return True


def _markdown_hint_score(text: str) -> int:
    """根据 Markdown 语法特征粗略打分。"""
    score = 0
    for hint in MARKDOWN_HINTS:
        if hint in text:
            score += 1
    return score


def is_plain_html_fragment(html: str) -> bool:
    """
    判断 HTML 片段是否只是带壳的 Markdown / 纯文本。

    复制按钮经常返回“只有 span + 内联样式 + 纯文本”的 HTML，
    如果直接走 Pandoc HTML 流程会把 Markdown 符号原样贴进 Word。
    这里通过结构标签数量、内联标签检测、以及 Markdown 语法特征
    来辅助判断是否应该退回 Markdown 流程。
    """
    if not html or not html.strip():
        return True

    if BeautifulSoup is None:  # pragma: no cover - fallback
        # 简单兜底：没有解析器时，只要看不到典型结构标签就视为纯文本
        lowered = html.lower()
        return not any(tag in lowered for tag in ("<p", "<h1", "<ul", "<table", "<pre", "<code", "<blockquote"))

    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception as exc:  # pragma: no cover - depends on env parser
        if FeatureNotFound is not None and isinstance(exc, FeatureNotFound):
            soup = BeautifulSoup(html, "html.parser")
        else:
            soup = BeautifulSoup(html, "html.parser")
    semantic_count = _count_semantic_tags(soup)

    if semantic_count > 0:
        return False

    if _only_contains_inline_wrappers(soup):
        return True

    body = soup.body or soup
    text = body.get_text(separator="\n").strip()
    if not text:
        return True

    hint_score = _markdown_hint_score(text)
    # 当 HTML 中没有语义标签，但文本里充满 Markdown 符号时，也视作纯文本
    return hint_score >= 2
