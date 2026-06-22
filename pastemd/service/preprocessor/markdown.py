"""Markdown content preprocessor."""

from .base import BasePreprocessor
from ...utils.md_normalizer import fence_plain_text_code, normalize_markdown
from ...utils.latex import convert_latex_delimiters
from ...utils.logging import log


class MarkdownPreprocessor(BasePreprocessor):
    """Markdown 内容预处理器（无状态）"""

    def process(self, markdown: str, config: dict) -> str:
        """
        预处理 Markdown 内容

        处理步骤:
        1. 标准化 Markdown 语法
        2. 处理 LaTeX 数学公式
        3. 其他自定义处理...

        Args:
            markdown: 原始 Markdown 文本
            config: 配置字典

        Returns:
            预处理后的 Markdown 文本
        """
        log("Preprocessing Markdown content")

        markdown = fence_plain_text_code(markdown)

        # 1. 标准化 Markdown
        if config.get("normalize_markdown", True):
            markdown = normalize_markdown(markdown)

        # 2. 处理 LaTeX
        if config.get("latex_support", True):
            fix_single_dollar_block = config.get("fix_single_dollar_block", True)
            markdown = convert_latex_delimiters(markdown, fix_single_dollar_block)

        # 未来可扩展其他处理...

        return markdown
