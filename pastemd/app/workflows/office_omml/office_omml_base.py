# -*- coding: utf-8 -*-
"""Base workflow for Office apps with OMML formula support."""

from __future__ import annotations

from abc import ABC, abstractmethod

from pastemd.app.workflows.base import BaseWorkflow
from pastemd.core.errors import ClipboardError, PandocError
from pastemd.i18n import t
from pastemd.service.paste import RichTextPastePlacer
from pastemd.utils.clipboard import (
    get_clipboard_html,
    get_clipboard_text,
    is_clipboard_empty,
    read_markdown_files_from_clipboard,
)
from pastemd.utils.html_analyzer import is_plain_html_fragment
from pastemd.utils.markdown_utils import merge_markdown_contents
from pastemd.utils.html_formatter import extract_html_body
from pastemd.utils.omml import convert_html_mathml_to_omml, generate_office_html


class OfficeOmmlBaseWorkflow(BaseWorkflow, ABC):
    """Office OMML 工作流基类（OneNote/PowerPoint）。
    
    处理流程：
    1. 读取剪贴板 HTML/Markdown
    2. 转换为带 MathML 的 HTML（Pandoc --mathml）
    3. 将 MathML 替换为 OMML 条件注释
    4. 使用 RichTextPastePlacer 粘贴
    """

    def __init__(self):
        super().__init__()
        self._placer = RichTextPastePlacer()

    @property
    @abstractmethod
    def app_name(self) -> str:
        """Application display name."""
        ...

    @property
    def placer(self) -> RichTextPastePlacer:
        return self._placer

    def execute(self) -> None:
        content_type: str | None = None
        from_md_file = False
        md_file_count = 0

        try:
            content_type, content, from_md_file, md_file_count = self._read_clipboard()
            self._log(f"Clipboard content type: {content_type}")

            # Preprocess content
            if content_type == "markdown":
                content = self.markdown_preprocessor.process(content, self.config)
            elif content_type == "html":
                content = self.html_preprocessor.process(content, self.config)

            # Convert to HTML with MathML
            if content_type == "html":
                md_text = self.doc_generator.convert_html_to_markdown_text(
                    content, self.config
                )
            else:
                md_text = content

            # Generate HTML with MathML formulas (Pandoc --mathml is enabled)
            html_with_mathml = self.doc_generator.convert_markdown_to_html_text(
                md_text, self.config
            )

            # Strip standalone HTML wrapper to avoid nested documents
            html_body = extract_html_body(html_with_mathml)

            # Convert MathML to OMML conditional comments
            html_with_omml = self._convert_html_mathml_to_omml(html_body)

            # Wrap in Office HTML template
            office_html = generate_office_html(html_with_omml)

            # Place using rich text placer
            result = self.placer.place(
                content=md_text,
                config=self.config,
                html=office_html,
            )

            if result.success:
                if from_md_file:
                    if md_file_count > 1:
                        msg = t(
                            "workflow.md_file.insert_success_multi",
                            count=md_file_count,
                            app=self.app_name,
                        )
                    else:
                        msg = t("workflow.md_file.insert_success", app=self.app_name)
                elif content_type == "html":
                    msg = t("workflow.html.insert_success", app=self.app_name)
                else:
                    msg = t("workflow.word.insert_success", app=self.app_name)

                self._notify_success(msg)
            else:
                self._notify_error(result.error or t("workflow.generic.failure"))

        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            if content_type == "html":
                self._notify_error(t("workflow.html.convert_failed_generic"))
            else:
                self._notify_error(t("workflow.markdown.convert_failed"))
        except Exception as e:
            self._log(f"{self.app_name} workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))

    def _read_clipboard(self) -> tuple[str, str, bool, int]:
        """读取剪贴板,返回 (类型, 内容, 是否来自 MD 文件, MD 文件数量)"""
        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return ("html", html, False, 0)
        except ClipboardError:
            pass

        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            merged = merge_markdown_contents(files_data)
            return ("markdown", merged, True, len(files_data))

        if not is_clipboard_empty():
            return ("markdown", get_clipboard_text(), False, 0)

        raise ClipboardError("剪贴板为空或无有效内容")

    def _convert_html_mathml_to_omml(self, html_body: str) -> str:
        return convert_html_mathml_to_omml(html_body)
