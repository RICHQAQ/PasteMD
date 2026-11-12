# -*- coding: utf-8 -*-
"""File paste workflow for apps that accept file paste."""

from .extensible_base import ExtensibleWorkflow
from ....core.errors import ClipboardError, PandocError
from ....core.types import PlacementResult
from ....utils.clipboard import (
    get_clipboard_text,
    get_clipboard_html,
    is_clipboard_empty,
    read_markdown_files_from_clipboard,
)
from ....utils.html_analyzer import is_plain_html_fragment
from ....utils.markdown_utils import merge_markdown_contents
from ....service.spreadsheet.parser import parse_markdown_table
from ....service.spreadsheet.generator import SpreadsheetGenerator
from ....utils.fs import generate_output_path
from ....i18n import t
from ....service.paste import FilePastePlacer


class FileWorkflow(ExtensibleWorkflow):
    """文件粘贴工作流

    适用于支持文件粘贴/附件粘贴的应用：
    - 读取剪贴板 HTML/Markdown
    - 转换为 DOCX 或 XLSX
    - 复制文件到剪贴板并模拟粘贴
    """

    def __init__(self):
        super().__init__()
        self.placer = FilePastePlacer()

    @property
    def workflow_key(self) -> str:
        return "file"

    def execute(self) -> None:
        """执行文件粘贴工作流"""
        content_type: str | None = None
        try:
            content_type = self._detect_content_type()
            self._log(f"File workflow: content_type={content_type}")

            success_msg = t("workflow.file.paste_success")
            result: PlacementResult | None = None

            if content_type == "table":
                markdown_text, _ = self._read_markdown_content()
                table_data = parse_markdown_table(markdown_text)
                keep_format = self.config.get(
                    "excel_keep_format", self.config.get("keep_format", True)
                )
                xlsx_bytes = SpreadsheetGenerator.generate_xlsx_bytes(
                    table_data, keep_format=keep_format
                )
                output_path = generate_output_path(
                    keep_file=self.config.get("keep_file", False),
                    save_dir=self.config.get("save_dir", ""),
                    table_data=table_data,
                )
                self._write_output(output_path, xlsx_bytes)
                result = self.placer.place(
                    content=output_path,
                    config=self.config,
                    file_paths=[output_path],
                )
            else:
                html_text = ""
                md_text = ""
                source_filenames: list[str] = []
                if content_type == "html":
                    html_text = get_clipboard_html(self.config)
                    html_text = self.html_preprocessor.process(html_text, self.config)
                    docx_bytes = self.doc_generator.convert_html_to_docx_bytes(
                        html_text, self.config
                    )
                else:
                    _, files_data, read_errors = read_markdown_files_from_clipboard()
                    # 触发条件用“剪贴板里检测到的 md 文件数”（成功读取 + 读取失败）：
                    # 否则部分文件读失败时会静默退回合并模式，且失败文件被丢弃
                    if (
                        len(files_data) + len(read_errors) > 1
                        and self.config.get("md_multi_file_mode", "merge") == "separate"
                    ):
                        result, success_msg = self._paste_separate_documents(
                            files_data, read_errors
                        )
                    else:
                        md_text, source_filenames = self._read_markdown_content()
                        md_text = self.markdown_preprocessor.process(md_text, self.config)
                        docx_bytes = self.doc_generator.convert_markdown_to_docx_bytes(
                            md_text, self.config
                        )

                if result is None:
                    output_path = generate_output_path(
                        keep_file=True,
                        save_dir=self.config.get("save_dir", ""),
                        md_text=md_text,
                        html_text=html_text,
                        source_filenames=source_filenames,
                        md_name_mode=self.config.get("md_file_output_name_mode", "content"),
                    )
                    self._write_output(output_path, docx_bytes)
                    result = self.placer.place(
                        content=output_path,
                        config=self.config,
                        file_paths=[output_path],
                    )

            if result.success:
                self._notify_success(success_msg)
            else:
                self._notify_error(result.error or t("workflow.action.clipboard_failed"))

        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            msg = str(e)
            if "为空" in msg:
                self._notify_error(t("workflow.clipboard.empty"))
            else:
                self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            if content_type == "html":
                self._notify_error(t("workflow.html.convert_failed_generic"))
            else:
                self._notify_error(t("workflow.markdown.convert_failed"))
        except Exception as e:
            self._log(f"File workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))

    def _detect_content_type(self) -> str:
        """
        检测剪贴板内容类型

        Returns:
            "table" | "html" | "markdown"
        """
        if is_clipboard_empty():
            found, _, _ = read_markdown_files_from_clipboard()
            if not found:
                raise ClipboardError("剪贴板为空")

        markdown_text = ""
        if not is_clipboard_empty():
            markdown_text = get_clipboard_text()
        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            markdown_text = merge_markdown_contents(files_data)

        table_data = parse_markdown_table(markdown_text) if markdown_text else None
        if table_data:
            return "table"

        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return "html"
        except ClipboardError:
            pass

        return "markdown"

    def _paste_separate_documents(
        self,
        files_data: list[tuple[str, str]],
        read_errors: list[tuple[str, str]],
    ) -> tuple[PlacementResult, str]:
        """每个 MD 文件单独生成一份文档，再作为多个文件一起粘贴（多文件拆分模式）

        输出位置跟随“保留生成的临时文件”：勾选则写保存目录，否则写临时目录。

        Returns:
            (粘贴结果, 成功文案)
        """
        keep_file = self.config.get("keep_file", False)
        save_dir = self.config.get("save_dir", "")
        name_mode = self.config.get("md_file_output_name_mode", "content")

        output_paths: list[str] = []
        failures: list[tuple[str, str]] = list(read_errors)

        for filename, raw_content in files_data:
            try:
                md_text = self.markdown_preprocessor.process(raw_content, self.config)
                docx_bytes = self.doc_generator.convert_markdown_to_docx_bytes(
                    md_text, self.config
                )
                output_path = generate_output_path(
                    keep_file=keep_file,
                    save_dir=save_dir,
                    md_text=md_text,
                    source_filenames=[filename],
                    md_name_mode=name_mode,
                )
                self._write_output(output_path, docx_bytes)
                output_paths.append(output_path)
            except Exception as e:
                self._log(f"Failed to convert MD file '{filename}': {e}")
                failures.append((filename, str(e)))

        if not output_paths:
            return (
                PlacementResult(
                    success=False,
                    method="clipboard_file",
                    error=t(
                        "workflow.md_file.batch_failed_all",
                        failed_count=len(failures),
                    ),
                ),
                "",
            )

        result = self.placer.place(
            content=output_paths[0],
            config=self.config,
            file_paths=output_paths,
        )

        success_msg = t("workflow.file.paste_success_multi", count=len(output_paths))
        if failures:
            success_msg += "\n" + t(
                "workflow.md_file.batch_failure_line",
                failed_count=len(failures),
                failed_files=", ".join(name for name, _ in failures),
            )

        return result, success_msg

    def _read_markdown_content(self) -> tuple[str, list[str]]:
        """读取 Markdown 内容（含剪贴板文件）

        Returns:
            (Markdown 文本, 来源 MD 文件名列表)；来源为剪贴板文本时文件名列表为空
        """
        if not is_clipboard_empty():
            content = get_clipboard_text()
        else:
            content = ""

        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            return merge_markdown_contents(files_data), [name for name, _ in files_data]

        if content.strip():
            return content, []

        raise ClipboardError("剪贴板为空或无有效内容")

    def _write_output(self, output_path: str, data: bytes) -> None:
        """写入文件"""
        with open(output_path, "wb") as f:
            f.write(data)
