# -*- coding: utf-8 -*-
"""PowerPoint workflow with OMML formula support."""

from pastemd.utils.omml import convert_html_mathml_to_omml

from .office_omml_base import OfficeOmmlBaseWorkflow


class PowerPointWorkflow(OfficeOmmlBaseWorkflow):
    """PowerPoint 工作流

    支持粘贴 Markdown/HTML 内容到 PowerPoint，
    数学公式自动转换为可编辑的 Office 公式。
    """

    @property
    def app_name(self) -> str:
        return "PowerPoint"

    def _convert_html_mathml_to_omml(self, html_body: str) -> str:
        return convert_html_mathml_to_omml(html_body, skip_table_mathml=True)
