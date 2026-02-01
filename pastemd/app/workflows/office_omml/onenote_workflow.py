# -*- coding: utf-8 -*-
"""OneNote workflow with OMML formula support."""

from .office_omml_base import OfficeOmmlBaseWorkflow


class OneNoteWorkflow(OfficeOmmlBaseWorkflow):
    """OneNote 工作流

    支持粘贴 Markdown/HTML 内容到 OneNote，
    数学公式自动转换为可编辑的 Office 公式。
    """

    @property
    def app_name(self) -> str:
        return "OneNote"
