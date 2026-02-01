# -*- coding: utf-8 -*-
"""Office OMML workflows for OneNote and PowerPoint.

These workflows support pasting Markdown/HTML with LaTeX formulas as
native Office equations using OMML (Office MathML).
"""

from .office_omml_base import OfficeOmmlBaseWorkflow
from .onenote_workflow import OneNoteWorkflow
from .powerpoint_workflow import PowerPointWorkflow

__all__ = [
    "OfficeOmmlBaseWorkflow",
    "OneNoteWorkflow",
    "PowerPointWorkflow",
]
