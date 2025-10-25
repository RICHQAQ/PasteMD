"""Spreadsheet insertion domain."""

from .excel import MSExcelInserter
from .wps_excel import WPSExcelInserter

__all__ = ["MSExcelInserter", "WPSExcelInserter"]
