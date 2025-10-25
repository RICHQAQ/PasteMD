"""Excel and WPS spreadsheet inserters."""

from typing import List
from .base import BaseTableInserter
from .formatting import CellFormat
from ...core.errors import InsertError
from ...utils.logging import log


class BaseExcelInserter(BaseTableInserter):
    """Excel 表格插入器基类"""
    
    def insert(self, table_data: List[List[str]], keep_format: bool = True) -> bool:
        """
        将表格数据插入到 Excel 当前光标位置
        
        Args:
            table_data: 二维数组表格数据
            keep_format: 是否保留 Markdown 格式（粗体、斜体等）
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            import pythoncom
            from pywintypes import com_error
            
            # 初始化 COM
            pythoncom.CoInitialize()
            
            try:
                # 获取 Excel 应用实例
                excel = self._get_excel_application()
            except Exception as e:
                raise InsertError(f"未找到运行中的 {self.app_name}，请先打开。错误: {e}")
            
            try:
                # 保存原始设置
                original_screen_updating = excel.ScreenUpdating
                original_calculation = excel.Calculation
                original_events = excel.EnableEvents

                # 优化性能禁用屏幕更新、自动计算和事件
                excel.ScreenUpdating = False
                excel.Calculation = -4135  # xlCalculationManual
                excel.EnableEvents = False

                try:
                    # 获取当前活动的工作表
                    sheet = excel.ActiveSheet

                    # 获取当前选中的单元格（起始位置）
                    start_cell = excel.ActiveCell

                    # 检查是否有活动单元格
                    if start_cell is None:
                        raise InsertError(f"未选中任何单元格，请在 {self.app_name} 中点击要插入表格的起始位置")

                    start_row = start_cell.Row
                    start_col = start_cell.Column

                    # 预处理数据：解析格式并准备批量数据
                    rows_count = len(table_data)
                    cols_count = max(len(row) for row in table_data) if table_data else 0

                    # 准备纯文本数据用于批量插入
                    clean_data = []
                    format_info = []  # 存储格式信息 [(row, col, cell_format, clean_text), ...]

                    for i, row in enumerate(table_data):
                        clean_row = []
                        for j, cell_value in enumerate(row):
                            if keep_format:
                                cell_format = CellFormat(cell_value)
                                clean_text = cell_format.parse()
                                clean_row.append(clean_text)

                                # 只有当单元格有格式时才记录
                                if (cell_format.has_newline or
                                    cell_format.is_code_block or
                                    len(cell_format.segments) > 1 or
                                    (len(cell_format.segments) == 1 and any([
                                        cell_format.segments[0].bold,
                                        cell_format.segments[0].italic,
                                        cell_format.segments[0].strikethrough,
                                        cell_format.segments[0].is_code,
                                        cell_format.segments[0].hyperlink_url
                                    ]))):
                                    format_info.append((i, j, cell_format, clean_text))
                            else:
                                clean_row.append(self._clean_markdown_formatting(cell_value))

                        # 补齐行长度
                        while len(clean_row) < cols_count:
                            clean_row.append('')
                        clean_data.append(clean_row)

                    # 批量写入数据（显著提升性能）
                    end_row = start_row + rows_count - 1
                    end_col = start_col + cols_count - 1
                    target_range = sheet.Range(
                        sheet.Cells(start_row, start_col),
                        sheet.Cells(end_row, end_col)
                    )
                    target_range.Value = clean_data

                    # 应用格式（如果需要）
                    if keep_format and format_info:
                        for i, j, cell_format, clean_text in format_info:
                            cell = sheet.Cells(start_row + i, start_col + j)

                            try:
                                # 如果包含换行,启用单元格自动换行
                                if cell_format.has_newline:
                                    cell.WrapText = True
                                
                                if cell_format.is_code_block:
                                    # 代码块使用等宽字体、浅灰色背景
                                    cell.Font.Name = "Consolas"
                                    cell.Interior.Color = 0xF0F0F0  # 浅灰色
                                    cell.WrapText = True
                                    # 设置垂直对齐为顶部
                                    cell.VerticalAlignment = -4160  # xlTop
                                else:
                                    # 检查是否整个单元格都是一个超链接
                                    if (len(cell_format.segments) == 1
                                            and cell_format.segments[0].hyperlink_url
                                            and not any([cell_format.segments[0].bold,
                                                        cell_format.segments[0].italic,
                                                        cell_format.segments[0].strikethrough])):
                                        # 单个超链接,使用 Hyperlinks.Add
                                        segment = cell_format.segments[0]
                                        try:
                                            sheet.Hyperlinks.Add(
                                                Anchor=cell,
                                                Address=segment.hyperlink_url,
                                                TextToDisplay=segment.text
                                            )
                                        except com_error as e:
                                            log(f"Failed to add hyperlink: {e}")
                                    else:
                                        char_index = 1  # Excel 字符索引从1开始
                                        
                                        # 检查是否有超链接(有超链接时不能使用 GetCharacters)
                                        has_hyperlink = any(seg.hyperlink_url for seg in cell_format.segments)
                                        
                                        if has_hyperlink:
                                            # 有超链接时,只设置文本,不设置富文本格式
                                            # 因为 Hyperlinks.Add 和 GetCharacters 不兼容
                                            for segment in cell_format.segments:
                                                if segment.hyperlink_url:
                                                    # 为链接部分添加超链接
                                                    # 注意: Excel 单元格只能有一个超链接,这里取第一个
                                                    try:
                                                        sheet.Hyperlinks.Add(
                                                            Anchor=cell,
                                                            Address=segment.hyperlink_url,
                                                            TextToDisplay=clean_text
                                                        )
                                                        break
                                                    except com_error as e:
                                                        log(f"Failed to add hyperlink: {e}")
                                        else:
                                            # 没有超链接,可以使用富文本格式
                                            for segment in cell_format.segments:
                                                if not segment.text:
                                                    continue
                                                
                                                seg_len = len(segment.text)
                                                # 获取字符范围
                                                chars = cell.GetCharacters(char_index, seg_len)
                                                
                                                # 应用格式
                                                if segment.is_code:
                                                    chars.Font.Name = "Consolas"
                                                if segment.bold:
                                                    chars.Font.Bold = True
                                                if segment.italic:
                                                    chars.Font.Italic = True
                                                if segment.strikethrough:
                                                    chars.Font.Strikethrough = True
                                                
                                                char_index += seg_len
                                        
                                        # 如果有行内代码,设置整个单元格背景
                                        if any(seg.is_code for seg in cell_format.segments):
                                            cell.Interior.Color = 0xF0F0F0  # 浅灰色
                                        
                            except com_error as e:
                                # 格式应用失败，记录但继续
                                log(f"Failed to apply format to cell ({i},{j}): {e}")

                    # 选中插入的区域
                    target_range.Select()

                    log(f"Successfully inserted table to {self.app_name}: {rows_count} rows x {cols_count} cols, keep_format={keep_format}")
                    return True

                finally:
                    # 恢复原始设置
                    excel.ScreenUpdating = original_screen_updating
                    excel.Calculation = original_calculation
                    excel.EnableEvents = original_events

            finally:
                pythoncom.CoUninitialize()
                
        except InsertError:
            raise
        except Exception as e:
            log(f"Failed to insert table to {self.app_name}: {e}")
            raise InsertError(f"{self.app_name} 插入失败: {e}")
    
    def _get_application(self):
        """
        获取 Excel 应用程序实例（尝试所有可能的 ProgID）
        
        Returns:
            Excel 应用程序对象
            
        Raises:
            Exception: 无法获取实例时
        """
        import win32com.client
        
        # 尝试所有可能的 ProgID
        for prog_id in self.prog_ids:
            try:
                # 尝试连接现有实例
                excel = win32com.client.GetActiveObject(prog_id)
                log(f"Successfully connected to {prog_id}")
                return excel
            except Exception as e:
                log(f"Failed to connect to {prog_id}: {e}")
                continue
        
        raise Exception(f"未找到运行中的 {self.app_name}，请先打开")
    
    # 保持向后兼容的别名
    _get_excel_application = _get_application
    
    def _clean_markdown_formatting(self, text: str) -> str:
        """
        清除 Markdown 格式符号(使用字符级解析)
        
        Args:
            text: 包含 Markdown 格式的文本
            
        Returns:
            清除格式后的纯文本
        """
        # 直接使用 CellFormat 来解析
        cell_format = CellFormat(text)
        return cell_format.parse()

    def _refalsh_app(self) -> object:
        """
        刷新应用程序状态（如果需要）
        
        Returns:
            刷新后的应用程序对象
        """
        return self._get_application()
    
    
class MSExcelInserter(BaseExcelInserter):
    """Microsoft Excel 插入器"""
    
    def __init__(self):
        super().__init__(prog_id="Excel.Application", app_name="Excel")



