"""Markdown table parser."""

import csv
import io
import re
from typing import List, Optional


def _split_table_cells(line: str) -> List[str]:
    """
    按 | 分割表格单元格,正确处理转义的竖线
    
    Args:
        line: 表格行文本
        
    Returns:
        单元格列表
    """
    cells = []
    current_cell = []
    i = 0
    
    while i < len(line):
        if i > 0 and line[i] == '|' and line[i - 1] == '\\':
            # 转义的竖线,替换前一个反斜杠并添加竖线
            current_cell[-1] = '|'
            i += 1
        elif line[i] == '|':
            # 未转义的竖线,分割单元格
            cells.append(''.join(current_cell).strip())
            current_cell = []
            i += 1
        else:
            current_cell.append(line[i])
            i += 1
    
    # 添加最后一个单元格
    if current_cell or cells:  # 如果有内容或已经有单元格
        cells.append(''.join(current_cell).strip())
    
    return cells


def parse_markdown_table(md_text: str) -> Optional[List[List[str]]]:
    """
    解析 Markdown 表格为二维数组
    
    Args:
        md_text: Markdown 文本内容
        
    Returns:
        二维数组，每个元素代表一行的单元格内容；如果不是表格则返回 None
    """
    lines = md_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    table_data = []
    separator_found = False
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # 跳过空行
        if not line:
            continue
            
        # 检查是否为表格行（以 | 开头或结尾）
        if not (line.startswith('|') or line.endswith('|') or '|' in line):
            # 如果已经找到分隔符，说明表格结束
            if separator_found:
                break
            # 否则不是表格
            return None
        
        # 检查是否为分隔符行（如 |---|---|）
        if re.match(r'^\s*\|?\s*[-:]+\s*(\|\s*[-:]+\s*)+\|?\s*$', line):
            separator_found = True
            continue
        
        # 使用新的分割方法解析单元格
        cells = _split_table_cells(line)
        
        # 移除首尾的空元素（如果行是 |a|b| 格式）
        if cells and cells[0] == '':
            cells = cells[1:]
        if cells and cells[-1] == '':
            cells = cells[:-1]
        
        if cells:
            table_data.append(cells)
    
    # 必须找到分隔符才认为是有效表格
    if not separator_found or not table_data:
        return None
    
    return table_data


def parse_csv_table(text: str) -> Optional[List[List[str]]]:
    """
    解析 CSV 格式文本为二维数组
    
    Args:
        text: CSV 格式文本内容
        
    Returns:
        二维数组，每个元素代表一行的单元格内容；如果不是有效 CSV 则返回 None
    """
    if not text or not text.strip():
        return None
    
    lines = text.strip().split('\n')
    # 至少需要一行数据
    if len(lines) < 1:
        return None
    
    # 检查是否包含逗号（CSV 的基本特征）
    if ',' not in text:
        return None
    
    try:
        reader = csv.reader(io.StringIO(text))
        table_data = []
        first_row_cols = None
        
        for row in reader:
            # 跳过空行
            if not row or all(cell.strip() == '' for cell in row):
                continue
            
            # 记录第一行的列数
            if first_row_cols is None:
                first_row_cols = len(row)
                # 至少需要 2 列才认为是表格
                if first_row_cols < 2:
                    return None
            
            # 去除每个单元格的首尾空白
            cleaned_row = [cell.strip() for cell in row]
            table_data.append(cleaned_row)
        
        # 至少需要 1 行数据
        if len(table_data) < 1:
            return None
        
        return table_data
        
    except csv.Error:
        return None


def parse_table(text: str) -> Optional[List[List[str]]]:
    """
    统一的表格解析入口，先尝试 Markdown 格式，再尝试 CSV 格式
    
    Args:
        text: 表格文本内容
        
    Returns:
        二维数组，每个元素代表一行的单元格内容；如果无法解析则返回 None
    """
    if not text:
        return None
    
    # 先尝试 Markdown 表格
    result = parse_markdown_table(text)
    if result:
        return result
    
    # 再尝试 CSV 格式
    result = parse_csv_table(text)
    if result:
        return result
    
    return None

