"""Cell formatting utilities for spreadsheet insertion."""

import re
from typing import List, Optional


class TextSegment:
    """文本片段,带有格式信息"""
    def __init__(self, text: str, bold: bool = False, italic: bool = False,
                 strikethrough: bool = False, is_code: bool = False,
                 hyperlink_url: Optional[str] = None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.strikethrough = strikethrough
        self.is_code = is_code
        self.hyperlink_url = hyperlink_url  # 如果非空,表示这是一个超链接


class CellFormat:
    """单元格格式信息"""
    def __init__(self, text: str):
        self.text = text
        self.is_code_block = False
        self.has_newline = False
        self.segments = []  # List[TextSegment]
        self.clean_text = text
    
    def parse(self) -> str:
        """解析 Markdown 格式并生成文本片段(字符级解析)"""
        text = self.text
        
        # 处理 HTML 标签和换行
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        if '\n' in text:
            self.has_newline = True
        
        # 检查是否包含代码块标签
        if '<pre>' in text.lower() or '<code>' in text.lower():
            self.is_code_block = True
            # 提取代码块内容
            text = re.sub(r'<pre>(.*?)</pre>',
                          lambda m: re.sub(r'<br\s*/?>', '\n', m.group(1), flags=re.IGNORECASE),
                          text, flags=re.DOTALL | re.IGNORECASE)
            text = re.sub(r'<code>(.*?)</code>',
                          lambda m: re.sub(r'<br\s*/?>', '\n', m.group(1), flags=re.IGNORECASE),
                          text, flags=re.DOTALL | re.IGNORECASE)
            self.clean_text = text.strip()
            self.segments = [TextSegment(self.clean_text, is_code=True)]
            return self.clean_text
        
        # 解析为文本片段
        self.segments = self._parse_segments(text)
        self.clean_text = ''.join(seg.text for seg in self.segments)
        return self.clean_text
    
    def _parse_segments(self, text: str, bold: bool = False, italic: bool = False,
                        strikethrough: bool = False) -> List[TextSegment]:
        """
        解析文本为带格式的片段列表
        
        Args:
            text: 要解析的文本
            bold: 当前是否在粗体环境中
            italic: 当前是否在斜体环境中
            strikethrough: 当前是否在删除线环境中
        """
        segments = []
        i = 0
        current_text = []
        
        def flush_current():
            """将当前累积的文本保存为片段"""
            if current_text:
                text_str = ''.join(current_text)
                if text_str:
                    segments.append(TextSegment(text_str, bold, italic, strikethrough))
                current_text.clear()
        
        while i < len(text):
            # 处理转义字符
            if text[i] == '\\' and i + 1 < len(text):
                current_text.append(text[i + 1])
                i += 2
                continue
            
            # 检测行内代码 `...`
            if text[i] == '`':
                end = text.find('`', i + 1)
                if end != -1:
                    flush_current()
                    # 代码块内的文本不做任何解析
                    code_text = text[i + 1:end]
                    segments.append(TextSegment(code_text, is_code=True))
                    i = end + 1
                    continue
            
            # 检测删除线 ~~...~~
            if text[i:i + 2] == '~~' and not strikethrough:
                end = text.find('~~', i + 2)
                if end != -1:
                    flush_current()
                    inner = text[i + 2:end]
                    # 递归解析,传递删除线状态
                    segments.extend(self._parse_segments(inner, bold, italic, True))
                    i = end + 2
                    continue
            
            # 检测粗斜体 ***...*** (必须在 ** 之前检测)
            if text[i:i + 3] == '***' and not bold and not italic:
                end = text.find('***', i + 3)
                if end != -1:
                    flush_current()
                    inner = text[i + 3:end]
                    # 递归解析,同时传递粗体和斜体状态
                    segments.extend(self._parse_segments(inner, True, True, strikethrough))
                    i = end + 3
                    continue
            
            # 检测粗体 **...** 或 __...__
            if text[i:i + 2] == '**' and not bold:
                # 查找匹配的 **,需要跳过单个 *
                end = i + 2
                while end < len(text) - 1:
                    if text[end:end + 2] == '**':
                        flush_current()
                        inner = text[i + 2:end]
                        # 递归解析,传递粗体状态
                        segments.extend(self._parse_segments(inner, True, italic, strikethrough))
                        i = end + 2
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测粗斜体 ___...___ (必须在 __ 之前检测)
            if text[i:i + 3] == '___' and not bold and not italic:
                end = text.find('___', i + 3)
                if end != -1:
                    flush_current()
                    inner = text[i + 3:end]
                    # 递归解析,同时传递粗体和斜体状态
                    segments.extend(self._parse_segments(inner, True, True, strikethrough))
                    i = end + 3
                    continue
            
            if text[i:i + 2] == '__' and not bold:
                # 查找匹配的 __,需要跳过单个 _
                end = i + 2
                while end < len(text) - 1:
                    if text[end:end + 2] == '__':
                        flush_current()
                        inner = text[i + 2:end]
                        segments.extend(self._parse_segments(inner, True, italic, strikethrough))
                        i = end + 2
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测斜体 *...* 或 _..._ (移除 not italic 限制,允许在粗体内使用斜体)
            if text[i] == '*' and (i + 1 >= len(text) or text[i + 1] != '*'):
                end = i + 1
                # 查找匹配的结束星号
                while end < len(text):
                    if text[end] == '*' and (end + 1 >= len(text) or text[end + 1] != '*'):
                        flush_current()
                        inner = text[i + 1:end]
                        # 递归解析,传递斜体状态
                        segments.extend(self._parse_segments(inner, bold, True, strikethrough))
                        i = end + 1
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            if text[i] == '_' and (i + 1 >= len(text) or text[i + 1] != '_'):
                end = i + 1
                while end < len(text):
                    if text[end] == '_' and (end + 1 >= len(text) or text[end + 1] != '_'):
                        flush_current()
                        inner = text[i + 1:end]
                        segments.extend(self._parse_segments(inner, bold, True, strikethrough))
                        i = end + 1
                        break
                    end += 1
                else:
                    # 没找到配对的,当普通字符处理
                    current_text.append(text[i])
                    i += 1
                continue
            
            # 检测链接 [text](url)
            if text[i] == '[':
                close_bracket = text.find(']', i + 1)
                if close_bracket != -1 and close_bracket + 1 < len(text) and text[close_bracket + 1] == '(':
                    close_paren = text.find(')', close_bracket + 2)
                    if close_paren != -1:
                        flush_current()
                        # 提取链接文本和URL
                        link_text = text[i + 1:close_bracket]
                        link_url = text[close_bracket + 2:close_paren]
                        # 解析链接文本的格式,并添加超链接URL
                        link_segments = self._parse_segments(link_text, bold, italic, strikethrough)
                        # 为每个片段添加超链接URL
                        for seg in link_segments:
                            seg.hyperlink_url = link_url
                        segments.extend(link_segments)
                        i = close_paren + 1
                        continue
            
            # 普通字符
            current_text.append(text[i])
            i += 1
        
        flush_current()
        return segments
