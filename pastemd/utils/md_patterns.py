"""Markdown 处理共用的正则模式常量。

集中存放 Markdown 变换模块（围栏识别、行内代码定位、图片路径锚定等）
引用的编译正则，业务模块内只保留函数逻辑。
"""

import re

# 围栏标记：3 个及以上连续反引号或波浪线
FENCE_MARKER_RE = re.compile(r'(`{3,}|~{3,})')

# 有序列表 / 无序列表标记（后随空白），如 "- "、"1. "
LIST_MARKER_RE = re.compile(r'(?:[-*+]|\d{1,9}[.)])\s')

# 行内代码 span：`...` / ``...``（定界符等长配对，覆盖多反引号场景）
INLINE_CODE_SPAN_RE = re.compile(r'(`+)[^`]*\1')

# 图片引用：![alt](<路径> "标题")，title 支持双引号或单引号
IMAGE_REF_RE = re.compile(
    r'((?<!\\)!\[[^\]]*\]\()(<?)([^<>()\n]+?)(>?)(\s+(?:"[^"]*"|\'[^\']*\'))?\)'
)

# 非相对路径：URL（含协议）、盘符/UNC 绝对路径、站点根相对路径
NOT_RELATIVE_PATH_RE = re.compile(r'^(?:[a-zA-Z][a-zA-Z0-9+.\-]*:|[a-zA-Z]:[\\/]|[\\/])')
