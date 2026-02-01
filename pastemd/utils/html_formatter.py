"""Utilities for cleaning and formatting HTML fragments before conversion."""

from __future__ import annotations

import re
from typing import Dict, Optional

from bs4 import BeautifulSoup, NavigableString, Tag

_CSS_CLASS_RE = re.compile(r"\.(?P<class>[A-Za-z0-9_-]+)\s*\{(?P<body>[^}]*)\}", re.DOTALL)


def extract_html_body(html: str) -> str:
    """Extract body content from a standalone HTML document."""
    body_match = re.search(r"<body[^>]*>(.*?)</body>", html, re.DOTALL | re.IGNORECASE)
    if body_match:
        return body_match.group(1).strip()

    html = re.sub(r"<!DOCTYPE[^>]*>", "", html, flags=re.IGNORECASE)
    html = re.sub(r"<html[^>]*>|</html>", "", html, flags=re.IGNORECASE)
    html = re.sub(r"<head[^>]*>.*?</head>", "", html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r"<body[^>]*>|</body>", "", html, flags=re.IGNORECASE)
    return html.strip()


def clean_html_content(soup: BeautifulSoup, options: Optional[Dict[str, object]] = None) -> None:
    """
    清理 HTML 内容，移除不可用元素，并按配置应用格式化规则。

    Args:
        html: 原始 HTML 内容。
        options: 可选格式化配置，如 ``{"strikethrough_to_del": True}``。

    Returns:
        清理后的 HTML 内容。
    """

    options = options or {}

    # 删除所有 <svg> 标签
    for svg in soup.find_all("svg"):
        svg.decompose()

    # 删除 src 指向 .svg 的 <img> 标签
    for img in soup.find_all("img", src=True):
        if img["src"].lower().endswith(".svg"):
            img.decompose()
    
    # 清理 LaTeX 公式块中的 <br> 标签
    _clean_latex_br_tags(soup)


def convert_css_font_to_semantic(soup: BeautifulSoup) -> None:
    """
    将 CSS 中的粗体/斜体类映射为 <strong>/<em>，以便 Pandoc 保留样式。

    主要用于 Excel/WPS 复制的 HTML：样式往往只写在 <style> 的 class 中，
    直接转 Markdown 会丢失加粗/斜体信息。
    """
    css_text_parts = []
    for style in soup.find_all("style"):
        css_text_parts.append(style.get_text() or "")
    css_text = "\n".join(css_text_parts)
    if not css_text.strip():
        return

    class_styles: dict[str, tuple[bool, bool]] = {}
    for match in _CSS_CLASS_RE.finditer(css_text):
        class_name = match.group("class")
        body = match.group("body").lower()

        bold = False
        italic = False
        weight_match = re.search(r"font-weight\s*:\s*([^;]+)", body)
        if weight_match:
            value = weight_match.group(1).strip()
            if value in ("bold", "bolder"):
                bold = True
            elif value.isdigit() and int(value) >= 600:
                bold = True

        style_match = re.search(r"font-style\s*:\s*([^;]+)", body)
        if style_match:
            value = style_match.group(1).strip()
            if "italic" in value or "oblique" in value:
                italic = True

        if bold or italic:
            class_styles[class_name] = (bold, italic)

    if not class_styles:
        return

    def _build_wrapper(current_bold: bool, current_italic: bool) -> tuple[Tag, Tag]:
        if current_bold and current_italic:
            strong = soup.new_tag("strong")
            em = soup.new_tag("em")
            strong.append(em)
            return strong, em
        if current_bold:
            strong = soup.new_tag("strong")
            return strong, strong
        em = soup.new_tag("em")
        return em, em

    for tag in soup.find_all(class_=True):
        classes = tag.get("class") or []
        bold = False
        italic = False
        for class_name in classes:
            if class_name in class_styles:
                class_bold, class_italic = class_styles[class_name]
                bold = bold or class_bold
                italic = italic or class_italic

        if not (bold or italic):
            continue

        if tag.name in ("table", "tbody", "thead", "tfoot", "tr"):
            continue

        if tag.name in ("td", "th"):
            if not tag.contents:
                continue
            wrapper, inner = _build_wrapper(bold, italic)
            for child in list(tag.contents):
                inner.append(child.extract())
            tag.append(wrapper)
            continue

        if tag.name in ("strong", "em"):
            if tag.name == "strong" and bold and not italic:
                continue
            if tag.name == "em" and italic and not bold:
                continue
            # 需要补充另一种样式，直接包裹内容
            if tag.name == "strong" and italic:
                wrapper = soup.new_tag("em")
                for child in list(tag.contents):
                    wrapper.append(child.extract())
                tag.append(wrapper)
            elif tag.name == "em" and bold:
                wrapper = soup.new_tag("strong")
                for child in list(tag.contents):
                    wrapper.append(child.extract())
                tag.append(wrapper)
            continue

        wrapper, inner = _build_wrapper(bold, italic)
        for child in list(tag.contents):
            inner.append(child.extract())
        tag.replace_with(wrapper)


def promote_bold_first_row_to_header(soup: BeautifulSoup) -> None:
    """
    将表格首行的粗体单元格提升为表头 (<th>)。

    主要用于 Excel/WPS 复制的 HTML：表头通常是加粗文本但仍是 <td>，
    Pandoc 无法识别为表头，导致 Markdown 不生成表头分隔线。
    """

    def _meaningful_children(tag: Tag) -> list:
        return [
            child
            for child in tag.contents
            if not (isinstance(child, NavigableString) and not str(child).strip())
        ]

    def _cell_is_bold(cell: Tag) -> bool:
        children = _meaningful_children(cell)
        if len(children) != 1:
            return False
        child = children[0]
        return isinstance(child, Tag) and child.name in ("strong", "b")

    for table in soup.find_all("table"):
        if table.find("th"):
            continue

        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        header_row = rows[0]
        header_cells = header_row.find_all(["td", "th"], recursive=False)
        if not header_cells:
            continue

        if not all(_cell_is_bold(cell) for cell in header_cells):
            continue

        has_non_bold_cell = False
        for row in rows[1:]:
            for cell in row.find_all(["td", "th"], recursive=False):
                if not _cell_is_bold(cell):
                    has_non_bold_cell = True
                    break
            if has_non_bold_cell:
                break

        if not has_non_bold_cell:
            continue

        for cell in header_cells:
            cell.name = "th"


def convert_strikethrough_to_del(soup) -> None:
    """
    在 BeautifulSoup 解析树中查找文本节点，将 ``~~text~~`` 替换为 ``<del>text</del>``。

    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 递归处理所有文本节点
    for element in soup.find_all(text=True):
        if isinstance(element, NavigableString):
            if "~~" not in element:
                continue
            pattern = r'~~([^~]+?)~~'
            if not re.search(pattern, element):
                continue

            new_content = []
            last_end = 0
            for match in re.finditer(pattern, element):
                if match.start() > last_end:
                    new_content.append(element[last_end:match.start()])

                del_tag = soup.new_tag("del")
                del_tag.string = match.group(1)
                new_content.append(del_tag)
                last_end = match.end()

            if last_end < len(element):
                new_content.append(element[last_end:])

            parent = element.parent
            if not parent:
                continue
            index = parent.contents.index(element)
            element.extract()
            for i, item in enumerate(new_content):
                if isinstance(item, str):
                    parent.insert(index + i, NavigableString(item))
                else:
                    parent.insert(index + i, item)


def _clean_latex_br_tags(soup) -> None:
    """
    清理 HTML 中 LaTeX 公式块内的 <br> 标签。
    
    LaTeX 公式块通常包裹在 class="katex" 或 class="katex-display" 的元素中，
    公式内容的 <br> 标签会破坏 LaTeX 语法，需要移除或替换为换行符。
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """  
    # 查找所有包含 katex 的元素（行内公式和块级公式）
    katex_elements = soup.find_all(class_=re.compile(r'katex'))
    
    for katex_elem in katex_elements:
        # 在 katex 元素内查找所有 <br> 标签
        br_tags = katex_elem.find_all('br')
        
        for br in br_tags:
            # 删除 <br> 标签
            br.replace_with('')

    # 处理 $$ ... $$ 包裹的内容
    # 遍历可能的容器元素
    for tag in soup.find_all(['p', 'div', 'span', 'li', 'td', 'th']):
        # 快速检查：容器内必须有 br 且文本包含 $$
        if not tag.find('br', recursive=False) or '$$' not in tag.get_text():
            continue

        in_latex = False
        # 使用 list 复制 children，因为我们会修改 DOM (删除 br)
        for child in list(tag.children):
            if isinstance(child, NavigableString):
                # 统计文本中 $$ 的数量，如果是奇数个，说明状态切换
                if str(child).count('$$') % 2 == 1:
                    in_latex = not in_latex
            elif child.name == 'br':
                # 如果在公式内，移除 br
                if in_latex:
                    child.decompose()
            elif hasattr(child, 'get_text'):
                # 对于其他标签，检查其文本内容是否包含 $$ 导致状态切换
                # 假设 $$ 成对出现，若子元素包含奇数个 $$，则改变当前上下文状态
                if child.get_text().count('$$') % 2 == 1:
                    in_latex = not in_latex


def unwrap_all_p_div_inside_li(soup, unwrap_tags=("p", "div")) -> None:
    """
    清理所有 li 内部(任意深度)的 p/div：只要在 li 子树里就 unwrap。
    - 包括 ul 下的 li、再嵌套的 ul/li ... 全部处理
    - 采用“从深到浅”顺序 unwrap，避免结构变化导致漏处理
    """
    # 选中所有在 li 里面的 p/div（包括嵌套 li 内的）
    wrappers = soup.select(",".join(f"li {t}" for t in unwrap_tags))

    # 从深到浅排序：父链越长越深，先 unwrap 深层更安全
    wrappers.sort(key=lambda node: len(list(node.parents)), reverse=True)

    for node in wrappers:
        # node 可能已被前面的 unwrap 影响而脱离树，做个保护
        if isinstance(node, Tag) and node.parent is not None:
            node.unwrap()

    # 可选：把每个 li 头尾多余空白文本清一下
    for li in soup.find_all("li"):
        _trim_whitespace_text_nodes(li)


def remove_empty_paragraphs(soup) -> None:
    """
    删除空 <p>：
      - 只有空白
      - 或只有 &nbsp; / \u00a0
      - 或只包含空白的 span/br 等（尽量温和：只在“可判定为空”时删除）
    """
    for p in soup.find_all("p"):
        text = p.get_text(strip=True).replace("\u00a0", "").strip()
        # 如果完全没内容，并且没有 img/iframe 等“非文本但有意义”的元素
        has_meaningful_media = bool(p.find(["img", "iframe", "video", "audio", "svg"]))
        if (not text) and (not has_meaningful_media):
            p.decompose()


def _trim_whitespace_text_nodes(tag) -> None:
    """
    去掉某个 tag 开头/结尾的纯空白 NavigableString，避免 unwrap 后出现奇怪空白。
    """
    # 头部
    while tag.contents and isinstance(tag.contents[0], NavigableString) and not str(tag.contents[0]).strip():
        tag.contents[0].extract()
    # 尾部
    while tag.contents and isinstance(tag.contents[-1], NavigableString) and not str(tag.contents[-1]).strip():
        tag.contents[-1].extract()


def postprocess_pandoc_html_macwps(html: str) -> str:
    """
    后处理 Pandoc 输出的 HTML，修复格式问题。
    
    处理内容：
    1. 修复代码块格式（移除属性标记，恢复换行）
    2. 清理所有样式和多余属性（生成纯净 HTML）
    3. 修复粗体斜体嵌套
    4. 清理 Pandoc 扩展语法残留
    
    Args:
        html: Pandoc 输出的 HTML 文本
        
    Returns:
        后处理后的 HTML 文本
    """
    soup = BeautifulSoup(html, "html.parser")
    
    # 清理列表中的 p/div 包装
    unwrap_all_p_div_inside_li(soup)

    # 替换 del 为 s 标签（WPS 兼容）
    _replace_del_with_s(soup)

    # 修复粗体加斜体的嵌套标签（WPS 兼容性）
    _fix_bold_italic_nesting(soup)
    
    # 修复代码块格式
    _fix_pandoc_code_blocks(soup)
    
    # 清理 Pandoc 扩展语法残留（如 ::: 语法块）
    # _clean_pandoc_fenced_divs(soup)
    
    # 清理多余的属性（style, class, data-* 等）
    # _clean_pandoc_attributes(soup)

    _fix_task_list_math_issue(soup)
    
    return str(soup)


def _fix_bold_italic_nesting(soup) -> None:
    """
    修复粗体加斜体的嵌套标签，以兼容 WPS。
    
    将 <strong><em>text</em></strong> 或 <em><strong>text</strong></em>
    转换为 <span style="font-weight: bold; font-style: italic;">text</span>
    
    WPS 对嵌套的 <strong><em> 标签支持不好，只会显示斜体效果。
    使用 inline style 可以确保粗体和斜体效果同时生效。
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 处理 <strong><em>text</em></strong> 模式
    for strong in soup.find_all('strong'):
        # 检查 strong 标签是否只包含一个 em 标签
        children = [c for c in strong.children if c.name or (isinstance(c, NavigableString) and c.strip())]
        if len(children) == 1 and children[0].name == 'em':
            em = children[0]
            text = em.get_text()
            
            # 创建新的 span 标签，使用 inline style
            span = soup.new_tag('span', style='font-weight: bold; font-style: italic;')
            span.string = text
            
            # 替换原来的 strong 标签
            strong.replace_with(span)
    
    # 处理 <em><strong>text</strong></em> 模式
    for em in soup.find_all('em'):
        # 检查 em 标签是否只包含一个 strong 标签
        children = [c for c in em.children if c.name or (isinstance(c, NavigableString) and c.strip())]
        if len(children) == 1 and children[0].name == 'strong':
            strong = children[0]
            text = strong.get_text()
            
            # 创建新的 span 标签，使用 inline style
            span = soup.new_tag('span', style='font-weight: bold; font-style: italic;')
            span.string = text
            
            # 替换原来的 em 标签
            em.replace_with(span)


def _fix_pandoc_code_blocks(soup) -> None:
    """
    修复 Pandoc 输出的代码块格式问题。
    
    处理两种情况：
    1. 属性标记格式：<p><code>{.class! attr="value"} actual code here</code></p>
    2. 复杂结构格式：<div class="sourceCode"><pre><code><span>...</span></code></pre></div>
    
    统一转换为简单的：
    <pre style="white-space: pre-wrap;"><code>actual code here</code></pre>
    
    使用 white-space: pre-wrap 确保代码块在 WPS 中可以自动换行。
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 处理 Pandoc 生成的 div.sourceCode 复杂结构
    for div in soup.find_all('div', class_='sourceCode'):
        # 查找内部的 pre > code 结构
        pre = div.find('pre')
        if pre:
            code = pre.find('code')
            if code:
                # 提取所有文本内容（自动合并所有 span 标签中的文本）
                code_text = code.get_text()
                
                # 创建新的简化 pre > code 结构
                new_pre = soup.new_tag('pre', style='white-space: pre-wrap;')
                new_code = soup.new_tag('code')
                new_code.string = code_text
                new_pre.append(new_code)
                
                # 替换整个 div.sourceCode
                div.replace_with(new_pre)
    
    # 处理 <p> 标签中包含 <code> 的情况（属性标记格式）
    for p in soup.find_all('p'):
        # 获取 p 标签的所有子节点（排除纯空白文本节点）
        meaningful_contents = [
            c for c in p.contents 
            if c.name or (isinstance(c, NavigableString) and c.strip())
        ]
        
        # 检查 p 是否只包含一个 code 标签
        code_tags = p.find_all('code', recursive=False)
        if len(code_tags) == 1 and len(meaningful_contents) == 1:
            code = code_tags[0]
            code_text = code.get_text()
            
            # 检查是否包含 Pandoc 属性标记（以 { 开头）
            if code_text.strip().startswith('{'):
                # 尝试提取属性和实际代码
                # 格式：{.class! attr="value"} actual code here
                match = re.match(r'^\{[^}]+\}\s*(.+)$', code_text, re.DOTALL)
                if match:
                    actual_code = match.group(1)
                    
                    # 恢复代码中的换行
                    # Pandoc 将多行代码压缩成单行，用多个空格代替换行
                    # 检测连续的多个空格（通常是 4+ 空格），替换为换行+缩进
                    actual_code = re.sub(r'    +', '\n    ', actual_code)
                    
                    # 创建新的 pre > code 结构，添加 white-space: pre-wrap
                    pre = soup.new_tag('pre', style='white-space: pre-wrap;')
                    new_code = soup.new_tag('code')
                    new_code.string = actual_code
                    pre.append(new_code)
                    
                    # 替换原来的 p 标签
                    p.replace_with(pre)


def _clean_pandoc_attributes(soup) -> None:
    """
    清理 Pandoc 输出的 HTML 中的额外属性，生成纯净的 HTML。
    
    移除：
    - style 属性（所有内联样式）
    - class 属性（CSS 类名）
    - data-* 属性（自定义数据属性）
    - 其他非标准属性
    
    保留：
    - id（锚点）
    - href（链接）
    - src, alt（图片）
    - type（列表类型）
    - colspan, rowspan（表格）
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 定义需要保留的属性白名单
    allowed_attrs = {
        'a': ['href', 'id'],
        'img': ['src', 'alt'],
        'td': ['colspan', 'rowspan'],
        'th': ['colspan', 'rowspan'],
        'ol': ['type', 'start'],
        'ul': ['type'],
        # 对于标题和其他标签，只保留 id
        'h1': ['id'], 'h2': ['id'], 'h3': ['id'], 
        'h4': ['id'], 'h5': ['id'], 'h6': ['id'],
    }
    
    for tag in soup.find_all(True):
        # 获取该标签类型允许的属性
        allowed = allowed_attrs.get(tag.name, ['id'])
        
        # 找出需要删除的属性
        attrs_to_del = [attr for attr in list(tag.attrs.keys()) if attr not in allowed]
        
        # 删除不允许的属性
        for attr in attrs_to_del:
            del tag.attrs[attr]


def _replace_del_with_s(soup) -> None:
    """
    将 <del> 标签替换为 <s> 标签（WPS 更兼容删除线）。
    保留内容与属性（后续 clean 会清掉 data-*）。
    """
    for tag in soup.find_all("del"):
        tag.name = "s"


def _clean_pandoc_fenced_divs(soup) -> None:
    """
    清理 Pandoc fenced divs 扩展语法产生的残留文本。
    
    Pandoc 在某些情况下会将 ::: 语法块转换为文本，需要清理。
    例如：:::::::: {.class} -> 应该被移除
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 查找包含 Pandoc 扩展语法的文本节点
    for text_node in soup.find_all(text=True):
        if isinstance(text_node, NavigableString):
            text = str(text_node)
            # 匹配以 : 开头的 Pandoc 扩展语法
            if re.match(r'^:+\s*\{[^}]*\}', text.strip()):
                # 完全移除这类文本
                text_node.extract()
            elif text.strip().startswith(':::'):
                # 移除包含 ::: 的行
                text_node.extract()


def clean_html_for_wps(html: str) -> str:
    """
    弃用
    专门为 WPS 工作流清理输入的 HTML，去除所有样式和扩展属性。
    
    这个函数在 Pandoc 转换之前调用，清理输入 HTML 中的：
    - style 属性（内联样式）
    - class 属性（CSS 类名）
    - 所有 data-* 自定义属性
    - 其他非标准属性（如 path-to-node, _ngcontent-*, hveid, ved 等）
    - 保护任务列表标记，避免被 Pandoc 误识别为数学公式
    
    只保留必要的语义属性（id, href, src, alt 等）。
    
    Args:
        html: 原始 HTML 字符串
        
    Returns:
        清理后的 HTML 字符串
    """
    soup = BeautifulSoup(html, "html.parser")
    
    # 先保护任务列表标记，避免 Pandoc 将 [x] 转义为 \[x\]，导致被识别为数学公式
    _protect_task_list_brackets(soup)
    
    # 定义需要保留的属性白名单
    allowed_attrs = {
        'a': ['href', 'id', 'title'],
        'img': ['src', 'alt', 'title'],
        'td': ['colspan', 'rowspan'],
        'th': ['colspan', 'rowspan'],
        'ol': ['type', 'start'],
        'ul': ['type'],
        # 对于其他标签，只保留 id 和 title
    }
    
    for tag in soup.find_all(True):
        # 获取该标签类型允许的属性
        allowed = allowed_attrs.get(tag.name, ['id', 'title'])
        
        # 找出需要删除的属性
        attrs_to_del = [attr for attr in list(tag.attrs.keys()) if attr not in allowed]
        
        # 删除不允许的属性
        for attr in attrs_to_del:
            del tag.attrs[attr]
    
    return str(soup)


def _remove_col_tags(soup) -> None:
    """
    移除 HTML 中的 <col> 标签。
    Pandoc 处理带有 span 属性的 <col> 标签时可能会导致表格转换错误。
    """
    for col in soup.find_all("col"):
        col.decompose()


def protect_brackets(html: str) -> str:
    """
    保护 HTML 转 md中的任务列表，$$等数学公式，避免被 Pandoc 转义和误识别。
    同时移除 <col> 标签以修复 Excel 表格转换问题。
    
    将 [x] 和 [ ] 替换为特殊标记：
    - [x] -> {{TASK_CHECKED}}
    - [ ] -> {{TASK_UNCHECKED}}
    
    这些特殊标记不会被 Pandoc 识别为 Markdown 语法或数学公式。
    
    Args:
        html: 原始 HTML 字符串
        
    Returns:
        处理后的 HTML 字符串
    """
    soup = BeautifulSoup(html, "html.parser")
    _remove_col_tags(soup)
    _protect_task_list_brackets(soup)
    return str(soup)


def _protect_task_list_brackets(soup) -> None:
    """
    保护 HTML 中的任务列表标记，避免被 Pandoc 转义和误识别。
    
    将 [x] 和 [ ] 替换为特殊标记：
    - [x] -> {{TASK_CHECKED}}
    - [ ] -> {{TASK_UNCHECKED}}
    
    这些特殊标记不会被 Pandoc 识别为 Markdown 语法或数学公式。
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 遍历所有文本节点
    for text_node in soup.find_all(text=True):
        if isinstance(text_node, NavigableString):
            text = str(text_node)
            # 只处理包含任务列表标记的文本
            if '[x]' in text or '[ ]' in text or '[X]' in text:
                # 替换为特殊标记
                text = text.replace('[x]', '{{TASK_CHECKED}}')
                text = text.replace('[ ]', '{{TASK_UNCHECKED}}')
                text_node.replace_with(text)


def _restore_task_list_brackets(soup) -> None:
    """
    将 HTML 中的 input checkbox 标签直接替换为 [x] 或 [ ] 文本。
    """
    # 1. 寻找所有的 input 标签
    for checkbox in soup.find_all('input'):
        # 更加鲁棒的判断：如果是 checkbox 或者它带有 checked 属性
        is_checkbox = checkbox.get('type') == 'checkbox'
        if is_checkbox:
            # 判断是否选中
            is_checked = checkbox.has_attr('checked')
            replacement_text = "[x] " if is_checked else "[ ] "
            
            # 核心修复：使用 NavigableString 确保替换为纯文本
            checkbox.replace_with(NavigableString(replacement_text))


def _fix_task_list_math_issue(soup) -> None:
    """
    恢复被保护的任务列表标记。
    
    在 clean_html_for_wps 中，任务列表标记被替换为特殊标记以避免被 Pandoc 误识别。
    这个函数将特殊标记恢复为原始的方括号格式。
    
    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """
    # 恢复任务列表标记
    _restore_task_list_brackets(soup)
