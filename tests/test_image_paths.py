# -*- coding: utf-8 -*-
"""resolve_markdown_image_paths：相对路径图片锚定到源文件目录"""

from pathlib import Path

from pastemd.utils.md_image_paths import resolve_markdown_image_paths


def _make_tree(tmp_path: Path) -> Path:
    (tmp_path / "images").mkdir()
    (tmp_path / "images" / "fig1.png").write_bytes(b"png")
    (tmp_path / "my images").mkdir()
    (tmp_path / "my images" / "fig 2.png").write_bytes(b"png")
    return tmp_path


def test_relative_path_resolved_to_absolute(tmp_path):
    base = _make_tree(tmp_path)
    md = "前文\n\n![图1](images/fig1.png)\n\n后文"
    out = resolve_markdown_image_paths(md, str(base))
    assert out == f"前文\n\n![图1]({(base / 'images' / 'fig1.png').as_posix()})\n\n后文"


def test_remote_url_untouched(tmp_path):
    md = "![logo](https://example.com/a.png)"
    assert resolve_markdown_image_paths(md, str(tmp_path)) == md


def test_absolute_path_untouched(tmp_path):
    md = "![x](D:/data/fig.png) ![y](/root/fig.png)"
    assert resolve_markdown_image_paths(md, str(tmp_path)) == md


def test_backslash_relative_path_resolved(tmp_path):
    base = _make_tree(tmp_path)
    md = "![图1](images\\fig1.png)"
    out = resolve_markdown_image_paths(md, str(base))
    assert (base / "images" / "fig1.png").as_posix() in out


def test_path_inside_code_fence_untouched(tmp_path):
    base = _make_tree(tmp_path)
    md = "```\n![图1](images/fig1.png)\n```"
    assert resolve_markdown_image_paths(md, str(base)) == md


def test_path_with_spaces_wrapped_in_angle_brackets(tmp_path):
    base = _make_tree(tmp_path)
    md = "![图2](my images/fig 2.png)"
    out = resolve_markdown_image_paths(md, str(base))
    assert out == f"![图2](<{(base / 'my images' / 'fig 2.png').as_posix()}>)"


def test_url_encoded_path_resolved(tmp_path):
    base = _make_tree(tmp_path)
    md = "![图2](my%20images/fig%202.png)"
    out = resolve_markdown_image_paths(md, str(base))
    assert (base / "my images" / "fig 2.png").as_posix() in out


def test_missing_file_left_as_is(tmp_path):
    md = "![缺失](images/nope.png)"
    assert resolve_markdown_image_paths(md, str(tmp_path)) == md


def test_title_is_preserved(tmp_path):
    base = _make_tree(tmp_path)
    md = '![图1](images/fig1.png "标题")'
    out = resolve_markdown_image_paths(md, str(base))
    assert out == f'![图1]({(base / "images" / "fig1.png").as_posix()} "标题")'


def test_link_not_rewritten(tmp_path):
    base = _make_tree(tmp_path)
    md = "[说明](images/fig1.png)"
    assert resolve_markdown_image_paths(md, str(base)) == md


def test_image_syntax_inside_inline_code_not_rewritten(tmp_path):
    # 讲 Markdown 语法的文档里，`![x](a.png)` 是示例文本
    base = _make_tree(tmp_path)
    md = "用 `![x](images/fig1.png)` 语法插入图片"
    assert resolve_markdown_image_paths(md, str(base)) == md


def test_image_syntax_inside_double_backtick_code_not_rewritten(tmp_path):
    base = _make_tree(tmp_path)
    md = 'a `` ![x](images/fig1.png) `` b'
    assert resolve_markdown_image_paths(md, str(base)) == md


def test_unbalanced_angle_bracket_not_rewritten(tmp_path):
    base = _make_tree(tmp_path)
    assert resolve_markdown_image_paths("![x](<images/fig1.png)", str(base)) == "![x](<images/fig1.png)"
    assert resolve_markdown_image_paths("![x](images/fig1.png>)", str(base)) == "![x](images/fig1.png>)"


def test_path_surrounding_spaces_are_stripped(tmp_path):
    # pandoc 会剥离目的地址两侧空白，应视为合法图片引用
    base = _make_tree(tmp_path)
    out = resolve_markdown_image_paths("![x]( images/fig1.png )", str(base))
    assert (base / "images" / "fig1.png").as_posix() in out


def test_single_quoted_title_is_preserved(tmp_path):
    base = _make_tree(tmp_path)
    out = resolve_markdown_image_paths("![x](images/fig1.png '标题')", str(base))
    assert (base / "images" / "fig1.png").as_posix() in out
    assert "'标题'" in out


def test_escaped_bang_link_not_rewritten(tmp_path):
    # \![a](x.png) 是字面 ! + 普通链接，不是图片
    base = _make_tree(tmp_path)
    md = "\\![a](images/fig1.png)"
    assert resolve_markdown_image_paths(md, str(base)) == md


def test_empty_inputs_untouched():
    assert resolve_markdown_image_paths("", "/tmp") == ""
    assert resolve_markdown_image_paths("![x](a.png)", "") == "![x](a.png)"
