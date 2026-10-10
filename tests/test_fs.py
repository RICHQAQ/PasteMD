import os
from datetime import datetime

from pastemd.utils.fs import generate_output_path, generate_unique_path, sanitize_filename


def test_sanitize_filename_keeps_regular_names():
    assert sanitize_filename("Chapter 1: Intro") == "Chapter 1_ Intro"


def test_sanitize_filename_avoids_windows_reserved_names():
    assert sanitize_filename("CON") == "CON_"
    assert sanitize_filename("con") == "con_"
    assert sanitize_filename("AUX.txt") == "AUX_.txt"
    assert sanitize_filename("LPT9") == "LPT9_"


def test_sanitize_filename_strips_trailing_dots_and_spaces():
    assert sanitize_filename("report. ") == "report"
    assert sanitize_filename("...") == "document"


def test_generate_output_path_defaults_to_content_title(tmp_path):
    """默认（content）不受来源文件名影响。"""
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        md_text="# 内容标题\n\n正文",
        source_filenames=["我的笔记.md"],
    )
    assert os.path.basename(path) == "内容标题.docx"


def test_generate_output_path_uses_md_file_name_in_original_mode(tmp_path):
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        md_text="# 内容标题\n\n正文",
        source_filenames=["我的笔记.md"],
        md_name_mode="original",
    )
    assert os.path.basename(path) == "我的笔记.docx"


def test_generate_output_path_original_mode_joins_multiple_sources(tmp_path):
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        md_text="# 内容标题\n\n正文",
        source_filenames=["第一篇.md", "第二章.md"],
        md_name_mode="original",
    )
    assert os.path.basename(path) == "第一篇_第二章.docx"


def test_generate_output_path_original_mode_falls_back_to_content(tmp_path):
    """原文件名清理后为空时，回退到内容命名。"""
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        md_text="# 内容标题\n\n正文",
        source_filenames=["??.md"],
        md_name_mode="original",
    )
    assert os.path.basename(path) == "内容标题.docx"


def test_generate_output_path_original_mode_ignores_text_clipboard(tmp_path):
    """剪贴板是纯文本（无来源文件）时，original 模式不改内容命名。"""
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        md_text="# 内容标题\n\n正文",
        source_filenames=[],
        md_name_mode="original",
    )
    assert os.path.basename(path) == "内容标题.docx"


def test_generate_output_path_table_ignores_name_mode(tmp_path):
    """表格输出仍按表头命名，不受命名方式影响。"""
    path = generate_output_path(
        keep_file=True,
        save_dir=str(tmp_path),
        table_data=[["姓名", "年龄"], ["张三", "20"]],
        source_filenames=["表格.md"],
        md_name_mode="original",
    )
    assert os.path.basename(path) == "姓名_年龄.xlsx"

def test_generate_unique_path_avoids_taken_timestamp_candidate(tmp_path, monkeypatch):
    """同一秒内带时间戳的候选名也被占用时，继续加序号而不是覆盖。"""
    import pastemd.utils.fs as fs_module

    class _FixedDateTime:
        @staticmethod
        def now():
            return datetime(2026, 10, 9, 23, 59, 0)

    monkeypatch.setattr(fs_module, "datetime", _FixedDateTime)
    (tmp_path / "笔记.docx").write_bytes(b"x")
    (tmp_path / "笔记_20261009_235900.docx").write_bytes(b"y")

    result = generate_unique_path(str(tmp_path / "笔记.docx"))

    assert os.path.basename(result) == "笔记_20261009_235900_1.docx"
    assert not os.path.exists(result)


