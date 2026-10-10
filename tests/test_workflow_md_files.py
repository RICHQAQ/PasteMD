"""工作流级回归测试：多 MD 文件拆分/合并、命名与落地位置。

打桩剪贴板与 Pandoc，不触碰真实剪贴板、不调用真实转换。
覆盖两个曾经全绿放行严重回归的路径：FileWorkflow 的 markdown 主路径、以及
"检测到多个 md 但其中一个读失败" 时 separate 是否仍然生效。
"""

import os
import types

import pytest

from pastemd.config.defaults import DEFAULT_CONFIG
from pastemd.core.state import app_state
from pastemd.core.types import PlacementResult


class _Notifier:
    def __init__(self):
        self.messages = []

    def notify(self, title, message, ok=True, **kwargs):
        self.messages.append((ok, message))

    @property
    def texts(self):
        return [message for _, message in self.messages]


class _DocGenerator:
    def __init__(self):
        self.markdown_calls = []
        self.html_calls = []

    def convert_markdown_to_docx_bytes(self, md_text, config):
        self.markdown_calls.append(md_text)
        return b"DOCX"

    def convert_html_to_docx_bytes(self, html_text, config):
        self.html_calls.append(html_text)
        return b"DOCX"


class _Placer:
    def __init__(self):
        self.calls = []

    def place(self, content, config, file_paths=None, **kwargs):
        self.calls.append(list(file_paths or []))
        return PlacementResult(success=True, method="clipboard_file")


def _config(**overrides):
    config = dict(DEFAULT_CONFIG)
    config.update(overrides)
    return config


def _stub_clipboard(module, monkeypatch, files_data, errors, text, html):
    files_data = list(files_data)
    errors = list(errors)
    monkeypatch.setattr(
        module,
        "read_markdown_files_from_clipboard",
        lambda: (bool(files_data), list(files_data), list(errors)),
    )
    monkeypatch.setattr(module, "get_clipboard_text", lambda: text)
    monkeypatch.setattr(module, "get_clipboard_html", lambda cfg: html)
    monkeypatch.setattr(
        module, "is_clipboard_empty", lambda: not files_data and not text and not html
    )
    # 空 html 视为纯文本片段，非空 html 视为富文本来源
    monkeypatch.setattr(module, "is_plain_html_fragment", lambda value: not value)


@pytest.fixture()
def file_workflow(monkeypatch):
    import pastemd.app.workflows.extensible.file_workflow as fw_mod

    def build(config, files_data=(), errors=(), text="", html=""):
        _stub_clipboard(fw_mod, monkeypatch, files_data, errors, text, html)
        monkeypatch.setattr(app_state, "config", config)
        workflow = fw_mod.FileWorkflow()
        notifier = _Notifier()
        doc_gen = _DocGenerator()
        placer = _Placer()
        workflow.notification_manager = notifier
        workflow._doc_generator = doc_gen
        workflow.placer = placer
        return workflow, notifier, doc_gen, placer

    return build


@pytest.fixture()
def fallback_workflow(monkeypatch):
    import pastemd.app.workflows.fallback.fallback_workflow as fb_mod

    def build(config, files_data=(), errors=(), text="", html=""):
        _stub_clipboard(fb_mod, monkeypatch, files_data, errors, text, html)
        monkeypatch.setattr(app_state, "config", config)
        workflow = fb_mod.FallbackWorkflow()
        notifier = _Notifier()
        doc_gen = _DocGenerator()
        workflow.notification_manager = notifier
        workflow.output_executor = fb_mod.OutputExecutor(notifier)
        workflow._doc_generator = doc_gen
        return workflow, notifier, doc_gen

    return build


# ---------------------------------------------------------------------------
# FileWorkflow：默认 merge 主路径（曾因 docx_bytes 未赋值而 100% 崩溃）
# ---------------------------------------------------------------------------

def test_file_workflow_plain_markdown_text_merges(file_workflow, tmp_path):
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(save_dir=str(tmp_path), keep_file=True),
        text="# Hello\nworld",
    )

    workflow.execute()

    assert len(doc_gen.markdown_calls) == 1
    assert "Hello" in doc_gen.markdown_calls[0]
    assert len(placer.calls) == 1
    assert len(placer.calls[0]) == 1
    assert os.path.exists(placer.calls[0][0])
    assert notifier.messages[-1][0] is True


def test_file_workflow_single_md_file_merges(file_workflow, tmp_path):
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(save_dir=str(tmp_path), keep_file=True, md_file_output_name_mode="original"),
        files_data=[("a.md", "# a")],
    )

    workflow.execute()

    assert len(doc_gen.markdown_calls) == 1
    assert [os.path.basename(p) for p in placer.calls[0]] == ["a.docx"]


def test_file_workflow_multiple_md_files_merge_into_one(file_workflow, tmp_path):
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(save_dir=str(tmp_path), keep_file=True, md_file_output_name_mode="original"),
        files_data=[("a.md", "# a"), ("b.md", "# b")],
    )

    workflow.execute()

    assert len(doc_gen.markdown_calls) == 1
    assert [os.path.basename(p) for p in placer.calls[0]] == ["a_b.docx"]


def test_file_workflow_html_source_still_converts(file_workflow, tmp_path):
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(save_dir=str(tmp_path), keep_file=True),
        html="<h1>标题</h1>",
    )

    workflow.execute()

    assert len(doc_gen.html_calls) == 1
    assert len(placer.calls) == 1
    assert notifier.messages[-1][0] is True


# ---------------------------------------------------------------------------
# FileWorkflow：separate 模式
# ---------------------------------------------------------------------------

def test_file_workflow_multiple_md_files_separate(file_workflow, tmp_path):
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a"), ("b.md", "# b")],
    )

    workflow.execute()

    assert len(doc_gen.markdown_calls) == 2
    assert [os.path.basename(p) for p in placer.calls[0]] == ["a.docx", "b.docx"]


def test_file_workflow_separate_triggers_when_one_file_unreadable(file_workflow, tmp_path):
    """检测到 2 个 md、其中 1 个读失败时，仍按 separate 输出并提示失败文件。"""
    workflow, notifier, doc_gen, placer = file_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a")],
        errors=[("b.md", "boom")],
    )

    workflow.execute()

    assert len(doc_gen.markdown_calls) == 1
    assert [os.path.basename(p) for p in placer.calls[0]] == ["a.docx"]
    assert any("b.md" in text for text in notifier.texts)


# ---------------------------------------------------------------------------
# FallbackWorkflow：合并 / 拆分 / 落地位置 / 打开阈值
# ---------------------------------------------------------------------------

def test_fallback_merge_default_produces_single_file(fallback_workflow, tmp_path):
    workflow, notifier, doc_gen = fallback_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            no_app_action="save",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a"), ("b.md", "# b")],
    )

    workflow.execute()

    assert sorted(os.listdir(tmp_path)) == ["a_b.docx"]


def test_fallback_separate_keep_file_on_writes_save_dir(fallback_workflow, tmp_path):
    workflow, notifier, doc_gen = fallback_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            no_app_action="save",
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a"), ("b.md", "# b")],
    )

    workflow.execute()

    assert sorted(os.listdir(tmp_path)) == ["a.docx", "b.docx"]


def test_fallback_separate_keep_file_off_writes_temp_dir(
    fallback_workflow, tmp_path, monkeypatch
):
    import pastemd.utils.fs as fs_mod

    temp_dir = tmp_path / "temp"
    temp_dir.mkdir()
    save_dir = tmp_path / "save"
    monkeypatch.setattr(
        fs_mod, "tempfile", types.SimpleNamespace(gettempdir=lambda: str(temp_dir))
    )

    workflow, notifier, doc_gen = fallback_workflow(
        _config(
            save_dir=str(save_dir),
            keep_file=False,
            no_app_action="save",
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a"), ("b.md", "# b")],
    )

    workflow.execute()

    assert sorted(os.listdir(temp_dir)) == ["a.docx", "b.docx"]
    assert not save_dir.exists() or os.listdir(save_dir) == []


def test_fallback_separate_triggers_when_one_file_unreadable(fallback_workflow, tmp_path):
    """检测到 2 个 md、其中 1 个读失败时，仍按 separate 输出并汇总失败。"""
    workflow, notifier, doc_gen = fallback_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            no_app_action="save",
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[("a.md", "# a")],
        errors=[("b.md", "boom")],
    )

    workflow.execute()

    assert sorted(os.listdir(tmp_path)) == ["a.docx"]
    assert any("b.md" in text for text in notifier.texts)


def test_fallback_separate_open_skips_when_over_limit(
    fallback_workflow, tmp_path, monkeypatch
):
    import pastemd.app.workflows.fallback.output_executor as oe_mod

    opened = []
    monkeypatch.setattr(
        oe_mod.AppLauncher,
        "awaken_and_open_document",
        staticmethod(lambda path: opened.append(path) or True),
    )

    workflow, notifier, doc_gen = fallback_workflow(
        _config(
            save_dir=str(tmp_path),
            keep_file=True,
            no_app_action="open",
            md_multi_file_mode="separate",
            md_file_output_name_mode="original",
        ),
        files_data=[(f"f{i}.md", f"# f{i}") for i in range(6)],
    )

    workflow.execute()

    assert opened == []
    assert len(os.listdir(tmp_path)) == 6
    assert any("6" in text for text in notifier.texts)
