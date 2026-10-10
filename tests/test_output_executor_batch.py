"""多文件分开输出（批量落地）的单元测试。"""

import os

from pastemd.app.workflows.fallback.output_executor import OutputExecutor
from pastemd.i18n import t
from pastemd.utils.fs import generate_output_path


class _FakeNotifier:
    def __init__(self):
        self.messages = []

    def notify(self, title, message, ok=True, **kwargs):
        self.messages.append((message, ok))


def _make_items(save_dir, names):
    """按原文件名模式预生成 items（此时路径可能重复，靠批量去重错开）。"""
    items = []
    for name in names:
        path = generate_output_path(
            keep_file=True,
            save_dir=str(save_dir),
            md_text=f"# {name}\n正文",
            source_filenames=[f"{name}.md"],
            md_name_mode="original",
        )
        items.append((b"docx-bytes", path, f"{name}.md"))
    return items


def test_batch_writes_separate_files_and_dedupes_same_name(tmp_path):
    notifier = _FakeNotifier()
    executor = OutputExecutor(notifier)
    items = _make_items(tmp_path, ["笔记", "笔记"])

    result = executor.execute_docx_batch(action="save", items=items)

    assert result["failures"] == []
    assert len(result["success_paths"]) == 2
    assert len(set(result["success_paths"])) == 2
    assert all(os.path.exists(path) for path in result["success_paths"])
    assert len(os.listdir(tmp_path)) == 2


def test_batch_open_over_limit_saves_without_opening(tmp_path, monkeypatch):
    import pastemd.app.workflows.fallback.output_executor as executor_module

    opened = []
    monkeypatch.setattr(
        executor_module.AppLauncher,
        "awaken_and_open_document",
        lambda path: opened.append(path) or True,
    )
    notifier = _FakeNotifier()
    executor = OutputExecutor(notifier)
    items = _make_items(tmp_path, [f"文件{i}" for i in range(6)])

    result = executor.execute_docx_batch(action="open", items=items)

    assert opened == []
    assert len(result["success_paths"]) == 6
    assert notifier.messages[-1][0] == t(
        "workflow.md_file.batch_generated_only", count=6
    )


def test_batch_open_at_limit_opens_every_file(tmp_path, monkeypatch):
    import pastemd.app.workflows.fallback.output_executor as executor_module

    opened = []
    monkeypatch.setattr(
        executor_module.AppLauncher,
        "awaken_and_open_document",
        lambda path: opened.append(path) or True,
    )
    notifier = _FakeNotifier()
    executor = OutputExecutor(notifier)
    items = _make_items(tmp_path, [f"文件{i}" for i in range(5)])

    result = executor.execute_docx_batch(action="open", items=items)

    assert len(opened) == 5
    assert len(result["success_paths"]) == 5
    assert len(notifier.messages) == 1


def test_batch_reports_failures_in_single_notification(tmp_path):
    notifier = _FakeNotifier()
    executor = OutputExecutor(notifier)
    items = _make_items(tmp_path, ["甲", "乙"])

    result = executor.execute_docx_batch(
        action="save", items=items, pre_failures=[("丙.md", "read failed")]
    )

    assert len(result["failures"]) == 1
    assert len(notifier.messages) == 1
    assert "丙.md" in notifier.messages[0][0]


def test_batch_all_failed_notifies_error():
    notifier = _FakeNotifier()
    executor = OutputExecutor(notifier)

    result = executor.execute_docx_batch(
        action="save", items=[], pre_failures=[("a.md", "boom"), ("b.md", "boom")]
    )

    assert result["success_paths"] == []
    assert notifier.messages[-1][1] is False
    assert notifier.messages[-1][0] == (
        t("workflow.md_file.batch_failed_all", failed_count=2)
        + "\n"
        + t("workflow.md_file.batch_failure_line", failed_count=2, failed_files="a.md, b.md")
    )
