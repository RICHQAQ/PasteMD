"""懒加载标签页重构的单元测试。

覆盖三个关键路径（M1 审查建议）：
1. 标签页注册表 & index→key 反查（无 GUI 依赖）
2. 懒加载 swap：选中恢复 + 无连锁懒加载（真实 Tk widget）
3. _on_save collect 等价性 + paste_delay_s 无条件 sanitize

GUI 相关测试需要可用的 Tk 显示；无显示环境（如 CI headless）跳过。
"""

import copy
import tkinter as tk
from dataclasses import replace
from tkinter import ttk

import pytest

from pastemd.config.defaults import DEFAULT_CONFIG
from pastemd.presentation.settings.dialog import SettingsDialog


# ---------------------------------------------------------------------------
# 1. 标签页注册表（无 GUI）
# ---------------------------------------------------------------------------

def test_build_tab_specs_registry():
    """注册表返回全部 6 个 spec，顺序即展示顺序，字段完整。"""
    obj = object.__new__(SettingsDialog)
    specs = obj._build_tab_specs()
    assert list(specs.keys()) == [
        "general", "conversion", "advanced", "experimental", "extensions", "permissions",
    ]
    for key, spec in specs.items():
        assert spec.label_key
        assert callable(spec.creator)
    # general 非懒加载；conversion 懒加载；permissions 仅当前平台启用（macOS True/其他 False）
    assert specs["general"].lazy_on_windows is False
    assert specs["conversion"].lazy_on_windows is True
    from pastemd.utils.system_detect import is_macos as _is_macos
    assert specs["permissions"].enabled() is _is_macos()


def test_tab_label_resolves_via_registry():
    """_tab_label 从注册表 label_key 解析，消除 creator 硬编码。"""
    obj = object.__new__(SettingsDialog)
    obj._tab_specs = obj._build_tab_specs()
    from pastemd.i18n import t
    for key, spec in obj._tab_specs.items():
        assert obj._tab_label(key) == t(spec.label_key)
        assert obj._tab_label(key)  # 非空


# ---------------------------------------------------------------------------
# 2. 懒加载 swap（真实 Tk widget）
# ---------------------------------------------------------------------------

# 模块级单例 Tk root：避免 pytest 中多次创建/销毁 Tk 导致 TCL 资源耗尽。
# GUI 测试在无显示/无 TCL 环境自动跳过。
_root = None


def _get_root():
    global _root
    if _root is None or not _root.winfo_exists():
        _root = tk.Tk()
        _root.withdraw()
    return _root


def _build_minimal_dialog():
    """构造最小 SettingsDialog（绕过 __init__ 副作用），Windows 懒加载初始化。"""
    root = _get_root()
    obj = object.__new__(SettingsDialog)
    obj.root = root
    obj.notebook = ttk.Notebook(root)
    obj.notebook.pack()
    obj._tab_map = {}
    obj._tab_created = set()
    obj._suppress_tab_change = False
    obj.current_config = copy.deepcopy(DEFAULT_CONFIG)
    obj._tab_specs = obj._build_tab_specs()

    # 模拟 Windows _create_widgets 的创建循环
    for key, spec in obj._tab_specs.items():
        if not spec.enabled():
            continue
        if spec.lazy_on_windows:
            placeholder = ttk.Frame(obj.notebook)
            obj.notebook.add(placeholder, text=obj._tab_label(key))
            obj._tab_map[key] = placeholder
        else:
            frame = ttk.Frame(obj.notebook)
            obj.notebook.add(frame, text=obj._tab_label(key))
            obj._tab_map[key] = frame
            obj._tab_created.add(key)

    obj.notebook.bind("<<NotebookTabChanged>>", obj._on_tab_changed)
    return obj


def _tk_available():
    try:
        _get_root()
        return True
    except Exception:
        return False


_needs_tk = pytest.mark.skipif(
    not _tk_available(),
    reason="Tk not available in this environment (headless or missing TCL)",
)


@pytest.fixture()
def tk_dialog():
    dlg = _build_minimal_dialog()
    yield dlg
    # 销毁独立 widget，但保留模块级 root 供后续测试复用
    try:
        dlg.notebook.destroy()
    except Exception:
        pass


def _is_macos_import():
    from pastemd.utils.system_detect import is_macos as _m
    return _m()


# Windows 懒加载布局测试：仅非 macOS 平台运行（macOS 全量即时创建含 permissions）
_windows_only = pytest.mark.skipif(
    _is_macos_import(),
    reason="Windows 懒加载布局仅在非 macOS 平台成立",
)


@_needs_tk
@_windows_only
def test_windows_initial_state(tk_dialog):
    """Windows 初始：general 即时 + 4 懒加载占位 = 5 标签，permissions 排除。"""
    dlg = tk_dialog
    assert len(dlg.notebook.tabs()) == 5
    assert dlg._tab_created == {"general"}
    assert "permissions" not in dlg._tab_map


@_needs_tk
@_windows_only
def test_get_tab_key_by_index_reverse(tk_dialog):
    """_get_tab_key_by_index 从实际 tabs 反查，顺序无关。"""
    dlg = tk_dialog
    assert dlg._get_tab_key_by_index(0) == "general"
    assert dlg._get_tab_key_by_index(1) == "conversion"
    assert dlg._get_tab_key_by_index(4) == "extensions"
    assert dlg._get_tab_key_by_index(5) is None  # 越界


@_needs_tk
def test_lazy_swap_keeps_selection_and_no_chain(tk_dialog):
    """懒加载 swap 后：选中停在真实 frame，无连锁懒加载（回归 145be61 bug）。

    走真实触发路径：先选中 conversion 占位符（模拟用户点击），再 swap。
    这样 forget 时占位符是选中态，会触发 <<NotebookTabChanged>> 重入，
    确保"无连锁"断言能捕获回归（若只 swap 未选中的占位符，断言是空洞守卫）。
    """
    dlg = tk_dialog
    initial_count = len(dlg.notebook.tabs())

    def creator():
        real = ttk.Frame(dlg.notebook)
        dlg.notebook.add(real, text="conversion_real")
        dlg._tab_map["conversion"] = real

    # 1) 选中 conversion 占位符（模拟用户点击），建立"选中态占位符"前提
    dlg.notebook.select(dlg._tab_map["conversion"])
    # 清掉 select 占位符时可能排队的残留事件
    dlg.root.update()

    # 2) 真实 _do_lazy_swap（此时占位符是选中态，forget 会触发重入）
    dlg._do_lazy_swap("conversion", creator)
    for _ in range(8):
        dlg.root.update()

    selected = dlg.notebook.select()
    assert str(selected) == str(dlg._tab_map["conversion"])  # 选中停在真实 frame
    assert len(dlg.notebook.tabs()) == initial_count  # 占位被替换而非追加（平台无关）
    assert dlg._tab_created == {"general", "conversion"}  # 无连锁懒加载


@_needs_tk
def test_lazy_swap_creator_failure_keeps_placeholder(tk_dialog):
    """creator 抛异常时：_tab_created 不含该 key，占位符保留可重试（m3）。"""
    dlg = tk_dialog
    ph_before = dlg._tab_map["advanced"]

    def boom_creator():
        raise RuntimeError("creation failed")

    with pytest.raises(RuntimeError):
        dlg._do_lazy_swap("advanced", boom_creator)

    assert "advanced" not in dlg._tab_created
    assert dlg._tab_map["advanced"] is ph_before  # 占位符未动


@_needs_tk
def test_lazy_swap_failure_after_add_cleans_orphan(tk_dialog):
    """creator 已 add frame 但尚未更新映射就失败时，不留下孤儿 tab。

    用相对初始 tab 数做断言（delta），避免依赖平台特定数量（macOS 含 permissions
    时为 6，Windows 为 5）。
    """
    dlg = tk_dialog
    initial_count = len(dlg.notebook.tabs())
    ph_before = dlg._tab_map["advanced"]

    def partial_creator():
        real = ttk.Frame(dlg.notebook)
        dlg.notebook.add(real, text="partial")
        raise RuntimeError("failed after add")

    with pytest.raises(RuntimeError):
        dlg._do_lazy_swap("advanced", partial_creator)

    assert len(dlg.notebook.tabs()) == initial_count  # 无孤儿 tab 残留
    assert dlg._tab_map["advanced"] is ph_before
    assert "advanced" not in dlg._tab_created


@_needs_tk
def test_lazy_swap_failure_not_added_cleans_mapping(tk_dialog):
    """creator 建 frame 存入 _tab_map 但未 add 到 notebook 就失败时，清理半成品映射。

    覆盖 _do_lazy_swap 恢复逻辑的 not-added 兜底分支（frame 不在 tabs_before 差异集里，
    走 current 的 _discard_tab_widget 清理），锁定依赖 Tcl 语义的兜底路径。
    """
    dlg = tk_dialog
    initial_count = len(dlg.notebook.tabs())
    ph_before = dlg._tab_map["advanced"]

    def not_added_creator():
        real = ttk.Frame(dlg.notebook)  # 只创建 frame 并写入映射，不 add 到 notebook
        dlg._tab_map["advanced"] = real
        raise RuntimeError("failed before add")

    with pytest.raises(RuntimeError):
        dlg._do_lazy_swap("advanced", not_added_creator)

    assert len(dlg.notebook.tabs()) == initial_count  # 未 add 的 frame 不产生孤儿 tab
    assert dlg._tab_map["advanced"] is ph_before  # 映射恢复为占位符
    assert "advanced" not in dlg._tab_created


# ---------------------------------------------------------------------------
# 3. collect 等价性 & sanitize
# ---------------------------------------------------------------------------

def test_sanitize_paste_delay():
    """paste_delay_s 无条件 sanitize：合法值保留、坏值回落、负数归零、NaN/Inf 回落。"""
    assert SettingsDialog._sanitize_paste_delay(0.3) == 0.3
    assert SettingsDialog._sanitize_paste_delay("0.8") == 0.8
    assert SettingsDialog._sanitize_paste_delay(-5) == 0.0
    assert SettingsDialog._sanitize_paste_delay("abc") == 0.3
    assert SettingsDialog._sanitize_paste_delay(None) == 0.3
    assert SettingsDialog._sanitize_paste_delay("") == 0.3
    # NaN/±Inf 会绕过 float() 异常但让 time.sleep 崩溃，须回落默认值
    assert SettingsDialog._sanitize_paste_delay(float("nan")) == 0.3
    assert SettingsDialog._sanitize_paste_delay(float("inf")) == 0.3
    assert SettingsDialog._sanitize_paste_delay(float("-inf")) == 0.3
    assert SettingsDialog._sanitize_paste_delay("1e400") == 0.3


def _build_save_ready_dialog():
    """构造可调用真实 _on_save 的 dialog：general 页已创建且具备所需 Tk var。

    minimal dialog 的 general 只有占位 frame，没有真实 var；这里补齐 _collect_general
    依赖的 Tk var，使真实 _on_save 能跑通 general 收集路径。
    """
    dlg = _build_minimal_dialog()
    dlg._tab_created = {"general"}
    # _collect_general 依赖的 var
    dlg.lang_var = tk.StringVar(value="en-US")
    dlg.lang_map = {"en-US": "en-US"}
    dlg.save_dir_var = tk.StringVar(value=dlg.current_config.get("save_dir", ""))
    dlg.keep_file_var = tk.BooleanVar(value=False)
    dlg.notify_var = tk.BooleanVar(value=True)
    dlg.startup_notify_var = tk.BooleanVar(value=True)
    dlg.no_app_action_var = tk.StringVar(value="打开")
    dlg.move_cursor_var = tk.BooleanVar(value=True)
    return dlg


@_needs_tk
def test_on_save_skips_uncreated_tabs_and_collects_created():
    """真实 _on_save：只收集已创建的 general 页，未创建页键保持原值；sanitize 生效。"""
    dlg = _build_save_ready_dialog()
    # 打桩保存与消息框，捕获实际写入的配置
    captured = {}
    dlg.config_loader = type("L", (), {"save": lambda self, cfg: captured.update(cfg)})()
    dlg._show_topmost_message = lambda *a, **k: None
    dlg._confirm_keep_formula_enable = lambda: True
    dlg._call_on_close_callback = lambda: None
    dlg._safe_destroy = lambda: None
    dlg.on_save_callback = None

    # 用户改了 save_dir（general 页）-> 应被收集
    dlg.save_dir_var.set("/custom/save")
    # 模拟手改配置中的坏 paste_delay_s -> 应被无条件 sanitize
    dlg.current_config["paste_delay_s"] = "abc"

    dlg._on_save()

    # general 页修改被收集
    assert captured.get("save_dir") == "/custom/save"
    # 未创建页的键保持 current_config 原值（不被覆盖也不报错）
    assert captured.get("pandoc_path") == dlg.current_config["pandoc_path"]
    assert captured.get("Keep_original_formula") == dlg.current_config["Keep_original_formula"]
    # paste_delay_s 被无条件 sanitize（坏值回落默认）
    assert captured.get("paste_delay_s") == float(DEFAULT_CONFIG.get("paste_delay_s", 0.3))


@_needs_tk
def test_on_save_collection_error_does_not_save_or_mutate_current_config():
    """collector 异常时保存失败可见，且不会写文件或污染原配置。"""
    dlg = _build_save_ready_dialog()
    original_html_formatting = copy.deepcopy(dlg.current_config["html_formatting"])
    save_calls = []
    messages = []

    def bad_collect(config):
        config["html_formatting"]["strikethrough_to_del"] = False
        raise AttributeError("collector bug")

    dlg._tab_specs["general"] = replace(dlg._tab_specs["general"], collect=bad_collect)
    dlg.config_loader = type("L", (), {"save": lambda self, cfg: save_calls.append(cfg)})()
    dlg._show_topmost_message = lambda *args: messages.append(args)
    dlg._confirm_keep_formula_enable = lambda: True
    dlg._call_on_close_callback = lambda: None
    dlg._safe_destroy = lambda: None
    dlg.on_save_callback = None

    dlg._on_save()

    assert save_calls == []
    assert dlg.current_config["html_formatting"] == original_html_formatting
    assert messages[-1][2] == "error"
