"""历史记录浏览对话框."""

import json
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dataclasses import dataclass
from typing import Optional, Callable

from ...utils.system_detect import is_macos

from ...service.history import HistoryManager
from ...service.notification.manager import NotificationManager
from ...utils.clipboard import set_clipboard_text
from ...utils.logging import log
from ...i18n import t

# ===== 单一起源：代码值 → 显示名 =====
# 目标应用
_TARGET_DISPLAY = {
    "word": "Word", "wps": "WPS", "excel": "Excel",
    "wps_excel": "WPS Excel", "onenote": "OneNote",
    "powerpoint": "PowerPoint", "youdao": "Youdao", "": "—",
}
# 兼容别名（详情页等外部模块仍可引用 TARGET_DISPLAY）
TARGET_DISPLAY = _TARGET_DISPLAY
# 工作流（含目标应用 + 扩展工作流 + fallback）
_WORKFLOW_DISPLAY = {
    "word": "Word", "wps": "WPS", "excel": "Excel",
    "wps_excel": "WPS Excel", "onenote": "OneNote",
    "powerpoint": "PowerPoint", "youdao": "Youdao",
    "": "Fallback", "fallback": "Fallback", "html": "HTML+MD",
    "md": "Markdown", "latex": "LaTeX", "file": "File",
}
CONTENT_TYPE_DISPLAY = {
    "markdown": "MD", "html": "HTML", "latex": "LaTeX",
    "table": "Table", "file": "File", "richtext": "RichText",
}

@dataclass
class FilterState:
    status_filter: str = ""
    target_filter: str = ""
    workflow_filter: str = ""
    content_type_filter: str = ""
    keyword: str = ""
    date_from: str = ""
    date_to: str = ""

ROW_HEIGHT = 24

# DatePicker 下拉值常量
_MONTHS = [f"{m:02d}" for m in range(1, 13)]
_DAYS   = [f"{d:02d}" for d in range(1, 32)]

# 跨平台等宽字体：Consolas (Windows) / Menlo (macOS) / TkFixedFont (兜底)
MONO_FONT = ("Menlo", 10) if is_macos() else ("Consolas", 10)

# 过滤下拉映射（显示名 → 代码值），从单一起源推导
def _reverse_display(m):
    """反转 (代码→显示) 为 (显示→代码)，排除空 key 以保留「无」入口。"""
    return {v: k for k, v in m.items() if k}

FILTER_STATUS_MAP = {"无": "", "✓ success": "success", "✗ fail": "fail"}
FILTER_TARGET_MAP = {"无": "", **_reverse_display(_TARGET_DISPLAY)}
FILTER_WORKFLOW_MAP = {"无": "", **_reverse_display(_WORKFLOW_DISPLAY)}
# 内容类型：树视图用短标签、下拉用全名，故映射不互通，独立维护
FILTER_CTYPE_MAP = {
    "无": "",
    "Markdown": "markdown", "HTML": "html", "LaTeX": "latex",
    "Table": "table", "File": "file", "RichText": "richtext",
}


class HistoryDialog:
    """粘贴历史浏览/搜索/管理窗口。"""

    def __init__(
        self,
        history_manager: HistoryManager,
        notification_manager: NotificationManager,
        on_close: Optional[Callable[[], None]] = None,
    ):
        self._hm = history_manager
        self._nm = notification_manager
        self._on_close = on_close
        self._root: Optional[tk.Toplevel] = None
        self._tree: Optional[ttk.Treeview] = None
        self._search_var: Optional[tk.StringVar] = None
        self._filter_status: Optional[tk.StringVar] = None
        self._filter_target: Optional[tk.StringVar] = None
        self._filter_workflow: Optional[tk.StringVar] = None
        self._filter_type: Optional[tk.StringVar] = None
        # 日期 Combobox 变量 (年/月/日) - from
        self._df_year: Optional[tk.StringVar] = None
        self._df_month: Optional[tk.StringVar] = None
        self._df_day: Optional[tk.StringVar] = None
        self._dt_year: Optional[tk.StringVar] = None
        self._dt_month: Optional[tk.StringVar] = None
        self._dt_day: Optional[tk.StringVar] = None
        # 过滤栏折叠
        self._filters_expanded = False
        self._filter_frame: Optional[ttk.Frame] = None
        self._filter_toggle_btn: Optional[ttk.Button] = None
        self._sort_col = "created_at"
        self._sort_order = "DESC"
        self._page_size = 100
        self._current_page = 0
        self._suppress_trace = False
        self._search_after_id = None

    # ------------------------------------------------------------------
    # Show / Focus
    # ------------------------------------------------------------------

    def show(self) -> None:
        self._root = tk.Toplevel()
        self._root.title(t("history.dialog.title"))
        # 获取屏幕尺寸，窗口占 80% 宽 x 75% 高
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()
        w = max(900, int(sw * 0.65))
        h = max(580, int(sh * 0.6))
        self._root.geometry(f"{w}x{h}")
        self._root.minsize(700, 450)
        self._root.protocol("WM_DELETE_WINDOW", self._on_window_close)

        try:
            from ...config.paths import get_app_png_path
            ico = get_app_png_path()
            if ico:
                img = tk.PhotoImage(file=ico)
                self._root.iconphoto(False, img)
        except Exception:
            pass

        self._build_ui()
        self._refresh_list()

    def lift(self) -> None:
        if self._root:
            try:
                self._root.lift()
                self._root.focus_force()
            except Exception:
                pass

    def focus_force(self) -> None:
        if self._root:
            try:
                self._root.focus_force()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Build UI
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        root = self._root
        if root is None:
            return

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=0)
        root.rowconfigure(1, weight=0)
        root.rowconfigure(2, weight=1)
        root.rowconfigure(3, weight=0)

        self._build_search_bar(root)
        self._build_filter_panel(root)
        self._build_tree(root)
        self._build_bottom_bar(root)
        self._update_stats()

        if self._filter_frame:
            self._filter_frame.grid_remove()

    # ---- sub-builder methods ----

    def _build_search_bar(self, root) -> None:
        tb0 = ttk.Frame(root, padding=(8, 8, 8, 0))
        tb0.grid(row=0, column=0, sticky="ew")
        tb0.columnconfigure(1, weight=1)

        ttk.Label(tb0, text=t("history.search.label"), font=("", 10, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        self._search_var = tk.StringVar()
        self._search_entry = ttk.Entry(tb0, textvariable=self._search_var)
        self._search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self._search_entry.bind("<Return>", lambda _e: self._do_search())
        self._search_entry.bind("<KeyRelease>", self._on_search_typing)
        ttk.Button(tb0, text=t("history.search.search"), command=self._do_search).grid(
            row=0, column=2, sticky="w", padx=(0, 4))
        ttk.Button(tb0, text=t("history.filter.clear"), command=self._clear_filters).grid(
            row=0, column=3, sticky="w", padx=(0, 4))

        self._filter_toggle_btn = ttk.Button(tb0, text="▸ " + t("history.filter.toggle"),
                                             command=self._toggle_filters, width=6)
        self._filter_toggle_btn.grid(row=0, column=4, sticky="w", padx=(0, 8))

        self._stats_label = ttk.Label(tb0, text="", font=("", 8))
        self._stats_label.grid(row=0, column=5, sticky="e")

    def _build_filter_panel(self, root) -> None:
        self._filter_frame = ttk.Frame(root, padding=(8, 4, 8, 2))
        self._filter_frame.grid(row=1, column=0, sticky="ew")
        self._filter_frame.columnconfigure(0, weight=1)

        self._filter_status = tk.StringVar(value="无")
        self._filter_target = tk.StringVar(value="无")
        self._filter_workflow = tk.StringVar(value="无")
        self._filter_type = tk.StringVar(value="无")

        def _on_filter_change(*_a):
            if self._suppress_trace: return
            self._current_page = 0
            self._refresh_list()

        for v in (self._filter_status, self._filter_target, self._filter_workflow, self._filter_type):
            v.trace_add("write", _on_filter_change)

        # Row 0: status / target / type / workflow
        ftop = ttk.Frame(self._filter_frame)
        ftop.grid(row=0, column=0, sticky="ew")

        for col, (label, var, values) in enumerate([
            (t("history.filter.status"),       self._filter_status,    ["无", "✓ success", "✗ fail"]),
            (t("history.filter.target"),       self._filter_target,    list(FILTER_TARGET_MAP.keys())),
            (t("history.filter.content_type"), self._filter_type,      list(FILTER_CTYPE_MAP.keys())),
            (t("history.filter.workflow"),     self._filter_workflow,  list(FILTER_WORKFLOW_MAP.keys())),
        ]):
            ttk.Label(ftop, text=label).grid(row=0, column=col * 2, sticky="w", padx=(0, 2))
            ttk.Combobox(ftop, textvariable=var, values=values,
                         state="readonly", width=(9 if col == 0 else (7 if col == 2 else 8))
                         ).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 10) if col < 3 else (0, 0))

        # Row 1: date range
        self._df_year  = tk.StringVar(value="")
        self._df_month = tk.StringVar(value="")
        self._df_day   = tk.StringVar(value="")
        self._dt_year  = tk.StringVar(value="")
        self._dt_month = tk.StringVar(value="")
        self._dt_day   = tk.StringVar(value="")

        def _on_date_combo_change(*_a):
            if self._suppress_trace: return
            self._current_page = 0
            self._refresh_list()

        for v in (self._df_year, self._df_month, self._df_day,
                  self._dt_year, self._dt_month, self._dt_day):
            v.trace_add("write", _on_date_combo_change)

        fbot = ttk.Frame(self._filter_frame)
        fbot.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        ttk.Label(fbot, text=t("history.filter.date_from"), font=("", 8)).grid(
            row=0, column=0, sticky="w", padx=(0, 2))
        self._build_date_row(fbot, 0, 1,
                             self._df_year, self._df_month, self._df_day,
                             padx_end=(0, 12))

        ttk.Label(fbot, text=t("history.filter.date_to"), font=("", 8)).grid(
            row=0, column=4, sticky="w", padx=(0, 2))
        self._build_date_row(fbot, 0, 5,
                             self._dt_year, self._dt_month, self._dt_day,
                             padx_end=(0, 0))

    def _build_tree(self, root) -> None:
        tree_frame = ttk.Frame(root)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(6, 0))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=ROW_HEIGHT)

        self._tree = ttk.Treeview(tree_frame,
            columns=("time", "target", "type", "preview", "filters", "status"),
            show="headings", selectmode="extended")
        sort_cols = {"time": "created_at", "target": "target_app", "type": "content_type", "status": "status"}
        for col_id, text, pct in [
            ("time",    t("history.column.time"),    0.12),
            ("target",  t("history.column.target"),  0.07),
            ("type",    t("history.column.type"),    0.05),
            ("preview", t("history.column.preview"), 0.50),
            ("filters", t("history.column.filters"), 0.18),
            ("status",  t("history.column.status"),  0.04),
        ]:
            sa = sort_cols.get(col_id)
            cmd = (lambda a=sa: self._sort(a)) if sa else (lambda: None)
            self._tree.heading(col_id, text=text, command=cmd)
            self._tree.column(col_id, width=int(pct * 900), minwidth=30, stretch=True)

        self._tree.column("time", minwidth=100)
        self._tree.column("preview", minwidth=150)
        self._tree.column("filters", minwidth=80)
        self._tree.column("status", anchor="center")
        self._tree.tag_configure("success", foreground="#000000")
        self._tree.tag_configure("fail", foreground="#cc0000")
        self._tree.tag_configure("pinned", foreground="#0055aa")

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self._tree.bind("<Double-Button-1>", self._on_double_click)
        self._tree.bind("<Button-2>" if sys.platform == "darwin" else "<Button-3>", self._on_right_click)
        self._tree.bind("<Delete>", self._on_delete_key)
        root.bind("<Escape>", lambda _e: self._on_window_close())
        root.bind("<Command-f>" if sys.platform == "darwin" else "<Control-f>",
                  lambda _e: self._search_entry.focus_set())

        self._tree.heading("time", text=t("history.column.time") + " ▼")

    def _build_bottom_bar(self, root) -> None:
        bottom = ttk.Frame(root, padding=(8, 4))
        bottom.grid(row=3, column=0, sticky="ew")
        ttk.Button(bottom, text=t("history.button.clear_all"), command=self._clear_all).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bottom, text=t("history.button.copy_selected"), command=self._copy_selected).pack(side=tk.LEFT)
        self._page_label = ttk.Label(bottom, text="")
        self._page_label.pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(bottom, text=">", width=2, command=self._next_page).pack(side=tk.RIGHT)
        ttk.Button(bottom, text="<", width=2, command=self._prev_page).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Date selector helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _build_date_row(parent: ttk.Frame, row: int, start_col: int,
                        yv: tk.StringVar, mv: tk.StringVar, dv: tk.StringVar,
                        padx_end=(0, 0)) -> None:
        years = [""] + [str(y) for y in range(2024, 2031)]

        cy = ttk.Combobox(parent, textvariable=yv, values=years, state="readonly", width=5)
        cy.grid(row=row, column=start_col, sticky="w", padx=(0, 1))
        cm = ttk.Combobox(parent, textvariable=mv, values=[""] + _MONTHS,
                          state="readonly", width=3)
        cm.grid(row=row, column=start_col + 1, sticky="w", padx=(0, 1))
        cd = ttk.Combobox(parent, textvariable=dv, values=[""] + _DAYS,
                          state="readonly", width=3)
        cd.grid(row=row, column=start_col + 2, sticky="w", padx=padx_end)

    def _toggle_filters(self) -> None:
        """折叠/展开筛选面板。"""
        self._filters_expanded = not self._filters_expanded
        if self._filter_frame:
            if self._filters_expanded:
                self._filter_frame.grid()
            else:
                self._filter_frame.grid_remove()
        if self._filter_toggle_btn:
            arrow = "▾ " if self._filters_expanded else "▸ "
            self._filter_toggle_btn.config(text=arrow + t("history.filter.toggle"))

    # ------------------------------------------------------------------
    # Data
    # ------------------------------------------------------------------

    def _refresh_list(self) -> None:
        tree = self._tree
        if tree is None:
            return
        for item in tree.get_children():
            tree.delete(item)

        f = self._effective_filters()
        offset = self._current_page * self._page_size

        kwargs = dict(status_filter=f.status_filter, target_filter=f.target_filter,
                      workflow_filter=f.workflow_filter, content_type_filter=f.content_type_filter,
                      date_from=f.date_from, date_to=f.date_to)
        if f.keyword:
            rows = self._hm.search(keyword=f.keyword, limit=self._page_size,
                                   sort_by=self._sort_col, order=self._sort_order,
                                   offset=offset, **kwargs)
        else:
            rows = self._hm.query(limit=self._page_size, offset=offset,
                                  sort_by=self._sort_col, order=self._sort_order, **kwargs)

        for row in rows:
            is_pinned = row.get("pinned", False)
            prefix = "📌 " if is_pinned else ""
            status_icon = "✓" if row["status"] == "success" else "✗"
            filters_str = self._format_filters_preview(row.get("filters_json", "[]"))
            preview_text = (row.get("preview", "") or "")[:160]
            # 压缩换行符为空格，避免 TreeView 行高异常
            preview_text = " ".join(preview_text.split())

            tree.insert("", tk.END, iid=str(row["id"]),
                        values=(prefix + row["created_at"],
                                TARGET_DISPLAY.get(row.get("target_app", ""), row.get("target_app", "—")),
                                CONTENT_TYPE_DISPLAY.get(row.get("content_type", ""), row.get("content_type", "")),
                                preview_text,
                                filters_str,
                                status_icon),
                        tags=(row["status"], "pinned" if is_pinned else ""))

        total = self._hm.count(status_filter=f.status_filter, target_filter=f.target_filter,
                               workflow_filter=f.workflow_filter, content_type_filter=f.content_type_filter,
                               date_from=f.date_from, date_to=f.date_to,
                               keyword=f.keyword)
        if total == 0:
            self._page_label.config(text=t("history.empty"))
        else:
            self._page_label.config(text=f"{offset + 1}-{offset + len(rows)} / {total}")

    @staticmethod
    def _format_filters_preview(filters_json: str) -> str:
        try:
            fl = json.loads(filters_json) if filters_json else []
            if not fl:
                return "—"
            return ", ".join(fl[:6]) + ("…" if len(fl) > 6 else "")
        except Exception:
            return "—"

    @staticmethod
    def _format_workflow_key(key: str) -> str:
        """将 workflow_key 转为可读显示名。"""
        return _WORKFLOW_DISPLAY.get(key, key) or "—"

    @staticmethod
    def _format_output_bytes(b: int) -> str:
        """人性化显示字节数。"""
        if b <= 0:
            return "—"
        if b < 1024:
            return f"{b} B"
        if b < 1024 * 1024:
            return f"{b / 1024:.1f} KB"
        return f"{b / (1024 * 1024):.1f} MB"

    @staticmethod
    def _format_output_file(path: str) -> str:
        """有文件路径时显示文件名，否则 —。"""
        import os as _os
        if not path or not path.strip():
            return "—"
        filename = _os.path.basename(path)
        return filename if filename else path

    def _effective_filters(self) -> FilterState:
        """返回当前生效的过滤条件。"""
        keyword = (self._search_var.get() or "").strip()
        return FilterState(
            status_filter=FILTER_STATUS_MAP.get(self._filter_status.get() or "", ""),
            target_filter=FILTER_TARGET_MAP.get(self._filter_target.get() or "", ""),
            workflow_filter=FILTER_WORKFLOW_MAP.get(self._filter_workflow.get() or "", ""),
            content_type_filter=FILTER_CTYPE_MAP.get(self._filter_type.get() or "", ""),
            keyword=keyword,
            date_from=self._get_date_var(self._df_year, self._df_month, self._df_day),
            date_to=self._get_date_var(self._dt_year, self._dt_month, self._dt_day),
        )

    @staticmethod
    def _get_date_var(year_var, month_var, day_var) -> str:
        """从年/月/日 Combobox 拼出 YYYY-MM-DD，不全则返回 ""."""
        y = (year_var.get() or "").strip()
        m = (month_var.get() or "").strip()
        d = (day_var.get() or "").strip()
        if y and m and d:
            return f"{y}-{m}-{d}"
        return ""

    def _update_stats(self) -> None:
        try:
            s = self._hm.get_stats()
            self._stats_label.config(
                text=f"{t('history.stats.total')}: {s['total']}  "
                     f"{t('history.stats.today')}: {s['today']}  {s['file_size_kb']}KB"
            )
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _do_search(self) -> None:
        if self._root is None:
            return
        self._current_page = 0
        self._refresh_list()

    def _on_search_typing(self, event) -> None:
        if self._search_after_id is not None:
            try:
                self._root.after_cancel(self._search_after_id)
            except Exception:
                pass
        if self._root is not None:
            self._search_after_id = self._root.after(300, self._do_search)

    def _clear_filters(self) -> None:
        self._suppress_trace = True
        try:
            self._search_var.set("")
            self._filter_status.set("无")
            self._filter_target.set("无")
            self._filter_workflow.set("无")
            self._filter_type.set("无")
            for v in (self._df_year, self._df_month, self._df_day,
                      self._dt_year, self._dt_month, self._dt_day):
                if v:
                    v.set("")
        finally:
            self._suppress_trace = False
        self._current_page = 0
        self._refresh_list()

    def _on_delete_key(self, event) -> None:
        sel = self._tree.selection()
        if sel:
            for iid in sel:
                self._hm.delete_entry(int(iid))
            self._hm.flush()
            self._refresh_list()
            self._update_stats()

    def _sort(self, column: str) -> None:
        self._sort_order = "ASC" if (self._sort_col == column and self._sort_order == "DESC") else "DESC"
        self._sort_col = column

        # 先清除所有排序列的箭头（恢复基础文字）
        col_map = {"created_at": "time", "target_app": "target", "content_type": "type", "status": "status"}
        base_map = {
            "time": t("history.column.time"), "target": t("history.column.target"),
            "type": t("history.column.type"), "status": t("history.column.status"),
        }
        for db_col, disp_col in col_map.items():
            self._tree.heading(disp_col, text=base_map.get(disp_col, ""))

        # 再给当前排序列加上箭头
        arrow = " ▼" if self._sort_order == "DESC" else " ▲"
        dc = col_map.get(column, column)
        base = base_map.get(dc, "")
        self._tree.heading(dc, text=base + arrow)
        self._current_page = 0
        self._refresh_list()

    def _next_page(self) -> None:
        f = self._effective_filters()
        total = self._hm.count(status_filter=f.status_filter, target_filter=f.target_filter,
                               workflow_filter=f.workflow_filter, content_type_filter=f.content_type_filter,
                               date_from=f.date_from, date_to=f.date_to,
                               keyword=f.keyword)
        if (self._current_page + 1) * self._page_size < total:
            self._current_page += 1
            self._refresh_list()

    def _prev_page(self) -> None:
        if self._current_page > 0:
            self._current_page -= 1
            self._refresh_list()

    def _on_double_click(self, event) -> None:
        sel = self._tree.selection()
        if sel:
            self._show_detail(int(sel[0]))

    def _on_right_click(self, event) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        menu = tk.Menu(self._root, tearoff=0)
        menu.add_command(label=t("history.context.view_detail"), command=lambda: self._show_detail(int(sel[0])))
        menu.add_command(label=t("history.context.copy_content"), command=lambda: self._copy_entry_content(int(sel[0])))
        menu.add_command(label=t("history.context.toggle_pin"), command=lambda: self._toggle_pin(int(sel[0])))
        menu.add_separator()
        menu.add_command(label=t("history.context.delete"), command=lambda: self._delete_entry(int(sel[0])))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _show_detail(self, entry_id: int) -> None:
        entry = self._hm.get_entry(entry_id)
        if entry is None:
            return

        detail = tk.Toplevel(self._root)
        detail.title(t("history.detail.title", id=entry_id))
        sw = detail.winfo_screenwidth()
        sh = detail.winfo_screenheight()
        detail.geometry(f"{(max(750, int(sw * 0.55)))}x{max(620, int(sh * 0.6))}")
        detail.minsize(500, 400)
        detail.columnconfigure(0, weight=1)
        detail.rowconfigure(3, weight=1)  # content area gets the flex space

        try:
            from ...config.paths import get_app_png_path
            img = tk.PhotoImage(file=get_app_png_path())
            detail.iconphoto(False, img)
        except Exception:
            pass

        # ── 基础元数据 (2 列网格) ──
        meta_frame = ttk.Frame(detail, padding=(12, 10))
        meta_frame.grid(row=0, column=0, sticky="ew")
        for c in range(4):
            meta_frame.columnconfigure(c, weight=1)

        fields_left = [
            ("history.detail.time",          entry.get("created_at", "")),
            ("history.detail.target_app",    TARGET_DISPLAY.get(entry.get("target_app", ""), entry.get("target_app", "—"))),
            ("history.detail.window",        entry.get("window_title", "") or "—"),
            ("history.detail.workflow",      self._format_workflow_key(entry.get("workflow_key", ""))),
            ("history.detail.source_format", CONTENT_TYPE_DISPLAY.get(entry.get("source_format", ""), entry.get("source_format", ""))),
        ]
        fields_right = [
            ("history.detail.content_type",  CONTENT_TYPE_DISPLAY.get(entry.get("content_type", ""), entry.get("content_type", ""))),
            ("history.detail.output_bytes",  self._format_output_bytes(entry.get("output_bytes", 0))),
            ("history.detail.output_file",   self._format_output_file(entry.get("output_file_path", "") or "")),
            ("history.detail.status",
             "✓ " + t("history.status.success") if entry["status"] == "success"
             else "✗ " + t("history.status.fail")),
        ]

        max_rows = max(len(fields_left), len(fields_right))
        for i in range(max_rows):
            if i < len(fields_left):
                ttk.Label(meta_frame, text=t(fields_left[i][0]) + ": " + fields_left[i][1],
                         font=("", 9)).grid(row=i, column=0, columnspan=2, sticky="w", pady=1)
            if i < len(fields_right):
                ttk.Label(meta_frame, text=t(fields_right[i][0]) + ": " + fields_right[i][1],
                         font=("", 9)).grid(row=i, column=2, columnspan=2, sticky="w", pady=1)

        # 错误信息
        if entry.get("error_msg"):
            ttk.Label(meta_frame, text=t("history.detail.error") + ": " + entry["error_msg"][:300],
                     foreground="#cc0000", font=("", 9)).grid(
                         row=max_rows, column=0, columnspan=4, sticky="w", pady=(2, 0))

        # ── 转换管道 + 过滤器 (醒目的 Frame) ──
        extra_frame = ttk.Frame(detail, padding=(12, 2))
        extra_frame.grid(row=1, column=0, sticky="ew")

        pipeline_text = ""
        try:
            pipe = json.loads(entry.get("conversion_pipeline", "{}"))
            if isinstance(pipe, dict):
                parts = []
                if pipe.get("input"):
                    parts.append(f"[{pipe['input']}]")
                if pipe.get("steps"):
                    parts.append(" → ".join(pipe["steps"]))
                pipeline_text = "  ".join(parts)
        except Exception:
            pass
        if pipeline_text:
            ttk.Label(extra_frame, text=f"🔄 {t('history.detail.pipeline')}{pipeline_text}",
                     font=("", 9), foreground="#1a6e8e").pack(anchor="w")

        try:
            fl = json.loads(entry.get("filters_json", "[]"))
        except Exception:
            fl = []
        if fl:
            # 用 Frame 包裹成一排 tag 样式
            filter_row = ttk.Frame(extra_frame)
            filter_row.pack(anchor="w", fill=tk.X, pady=(2, 0))
            ttk.Label(filter_row, text=f"⚙ {t('history.detail.filters')}",
                     font=("", 9), foreground="#555").pack(side=tk.LEFT)
            # 每 6 个换行避免溢出
            for idx, f in enumerate(fl):
                if idx > 0 and idx % 8 == 0:
                    ttk.Label(filter_row, text="").pack(side=tk.LEFT)
                tag = ttk.Label(filter_row, text=f, font=("", 8),
                               background="#e8e8e8", relief="solid", padding=(3, 1))
                tag.pack(side=tk.LEFT, padx=2, pady=1)

        # ── 内容区 ──
        ttk.Separator(detail, orient=tk.HORIZONTAL).grid(row=2, column=0, sticky="ew", padx=8, pady=(6, 4))

        content_frame = ttk.Frame(detail, padding=(12, 4))
        content_frame.grid(row=3, column=0, sticky="nsew")
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        ct = entry.get("full_content", "") or entry.get("preview", "")
        text_widget = tk.Text(content_frame, wrap=tk.WORD, font=MONO_FONT)
        text_widget.insert("1.0", ct)
        text_widget.configure(state=tk.DISABLED)

        sb = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=sb.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # 按钮
        btn_frame = ttk.Frame(detail, padding=(12, 8))
        btn_frame.grid(row=4, column=0, sticky="ew")
        ttk.Button(btn_frame, text=t("history.detail.copy"),
                   command=lambda: self._copy_to_clipboard(ct)).pack(side=tk.LEFT, padx=(0, 8))

        # HTML 源码下拉按钮（仅当有 HTML 来源时显示）
        original_html_str = entry.get("original_html", "") or ""
        if original_html_str.strip():
            full_text = entry.get("full_content", "") or ""
            self._build_html_dropdown(btn_frame, original_html_str, full_text)

        ttk.Button(btn_frame, text=t("history.detail.close"), command=detail.destroy).pack(side=tk.RIGHT)

    def _copy_html_to_clipboard(self, html: str, plain_text: str = "") -> None:
        """以 CF_HTML + CF_TEXT 写回剪贴板，复现网页复制时的原始状态。"""
        try:
            from ...utils.clipboard import set_clipboard_rich_text
            set_clipboard_rich_text(html=html, text=plain_text.strip() if plain_text else None)
            self._nm.notify("PasteMD", t("history.copied_html"), ok=True)
        except Exception as e:
            log(f"HTML clipboard copy failed: {e}")
            self._copy_to_clipboard(html or "")

    def _build_html_dropdown(self, parent: ttk.Frame, html: str, plain_text: str) -> None:
        """创建 HTML 源码的下拉菜单按钮（复制 / 导出文件）。"""
        mb = ttk.Menubutton(parent, text=t("history.detail.copy_html"), direction="above")
        mb.pack(side=tk.LEFT, padx=(0, 8))
        menu = tk.Menu(mb, tearoff=0)
        menu.add_command(
            label=t("history.detail.copy_html_clipboard"),
            command=lambda: self._copy_html_to_clipboard(html, plain_text),
        )
        menu.add_command(
            label=t("history.detail.copy_html_export"),
            command=lambda: self._export_html_file(html),
        )
        mb.configure(menu=menu)

    def _export_html_file(self, html: str) -> None:
        """导出原始 HTML 到文件。"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            title=t("history.detail.copy_html_export"),
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(html)
            self._nm.notify("PasteMD", t("history.exported_html", path=file_path), ok=True)
        except Exception as e:
            log(f"HTML export failed: {e}")
            self._nm.notify("PasteMD", t("history.exported_html_failed"), ok=False)

    def _copy_to_clipboard(self, content: str) -> None:
        try:
            set_clipboard_text(content)
            self._nm.notify("PasteMD", t("history.copied"), ok=True)
        except Exception as e:
            log(f"History copy failed: {e}")
            try:
                if self._root is not None and self._root.winfo_exists():
                    self._root.clipboard_clear()
                    self._root.clipboard_append(content)
                    self._nm.notify("PasteMD", t("history.copied"), ok=True)
            except Exception as fallback_exc:
                log(f"History Tk clipboard fallback failed: {fallback_exc}")

    def _copy_entry_content(self, entry_id: int) -> None:
        entry = self._hm.get_entry(entry_id)
        if entry:
            self._copy_to_clipboard(entry.get("full_content", "") or entry.get("preview", ""))

    def _copy_selected(self) -> None:
        sel = self._tree.selection()
        if sel:
            self._copy_entry_content(int(sel[0]))

    def _toggle_pin(self, entry_id: int) -> None:
        self._hm.toggle_pin(entry_id)
        self._hm.flush()
        self._refresh_list()

    def _delete_entry(self, entry_id: int) -> None:
        self._hm.delete_entry(entry_id)
        self._hm.flush()
        self._refresh_list()
        self._update_stats()

    def _clear_all(self) -> None:
        if not messagebox.askyesno(t("history.clear.title"), t("history.clear.confirm")):
            return
        self._hm.clear_all()
        self._hm.flush()
        self._current_page = 0
        self._refresh_list()
        self._update_stats()
        self._nm.notify("PasteMD", t("history.clear.done"), ok=True)

    def _on_window_close(self) -> None:
        if self._on_close:
            self._on_close()
        if self._root:
            self._root.destroy()
            self._root = None
