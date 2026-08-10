"""Settings configuration dialog."""

import copy
import os
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Callable, Dict, Any

from ...config.paths import get_app_icon_path
from ...utils.logging import log
from ...utils.dpi import get_dpi_scale
from ...utils.system_detect import is_windows, is_macos
from ...i18n import t, iter_languages, get_language_label, get_no_app_action_map
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.defaults import DEFAULT_CONFIG
from .extensions_tab import ExtensionsTab

if is_macos():
    from .permissions import MacOSPermissionsTab


class SettingsDialog:
    """设置对话框"""
    
    def __init__(
        self,
        on_save: Callable[[], None],
        on_close: Optional[Callable[[], None]] = None,
        initial_tab: Optional[str] = None,
    ):
        """
        初始化设置对话框
        
        Args:
            on_save: 保存回调函数
            on_close: 关闭回调函数
        """
        self.on_save_callback = on_save
        self.on_close_callback = on_close
        self._close_callback_called = False
        self._open_hotkey_dialog: Optional[Callable[[], None]] = None
        self._initial_tab = initial_tab
        self._tab_map: Dict[str, tk.Widget] = {}
        self._permissions_tab = None
        self.config_loader = ConfigLoader()
        
        # 加载当前配置的副本，避免直接修改 app_state
        self.current_config = copy.deepcopy(app_state.config)
        
        self._conversion_filter_keys = [
            "md_to_docx",
            "html_to_docx",
            "html_to_md",
            "md_to_html",
            "md_to_rtf",
            "md_to_latex",
        ]
        self.filter_conversion_options = [
            ("global", t("settings.conversion.filter_conversion_global")),
            ("md_to_docx", t("settings.conversion.filter_conversion_md_to_docx")),
            ("html_to_docx", t("settings.conversion.filter_conversion_html_to_docx")),
            ("html_to_md", t("settings.conversion.filter_conversion_html_to_md")),
            ("md_to_html", t("settings.conversion.filter_conversion_md_to_html")),
            ("md_to_rtf", t("settings.conversion.filter_conversion_md_to_rtf")),
            ("md_to_latex", t("settings.conversion.filter_conversion_md_to_latex")),
        ]
        self._filter_conversion_label_to_key = {
            label: key for key, label in self.filter_conversion_options
        }
        self.horizontal_rule_style_options = [
            ("default", t("settings.conversion.horizontal_rule_style_default")),
            ("paragraph_border", t("settings.conversion.horizontal_rule_style_paragraph_border")),
        ]
        self._horizontal_rule_style_label_to_key = {
            label: key for key, label in self.horizontal_rule_style_options
        }
        self._horizontal_rule_style_key_to_label = {
            key: label for key, label in self.horizontal_rule_style_options
        }

        raw_filters_by_conversion = self.current_config.get("pandoc_filters_by_conversion")
        if not isinstance(raw_filters_by_conversion, dict):
            raw_filters_by_conversion = {}

        raw_global_filters = self.current_config.get("pandoc_filters") or []
        if isinstance(raw_global_filters, str):
            raw_global_filters = [raw_global_filters]
        elif not isinstance(raw_global_filters, (list, tuple)):
            raw_global_filters = []
        self.global_filters: list[dict] = [
            item for item in (self._parse_filter_item(f) for f in raw_global_filters)
            if item is not None
        ]

        self.filters_by_conversion: Dict[str, list[dict]] = {}
        for key in self._conversion_filter_keys:
            raw_list = raw_filters_by_conversion.get(key, [])
            if isinstance(raw_list, str):
                raw_list = [raw_list]
            elif not isinstance(raw_list, (list, tuple)):
                raw_list = []
            self.filters_by_conversion[key] = [
                item for item in (self._parse_filter_item(f) for f in raw_list)
                if item is not None
            ]

        default_label = self.filter_conversion_options[0][1]
        self.filter_conversion_var = tk.StringVar(value=default_label)
        
        if app_state.root:
            self.root = tk.Toplevel(app_state.root)
        else:
            self.root = tk.Tk()
            
        self.root.title(t("settings.dialog.title"))
        # 默认不强制置顶，便于最小化
        self.root.attributes("-topmost", False)
        # 确保有最小化按钮
        self.root.resizable(True, True)
        
        # Windows 特有属性
        if is_windows():
            self.root.attributes("-toolwindow", False)
            # 设置图标
            try:
                icon_path = get_app_icon_path()
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
            except Exception as e:
                log(f"Failed to set settings dialog icon: {e}")
        
        # 适配高分屏和屏幕尺寸
        scale = get_dpi_scale()
        
        # 在 macOS 上根据屏幕大小自适应
        if not is_windows():
            # 获取屏幕尺寸
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # 窗口大小设置为屏幕的 50%，但不超过 700x600，不小于 500x450
            width = max(500, min(700, int(screen_width * 0.5)))
            height = max(450, min(600, int(screen_height * 0.5)))
        else:
            # Windows 保持原来的固定大小
            width = int(600 * scale)
            height = int(500 * scale)
        
        self.root.geometry(f"{width}x{height}")
        
        # 设置关闭窗口时的处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 窗口居中
        self._center_window()
        
        # 延迟创建 UI 组件：窗口先显示再填充，避免低配机器上黑屏/卡顿
        self.root.after_idle(self._lazy_create_widgets)

    def set_open_hotkey_dialog(self, callback: Optional[Callable[[], None]]) -> None:
        """Inject a callback to open the hotkey dialog (provided by TrayMenuManager)."""
        self._open_hotkey_dialog = callback
        try:
            btn = getattr(self, "set_hotkey_btn", None)
            if btn is not None:
                btn.configure(state="normal" if callback else "disabled")
        except Exception:
            pass

    def refresh_hotkey_display(self) -> None:
        """Refresh the hotkey label from global config/state."""
        try:
            if hasattr(self, "hotkey_display_var"):
                value = app_state.config.get("hotkey") or getattr(app_state, "hotkey_str", "")
                self.hotkey_display_var.set(str(value))
        except Exception as e:
            log(f"Failed to refresh hotkey display: {e}")

    def _call_on_close_callback(self):
        """确保关闭回调只调用一次"""
        if self._close_callback_called:
            return
        self._close_callback_called = True
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                log(f"Error in close callback: {e}")
    
    def is_alive(self) -> bool:
        """判断窗口是否仍存在"""
        try:
            return bool(self.root.winfo_exists())
        except Exception:
            return False
    
    def restore_and_focus(self):
        """从最小化恢复并置于前台"""
        if not self.is_alive():
            return
        
        try:
            self.root.deiconify()
            original_topmost = bool(self.root.attributes("-topmost"))
            # 通过临时取消/恢复置顶让窗口浮到最前
            self.root.attributes("-topmost", False)
            self.root.lift()
            self.root.focus_force()
            self.root.attributes("-topmost", original_topmost)
        except Exception as e:
            log(f"Failed to restore settings dialog: {e}")
        
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def _lazy_create_widgets(self):
        """延迟创建 UI 组件（窗口显示后再创建，避免初始化黑屏）"""
        try:
            if not self.is_alive():
                return
            self._create_widgets()
            if self._initial_tab and self._initial_tab != "general":
                # 标签页可能已由即时创建路径创建（macOS），也可能还未创建（Windows 懒加载）
                if self._initial_tab not in self._tab_created:
                    self._lazy_create_tab_at(self._initial_tab)
                self.select_tab(self._initial_tab)
        except Exception as e:
            log(f"Failed to create settings widgets lazily: {e}")

    def _create_widgets(self):
        """创建UI组件"""
        # 底部按钮栏（先创建，确保不被覆盖）
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)

        cancel_btn = ttk.Button(
            button_frame,
            text=t("settings.buttons.cancel"),
            command=self._on_cancel,
            width=10
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

        save_btn = ttk.Button(
            button_frame,
            text=t("settings.buttons.save"),
            command=self._on_save,
            width=10
        )
        save_btn.pack(side=tk.RIGHT, padx=5)

        # 创建 Notebook (选项卡容器)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        # 标签页跟踪
        self._tab_created = set()
        self._extensions_tab = None
        self._permissions_tab = None

        # General 标签页 - 始终即时创建（默认首个可见）
        self._create_general_tab()
        self._tab_created.add("general")

        if is_macos():
            # macOS 性能足够，全部标签页即时创建，避免懒加载的标签切换延迟
            self._create_conversion_tab()
            self._tab_created.add("conversion")
            self._create_advanced_tab()
            self._tab_created.add("advanced")
            self._create_experimental_tab()
            self._tab_created.add("experimental")

            try:
                self._extensions_tab = ExtensionsTab(self.notebook, self.current_config)
                self._tab_map["extensions"] = self._extensions_tab.frame
                self._tab_created.add("extensions")
            except Exception as e:
                log(f"Failed to create extensions tab: {e}")

            try:
                self._permissions_tab = MacOSPermissionsTab(self.notebook, self.root)
                self._tab_map["permissions"] = self._permissions_tab.frame
                self._tab_created.add("permissions")
            except Exception as e:
                log(f"Failed to create permissions tab: {e}")
        else:
            # Windows 低配机器：懒加载，仅加占位符，点击标签时才创建
            self._lazy_tab_keys = ["conversion", "advanced", "experimental"]
            self._lazy_tab_labels = {
                "conversion": lambda: t("settings.tab.conversion"),
                "advanced":   lambda: t("settings.tab.advanced"),
                "experimental": lambda: t("settings.tab.experimental"),
            }
            for key in self._lazy_tab_keys:
                placeholder = ttk.Frame(self.notebook)
                self.notebook.add(placeholder, text=self._lazy_tab_labels[key]())
                self._tab_map[key] = placeholder

            # Extensions 标签页占位
            self._extensions_placeholder = ttk.Frame(self.notebook)
            self.notebook.add(self._extensions_placeholder, text=t("settings.tab.extensions"))
            self._tab_map["extensions"] = self._extensions_placeholder

            # 监听标签页切换，触发懒加载
            self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # 避免首次打开时就选中首个输入框
        self.root.after_idle(self._clear_initial_selection)

    def _get_tab_key_by_index(self, idx: int) -> str | None:
        """根据 notebook 索引获取标签页键名（仅用于 Windows 懒加载路径）"""
        tab_order = ["general", *self._lazy_tab_keys, "extensions"]
        if 0 <= idx < len(tab_order):
            return tab_order[idx]
        return None

    def _on_tab_changed(self, event=None) -> None:
        """标签页切换时触发懒加载"""
        try:
            selected = self.notebook.select()
            idx = self.notebook.index(selected)
            key = self._get_tab_key_by_index(idx)
            if key is None or key in self._tab_created:
                return
            self._lazy_create_tab_at(key)
        except Exception as e:
            log(f"Failed to lazy-create tab: {e}")

    def _lazy_create_tab_at(self, key: str) -> None:
        """创建指定标签页的真实内容，替换占位框架"""
        if key in self._tab_created:
            return
        if key == "conversion":
            self._do_lazy_swap(key, self._create_conversion_tab)
        elif key == "advanced":
            self._do_lazy_swap(key, self._create_advanced_tab)
        elif key == "experimental":
            self._do_lazy_swap(key, self._create_experimental_tab)
        elif key == "extensions":
            self._do_lazy_swap_extensions()

    def _do_lazy_swap(self, key: str, creator) -> None:
        """通用懒加载替换：执行 creator → 移动新框架到位 → 移除占位符"""
        placeholder = self._tab_map.get(key)
        if placeholder is None:
            return
        try:
            idx = self.notebook.index(placeholder)
        except Exception:
            return

        # 先标记已创建，防止 forget 触发的 <<NotebookTabChanged>> 事件重入
        self._tab_created.add(key)

        # 调用真实创建方法（会在 notebook 末尾添加新框架）
        creator()

        # 刚创建的框架在末尾，_tab_map[key] 已被更新
        real_frame = self._tab_map.get(key)
        if real_frame is None or real_frame is placeholder:
            return

        # 移动到占位符原本的位置（占位符被推到 idx+1）
        self.notebook.insert(idx, real_frame)
        self.notebook.forget(placeholder)
        placeholder.destroy()

        log(f"Lazy tab created: {key}")

    def _do_lazy_swap_extensions(self) -> None:
        """懒加载创建 Extensions 标签页"""
        placeholder = getattr(self, "_extensions_placeholder", None)
        if placeholder is None:
            return
        try:
            idx = self.notebook.index(placeholder)
        except Exception:
            return

        # 先标记已创建，防止 forget 触发的 <<NotebookTabChanged>> 事件重入
        self._tab_created.add("extensions")

        try:
            self._extensions_tab = ExtensionsTab(self.notebook, self.current_config)
            real_frame = self._extensions_tab.frame
            self._tab_map["extensions"] = real_frame
            self.notebook.insert(idx, real_frame)
            self.notebook.forget(placeholder)
            placeholder.destroy()
            log("Lazy tab created: extensions")
        except Exception as e:
            log(f"Failed to create extensions tab: {e}")
            self._extensions_tab = None

    def _create_general_tab(self):
        """创建常规设置选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(1, weight=1)
        self.notebook.add(frame, text=t("settings.tab.general"))
        self._tab_map["general"] = frame
        
        # 保存目录
        ttk.Label(frame, text=t("settings.general.save_dir")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # 文件路径输入框
        self.save_dir_var = tk.StringVar(value=self.current_config.get("save_dir", ""))
        self.save_dir_entry = ttk.Entry(frame, textvariable=self.save_dir_var, width=50)
        self.save_dir_entry.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        self.save_dir_entry.bind("<FocusIn>", self._on_focus_in)
        
        # 多按钮放在输入框下方
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(button_frame, text=t("settings.general.browse"), command=self._browse_save_dir).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text=t("settings.general.restore_default"), command=self._restore_default_save_dir).pack(side=tk.LEFT, padx=2)
        
        # 无应用时动作下拉框
        ttk.Label(frame, text=t("settings.general.no_app_action")).grid(row=2, column=0, sticky=tk.W, pady=5)

        # 获取当前动作设置
        current_action = self.current_config.get("no_app_action", "open")
        action_map = get_no_app_action_map()
        self.no_app_action_var = tk.StringVar(value=action_map.get(current_action, t("action.open")))
        self.no_app_action_combo = ttk.Combobox(frame, textvariable=self.no_app_action_var, state="readonly")
        self.no_app_action_combo['values'] = list(action_map.values())
        self.no_app_action_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        # 复选框选项
        self.keep_file_var = tk.BooleanVar(value=self.current_config.get("keep_file", False))
        ttk.Checkbutton(frame, text=t("settings.general.keep_file"), variable=self.keep_file_var).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.notify_var = tk.BooleanVar(value=self.current_config.get("notify", True))
        ttk.Checkbutton(frame, text=t("settings.general.notify"), variable=self.notify_var).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.startup_notify_var = tk.BooleanVar(value=self.current_config.get("startup_notify", True))
        ttk.Checkbutton(frame, text=t("settings.general.startup_notify"), variable=self.startup_notify_var).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)

        if is_windows():
            self.move_cursor_var = tk.BooleanVar(value=self.current_config.get("move_cursor_to_end", True))
            ttk.Checkbutton(frame, text=t("settings.general.move_cursor"), variable=self.move_cursor_var).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=5)
            
        hotkey_row = 6 if not is_windows() else 7
        language_row = hotkey_row + 1

        # 热键（从设置页直接打开热键录制）
        ttk.Label(frame, text=t("settings.general.hotkey")).grid(row=hotkey_row, column=0, sticky=tk.W, pady=(10, 5))
        self.hotkey_display_var = tk.StringVar(
            value=str(self.current_config.get("hotkey") or getattr(app_state, "hotkey_str", ""))
        )
        ttk.Label(frame, textvariable=self.hotkey_display_var).grid(row=hotkey_row, column=1, sticky=tk.W, padx=5, pady=(10, 5))
        set_hotkey_btn = ttk.Button(
            frame,
            text=t("settings.general.set_hotkey"),
            command=self._on_open_hotkey,
            width=12,
        )
        set_hotkey_btn.grid(row=hotkey_row, column=2, sticky=tk.E, padx=5, pady=(10, 5))
        self.set_hotkey_btn = set_hotkey_btn
        if not self._open_hotkey_dialog:
            try:
                set_hotkey_btn.configure(state="disabled")
            except Exception:
                pass

        # 语言设置（移动到常规页最下方）
        ttk.Label(frame, text=t("settings.general.language")).grid(row=language_row, column=0, sticky=tk.W, pady=(15, 5))
        
        # 获取当前语言代码和对应的显示名称
        current_code = self.current_config.get("language", "en-US")
        current_label = get_language_label(current_code)
        
        self.lang_var = tk.StringVar(value=current_label)
        self.lang_combo = ttk.Combobox(frame, textvariable=self.lang_var, state="readonly")
        
        # 构建语言列表
        langs = []
        self.lang_map: Dict[str, str] = {}
        for code, label in iter_languages():
            langs.append(label)
            self.lang_map[label] = code
        
        self.lang_combo['values'] = langs
        self.lang_combo.grid(row=language_row, column=1, sticky=tk.W, padx=5, pady=(15, 5))
        # 绑定 FocusIn，避免自动全选
        self.lang_combo.bind("<FocusIn>", self._on_focus_in)

    def _on_open_hotkey(self) -> None:
        """Open the hotkey dialog from Settings, then refresh the displayed hotkey."""
        cb = self._open_hotkey_dialog
        if not cb:
            return

        try:
            cb()
            # Hotkey dialog is non-blocking; refresh shortly after user action.
            self.root.after(300, self.refresh_hotkey_display)
            self.root.after(1200, self.refresh_hotkey_display)
        except Exception as e:
            log(f"Failed to open hotkey dialog from settings: {e}")

    def _create_conversion_tab(self):
        """创建转换设置选项卡"""
        frame = ttk.Frame(self.notebook)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.notebook.add(frame, text=t("settings.tab.conversion"))
        self._tab_map["conversion"] = frame

        canvas = tk.Canvas(frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        canvas.grid(row=0, column=0, sticky=tk.NSEW)

        content = ttk.Frame(canvas, padding=10)
        content.columnconfigure(1, weight=1)

        window_id = canvas.create_window((0, 0), window=content, anchor="nw")

        def _on_content_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        content.bind("<Configure>", _on_content_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120), "units")

        def _on_enter(_event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _on_leave(_event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _on_enter)
        canvas.bind("<Leave>", _on_leave)
        content.bind("<Enter>", _on_enter)
        content.bind("<Leave>", _on_leave)
        
        # Pandoc 路径
        ttk.Label(content, text=t("settings.conversion.pandoc_path")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.pandoc_path_var = tk.StringVar(value=self.current_config.get("pandoc_path", "pandoc"))
        self.pandoc_entry = ttk.Entry(content, textvariable=self.pandoc_path_var, width=50)
        self.pandoc_entry.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        self.pandoc_entry.bind("<FocusIn>", self._on_focus_in)
        
        # 多按钮放在输入框下方
        pandoc_button_frame = ttk.Frame(content)
        pandoc_button_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(pandoc_button_frame, text=t("settings.general.browse"), command=self._browse_pandoc).pack(side=tk.LEFT, padx=2)
        ttk.Button(pandoc_button_frame, text=t("settings.general.restore_default"), command=self._restore_default_pandoc_path).pack(side=tk.LEFT, padx=2)
        
        # Reference Docx
        ttk.Label(content, text=t("settings.conversion.reference_docx")).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        ref_docx = self.current_config.get("reference_docx")
        self.ref_docx_var = tk.StringVar(value=ref_docx if ref_docx else "")
        self.ref_docx_entry = ttk.Entry(content, textvariable=self.ref_docx_var, width=50)
        self.ref_docx_entry.grid(row=2, column=1, sticky=tk.EW, pady=5, padx=5)
        self.ref_docx_entry.bind("<FocusIn>", self._on_focus_in)
        
        ref_button_frame = ttk.Frame(content)
        ref_button_frame.grid(row=3, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(ref_button_frame, text=t("settings.general.browse"), command=self._browse_ref_docx).pack(side=tk.LEFT, padx=2)
        ttk.Button(ref_button_frame, text=t("settings.general.clear"), command=self._clear_ref_docx).pack(side=tk.LEFT, padx=2)
        
        # Pandoc Filters 配置区
        current_row = self._create_filters_section(content, row=4)
        
        # HTML 格式化
        ttk.Label(content, text=t("settings.conversion.html_formatting"), font=("", 10, "bold")).grid(row=current_row, column=0, columnspan=3, sticky=tk.W, pady=(15, 5))
        
        html_fmt = self.current_config.get("html_formatting", {})
        self.strikethrough_var = tk.BooleanVar(value=html_fmt.get("strikethrough_to_del", True))
        ttk.Checkbutton(content, text=t("settings.conversion.strikethrough"), variable=self.strikethrough_var).grid(row=current_row+1, column=0, columnspan=3, sticky=tk.W, pady=2)

        ttk.Label(content, text=t("settings.conversion.horizontal_rule_style")).grid(row=current_row+2, column=0, sticky=tk.W, pady=(8, 2))
        horizontal_rule_style = self.current_config.get("horizontal_rule_style", "default")
        horizontal_rule_label = self._horizontal_rule_style_key_to_label.get(
            horizontal_rule_style,
            self._horizontal_rule_style_key_to_label["default"],
        )
        self.horizontal_rule_style_var = tk.StringVar(value=horizontal_rule_label)
        self.horizontal_rule_style_combo = ttk.Combobox(
            content,
            textvariable=self.horizontal_rule_style_var,
            values=[label for _, label in self.horizontal_rule_style_options],
            state="readonly",
            width=28,
        )
        self.horizontal_rule_style_combo.grid(row=current_row+2, column=1, sticky=tk.W, pady=(8, 2), padx=5)
        
        ttk.Label(content, text=t("settings.conversion.first_paragraph_heading"), font=("", 10, "bold")).grid(row=current_row+3, column=0, columnspan=3, sticky=tk.W, pady=(12, 5))
        
        # 其他转换选项
        self.md_indent_var = tk.BooleanVar(value=self.current_config.get("md_disable_first_para_indent", True))
        ttk.Checkbutton(content, text=t("settings.conversion.md_indent"), variable=self.md_indent_var).grid(row=current_row+4, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        self.html_indent_var = tk.BooleanVar(value=self.current_config.get("html_disable_first_para_indent", True))
        ttk.Checkbutton(content, text=t("settings.conversion.html_indent"), variable=self.html_indent_var).grid(row=current_row+5, column=0, columnspan=3, sticky=tk.W, pady=2)

        self.markdown_hard_line_breaks_var = tk.BooleanVar(value=self.current_config.get("markdown_hard_line_breaks", False))
        ttk.Checkbutton(
            content,
            text=t("settings.conversion.markdown_hard_line_breaks"),
            variable=self.markdown_hard_line_breaks_var,
        ).grid(row=current_row+6, column=0, columnspan=3, sticky=tk.W, pady=(8, 2))
        ttk.Label(
            content,
            text=t("settings.conversion.markdown_hard_line_breaks_note"),
            foreground="gray",
            font=("", 8),
        ).grid(row=current_row+7, column=0, columnspan=3, sticky=tk.W, padx=(20, 0), pady=(0, 5))

    def _create_advanced_tab(self):
        """创建高级设置选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text=t("settings.tab.advanced"))
        self._tab_map["advanced"] = frame
        frame.columnconfigure(1, weight=1)
        
        # Excel 选项
        self.excel_enable_var = tk.BooleanVar(value=self.current_config.get("enable_excel", True))
        ttk.Checkbutton(frame, text=t("settings.advanced.excel_enable"), variable=self.excel_enable_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.excel_format_var = tk.BooleanVar(value=self.current_config.get("excel_keep_format", True))
        ttk.Checkbutton(frame, text=t("settings.advanced.excel_format"), variable=self.excel_format_var).grid(row=1, column=0, sticky=tk.W, pady=5)

        # 粘贴延迟
        ttk.Label(frame, text=t("settings.advanced.paste_delay")).grid(row=2, column=0, sticky=tk.W, pady=(10, 5))
        self.paste_delay_var = tk.StringVar(value=str(self.current_config.get("paste_delay_s", 0.3)))
        paste_delay_entry = ttk.Entry(frame, textvariable=self.paste_delay_var, width=10)
        paste_delay_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=(10, 5))
        paste_delay_entry.bind("<FocusIn>", self._on_focus_in)
        ttk.Label(
            frame,
            text=t("settings.advanced.paste_delay_note"),
            foreground="gray",
            font=("", 8),
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=(0, 5), pady=(0, 5))

    def _create_experimental_tab(self):
        """创建实验性功能选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text=t("settings.tab.experimental"))
        self._tab_map["experimental"] = frame
        frame.columnconfigure(0, weight=1)
        
        self.keep_formula_var = tk.BooleanVar(value=self.current_config.get("Keep_original_formula", False))
        ttk.Checkbutton(frame, text=t("settings.conversion.keep_formula"), variable=self.keep_formula_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # LaTeX 自动替换开关
        self.enable_latex_replacements_var = tk.BooleanVar(value=self.current_config.get("enable_latex_replacements", True))
        latex_check = ttk.Checkbutton(frame, text=t("settings.conversion.enable_latex_replacements"), variable=self.enable_latex_replacements_var)
        latex_check.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # 说明文本（灰色小字）
        note_text = t("settings.conversion.latex_replacements_note")
        note_label = ttk.Label(frame, text=note_text, foreground="gray", font=("", 8))
        note_label.grid(row=2, column=0, sticky=tk.W, padx=(20, 0), pady=(0, 5))

        # 单行 $ 块级公式修复开关
        self.fix_single_dollar_block_var = tk.BooleanVar(value=self.current_config.get("fix_single_dollar_block", True))
        fix_dollar_check = ttk.Checkbutton(frame, text=t("settings.conversion.fix_single_dollar_block"), variable=self.fix_single_dollar_block_var)
        fix_dollar_check.grid(row=3, column=0, sticky=tk.W, pady=5)

        # 说明文本
        fix_dollar_note = t("settings.conversion.fix_single_dollar_block_note")
        fix_dollar_label = ttk.Label(frame, text=fix_dollar_note, foreground="gray", font=("", 8))
        fix_dollar_label.grid(row=4, column=0, sticky=tk.W, padx=(20, 0), pady=(0, 5))

        self.docx_auto_table_layout_var = tk.BooleanVar(value=self.current_config.get("docx_auto_table_layout", False))
        ttk.Checkbutton(
            frame,
            text=t("settings.conversion.docx_auto_table_layout"),
            variable=self.docx_auto_table_layout_var,
        ).grid(row=5, column=0, sticky=tk.W, pady=5)

        ttk.Label(
            frame,
            text=t("settings.conversion.docx_auto_table_layout_note"),
            foreground="gray",
            font=("", 8),
        ).grid(row=6, column=0, sticky=tk.W, padx=(20, 0), pady=(0, 5))

        # Pandoc request headers（用于下载远程图片等资源时附加请求头）
        self._pandoc_request_headers_example = [
            "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ]
        headers_value = self.current_config.get("pandoc_request_headers")
        if isinstance(headers_value, str):
            initial_headers = [headers_value]
        elif isinstance(headers_value, (list, tuple)):
            initial_headers = [h for h in headers_value if isinstance(h, str) and h.strip()]
        else:
            initial_headers = []

        self.pandoc_request_headers_enable_var = tk.BooleanVar(value=bool(initial_headers))

        ttk.Label(
            frame,
            text=t("settings.conversion.pandoc_request_headers"),
            font=("", 10, "bold"),
        ).grid(row=7, column=0, sticky=tk.W, pady=(10, 5))

        ttk.Checkbutton(
            frame,
            text=t("settings.conversion.pandoc_request_headers_enable"),
            variable=self.pandoc_request_headers_enable_var,
            command=self._toggle_pandoc_request_headers_state,
        ).grid(row=8, column=0, sticky=tk.W, pady=2)

        if not initial_headers:
            initial_headers = self._pandoc_request_headers_example

        headers_frame = ttk.Frame(frame)
        headers_frame.grid(row=9, column=0, sticky=tk.EW, padx=(20, 0), pady=(0, 5))
        headers_frame.columnconfigure(0, weight=1)

        self.pandoc_request_headers_text = tk.Text(headers_frame, height=4, wrap="word")
        self.pandoc_request_headers_text.grid(row=0, column=0, sticky=tk.EW)

        headers_scroll = ttk.Scrollbar(
            headers_frame, orient="vertical", command=self.pandoc_request_headers_text.yview
        )
        headers_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.pandoc_request_headers_text.configure(yscrollcommand=headers_scroll.set)

        self._set_pandoc_request_headers_text(initial_headers)
        self._toggle_pandoc_request_headers_state()

        ttk.Label(
            frame,
            text=t("settings.conversion.pandoc_request_headers_note"),
            foreground="gray",
            font=("", 8),
        ).grid(row=10, column=0, sticky=tk.W, padx=(20, 0), pady=(0, 5))

        ttk.Button(
            frame,
            text=t("settings.conversion.pandoc_request_headers_fill_example"),
            command=self._fill_pandoc_request_headers_example,
        ).grid(row=11, column=0, sticky=tk.W, padx=(20, 0), pady=(0, 5))

    def _set_pandoc_request_headers_text(self, headers: list[str]) -> None:
        try:
            self.pandoc_request_headers_text.configure(state=tk.NORMAL)
            self.pandoc_request_headers_text.delete("1.0", tk.END)
            self.pandoc_request_headers_text.insert(tk.END, "\n".join(headers))
        except Exception as e:
            log(f"Failed to set pandoc_request_headers text: {e}")

    def _fill_pandoc_request_headers_example(self) -> None:
        self._set_pandoc_request_headers_text(self._pandoc_request_headers_example)
        self.pandoc_request_headers_enable_var.set(True)
        self._toggle_pandoc_request_headers_state()

    def _toggle_pandoc_request_headers_state(self) -> None:
        if not hasattr(self, "pandoc_request_headers_text"):
            return
        enabled = bool(self.pandoc_request_headers_enable_var.get())
        self.pandoc_request_headers_text.configure(
            state=(tk.NORMAL if enabled else tk.DISABLED)
        )

    def _should_confirm_keep_formula_enable(self) -> bool:
        """Return True when enabling keep-formula requires an explicit warning."""
        original_value = bool(self.current_config.get("Keep_original_formula", False))
        new_value = bool(self._get_var_value("keep_formula_var", original_value))
        return new_value and original_value != new_value

    def _get_var_value(self, attr_name: str, fallback: Any) -> Any:
        """Return a Tk variable value when its lazy tab exists, otherwise fallback."""
        var = getattr(self, attr_name, None)
        if var is None:
            return fallback
        return var.get()

    def _confirm_keep_formula_enable(self) -> bool:
        """
        Warn before saving when the keep-formula experimental option is enabled.
        Note: This is a temporary implementation. Please be aware that future features of a similar nature may require special confirmation and will be distinctly labeled.
        """
        if not self._should_confirm_keep_formula_enable():
            return True

        return self._ask_topmost_yes_no(
            t("settings.title.warning"),
            t("settings.conversion.keep_formula_enable_warning"),
            icon="warning",
            default="no",
        )

    def _browse_save_dir(self):
        """浏览保存目录"""
        path = filedialog.askdirectory(
            title=t("settings.dialog.select_save_dir"),
            initialdir=os.path.expandvars(self.save_dir_var.get())
        )
        
        if path:
            self.save_dir_var.set(path)

    def _restore_default_save_dir(self):
        """恢复默认保存目录（展开环境变量）"""
        default_save_dir = DEFAULT_CONFIG["save_dir"]
        # 展开环境变量，如 %USERPROFILE% 转换为实际路径
        expanded_save_dir = os.path.expandvars(default_save_dir)
        self.save_dir_var.set(expanded_save_dir)
    
    def _restore_default_pandoc_path(self):
        """恢复默认 Pandoc 路径"""
        default_path = DEFAULT_CONFIG.get("pandoc_path", "pandoc")
        self.pandoc_path_var.set(default_path)
    
    def _clear_ref_docx(self):
        """清除参考文档路径"""
        self.ref_docx_var.set("")
    
    def _clear_initial_selection(self):
        """打开窗口时避免一开始就选中首个输入框"""
        try:
            self.notebook.focus_set()
        except Exception as e:
            log(f"Failed to clear initial selection: {e}")

    def _browse_pandoc(self):
        """浏览 Pandoc 可执行文件"""
        initialdir: Optional[str] = None
        current_value = (self.pandoc_path_var.get() or "").strip()
        if current_value:
            candidate_dir = os.path.dirname(current_value)
            if candidate_dir:
                initialdir = candidate_dir

        kwargs: Dict[str, Any] = {
            "title": t("settings.dialog.select_pandoc"),
            "initialdir": initialdir,
        }
        if not is_macos():
            kwargs["filetypes"] = [
                (t("settings.file_type.executable"), "*.exe" if is_windows() else "*"),
                (t("settings.file_type.all_files"), "*.*" if is_windows() else "*"),
            ]
        path = filedialog.askopenfilename(**kwargs)
        
        if path:
            self.pandoc_path_var.set(path)

    def _browse_ref_docx(self):
        """浏览参考 docx 文件"""
        initialdir: Optional[str] = None
        current_value = (self.ref_docx_var.get() or "").strip()
        if current_value:
            candidate_dir = os.path.dirname(current_value)
            if candidate_dir:
                initialdir = candidate_dir

        kwargs: Dict[str, Any] = {
            "title": t("settings.dialog.select_ref_docx"),
            "initialdir": initialdir,
        }
        if not is_macos():
            kwargs["filetypes"] = [
                (t("settings.file_type.word_doc"), "*.docx"),
                (t("settings.file_type.all_files"), "*.*" if is_windows() else "*"),
            ]
        path = filedialog.askopenfilename(**kwargs)
        
        if path:
            self.ref_docx_var.set(path)

    def _on_save(self):
        """保存配置"""
        try:
            if not self._confirm_keep_formula_enable():
                return

            # 更新配置字典
            new_config = self.current_config.copy()
            
            # 将显示名称映射回代码
            selected_label = self.lang_var.get()
            new_config["language"] = self.lang_map.get(selected_label, "en-US")
            new_config["save_dir"] = self.save_dir_var.get()
            new_config["keep_file"] = self.keep_file_var.get()
            new_config["notify"] = self.notify_var.get()
            new_config["startup_notify"] = self.startup_notify_var.get()
            # Preserve the latest hotkey (may have been changed via HotkeyDialog while Settings is open).
            latest_hotkey = app_state.config.get("hotkey") or getattr(app_state, "hotkey_str", None)
            if latest_hotkey:
                new_config["hotkey"] = str(latest_hotkey)

            # 获取 action_map 用于映射
            action_map = get_no_app_action_map()

            # 映射显示文本回配置值
            reverse_action_map = {v: k for k, v in action_map.items()}
            selected_action_text = self.no_app_action_var.get()
            new_config["no_app_action"] = reverse_action_map.get(selected_action_text, "open")
            if is_windows():
                new_config["move_cursor_to_end"] = self.move_cursor_var.get()
            
            new_config["pandoc_path"] = self._get_var_value(
                "pandoc_path_var",
                self.current_config.get("pandoc_path", "pandoc"),
            )
            ref_docx = self._get_var_value(
                "ref_docx_var",
                self.current_config.get("reference_docx") or "",
            )
            new_config["reference_docx"] = ref_docx if ref_docx else None
            
            if not isinstance(new_config.get("html_formatting"), dict):
                new_config["html_formatting"] = {}
            current_html_formatting = self.current_config.get("html_formatting", {})
            if not isinstance(current_html_formatting, dict):
                current_html_formatting = {}
            new_config["html_formatting"]["strikethrough_to_del"] = self._get_var_value(
                "strikethrough_var",
                current_html_formatting.get("strikethrough_to_del", True),
            )
            horizontal_rule_var = getattr(self, "horizontal_rule_style_var", None)
            if horizontal_rule_var is None:
                new_config["horizontal_rule_style"] = self.current_config.get(
                    "horizontal_rule_style",
                    "default",
                )
            else:
                selected_horizontal_rule_label = horizontal_rule_var.get()
                new_config["horizontal_rule_style"] = self._horizontal_rule_style_label_to_key.get(
                    selected_horizontal_rule_label,
                    "default",
                )
            
            new_config["md_disable_first_para_indent"] = self._get_var_value(
                "md_indent_var",
                self.current_config.get("md_disable_first_para_indent", True),
            )
            new_config["html_disable_first_para_indent"] = self._get_var_value(
                "html_indent_var",
                self.current_config.get("html_disable_first_para_indent", True),
            )
            new_config["markdown_hard_line_breaks"] = self._get_var_value(
                "markdown_hard_line_breaks_var",
                self.current_config.get("markdown_hard_line_breaks", False),
            )
            new_config["Keep_original_formula"] = self._get_var_value(
                "keep_formula_var",
                self.current_config.get("Keep_original_formula", False),
            )
            new_config["enable_latex_replacements"] = self._get_var_value(
                "enable_latex_replacements_var",
                self.current_config.get("enable_latex_replacements", True),
            )
            new_config["fix_single_dollar_block"] = self._get_var_value(
                "fix_single_dollar_block_var",
                self.current_config.get("fix_single_dollar_block", True),
            )
            new_config["docx_auto_table_layout"] = self._get_var_value(
                "docx_auto_table_layout_var",
                self.current_config.get("docx_auto_table_layout", False),
            )

            # pandoc_request_headers（实验性功能）
            if getattr(self, "pandoc_request_headers_enable_var", None) is not None:
                if self.pandoc_request_headers_enable_var.get():
                    raw = self.pandoc_request_headers_text.get("1.0", tk.END).splitlines()
                    headers = [line.strip() for line in raw if isinstance(line, str) and line.strip()]
                    new_config["pandoc_request_headers"] = headers
                else:
                    # 由于 DEFAULT_CONFIG 中默认包含该字段，不建议删除 key；用空列表表示“禁用任何 header”。
                    new_config["pandoc_request_headers"] = []
            
            new_config["enable_excel"] = self._get_var_value(
                "excel_enable_var",
                self.current_config.get("enable_excel", True),
            )
            new_config["excel_keep_format"] = self._get_var_value(
                "excel_format_var",
                self.current_config.get("excel_keep_format", True),
            )
            try:
                paste_delay_value = float(
                    self._get_var_value(
                        "paste_delay_var",
                        self.current_config.get("paste_delay_s", 0.3),
                    )
                )
                if paste_delay_value < 0:
                    paste_delay_value = 0.0
            except (TypeError, ValueError):
                paste_delay_value = DEFAULT_CONFIG.get("paste_delay_s", 0.3)
            new_config["paste_delay_s"] = paste_delay_value
            
            # 保存 Pandoc Filters 列表（dict 格式，兼容旧版 str 格式）
            new_config["pandoc_filters_by_conversion"] = {
                key: list(self.filters_by_conversion.get(key, []))
                for key in self._conversion_filter_keys
            }
            new_config["pandoc_filters"] = list(self.global_filters)
            
            # 保存扩展选项卡配置
            if self._extensions_tab:
                ext_config = new_config.get("extensible_workflows", {})
                ext_config.update(self._extensions_tab.get_config())
                new_config["extensible_workflows"] = ext_config
            
            # 保存到文件
            self.config_loader.save(new_config)
            
            # 更新全局状态
            app_state.config = new_config
            
            # 显示成功消息（置顶）
            self._show_topmost_message(t("settings.title.success"), t("settings.success.saved"), "info")
            
            if self.on_save_callback:
                self.on_save_callback()
            self._call_on_close_callback()
            self._safe_destroy()
            
        except Exception as e:
            log(f"Failed to save settings: {e}")
            # 显示错误消息（置顶）
            self._show_topmost_message(t("settings.title.error"), t("settings.error.save_failed", error=str(e)), "error")

    def _on_cancel(self):
        self._call_on_close_callback()
        self._safe_destroy()

    def _on_close(self):
        self._call_on_close_callback()
        self._safe_destroy()

    def _safe_destroy(self):
        try:
            self.root.destroy()
        except Exception as e:
            log(f"Error destroying settings window: {e}")

    def select_tab(self, tab_key: str) -> None:
        """Select a tab by key."""
        tab = self._tab_map.get(tab_key)
        if tab is None:
            return
        try:
            self.notebook.select(tab)
        except Exception as e:
            log(f"Failed to select settings tab '{tab_key}': {e}")

    def _show_topmost_message(self, title, message, msg_type="info"):
        """显示置顶消息框（使用标准样式）"""
        # 创建临时窗口来控制消息框的置顶属性
        temp_root = tk.Toplevel(self.root)
        temp_root.withdraw()  # 隐藏临时窗口
        temp_root.attributes("-topmost", True)
        
        # 使用标准消息框
        if msg_type == "error":
            messagebox.showerror(title, message, parent=self.root)
        else:
            messagebox.showinfo(title, message, parent=self.root)
        
        # 清理临时窗口
        temp_root.destroy()

    def _ask_topmost_yes_no(
        self,
        title: str,
        message: str,
        *,
        icon: str = "question",
        default: str | None = None,
    ) -> bool:
        """显示置顶确认框并返回用户选择。"""
        temp_root = tk.Toplevel(self.root)
        temp_root.withdraw()
        temp_root.attributes("-topmost", True)

        try:
            kwargs: Dict[str, Any] = {
                "parent": self.root,
                "icon": icon,
            }
            if default:
                kwargs["default"] = default
            return bool(messagebox.askyesno(title, message, **kwargs))
        finally:
            temp_root.destroy()

    def show(self):
        """显示设置对话框（非阻塞模式）"""
        try:
            # 设置对话框在后台显示，让主事件循环继续运行（允许最小化）
            if app_state.root and app_state.root.winfo_exists():
                self.root.transient(app_state.root)
            else:
                self.root.transient(None)
            
            # 监听退出事件，在后台检查
            self._monitor_quit_event()
            
            # 显示窗口
            self.restore_and_focus()
            
        except Exception as e:
            log(f"Error in settings show: {e}")
            self._safe_destroy()
    
    def _monitor_quit_event(self):
        """监听退出事件，自动关闭设置对话框"""
        def check_quit():
            try:
                # 检查退出事件
                if app_state.quit_event and app_state.quit_event.is_set():
                    self._safe_destroy()
                    return
                
                # 继续监听
                self.root.after(200, check_quit)
            except Exception as e:
                log(f"Error in quit monitor: {e}")
                self.root.after(200, check_quit)
                
        if self.is_alive():
            # 开始监听
            self.root.after(200, check_quit)
    
    def _on_focus_in(self, event):
        """所有 Entry/Combobox 获得焦点时，自动清除选中并将光标移到末尾"""
        widget = event.widget

        def clear_sel():
            # 清除自动选中的部分
            try:
                if hasattr(widget, "select_clear"):
                    widget.select_clear()
                elif hasattr(widget, "selection_clear"):
                    widget.selection_clear(0, "end")
            except Exception:
                pass

            # 光标移到末尾
            try:
                if hasattr(widget, "icursor"):
                    widget.icursor("end")
            except Exception:
                pass

        # 立即清一次
        clear_sel()
        # 再用 after_idle 清一次，覆盖类绑定可能重新设置的选中
        try:
            self.root.after_idle(clear_sel)
        except Exception:
            pass

    def _create_filters_section(self, frame: ttk.Frame, row: int) -> int:
        """创建 Pandoc Filters 配置区"""
        # 分组标题
        ttk.Label(frame, text=t("settings.conversion.pandoc_filters"), font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=3, sticky=tk.W, pady=(15, 5)
        )

        # 资源链接
        link_frame = ttk.Frame(frame)
        link_frame.grid(row=row+1, column=0, columnspan=3, sticky=tk.W, padx=(0, 5), pady=(0, 5))
        
        # 创建超链接
        link1_text = t("settings.conversion.filter_resources_link1_text")
        link1_url = "https://github.com/jgm/pandoc/wiki/Pandoc-Filters"
        link1 = self._create_hyperlink_label(link_frame, link1_text, link1_url)
        link1.pack(side=tk.LEFT)
        
        # 分隔符
        separator = ttk.Label(link_frame, text=" | ")
        separator.pack(side=tk.LEFT)
        
        # 第二个链接
        link2_text = t("settings.conversion.filter_resources_link2_text")
        link2_url = "https://pandoc.org/lua-filters.html"
        link2 = self._create_hyperlink_label(link_frame, link2_text, link2_url)
        link2.pack(side=tk.LEFT)

        # 转换类型选择
        conversion_frame = ttk.Frame(frame)
        conversion_frame.grid(row=row+2, column=0, columnspan=3, sticky=tk.W, padx=(0, 5), pady=(0, 5))
        ttk.Label(conversion_frame, text=t("settings.conversion.filter_conversion_label")).pack(
            side=tk.LEFT
        )
        conversion_labels = [label for _, label in self.filter_conversion_options]
        self.filter_conversion_combo = ttk.Combobox(
            conversion_frame,
            textvariable=self.filter_conversion_var,
            values=conversion_labels,
            state="readonly",
            width=22,
        )
        self.filter_conversion_combo.pack(side=tk.LEFT, padx=(6, 0))
        self.filter_conversion_combo.bind(
            "<<ComboboxSelected>>", lambda e: self._on_filter_conversion_changed()
        )

        # 列表框
        self.filters_listbox = tk.Listbox(
            frame, 
            height=5, 
            selectmode=tk.SINGLE,
            activestyle="none"
        )
        self.filters_listbox.grid(row=row+3, column=0, columnspan=2, sticky=tk.NSEW, padx=(0, 5), pady=5)
        
        # 按钮组
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row+3, column=2, sticky=tk.N, pady=5)
        
        # 按钮宽度统一
        btn_width = 8
        
        self.add_filter_btn = ttk.Button(
            button_frame, 
            text=t("settings.conversion.add_filter"), 
            command=self._on_add_filter,
            width=btn_width
        )
        self.add_filter_btn.pack(pady=2)
        
        self.remove_filter_btn = ttk.Button(
            button_frame, 
            text=t("settings.conversion.remove_filter"), 
            command=self._on_remove_filter,
            width=btn_width
        )
        self.remove_filter_btn.pack(pady=2)
        
        self.move_up_btn = ttk.Button(
            button_frame, 
            text=t("settings.conversion.move_up"), 
            command=self._on_move_filter_up,
            width=btn_width
        )
        self.move_up_btn.pack(pady=2)
        
        self.move_down_btn = ttk.Button(
            button_frame, 
            text=t("settings.conversion.move_down"), 
            command=self._on_move_filter_down,
            width=btn_width
        )
        self.move_down_btn.pack(pady=2)
        
        self.toggle_filter_btn = ttk.Button(
            button_frame,
            text=t("settings.conversion.disable_filter"),
            command=self._on_toggle_filter,
            width=btn_width
        )
        self.toggle_filter_btn.pack(pady=2)
        
        # 说明文本
        note_label = ttk.Label(
            frame, 
            text=t("settings.conversion.pandoc_filters_note"),
            foreground="gray",
            font=("", 8)
        )
        note_label.grid(row=row+4, column=0, columnspan=3, sticky=tk.W, padx=(0, 5), pady=(0, 8))
        
        # 绑定列表框选择事件
        self.filters_listbox.bind("<<ListboxSelect>>", lambda e: self._update_filter_buttons_state())
        
        # 绑定双击事件进行编辑
        self.filters_listbox.bind("<Double-Button-1>", lambda e: self._on_edit_filter())
        
        # 初始化列表内容
        self._refresh_filters_listbox()
        
        # 初始化按钮状态
        self._update_filter_buttons_state()
        
        # 返回下一个可用行号
        return row + 5

    @staticmethod
    def _parse_filter_item(item) -> Optional[dict]:
        """Parse a filter item, supporting both legacy (str) and new (dict) formats.

        Returns ``{"path": ..., "enabled": ...}`` or ``None`` if invalid.
        """
        if isinstance(item, str) and item.strip():
            return {"path": item.strip(), "enabled": True}
        if isinstance(item, dict):
            path = item.get("path", "")
            if isinstance(path, str) and path.strip():
                return {"path": path.strip(), "enabled": bool(item.get("enabled", True))}
        return None

    def _create_hyperlink_label(self, parent: tk.Widget, text: str, url: str) -> ttk.Label:
        """创建可点击的超链接标签"""
        link = ttk.Label(
            parent,
            text=text,
            foreground="#0066CC",
            cursor="hand2",
            font=("", 8, "underline")
        )
        
        # 绑定点击事件
        link.bind("<Button-1>", lambda e: webbrowser.open(url))
        
        return link

    def _get_current_filters(self) -> list[dict]:
        label = self.filter_conversion_var.get()
        key = self._filter_conversion_label_to_key.get(label, "global")
        if key == "global":
            return self.global_filters
        return self.filters_by_conversion.setdefault(key, [])

    def _on_filter_conversion_changed(self):
        self._refresh_filters_listbox()
        self._update_filter_buttons_state()

    def _on_add_filter(self):
        """处理添加 Filter 的操作"""
        # 打开文件选择对话框
        kwargs: Dict[str, Any] = {"title": t("settings.dialog.select_filter")}
        if not is_macos():
            kwargs["filetypes"] = [
                (t("settings.file_type.lua_script"), "*.lua"),
                (
                    t("settings.file_type.executable"),
                    "*.exe *.cmd *.bat" if is_windows() else "*.sh *.py *.rb *.pl *.js",
                ),
                (t("settings.file_type.all_files"), "*.*" if is_windows() else "*"),
            ]
        path = filedialog.askopenfilename(**kwargs)
        
        # 用户是否选择文件
        if path:
            # 将路径添加到当前转换类型列表（新格式 dict）
            self._get_current_filters().append({"path": path, "enabled": True})
            
            # 刷新列表框显示
            self._refresh_filters_listbox()
            
            # 更新按钮状态
            self._update_filter_buttons_state()

    def _on_remove_filter(self):
        """删除列表中选中的 Filter"""
        # 获取当前选中的索引
        selection = self.filters_listbox.curselection()
        
        # 检查是否有选中项
        if not selection:
            return
            
        index = selection[0]
        
        # 从当前转换类型列表中删除对应项
        self._get_current_filters().pop(index)
        
        # 刷新列表框显示
        self._refresh_filters_listbox()
        
        # 更新按钮状态
        self._update_filter_buttons_state()

    def _on_move_filter_up(self):
        """将选中的 Filter 向上移动一位"""
        # 获取选中索引
        selection = self.filters_listbox.curselection()
        
        # 检查是否有选中项
        if not selection:
            return
            
        index = selection[0]
        
        # 检查有效性
        current_list = self._get_current_filters()
        if index > 0:
            # 交换位置
            current_list[index], current_list[index-1] = \
                current_list[index-1], current_list[index]
            
            # 刷新显示
            self._refresh_filters_listbox()
            
            # 恢复选中
            self.filters_listbox.selection_set(index-1)
            
            # 更新按钮状态
            self._update_filter_buttons_state()

    def _on_toggle_filter(self):
        """切换选中 Filter 的启用/禁用状态"""
        selection = self.filters_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        current_list = self._get_current_filters()
        current_item = current_list[index]
        current_item["enabled"] = not current_item.get("enabled", True)
        
        # 刷新显示
        self._refresh_filters_listbox()
        
        # 恢复选中
        self.filters_listbox.selection_set(index)
        
        # 更新按钮状态
        self._update_filter_buttons_state()

    def _on_move_filter_down(self):
        """将选中的 Filter 向下移动一位"""
        # 获取选中索引
        selection = self.filters_listbox.curselection()
        
        # 检查是否有选中项
        if not selection:
            return
            
        index = selection[0]
        
        # 检查有效性
        current_list = self._get_current_filters()
        if index < len(current_list) - 1:
            # 交换位置
            current_list[index], current_list[index+1] = \
                current_list[index+1], current_list[index]
            
            # 刷新显示
            self._refresh_filters_listbox()
            
            # 恢复选中
            self.filters_listbox.selection_set(index+1)
            
            # 更新按钮状态
            self._update_filter_buttons_state()

    def _update_filter_buttons_state(self):
        """根据当前选择状态更新按钮的启用/禁用状态"""
        # 获取选中索引
        selection = self.filters_listbox.curselection()
        
        # 判断是否有选中
        has_selection = bool(selection)
        
        current_list = self._get_current_filters()
        if has_selection:
            index = selection[0]
            is_first = (index == 0)
            is_last = (index == len(current_list) - 1)
            enabled = current_list[index].get("enabled", True)
            
            # 设置按钮状态
            self.remove_filter_btn.config(state=tk.NORMAL)
            self.move_up_btn.config(state=tk.DISABLED if is_first else tk.NORMAL)
            self.move_down_btn.config(state=tk.DISABLED if is_last else tk.NORMAL)
            self.toggle_filter_btn.config(
                state=tk.NORMAL,
                text=t("settings.conversion.disable_filter" if enabled else "settings.conversion.enable_filter")
            )
        else:
            # 无选中项时禁用移除/上移/下移按钮
            self.remove_filter_btn.config(state=tk.DISABLED)
            self.move_up_btn.config(state=tk.DISABLED)
            self.move_down_btn.config(state=tk.DISABLED)
            self.toggle_filter_btn.config(state=tk.DISABLED, text=t("settings.conversion.disable_filter"))

    def _refresh_filters_listbox(self):
        """刷新列表框显示内容，同步 filters_list 数据"""
        try:
            # 清空列表框
            self.filters_listbox.delete(0, tk.END)
            
            # 遍历当前转换类型列表，显示完整路径
            for item in self._get_current_filters():
                path = item["path"]
                enabled = item.get("enabled", True)
                if enabled:
                    self.filters_listbox.insert(tk.END, path)
                else:
                    self.filters_listbox.insert(tk.END, f"✗ {path}")
                    idx = self.filters_listbox.size() - 1
                    self.filters_listbox.itemconfig(idx, fg="gray")
        except Exception as e:
            log(f"Failed to refresh filters listbox: {e}")

    def _on_edit_filter(self):
        """处理双击编辑 Filter 路径"""
        # 获取当前选中的索引
        selection = self.filters_listbox.curselection()
        
        # 检查是否有选中项
        if not selection:
            return
            
        index = selection[0]
        current_list = self._get_current_filters()
        current_path = current_list[index]["path"]
        
        # 创建编辑对话框
        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title(t("settings.dialog.edit_filter"))
        edit_dialog.transient(self.root)
        edit_dialog.grab_set()
        
        # 设置对话框大小和位置
        dialog_width = 500
        dialog_height = 120
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog_width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog_height // 2)
        edit_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        edit_dialog.resizable(False, False)
        
        # 创建标签和输入框
        frame = ttk.Frame(edit_dialog, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=t("settings.dialog.filter_path")).pack(anchor=tk.W, pady=(0, 5))
        
        path_var = tk.StringVar(value=current_path)
        path_entry = ttk.Entry(frame, textvariable=path_var, width=60)
        path_entry.pack(fill=tk.X, pady=(0, 10))
        path_entry.focus_set()
        path_entry.select_range(0, tk.END)
        
        # 按钮组
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        def save_edit():
            """保存编辑"""
            new_path = path_var.get().strip()
            if new_path:
                current_list[index]["path"] = new_path
                self._refresh_filters_listbox()
                # 恢复选中
                self.filters_listbox.selection_set(index)
            edit_dialog.destroy()
        
        def cancel_edit():
            """取消编辑"""
            edit_dialog.destroy()
        
        # 浏览按钮
        def browse_filter():
            """浏览文件"""
            current_value = path_var.get().strip()
            kwargs: Dict[str, Any] = {"title": t("settings.dialog.select_filter")}
            if current_value:
                candidate_dir = os.path.dirname(current_value)
                if candidate_dir:
                    kwargs["initialdir"] = candidate_dir
                kwargs["initialfile"] = os.path.basename(current_value)
            if not is_macos():
                kwargs["filetypes"] = [
                    (t("settings.file_type.lua_script"), "*.lua"),
                    (
                        t("settings.file_type.executable"),
                        "*.exe *.cmd *.bat" if is_windows() else "*.sh *.py *.rb *.pl *.js",
                    ),
                    (t("settings.file_type.all_files"), "*.*" if is_windows() else "*"),
                ]
            path = filedialog.askopenfilename(**kwargs)
            if path:
                path_var.set(path)
        
        ttk.Button(button_frame, text=t("settings.general.browse"), command=browse_filter).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text=t("settings.buttons.cancel"), command=cancel_edit).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text=t("settings.buttons.save"), command=save_edit).pack(side=tk.RIGHT, padx=(0, 5))
        
        # 绑定 Enter 键保存，Escape 键取消
        edit_dialog.bind("<Return>", lambda e: save_edit())
        edit_dialog.bind("<Escape>", lambda e: cancel_edit())
