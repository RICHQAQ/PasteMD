"""Workflow router - main entry point."""

import re
from ...core.state import app_state
from ...utils.detector import detect_active_app, get_frontmost_window_title
from ...utils.logging import log
from ...service.notification.manager import NotificationManager
from ...service.history.recorder import HistoryRecorder
from ...utils.clipboard import capture_clipboard_content
from ...i18n import t

from .word import WordWorkflow, WPSWorkflow
from .excel import ExcelWorkflow, WPSExcelWorkflow
from .fallback import FallbackWorkflow
from .office_omml import OneNoteWorkflow, PowerPointWorkflow
from .extensible import HtmlWorkflow, MdWorkflow, LatexWorkflow, FileWorkflow, YoudaoWorkflow


class WorkflowRouter:
    """工作流路由器（单例）"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if hasattr(self, "_initialized"):
            return

        self.core_workflows = {
            "word": WordWorkflow(),
            "wps": WPSWorkflow(),
            "excel": ExcelWorkflow(),
            "wps_excel": WPSExcelWorkflow(),
            "onenote": OneNoteWorkflow(),
            "powerpoint": PowerPointWorkflow(),
            "youdao": YoudaoWorkflow(),
            "": FallbackWorkflow(),
        }

        self.extensible_registry = {
            "html": HtmlWorkflow(),
            "md": MdWorkflow(),
            "latex": LatexWorkflow(),
            "file": FileWorkflow(),
        }

        self.notification_manager = NotificationManager()
        self._recorder = None
        self._initialized = True
        log("WorkflowRouter initialized")

    def set_history_manager(self, hm) -> None:
        self._recorder = HistoryRecorder(hm)

    def _build_dynamic_routes(self, window_title: str = "") -> dict:
        routes = dict(self.core_workflows)

        ext_config = app_state.config.get("extensible_workflows", {})
        for key, workflow in self.extensible_registry.items():
            cfg = ext_config.get(key, {})
            if cfg.get("enabled", False):
                for app in cfg.get("apps", []):
                    app_id = app.get("id", "") if isinstance(app, dict) else ""
                    if isinstance(app_id, str):
                        app_id = app_id.lower()
                    window_patterns = app.get("window_patterns", []) if isinstance(app, dict) else []

                    app_key = app_id
                    if not app_key:
                        continue

                    if window_patterns and window_title:
                        if self._match_window_patterns(window_title, window_patterns):
                            routes[app_key] = workflow
                            log(f"Registered extensible route (window matched): {app_key} -> {key}")
                    elif not window_patterns:
                        if app_key not in routes:
                            routes[app_key] = workflow
                            log(f"Registered extensible route: {app_key} -> {key}")

        return routes

    def _match_window_patterns(self, window_title: str, patterns: list) -> bool:
        for pattern in patterns:
            if not pattern:
                continue
            try:
                if re.search(pattern, window_title, re.IGNORECASE):
                    log(f"Window title '{window_title}' matched pattern '{pattern}'")
                    return True
            except re.error as e:
                log(f"Invalid regex pattern '{pattern}': {e}")
        return False

    def route(self) -> None:
        target_app = ""
        workflow_key = ""
        error_msg = ""
        window_title = ""
        content_preview = ""
        full_content = ""
        original_html = ""
        workflow = None
        try:
            target_app = detect_active_app()
            log(f"Detected target app: {target_app}")

            window_title = get_frontmost_window_title()
            log(f"Window title: {window_title}")

            # 仅在历史记录启用时预捕获剪贴板（避免不必要的 CF_HTML 解析开销）
            _history_enabled = (
                self._recorder is not None
                and app_state.config.get("history", {}).get("enabled", True)
            )
            if _history_enabled:
                (content_preview, full_content, original_html,
                 from_md_file, md_file_count) = capture_clipboard_content()
            else:
                content_preview = full_content = original_html = ""
                from_md_file = False
                md_file_count = 0

            routes = self._build_dynamic_routes(window_title)
            workflow = routes.get(target_app, routes[""])
            for k, v in {**routes, **dict.fromkeys(self.extensible_registry.keys(), None)}.items():
                if v is workflow:
                    workflow_key = k
                    break
            if not workflow_key:
                if workflow in self.extensible_registry.values():
                    workflow_key = list(self.extensible_registry.keys())[
                        list(self.extensible_registry.values()).index(workflow)
                    ]
                # 正向匹配失败 → 不可读的路径回退为空（fallback）
                elif target_app and target_app not in ("word", "wps", "excel", "wps_excel",
                                                       "onenote", "powerpoint", "youdao"):
                    workflow_key = ""
                else:
                    workflow_key = target_app or ""
            if _history_enabled:
                workflow._pre_captured_text = full_content
                workflow._pre_captured_html = original_html
                workflow._pre_captured_from_md = from_md_file
                workflow._pre_captured_md_count = md_file_count
            workflow.execute()

        except Exception as e:
            log(f"Router failed: {e}")
            import traceback
            traceback.print_exc()
            error_msg = str(e)[:500]
            self.notification_manager.notify("PasteMD", t("workflow.generic.failure"), ok=False)

        # 优先从工作流获取成功标志，否则从异常推导
        try:
            ok = workflow.success if workflow else False
        except Exception:
            ok = not error_msg

        if getattr(self, "_recorder", None):
            self._recorder.record(target_app, workflow_key, ok, error_msg,
                                  window_title, content_preview, full_content,
                                  workflow=workflow, original_html=original_html)

# 全局单例
router = WorkflowRouter()


def execute_paste_workflow():
    """热键入口函数"""
    router.route()
