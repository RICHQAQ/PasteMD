"""Workflow router - main entry point."""

import re
from ...core.state import app_state
from ...utils.detector import detect_active_app, get_frontmost_window_title
from ...utils.logging import log
from ...service.notification.manager import NotificationManager
from ...i18n import t

from .word import WordWorkflow, WPSWorkflow
from .excel import ExcelWorkflow, WPSExcelWorkflow
from .fallback import FallbackWorkflow
from .office_omml import OneNoteWorkflow, PowerPointWorkflow
from .extensible import HtmlWorkflow, MdWorkflow, LatexWorkflow, FileWorkflow


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
        
        # 核心工作流（不可配置）
        self.core_workflows = {
            "word": WordWorkflow(),
            "wps": WPSWorkflow(),
            "excel": ExcelWorkflow(),
            "wps_excel": WPSExcelWorkflow(),
            "onenote": OneNoteWorkflow(),
            "powerpoint": PowerPointWorkflow(),
            "": FallbackWorkflow(),  # 空字符串表示无应用/兜底
        }
        
        # 可扩展工作流注册表
        self.extensible_registry = {
            "html": HtmlWorkflow(),
            "md": MdWorkflow(),
            "latex": LatexWorkflow(),
            "file": FileWorkflow(),
        }
        
        self.notification_manager = NotificationManager()
        self._initialized = True
        log("WorkflowRouter initialized")
    
    def _build_dynamic_routes(self, window_title: str = "") -> dict:
        """根据配置动态构建路由表
        
        Args:
            window_title: 当前窗口标题，用于正则匹配
        """
        routes = dict(self.core_workflows)
        
        ext_config = app_state.config.get("extensible_workflows", {})
        for key, workflow in self.extensible_registry.items():
            cfg = ext_config.get(key, {})
            if cfg.get("enabled", False):
                # apps 是 [{"name": ..., "path": ..., "window_patterns": [...]}, ...] 格式
                for app in cfg.get("apps", []):
                    app_name = app.get("name") if isinstance(app, dict) else app
                    app_id = app.get("id", "") if isinstance(app, dict) else ""
                    if isinstance(app_id, str):
                        app_id = app_id.lower()
                    window_patterns = app.get("window_patterns", []) if isinstance(app, dict) else []
                    
                    app_key = app_id

                    if not app_key:
                        continue
                    
                    # 如果有窗口匹配模式，需要检查窗口标题
                    if window_patterns and window_title:
                        if self._match_window_patterns(window_title, window_patterns):
                            routes[app_key] = workflow
                            log(f"Registered extensible route (window matched): {app_key} -> {key}")
                        # 如果有模式但不匹配，不添加此路由
                    elif not window_patterns:
                        # 没有窗口模式，直接匹配应用名称
                        if app_key not in routes:
                            routes[app_key] = workflow
                            log(f"Registered extensible route: {app_key} -> {key}")
        
        return routes
    
    def _match_window_patterns(self, window_title: str, patterns: list) -> bool:
        """检查窗口标题是否匹配任意一个正则表达式模式
        
        Args:
            window_title: 窗口标题
            patterns: 正则表达式模式列表
            
        Returns:
            True 如果匹配任意一个模式
        """
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
        """主入口：检测应用 → 路由到工作流"""
        try:
            # 检测目标应用
            target_app = detect_active_app()
            log(f"Detected target app: {target_app}")
            
            # 获取窗口标题（用于正则匹配）
            window_title = get_frontmost_window_title()
            log(f"Window title: {window_title}")
            
            # 动态构建路由表并路由
            routes = self._build_dynamic_routes(window_title)
            workflow = routes.get(target_app, routes[""])
            workflow.execute()
        
        except Exception as e:
            log(f"Router failed: {e}")
            import traceback
            traceback.print_exc()
            self.notification_manager.notify("PasteMD", t("workflow.generic.failure"), ok=False)


# 全局单例
router = WorkflowRouter()


def execute_paste_workflow():
    """热键入口函数"""
    router.route()
