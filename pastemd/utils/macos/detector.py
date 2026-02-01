# -*- coding: utf-8 -*-
"""
macOS application detection utilities.

依赖:
    pip install pyobjc-framework-AppKit pyobjc-framework-Quartz

说明:
- Word / Excel：优先用 bundle id 精确识别
- WPS：mac 端通常是一个统一 App（WPS Office），再用"窗口标题"区分文字/表格
"""

from __future__ import annotations
import subprocess

from AppKit import NSWorkspace, NSRunningApplication
from Quartz import (
    CGWindowListCopyWindowInfo,
    kCGWindowListOptionOnScreenOnly,
    kCGWindowListExcludeDesktopElements,
)

from ..logging import log


def detect_active_app() -> str:
    """
    检测当前活跃的插入目标应用

    Returns:
        "word", "wps", "excel", "wps_excel" 或前台应用标识（用于可扩展工作流匹配）
    """
    # 直接使用 osascript，它在热键场景下更准确
    app = _get_frontmost_app_via_osascript()
    if app:
        name = (app.get("name") or "").lower()
        bundle_id = app.get("bundle_id") or ""
        if not bundle_id or name in ("electron",):
            ns_app = _get_frontmost_app()
            if ns_app and (ns_app.get("bundle_id") or ""):
                app = ns_app
    
    if not app:
        return ""

    name = (app.get("name") or "").lower()
    original_name = app.get("name") or ""
    bundle_id = app.get("bundle_id") or ""
    bundle_id_norm = bundle_id.lower() if bundle_id else ""
    pid = app.get("pid")

    log(f"前台应用: name={original_name}, bundle_id={bundle_id}, pid={pid}")


    if name in ("word", "microsoft word"):
        return "word"
    if name in ("excel", "microsoft excel"):
        return "excel"
    if name in ("onenote", "microsoft onenote"):
        return "onenote"
    if name in ("powerpoint", "microsoft powerpoint"):
        return "powerpoint"
    if "wps" in name or "kingsoft" in name:
        return detect_wps_type()

    # 兜底：返回原始应用名称（用于可扩展工作流匹配）
    if bundle_id_norm:
        return bundle_id_norm
    return original_name


def detect_wps_type() -> str:
    """
    检测 WPS 应用的具体类型 (文字/表格)
    macOS 不像 Windows 那样容易通过 COM 精确区分，因此主要依赖窗口标题。

    Returns:
        "wps" (文字), "wps_excel" (表格) 或空字符串
    """
    window_title = get_frontmost_window_title()
    log(f"WPS 窗口标题: {window_title}")

    # 如果标题拿不到，就只能保守默认文字
    if not window_title:
        log("无法获取窗口标题,默认识别为 WPS 文字")
        return "wps"

    title_l = window_title.lower()

    # 优先级1: 文件后缀判断（最明确）
    excel_extensions = [".et", ".xls", ".xlsx", ".csv"]
    for ext in excel_extensions:
        if ext in title_l:
            log(f"通过窗口标题后缀 '{ext}' 识别为 WPS 表格")
            return "wps_excel"

    word_extensions = [".doc", ".docx", ".wps"]
    for ext in word_extensions:
        if ext in title_l:
            log(f"通过窗口标题后缀 '{ext}' 识别为 WPS 文字")
            return "wps"

    # 优先级2: 关键词判断（不同语言/版本的 WPS 可能不同，可按你用户群继续补充）
    excel_keywords = [
        "wps spreadsheets",
        "表格",
        "工作簿",
        "spreadsheet",
        "sheet",
    ]
    for kw in excel_keywords:
        if kw.lower() in title_l:
            log(f"通过窗口标题关键词 '{kw}' 识别为 WPS 表格")
            return "wps_excel"

    word_keywords = [
        "wps writer",
        "文字",
        "文档",
        "writer",
        "document",
    ]
    for kw in word_keywords:
        if kw.lower() in title_l:
            log(f"通过窗口标题关键词 '{kw}' 识别为 WPS 文字")
            return "wps"

    log("无明确标识,默认识别为 WPS 文字")
    return "wps"


def _get_frontmost_app() -> dict | None:
    """通过 NSWorkspace 获取前台应用信息"""
    try:
        ws = NSWorkspace.sharedWorkspace()
        app = ws.frontmostApplication()
        if not app:
            return None
        return {
            "name": str(app.localizedName() or ""),
            "bundle_id": str(app.bundleIdentifier() or ""),
            "pid": int(app.processIdentifier()),
        }
    except Exception as e:
        log(f"获取前台应用失败(NSWorkspace): {e}")
        return None


def _get_frontmost_app_via_osascript() -> dict | None:
    """
    兜底方案：通过 AppleScript 获取 frontmost app 名称（非常稳定）
    通过 pid 反查 NSRunningApplication，获取更准确的 localizedName/bundle_id
    """
    try:
        pid_cmd = [
            "osascript",
            "-e",
            'tell application "System Events" to get unix id of first application process whose frontmost is true'
        ]
        pid_str = subprocess.check_output(
            pid_cmd,
            text=True,
            encoding="utf-8",
            errors="replace",
        ).strip()
        bundle_id = ""
        bundle_cmd = [
            "osascript",
            "-e",
            'tell application "System Events" to get bundle identifier of first application process whose frontmost is true'
        ]
        try:
            bundle_id = subprocess.check_output(
                bundle_cmd,
                text=True,
                encoding="utf-8",
                errors="replace",
            ).strip()
        except Exception:
            bundle_id = ""
        if pid_str:
            pid = int(pid_str)
            app = NSRunningApplication.runningApplicationWithProcessIdentifier_(pid)
            if app:
                app_name = str(app.localizedName() or "")
                app_bundle_id = str(app.bundleIdentifier() or "") or bundle_id
                return {
                    "name": app_name,
                    "bundle_id": app_bundle_id,
                    "pid": pid,
                }
        if bundle_id:
            apps = NSRunningApplication.runningApplicationsWithBundleIdentifier_(bundle_id) or []
            if apps:
                app = apps[0]
                return {
                    "name": str(app.localizedName() or ""),
                    "bundle_id": str(app.bundleIdentifier() or ""),
                    "pid": int(app.processIdentifier()),
                }

        name_cmd = [
            "osascript",
            "-e",
            'tell application "System Events" to get name of first application process whose frontmost is true'
        ]
        name = subprocess.check_output(
            name_cmd,
            text=True,
            encoding="utf-8",
            errors="replace",
        ).strip()
        if not name:
            return None
        return {"name": name, "bundle_id": bundle_id, "pid": None}
    except Exception as e:
        log(f"获取前台应用失败(osascript): {e}")
        return None


def get_frontmost_window_title() -> str:
    """
    尝试获取前台窗口标题
    先通过 osascript 获取前台应用的 pid，再查询该进程的窗口
    """
    try:
        # 先获取前台应用的 pid
        cmd = [
            "osascript",
            "-e",
            'tell application "System Events" to get unix id of first application process whose frontmost is true'
        ]
        pid_str = subprocess.check_output(
            cmd,
            text=True,
            encoding="utf-8",
            errors="replace",
        ).strip()
        if not pid_str:
            return ""
        
        frontmost_pid = int(pid_str)
        
        # 获取屏幕上所有窗口的基本信息（不包含桌面元素）
        options = kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements
        win_list = CGWindowListCopyWindowInfo(options, 0) or []

        # 只看前台进程的窗口
        candidates = []
        for w in win_list:
            try:
                owner_pid = int(w.get("kCGWindowOwnerPID", -1))
                layer = int(w.get("kCGWindowLayer", 999))
                title = w.get("kCGWindowName", "") or ""

                if layer != 0:
                    continue
                if owner_pid != frontmost_pid:
                    continue

                if title.strip():
                    candidates.append(title)
            except Exception:
                continue

        if candidates:
            return str(candidates[0])

        return ""
    except Exception as e:
        log(f"获取前台窗口标题失败: {e}")
        return ""


if __name__ == "__main__":
    import time
    from pynput import keyboard

    log("macOS 前台应用检测测试 - 按 Cmd+Shift+D 触发检测，按 Ctrl+C 退出")
    
    def on_activate():
        """热键触发时执行检测"""
        # 添加短暂延迟，避免热键按下时焦点切换的干扰
        time.sleep(0.1)
        
        print(f"\n{'='*60}")
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始检测")
        
        # 使用正常流程检测
        result = detect_active_app()
        
        print(f"检测结果: {result}")
        print(f"{'='*60}\n")
    
    # 设置热键 Cmd+Shift+D
    hotkey = keyboard.GlobalHotKeys({
        '<cmd>+<shift>+d': on_activate
    })
    
    try:
        hotkey.start()
        print("✓ 热键监听已启动")
        print("✓ 请切换到要检测的应用窗口")
        print("✓ 按 Cmd+Shift+D 触发检测")
        print("✓ 按 Ctrl+C 退出\n")
        hotkey.join()
    except KeyboardInterrupt:
        log("检测测试已手动终止")
        print("\n退出检测")
