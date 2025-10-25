"""Windows window and process API utilities."""

import os
from time import sleep
import psutil
import win32gui
import win32process
from ..logging import log


def get_foreground_window() -> int:
    """
    获取前台窗口句柄
    
    Returns:
        窗口句柄，失败时返回 0
    """
    try:
        return win32gui.GetForegroundWindow()
    except Exception as e:
        log(f"Failed to get foreground window: {e}")
        return 0


def get_foreground_process_name() -> str:
    """
    获取当前前台进程的名称
    
    Returns:
        进程名称（小写），失败时返回空字符串
    """
    try:
        hwnd = get_foreground_window()
        if not hwnd:
            return ""
        
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return os.path.basename(process.exe()).lower()
        
    except Exception as e:
        log(f"Failed to get foreground process: {e}")
        return ""


def get_foreground_window_title() -> str:
    """
    获取当前前台窗口标题
    
    Returns:
        窗口标题，失败时返回空字符串
    """
    try:
        hwnd = get_foreground_window()
        if not hwnd:
            return ""
        return win32gui.GetWindowText(hwnd)
    except Exception as e:
        log(f"Failed to get window title: {e}")
        return ""


def cleanup_background_wps_processes(ep:int = 0) -> int:
    """
    清理后台的 WPS 进程，保留前台的 WPS 进程
    
    策略：
    1. 查找有 /prometheus 参数的主 WPS 进程（真正的应用）
    2. 保留主进程及其所有子进程树
    3. 只清理孤立的、既没有任务栏窗口也不属于进程树的进程
    
    Returns:
        int: 清理的进程数量
    """
    try:
        cleaned_count = 0
        
        # 只处理这些进程（排除 wpscloudsvr 等服务进程）
        wps_process_names = ['wps.exe', 'kwps.exe', 'et.exe', 'ket.exe']
        
        # 获取所有 WPS 进程（包括服务进程，用于查找父子关系）
        all_wps_processes = []
        target_processes = []
        
        for proc in psutil.process_iter(['pid', 'name', 'cmdline', 'ppid']):
            try:
                proc_name = proc.info['name'].lower()
                # 收集所有 WPS 相关进程（包括 wpscloudsvr）
                if proc_name in wps_process_names or proc_name == 'wpscloudsvr.exe':
                    all_wps_processes.append(proc)
                    # 只对文档/表格进程进行清理判断
                    if proc_name in wps_process_names:
                        target_processes.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        if not target_processes:
            log("No WPS target processes found")
            return 0
        
        log(f"Found {len(target_processes)} WPS target process(es), {len(all_wps_processes)} total WPS-related")
        
        # 1. 找到所有需要保护的主进程
        # - wpscloudsvr.exe（服务进程，始终保护）
        # - 没有 /Automation 参数的文档/表格进程（用户打开的应用）
        # 注意：/Automation 表示 COM 自动化模式，应该被清理
        protected_pids = set()
        
        for proc in all_wps_processes:
            try:
                pid = proc.pid

                # 检查文档/表格进程：如果有 /Automation 参数，说明是 COM 自动化模式，不保护
                cmdline = proc.info.get('cmdline', [])
                cmdline_str = ' '.join(cmdline) if cmdline else ''
                has_automation = '/automation' in cmdline_str.lower()

                # 没有 /Automation 参数并且有窗口的才保护（用户正常打开的应用）
                if not has_automation and _has_main_user_window(pid):
                    protected_pids.add(pid)
                    log(f"Protected: user application {pid} (no /Automation)")
                else:
                    log(f"Skipped: automation process {pid} (has /Automation)")
            except:
                pass

        # 2. 保护所有主进程的子进程树
        def add_children(parent_pid):
            for proc in all_wps_processes:
                try:
                    if proc.info.get('ppid') == parent_pid and proc.pid not in protected_pids:
                        protected_pids.add(proc.pid)
                        log(f"Protected: child {proc.pid} (parent: {parent_pid})")
                        add_children(proc.pid)
                except:
                    pass

        for pid in list(protected_pids):
            add_children(pid)

        # 3. 清理不在保护列表中的目标进程（只清理文档/表格进程）
        for proc in target_processes:
            if proc.info['name'].lower() == 'wpscloudsvr.exe':
                continue  # 永远不清理服务进程
            try:
                pid = proc.pid
                proc_name = proc.info['name'].lower()
                
                if pid in protected_pids:
                    continue
                
                log(f"Cleaning up background process: {pid} ({proc_name})")
                try:
                    proc.terminate()
                    try:
                        proc.wait(timeout=2)
                    except psutil.TimeoutExpired:
                        proc.kill()
                    cleaned_count += 1
                except Exception as e:
                    log(f"Failed to terminate {pid}: {e}")
            except:
                pass
        
        if cleaned_count > 0:
            if ep < 3:  # 限制递归深度，避免无限循环
                sleep(0.15)  # 等待进程退出
                cleanup_background_wps_processes(ep+1)  # 递归清理，直到没有可清理的进程
            log(f"Cleaned up {cleaned_count} background process(es)")
        
        return cleaned_count
    
    except Exception as e:
        log(f"Error during cleanup: {e}")
        return 0


def _has_taskbar_window(pid: int) -> bool:
    """
    检查进程是否有任务栏窗口（即在任务管理器中显示为"应用"）
    
    任务栏窗口的条件（与任务管理器的逻辑一致）：
    1. 窗口可见或最小化
    2. 不是工具窗口（WS_EX_TOOLWINDOW）
    3. 有 WS_EX_APPWINDOW 样式，或者没有所有者窗口且可见
    4. 窗口有标题
    
    Args:
        pid: 进程 ID
        
    Returns:
        True 如果进程有任务栏窗口（是"应用"而非"后台进程"）
    """
    try:
        import win32con
        
        has_taskbar = False
        
        def window_callback(hwnd, extra):
            nonlocal has_taskbar
            try:
                # 获取窗口所属的进程 ID
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid != pid:
                    return True
                
                # 检查窗口是否可见或最小化
                is_visible = win32gui.IsWindowVisible(hwnd)
                is_minimized = win32gui.IsIconic(hwnd)
                
                if not is_visible and not is_minimized:
                    # 完全不可见且未最小化，不是任务栏窗口
                    return True
                
                # 获取窗口扩展样式
                ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                
                # 排除工具窗口（工具窗口不在任务栏显示）
                if ex_style & win32con.WS_EX_TOOLWINDOW:
                    return True
                
                # 检查是否有 WS_EX_APPWINDOW 样式（强制在任务栏显示）
                has_app_window = ex_style & win32con.WS_EX_APPWINDOW
                
                if has_app_window:
                    # 有 APPWINDOW 样式，是任务栏窗口
                    title = win32gui.GetWindowText(hwnd)
                    log(f"Found taskbar window (APPWINDOW) for PID {pid}: '{title}'")
                    has_taskbar = True
                    return False
                
                # 检查窗口是否有所有者
                owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
                
                # 如果窗口没有所有者且可见，通常会在任务栏显示
                if not owner and is_visible:
                    title = win32gui.GetWindowText(hwnd)
                    # 必须有标题才算任务栏窗口
                    if title:
                        log(f"Found taskbar window (no owner, visible) for PID {pid}: '{title}'")
                        has_taskbar = True
                        return False
                
            except Exception as e:
                log(f"Error checking taskbar window for PID {pid}: {e}")
                pass
            return True  # 继续枚举
        
        win32gui.EnumWindows(window_callback, None)
        return has_taskbar
    
    except Exception as e:
        log(f"Error checking taskbar window for PID {pid}: {e}")
        # 如果无法检查，假设进程有任务栏窗口（保守处理，避免误杀）
        return True


def _has_document_window(pid: int) -> bool:
    """
    检查进程是否有文档窗口（主窗口或最小化的窗口）
    
    只检查真正的文档窗口，忽略：
    - 完全不可见的窗口
    - 没有标题的窗口
    - 尺寸异常小的窗口（但最小化除外）
    
    Args:
        pid: 进程 ID
        
    Returns:
        True 如果进程有文档窗口
    """
    try:
        import win32con
        
        has_doc_window = False
        
        def window_callback(hwnd, extra):
            nonlocal has_doc_window
            try:
                # 获取窗口所属的进程 ID
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid != pid:
                    return True
                
                # 检查窗口是否最小化
                is_minimized = win32gui.IsIconic(hwnd)
                
                # 检查窗口是否可见
                is_visible = win32gui.IsWindowVisible(hwnd)
                
                # 最小化的窗口一定是文档窗口
                if is_minimized:
                    title = win32gui.GetWindowText(hwnd)
                    if title:  # 有标题的最小化窗口
                        log(f"Found minimized document window for PID {pid}: '{title}'")
                        has_doc_window = True
                        return False
                
                # 不可见的窗口跳过
                if not is_visible:
                    return True
                
                # 获取窗口标题
                title = win32gui.GetWindowText(hwnd)
                if not title:
                    # 没有标题的可见窗口，不是文档窗口
                    return True
                
                # 获取窗口样式
                ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                
                # 排除工具窗口（工具窗口不算文档窗口）
                if ex_style & win32con.WS_EX_TOOLWINDOW:
                    return True
                
                # 检查窗口尺寸（排除异常小的窗口）
                try:
                    rect = win32gui.GetWindowRect(hwnd)
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]
                    
                    # 文档窗口应该有合理的尺寸（至少 200x200）
                    if width < 200 or height < 200:
                        return True
                except Exception:
                    # 无法获取尺寸，跳过
                    return True
                
                # 找到了文档窗口
                log(f"Found document window for PID {pid}: '{title}' ({width}x{height})")
                has_doc_window = True
                return False  # 停止枚举
                
            except Exception as e:
                log(f"Error checking window for PID {pid}: {e}")
                pass
            return True  # 继续枚举
        
        win32gui.EnumWindows(window_callback, None)
        return has_doc_window
    
    except Exception as e:
        log(f"Error checking document window for PID {pid}: {e}")
        # 如果无法检查，假设进程有窗口（保守处理，避免误杀）
        return True


def _has_main_user_window(pid: int) -> bool:
    """
    检查进程是否有用户窗口（包括主窗口、工具窗口、最小化窗口等）
    
    只要进程有以下任意一种窗口，就认为是前台进程，不应该清理：
    1. 可见的窗口（IsWindowVisible = True）
    2. 最小化的窗口（IsIconic = True）
    3. 有标题的窗口（无论是主窗口还是工具窗口）
    
    只有完全没有任何用户可见窗口的进程，才认为是纯后台进程。
    
    Args:
        pid: 进程 ID
        
    Returns:
        True 如果进程有任何用户窗口
    """
    try:
        import win32con
        
        has_user_window = False
        
        def window_callback(hwnd, extra):
            nonlocal has_user_window
            try:
                # 获取窗口所属的进程 ID
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid != pid:
                    return True
                
                # 检查窗口是否最小化（最小化的窗口应该保留）
                is_minimized = win32gui.IsIconic(hwnd)
                if is_minimized:
                    title = win32gui.GetWindowText(hwnd)
                    log(f"Found minimized window for PID {pid}: '{title}'")
                    has_user_window = True
                    return False  # 停止枚举
                
                # 检查窗口是否可见
                is_visible = win32gui.IsWindowVisible(hwnd)
                if not is_visible:
                    # 不可见的窗口，跳过
                    return True
                
                # 窗口可见，获取窗口标题
                title = win32gui.GetWindowText(hwnd)
                
                # 获取窗口样式
                style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                
                # 检查是否是工具窗口
                is_tool_window = ex_style & win32con.WS_EX_TOOLWINDOW
                
                # 检查是否有标题栏（主窗口的特征）
                has_caption = style & win32con.WS_CAPTION
                
                # 只要有标题或者是工具窗口，就认为是用户窗口
                if title or is_tool_window or has_caption:
                    window_type = "tool window" if is_tool_window else "main window"
                    log(f"Found visible {window_type} for PID {pid}: '{title}'")
                    has_user_window = True
                    return False  # 停止枚举
                
            except Exception as e:
                log(f"Error checking window for PID {pid}: {e}")
                pass
            return True  # 继续枚举
        
        win32gui.EnumWindows(window_callback, None)
        return has_user_window
    
    except Exception as e:
        log(f"Error checking main window for PID {pid}: {e}")
        # 如果无法检查，假设进程有窗口（保守处理，避免误杀）
        return True
