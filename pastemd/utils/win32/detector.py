"""Windows application detection utilities."""

import win32com.client
from .window import get_foreground_process_name, get_foreground_window_title
from ..logging import log


def detect_active_app() -> str:
    """
    检测当前活跃的插入目标应用
    
    Returns:
        "word", "wps", "excel", "wps_excel" 或空字符串
    """
    process_name = get_foreground_process_name()
    log(f"前台进程名称: {process_name}")
    
    if "winword" in process_name:
        return "word"
    elif "excel" in process_name:
        return "excel"
    elif process_name == "et.exe":  # 独立的 WPS 表格进程(较少见)
        return "wps_excel"
    elif "wps" in process_name:  # WPS Office 统一进程
        # 需要进一步区分是文字还是表格
        return detect_wps_type()
    else:
        return ""


def detect_wps_type() -> str:
    """
    检测 WPS 应用的具体类型 (文字/表格)
    通过获取前台窗口的 COM 对象来精确判断
    
    Returns:
        "wps" (文字), "wps_excel" (表格) 或空字符串
    """
    window_title = get_foreground_window_title()
    log(f"WPS 窗口标题: {window_title}")
    
    # 方法1: 通过 COM 对象判断(最准确)
    # 尝试获取 WPS 表格的 COM 对象,并检查当前激活的窗口是否匹配
    excel_prog_ids = ["ket.Application", "ET.Application"]
    for prog_id in excel_prog_ids:
        try:
            app = win32com.client.GetActiveObject(prog_id)
            # 检查 COM 对象的活动窗口标题是否与前台窗口标题匹配
            try:
                com_caption = app.ActiveDocument.Name
                log(f"WPS 表格 COM 窗口标题: {com_caption}")
                # 比较窗口标题(去除空格和换行符)
                if _normalize_title(com_caption) in _normalize_title(window_title):
                    log("通过 COM 窗口标题匹配,确认为 WPS 表格")
                    return "wps_excel"
                else:
                    log("COM 窗口标题不匹配,WPS 表格不在前台")
            except Exception as e:
                log(f"无法获取 {prog_id} 的 Caption: {e}")
                # 如果无法获取 Caption,尝试其他方法
                pass
        except Exception:
            continue
    
    # 方法2: 尝试获取 WPS 文字的 COM 对象
    word_prog_ids = ["kwps.Application", "KWPS.Application"]
    for prog_id in word_prog_ids:
        try:
            app = win32com.client.GetActiveObject(prog_id)
            # 检查 COM 对象的活动窗口标题是否与前台窗口标题匹配
            try:
                # WPS 文字可能使用 Caption
                com_caption = app.ActiveDocument.Name
                log(f"WPS 文字 COM Caption: {com_caption}")
                # WPS 文字的 Caption 通常是 "Microsoft Word",不太有用
                # 但我们可以验证 ProgID 是否可用
                # 如果能连接到,说明 WPS 文字在运行
                log(f"成功连接到 {prog_id}")
                # 不能直接判断是否在前台,需要依靠窗口标题
                break
            except Exception as e:
                log(f"无法获取 {prog_id} 的 Caption: {e}")
        except Exception:
            continue
    
    # 方法3: 通过窗口标题关键词判断
    log("COM 检测失败,使用窗口标题判断")
    
    # 优先级1: 文件后缀判断（最明确）
    # WPS 表格的文件后缀
    excel_extensions = [
        ".et",
        ".xls",
        ".xlsx",
        ".csv",
    ]
    
    # 先检查表格文件后缀
    for ext in excel_extensions:
        if ext in window_title.lower():
            log(f"通过窗口标题后缀 '{ext}' 识别为 WPS 表格")
            return "wps_excel"
    
    # WPS 文字的文件后缀
    word_extensions = [
        ".doc",
        ".docx",
        ".wps",
    ]
    
    # 检查文字文件后缀
    for ext in word_extensions:
        if ext in window_title.lower():
            log(f"通过窗口标题后缀 '{ext}' 识别为 WPS 文字")
            return "wps"
    
    # 优先级2: 关键词判断
    # WPS 表格的关键词
    excel_keywords = [
        "WPS 表格",
        " - WPS Spreadsheets",
        " ET ",
        "工作簿",
    ]
    
    # 检查是否是 WPS 表格
    for keyword in excel_keywords:
        if keyword in window_title:
            log(f"通过窗口标题关键词 '{keyword}' 识别为 WPS 表格")
            return "wps_excel"
    
    # WPS 文字的关键词
    word_keywords = [
        "文字文稿",
        "WPS 文字",
        " - WPS Writer",
    ]
    
    # 检查是否是 WPS 文字
    for keyword in word_keywords:
        if keyword in window_title:
            log(f"通过窗口标题关键词 '{keyword}' 识别为 WPS 文字")
            return "wps"
    
    # 默认认为是 WPS 文字
    log("无明确标识,默认识别为 WPS 文字")
    return "wps"


def _normalize_title(title: str) -> str:
    """
    标准化窗口标题(去除空格、换行等)
    
    Args:
        title: 原始标题
        
    Returns:
        标准化后的标题
    """
    if not title:
        return ""
    return title.replace(" ", "").replace("\n", "").replace("\r", "").lower()


def _verify_wps_excel_running() -> bool:
    """
    验证 WPS 表格是否正在运行
    
    Returns:
        True 如果 WPS 表格在运行
    """
    excel_prog_ids = ["ket.Application", "ET.Application"]
    for prog_id in excel_prog_ids:
        try:
            app = win32com.client.GetActiveObject(prog_id)
            # 验证确实有活动工作表
            try:
                _ = app.ActiveSheet
                return True
            except Exception:
                continue
        except Exception:
            continue
    return False
