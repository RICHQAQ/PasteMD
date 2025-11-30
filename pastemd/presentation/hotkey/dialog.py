"""Hotkey configuration dialog."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable

from ...utils.logging import log
from ...domains.hotkey.recorder import HotkeyRecorder


class HotkeyDialog:
    """热键设置对话框"""
    
    def __init__(self, current_hotkey: str, on_save: Callable[[str], None], on_close: Optional[Callable[[], None]] = None):
        """
        初始化热键设置对话框
        
        Args:
            current_hotkey: 当前热键字符串
            on_save: 保存回调函数，接收新的热键字符串
            on_close: 关闭回调函数（用于恢复热键监听）
        """
        self.current_hotkey = current_hotkey
        self.on_save = on_save
        self.on_close_callback = on_close
        self.new_hotkey: Optional[str] = None
        
        self.root = tk.Tk()
        self.root.title("设置热键")
        self.root.geometry("450x300")
        self.root.resizable(False, False)
        
        # 设置关闭窗口时的处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 窗口居中
        self._center_window()
        
        # 创建UI组件
        self._create_widgets()
        
        # 热键录制器
        self.recorder = HotkeyRecorder()
    
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self):
        """创建UI组件"""
        # 标题
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame,
            text="设置全局热键",
            font=("Microsoft YaHei UI", 12, "bold")
        )
        title_label.pack()
        
        # 当前热键显示
        current_frame = ttk.Frame(self.root, padding="10")
        current_frame.pack(fill=tk.X)
        
        ttk.Label(current_frame, text="当前热键：").pack(side=tk.LEFT)
        ttk.Label(
            current_frame,
            text=self._format_hotkey(self.current_hotkey),
            font=("Consolas", 10, "bold")
        ).pack(side=tk.LEFT, padx=5)
        
        # 新热键输入
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        ttk.Label(input_frame, text="新热键：").pack(side=tk.LEFT)
        
        self.hotkey_entry = ttk.Entry(
            input_frame,
            font=("Consolas", 10),
            state="readonly",
            width=25
        )
        self.hotkey_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 录制按钮
        record_frame = ttk.Frame(self.root, padding="10")
        record_frame.pack(fill=tk.X)
        
        self.record_btn = ttk.Button(
            record_frame,
            text="点击录制热键",
            command=self._start_recording
        )
        self.record_btn.pack()
        
        # 按钮栏
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        cancel_btn = ttk.Button(
            button_frame,
            text="取消",
            command=self._on_cancel,
            width=12
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        self.save_btn = ttk.Button(
            button_frame,
            text="保存并应用",
            command=self._on_save,
            state=tk.DISABLED,
            width=12
        )
        self.save_btn.pack(side=tk.RIGHT, padx=5)
    
    def _format_hotkey(self, hotkey: str) -> str:
        """格式化热键显示"""
        return hotkey.replace("<", "").replace(">", "").replace("+", " + ").title()
    
    def _start_recording(self):
        """开始录制热键"""
        self.new_hotkey = None
        
        self.record_btn.config(text="正在录制... (按下组合键)", state=tk.DISABLED)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, "等待按键...")
        self.hotkey_entry.config(state="readonly")
        
        # 使用录制器开始录制
        self.recorder.start_recording(
            on_update=self._on_recording_update,
            on_finish=self._on_recording_finish
        )
    
    def _on_recording_update(self, display_text: str):
        """录制更新回调（实时显示）"""
        self.root.after(0, lambda: self._set_entry_text(display_text))
    
    def _on_recording_finish(self, hotkey_str: Optional[str], error: Optional[str]):
        """录制完成回调"""
        if error:
            # 有错误，显示警告
            self.root.after(0, lambda: messagebox.showwarning("无效热键", error))
            self.root.after(0, self._reset_recording)
        elif hotkey_str:
            # 成功录制
            self.new_hotkey = hotkey_str
            self.root.after(0, self._enable_save_button)
        else:
            # 未知情况
            self.root.after(0, self._reset_recording)
    
    def _set_entry_text(self, text: str):
        """设置输入框文本（线程安全）"""
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, text)
        self.hotkey_entry.config(state="readonly")
    
    def _enable_save_button(self):
        """启用保存按钮"""
        self.save_btn.config(state=tk.NORMAL)
        self.record_btn.config(text="重新录制", state=tk.NORMAL)
    
    def _reset_recording(self):
        """重置录制状态"""
        self.record_btn.config(text="点击录制热键", state=tk.NORMAL)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.config(state="readonly")
    
    def _on_save(self):
        """保存热键"""
        if not self.new_hotkey:
            messagebox.showwarning("提示", "请先录制新热键")
            return
        
        # 确认对话框
        confirm_msg = (
            f"确认将热键更改为：{self._format_hotkey(self.new_hotkey)}\n\n"
            f"原热键：{self._format_hotkey(self.current_hotkey)}\n"
            f"新热键：{self._format_hotkey(self.new_hotkey)}\n\n"
            "更改将立即生效，是否继续？"
        )
        
        if not messagebox.askyesno("确认更改", confirm_msg):
            return
        
        try:
            # 调用保存回调
            self.on_save(self.new_hotkey)
            messagebox.showinfo("成功", f"热键已更新为：{self._format_hotkey(self.new_hotkey)}\n\n请使用新热键测试功能。")
            self._cleanup()
            self._safe_destroy()
        except Exception as e:
            log(f"Failed to save hotkey: {e}")
            messagebox.showerror("错误", f"保存热键失败：{str(e)}")
    
    def _cleanup(self):
        """清理资源"""
        try:
            self.recorder.stop_recording()
        except Exception as e:
            log(f"Error stopping recorder: {e}")
    
    def _safe_destroy(self):
        """安全销毁窗口（避免线程问题）"""
        try:
            self.root.withdraw()  # 先隐藏窗口
            self.root.update_idletasks()  # 处理待处理的事件
        except Exception as e:
            log(f"Error withdrawing window: {e}")
        
        try:
            self.root.quit()
        except Exception as e:
            log(f"Error quitting mainloop: {e}")
        
        try:
            self.root.destroy()
        except Exception as e:
            # 忽略 Tcl_AsyncDelete 错误，这是 tkinter 在非主线程中的已知问题
            if "Tcl_AsyncDelete" not in str(e):
                log(f"Error destroying window: {e}")
    
    def _on_close(self):
        """窗口关闭事件"""
        self._cleanup()
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                log(f"Error in close callback: {e}")
        self._safe_destroy()
    
    def _on_cancel(self):
        """取消设置"""
        self._cleanup()
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                log(f"Error in close callback: {e}")
        self._safe_destroy()
    
    def show(self):
        """显示对话框"""
        try:
            self.root.mainloop()
        except Exception as e:
            log(f"Error in dialog mainloop: {e}")
        finally:
            # 确保清理监听器
            try:
                self._cleanup()
            except Exception as e:
                log(f"Error in cleanup: {e}")
