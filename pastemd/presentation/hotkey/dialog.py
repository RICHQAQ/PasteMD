"""Hotkey configuration dialog."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from pynput import keyboard

from ...utils.logging import log


class HotkeyDialog:
    """热键设置对话框"""
    
    def __init__(self, current_hotkey: str, on_save: Callable[[str], None]):
        """
        初始化热键设置对话框
        
        Args:
            current_hotkey: 当前热键字符串
            on_save: 保存回调函数，接收新的热键字符串
        """
        self.current_hotkey = current_hotkey
        self.on_save = on_save
        self.new_hotkey: Optional[str] = None
        self.recording = False
        self.pressed_keys = set()
        self.released_keys = set()  # 新增：记录已释放的键
        self.all_pressed_keys = set()  # 新增：记录所有按下过的键
        
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
        
        # 键盘监听器
        self.listener: Optional[keyboard.Listener] = None
    
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
        if self.recording:
            return
        
        self.recording = True
        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        self.new_hotkey = None
        
        self.record_btn.config(text="正在录制... (按下组合键)", state=tk.DISABLED)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, "等待按键...")
        self.hotkey_entry.config(state="readonly")
        
        # 启动键盘监听
        self.listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release
        )
        self.listener.start()
    
    def _on_key_press(self, key):
        """按键按下事件"""
        if not self.recording:
            return
        
        try:
            # 记录按下的键
            key_name = self._get_key_name(key)
            if key_name:
                self.pressed_keys.add(key_name)
                self.all_pressed_keys.add(key_name)  # 记录到总集合
                self._update_hotkey_display()
        except Exception as e:
            log(f"Error in key press handler: {e}")
    
    def _on_key_release(self, key):
        """按键释放事件"""
        if not self.recording:
            return False
        
        try:
            key_name = self._get_key_name(key)
            if key_name:
                self.released_keys.add(key_name)
                # 从当前按下的键中移除
                self.pressed_keys.discard(key_name)
                
                # 检查是否所有按下过的键都已释放
                if self.all_pressed_keys and self.all_pressed_keys == self.released_keys:
                    # 所有键都释放了，完成录制
                    self.root.after(100, self._finish_recording)
                    return False  # 停止监听
        except Exception as e:
            log(f"Error in key release handler: {e}")
        
        return True  # 继续监听
    
    def _get_key_name(self, key) -> Optional[str]:
        """获取键名称"""
        try:
            # 修饰键
            if key in [keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                return "ctrl"
            elif key in [keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r]:
                return "shift"
            elif key in [keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r]:
                return "alt"
            elif key == keyboard.Key.cmd:
                return "cmd"
            
            # 尝试获取键的名称（适用于特殊键）
            if hasattr(key, 'name'):
                return key.name.lower()
            
            # 普通键：优先使用 vk (虚拟键码)
            # 这样可以避免组合键时获取到控制字符
            if hasattr(key, 'vk'):
                # 将虚拟键码转换为字符
                vk = key.vk
                # A-Z: 65-90
                if 65 <= vk <= 90:
                    return chr(vk).lower()
                # 0-9: 48-57
                elif 48 <= vk <= 57:
                    return chr(vk)
                # 数字键盘 0-9: 96-105
                elif 96 <= vk <= 105:
                    return f"num{vk - 96}"
            
            # 最后尝试使用 char（仅当不是控制字符时）
            if hasattr(key, 'char') and key.char:
                # 过滤控制字符（ASCII < 32）
                if ord(key.char) >= 32:
                    return key.char.lower()
            
            return None
        except Exception as e:
            log(f"Error getting key name: {e}")
            return None
    
    def _update_hotkey_display(self):
        """更新热键显示"""
        if not self.all_pressed_keys:
            return
        
        # 排序：修饰键在前，普通键在后
        modifiers = []
        keys = []
        
        modifier_order = ['ctrl', 'shift', 'alt', 'cmd']
        for mod in modifier_order:
            if mod in self.all_pressed_keys:
                modifiers.append(mod)
        
        for key in self.all_pressed_keys:
            if key not in modifier_order:
                keys.append(key)
        
        all_keys = modifiers + sorted(keys)
        display_text = " + ".join(k.title() for k in all_keys)
        
        self.root.after(0, lambda: self._set_entry_text(display_text))
    
    def _set_entry_text(self, text: str):
        """设置输入框文本（线程安全）"""
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, text)
        self.hotkey_entry.config(state="readonly")
    
    def _finish_recording(self):
        """完成录制"""
        if not self.all_pressed_keys:
            self._reset_recording()
            return
        
        # 停止监听器
        if self.listener:
            try:
                self.listener.stop()
            except Exception as e:
                log(f"Error stopping listener: {e}")
            self.listener = None
        
        # 验证热键（至少需要一个修饰键和一个普通键）
        modifiers = {'ctrl', 'shift', 'alt', 'cmd'}
        has_modifier = bool(self.all_pressed_keys & modifiers)
        has_normal_key = bool(self.all_pressed_keys - modifiers)
        
        if not has_modifier:
            messagebox.showwarning(
                "无效热键",
                "热键必须包含至少一个修饰键（Ctrl、Shift、Alt）"
            )
            self._reset_recording()
            return
        
        if not has_normal_key:
            messagebox.showwarning(
                "无效热键",
                "热键必须包含至少一个普通键"
            )
            self._reset_recording()
            return
        
        # 生成热键字符串
        self.new_hotkey = self._generate_hotkey_string()
        
        # 更新UI
        self._enable_save_button()
    
    def _generate_hotkey_string(self) -> str:
        """生成热键字符串（pynput格式）"""
        # 排序：修饰键在前，普通键在后
        modifiers = []
        keys = []
        
        modifier_order = ['ctrl', 'shift', 'alt', 'cmd']
        for mod in modifier_order:
            if mod in self.all_pressed_keys:
                modifiers.append(f"<{mod}>")
        
        for key in self.all_pressed_keys:
            if key not in modifier_order:
                # 特殊键需要用尖括号包裹
                if len(key) > 1:
                    keys.append(f"<{key}>")
                else:
                    keys.append(key)
        
        return "+".join(modifiers + sorted(keys))
    
    def _enable_save_button(self):
        """启用保存按钮"""
        self.save_btn.config(state=tk.NORMAL)
        self.record_btn.config(text="重新录制", state=tk.NORMAL)
        self.recording = False
    
    def _reset_recording(self):
        """重置录制状态"""
        self.recording = False
        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        self.record_btn.config(text="点击录制热键", state=tk.NORMAL)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.config(state="readonly")
        
        if self.listener:
            try:
                self.listener.stop()
            except Exception as e:
                log(f"Error stopping listener: {e}")
            self.listener = None
    
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
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            log(f"Failed to save hotkey: {e}")
            messagebox.showerror("错误", f"保存热键失败：{str(e)}")
    
    def _cleanup(self):
        """清理资源"""
        if self.listener:
            try:
                self.listener.stop()
            except Exception as e:
                log(f"Error stopping listener: {e}")
            finally:
                self.listener = None
    
    def _on_close(self):
        """窗口关闭事件"""
        self._cleanup()
        self.root.quit()
        self.root.destroy()
    
    def _on_cancel(self):
        """取消设置"""
        self._cleanup()
        self.root.quit()
        self.root.destroy()
    
    def show(self):
        """显示对话框"""
        try:
            self.root.mainloop()
        finally:
            # 确保清理监听器
            self._cleanup()
