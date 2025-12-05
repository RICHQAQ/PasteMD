"""Hotkey recording functionality."""

from typing import Optional, Callable, Set
from pynput import keyboard

from ...utils.logging import log
from ...i18n import t


SYSTEM_HOTKEY_DESCRIPTIONS = {
    ('ctrl', 'c'): "hotkey.recorder.system.ctrl_c",
    ('ctrl', 'v'): "hotkey.recorder.system.ctrl_v",
    ('ctrl', 'x'): "hotkey.recorder.system.ctrl_x",
    ('ctrl', 'z'): "hotkey.recorder.system.ctrl_z",
    ('ctrl', 'y'): "hotkey.recorder.system.ctrl_y",
    ('ctrl', 'a'): "hotkey.recorder.system.ctrl_a",
    ('ctrl', 's'): "hotkey.recorder.system.ctrl_s",
    ('ctrl', 'f'): "hotkey.recorder.system.ctrl_f",
    ('ctrl', 'p'): "hotkey.recorder.system.ctrl_p",
    ('ctrl', 'n'): "hotkey.recorder.system.ctrl_n",
    ('ctrl', 'w'): "hotkey.recorder.system.ctrl_w",
    ('ctrl', 't'): "hotkey.recorder.system.ctrl_t",
    ('alt', 'f4'): "hotkey.recorder.system.alt_f4",
}


class HotkeyRecorder:
    """热键录制器 - 负责录制用户按下的组合键"""
    
    def __init__(self):
        self.recording = False
        self.pressed_keys: Set[str] = set()
        self.released_keys: Set[str] = set()
        self.all_pressed_keys: Set[str] = set()
        self.recording_listener: Optional[keyboard.Listener] = None  # 录制专用监听器
        self.on_update_callback: Optional[Callable[[str], None]] = None
        self.on_finish_callback: Optional[Callable[[Optional[str], Optional[str]], None]] = None
    
    @staticmethod
    def validate_hotkey_string(hotkey: str) -> Optional[str]:
        """
        静态方法：验证热键字符串是否有效
        
        Args:
            hotkey: 热键字符串，如 "<ctrl>+b" 或 "<ctrl>+<shift>+v"
        
        Returns:
            错误信息，如果验证通过则返回None
        """
        try:
            # 解析热键字符串
            parts = hotkey.lower().replace("<", "").replace(">", "").split("+")
            keys = set(parts)
            
            modifiers = {'ctrl', 'shift', 'alt', 'cmd'}
            has_modifier = bool(keys & modifiers)
            has_normal_key = bool(keys - modifiers)
            
            # 检查是否有修饰键
            if not has_modifier:
                return t("hotkey.recorder.error.no_modifier")
            
            # 检查是否有普通键
            if not has_normal_key:
                return t("hotkey.recorder.error.no_normal_key")
            
            # 检查是否仅使用 Shift 作为修饰键
            if keys & modifiers == {'shift'}:
                return t("hotkey.recorder.error.shift_only_short")
            
            # 检查是否使用了系统级快捷键
            for system_combo in SYSTEM_HOTKEY_DESCRIPTIONS:
                if keys == set(system_combo):
                    return t("hotkey.recorder.error.system_reserved_short", combo=hotkey.upper())
            
            return None
        except Exception as e:
            return t("hotkey.recorder.error.invalid_format", error=str(e))
    
    def start_recording(
        self,
        on_update: Optional[Callable[[str], None]] = None,
        on_finish: Optional[Callable[[Optional[str], Optional[str]], None]] = None
    ) -> None:
        """
        开始录制热键
        
        Args:
            on_update: 更新回调函数，参数为格式化的热键字符串（用于实时显示）
            on_finish: 完成回调函数，参数为(热键字符串, 错误信息)，成功时错误信息为None
        """
        if self.recording:
            return
        
        self.recording = True
        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        self.on_update_callback = on_update
        self.on_finish_callback = on_finish
        
        # 启动键盘监听（仅用于录制）
        self.recording_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release
        )
        self.recording_listener.start()
        log("Hotkey recording started")
    
    def stop_recording(self) -> None:
        """停止录制"""
        self.recording = False
        if self.recording_listener:
            try:
                self.recording_listener.stop()
            except Exception as e:
                log(f"Error stopping recorder listener: {e}")
            finally:
                self.recording_listener = None
        
        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        log("Hotkey recording stopped")
    
    def _on_key_press(self, key):
        """按键按下事件"""
        if not self.recording:
            return
        
        try:
            key_name = self._get_key_name(key)
            if key_name:
                self.pressed_keys.add(key_name)
                self.all_pressed_keys.add(key_name)
                self._notify_update()
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
                self.pressed_keys.discard(key_name)
                
                # 检查是否所有按下过的键都已释放
                if self.all_pressed_keys and self.all_pressed_keys == self.released_keys:
                    # 所有键都释放了，完成录制
                    self._finish_recording()
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
    
    def _notify_update(self) -> None:
        """通知更新（用于实时显示）"""
        if self.on_update_callback and self.all_pressed_keys:
            display_text = self._format_keys_for_display()
            try:
                self.on_update_callback(display_text)
            except Exception as e:
                log(f"Error in update callback: {e}")
    
    def _format_keys_for_display(self) -> str:
        """格式化按键用于显示"""
        if not self.all_pressed_keys:
            return ""
        
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
        return " + ".join(k.title() for k in all_keys)
    
    def _finish_recording(self) -> None:
        """完成录制并验证"""
        if not self.all_pressed_keys:
            self.stop_recording()
            if self.on_finish_callback:
                self.on_finish_callback(None, t("hotkey.recorder.error.no_key_detected"))
            return
        
        # 验证热键
        error = self._validate_hotkey()
        if error:
            hotkey_str = None
        else:
            hotkey_str = self._generate_hotkey_string()
        
        # 停止录制
        self.stop_recording()
        
        # 通知完成
        if self.on_finish_callback:
            try:
                self.on_finish_callback(hotkey_str, error)
            except Exception as e:
                log(f"Error in finish callback: {e}")
    
    def _validate_hotkey(self) -> Optional[str]:
        """
        验证热键是否有效
        
        Returns:
            错误信息，如果验证通过则返回None
        """
        modifiers = {'ctrl', 'shift', 'alt', 'cmd'}
        has_modifier = bool(self.all_pressed_keys & modifiers)
        has_normal_key = bool(self.all_pressed_keys - modifiers)
        
        # 检查是否有修饰键
        if not has_modifier:
            return t("hotkey.recorder.error.no_modifier")
        
        # 检查是否有普通键
        if not has_normal_key:
            return t("hotkey.recorder.error.no_normal_key")
        
        # 检查是否仅使用 Shift 作为修饰键
        # Shift 单独作为修饰键会与系统输入冲突
        if self.all_pressed_keys & modifiers == {'shift'}:
            return t("hotkey.recorder.error.shift_only_long")
        
        # 检查是否使用了系统级快捷键
        # pynput.GlobalHotKeys 无法拦截这些快捷键的系统默认行为
        for key_combo, desc_key in SYSTEM_HOTKEY_DESCRIPTIONS.items():
            if set(key_combo) == self.all_pressed_keys:
                return t("hotkey.recorder.error.system_reserved_long", combo=t(desc_key))
        
        return None
    
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
