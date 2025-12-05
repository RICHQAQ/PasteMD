"""Hotkey UI entry point."""

from ...domains.hotkey.manager import HotkeyManager
from ...domains.hotkey.debounce import DebounceManager
from ...domains.hotkey.recorder import HotkeyRecorder
from ...config.defaults import DEFAULT_CONFIG
from ...core.state import app_state
from ...utils.logging import log
from ...i18n import t


class HotkeyRunner:
    """热键运行器"""
    
    def __init__(self, controller_callback, notification_manager=None, config_loader=None):
        self.hotkey_manager = HotkeyManager()
        self.debounce_manager = DebounceManager()
        self.controller_callback = controller_callback
        self.notification_manager = notification_manager
        self.config_loader = config_loader
    
    def get_hotkey_manager(self) -> HotkeyManager:
        """获取热键管理器（用于暂停/恢复）"""
        return self.hotkey_manager
    
    def start(self) -> None:
        """启动热键监听"""
        hotkey = app_state.hotkey_str
        
        # 验证热键是否有效
        error = HotkeyRecorder.validate_hotkey_string(hotkey)
        if error:
            log(f"Invalid hotkey '{hotkey}': {error}. Resetting to default.")
            
            # 恢复默认热键
            default_hotkey = DEFAULT_CONFIG["hotkey"]
            app_state.hotkey_str = default_hotkey
            app_state.config["hotkey"] = default_hotkey
            hotkey = default_hotkey
            
            # 保存配置
            if self.config_loader:
                try:
                    self.config_loader.save(app_state.config)
                except Exception as e:
                    log(f"Failed to save corrected config: {e}")
            
            # 通知用户
            if self.notification_manager:
                self.notification_manager.notify(
                    f"PasteMD - {t('hotkey.runner.title_invalid_config')}",
                    t("hotkey.runner.invalid_config", error=error),
                    ok=False
                )
        
        def on_hotkey():
            if app_state.enabled:
                self.debounce_manager.trigger_async(self.controller_callback)
        
        try:
            self.hotkey_manager.bind(hotkey, on_hotkey)
        except Exception as e:

            log(f"Failed to bind hotkey '{hotkey}': {e}")
            
            # 如果绑定失败，尝试使用默认热键
            if hotkey != DEFAULT_CONFIG["hotkey"]:
                try:
                    default_hotkey = DEFAULT_CONFIG["hotkey"]
                    app_state.hotkey_str = default_hotkey
                    app_state.config["hotkey"] = default_hotkey
                    self.hotkey_manager.bind(default_hotkey, on_hotkey)
                    
                    if self.config_loader:
                        self.config_loader.save(app_state.config)
                    
                    if self.notification_manager:
                        self.notification_manager.notify(
                            f"PasteMD - {t('hotkey.runner.title_binding_failed')}",
                            t("hotkey.runner.binding_failed"),
                            ok=False
                        )
                except Exception as fallback_error:
                    log(f"Failed to bind default hotkey: {fallback_error}")
                    if self.notification_manager:
                        self.notification_manager.notify(
                            f"PasteMD - {t('hotkey.runner.title_serious_error')}",
                            t("hotkey.runner.serious_error"),
                            ok=False
                        )
    
    def stop(self) -> None:
        """停止热键监听"""
        self.hotkey_manager.unbind()
    
    def restart(self) -> None:
        """重启热键监听"""
        self.stop()
        self.start()
