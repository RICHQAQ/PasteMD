"""Configuration loading and saving."""

import json
import os

from .defaults import DEFAULT_CONFIG
from .paths import get_config_path
from ..core.types import ConfigDict
from ..core.errors import ConfigError
from ..utils.logging import log


class ConfigLoader:
    """配置加载器"""
    
    def __init__(self):
        self.config_path = get_config_path()
    
    def load(self) -> ConfigDict:
        """加载配置文件"""
        config = DEFAULT_CONFIG.copy()
        
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    user_config = json.load(f)
                
                # 合并用户配置
                for key, value in user_config.items():
                    config[key] = value
                    
            except Exception as e:
                log(f"Load config error: {e}")
                raise ConfigError(f"Failed to load config: {e}")
        
        # 展开环境变量
        config["save_dir"] = os.path.expandvars(config["save_dir"])
        
        return config
    
    def save(self, config: ConfigDict) -> None:
        """保存配置文件"""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log(f"Save config error: {e}")
            raise ConfigError(f"Failed to save config: {e}")
