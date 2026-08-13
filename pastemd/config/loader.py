import json
import os
import copy  # 用于深拷贝
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
        """加载配置文件并处理默认值补全"""
        # 1. 以默认配置为基准
        config = copy.deepcopy(DEFAULT_CONFIG)
        user_config_raw = {}
        config_needs_save = False

        # 2. 读取用户配置 (如果存在)
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    user_config_raw = json.load(f)
            except Exception as e:
                log(f"Load config error: {e}, backing up and utilizing default config.")
                self._backup_corrupted_config()
                config_needs_save = True
                # 这里不抛错，而是降级使用默认配置，防止程序直接崩溃
        else:
            # 如果文件不存在，肯定需要保存
            config_needs_save = True

        # 3. 智能合并：将用户配置覆盖到默认配置上，并检测是否有缺失的 Key
        # update_recursive 返回 True 表示结构发生了变化（即补全了新字段）
        if self._update_recursive(config, user_config_raw):
            config_needs_save = True

        # 4. 如果配置有更新（补全了新字段）或者文件原本不存在，才执行保存
        # 注意：这里保存的是 config (包含了默认值和用户值，但在 path 展开之前)
        if config_needs_save:
            log("Configuration updated/initialized, saving to disk...")
            self.save(config)

        # 5. 运行时处理 (Runtime Processing)
        # 这里做的修改只存在于内存中，不会被写回文件
        # 这样保持了配置文件里的 "$HOME" 或 "%APPDATA%" 原样
        if "save_dir" in config:
            config["save_dir"] = os.path.expandvars(config["save_dir"])

        return config

    def _update_recursive(self, target: dict, source: dict) -> bool:
        """
        递归合并字典，并返回是否有新字段被合并进去了。
        Target 是默认配置（基准），Source 是用户配置。
        """
        has_changes = False

        # 特殊处理：向后兼容迁移 - auto_open_on_no_app -> no_app_action
        if "auto_open_on_no_app" in source and "no_app_action" not in source:
            old_value = source["auto_open_on_no_app"]
            # 根据旧值设置新值
            target["no_app_action"] = "open" if old_value else "none"
            has_changes = True
            log(f"Migrated auto_open_on_no_app={old_value} to no_app_action='{target['no_app_action']}'")

        for key, value in source.items():
            if key in target:
                if isinstance(value, dict) and isinstance(target[key], dict):
                    # 如果双方都是字典，递归深入
                    if self._update_recursive(target[key], value):
                        has_changes = True
                else:
                    # 如果值不一样，更新它，但这不算结构变化（不需要为了值改变而重写文件，除非你想格式化）
                    # 但为了保持用户修改的值，我们需要覆盖
                    target[key] = value
            else:
                # 用户配置里有，但默认配置里没有的废弃字段，通常选择保留或剔除
                # 这里简单处理：保留用户多余的配置，但标记为 changed 以便同步格式
                target[key] = value
                # has_changes = True # 如果你想自动清理废弃字段，这里逻辑要反过来写

        # 反向检查：检查 target (默认配置) 里有，但 source (用户配置) 里没有的 key
        # 这才是"自动补全"的核心
        for key in target.keys():
            if key not in source:
                has_changes = True  # 发现了一个新配置项，需要保存！

        return has_changes

    def _backup_corrupted_config(self) -> None:
        """备份损坏的配置文件为 .bak"""
        backup_path = self.config_path + ".bak"
        try:
            if os.path.exists(backup_path):
                os.remove(backup_path)
            os.rename(self.config_path, backup_path)
            log(f"Corrupted config backed up to: {backup_path}")
        except Exception as e:
            log(f"Failed to backup corrupted config: {e}")

    def check_workflow_conflicts(self, config: ConfigDict) -> dict:
        """检查可扩展工作流中的跨工作流应用冲突
        
        Returns:
            dict: {app_name: [workflow_name1, workflow_name2, ...]} 包含冲突的应用
        """
        ext_config = config.get("extensible_workflows", {})
        app_workflows = {}  # {app_name: [workflow_key1, workflow_key2, ...]}
        
        # 收集所有工作流中的应用
        for workflow_key in ["html", "md", "latex", "file"]:
            workflow_config = ext_config.get(workflow_key, {})
            apps = workflow_config.get("apps", [])
            
            for app in apps:
                # 兼容旧格式（字符串）和新格式（字典）
                if isinstance(app, dict):
                    app_name = app.get("name", "")
                else:
                    app_name = str(app)
                
                if app_name:
                    if app_name not in app_workflows:
                        app_workflows[app_name] = []
                    app_workflows[app_name].append(workflow_key)
        
        # 找出存在冲突的应用（在多个工作流中出现）
        conflicts = {app: workflows 
                    for app, workflows in app_workflows.items() 
                    if len(workflows) > 1}
        
        return conflicts

    def save(self, config: ConfigDict) -> None:
        """保存配置文件"""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log(f"Save config error: {e}")
            raise ConfigError(f"Failed to save config: {e}")
