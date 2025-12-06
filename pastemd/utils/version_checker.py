"""版本更新检查器"""

import json
import urllib.request
import urllib.error
from typing import Optional, Dict, Any

from .logging import log


class VersionChecker:
    """检查 GitHub 最新版本"""
    
    GITHUB_API_URL = "https://api.github.richqaq.cn/repos/RICHQAQ/PasteMD/releases/latest"
    TIMEOUT = 5  # 超时时间（秒）
    
    def __init__(self, current_version: str):
        """
        初始化版本检查器
        
        Args:
            current_version: 当前应用版本号
        """
        self.current_version = current_version
        no_proxy_handler = urllib.request.ProxyHandler({})
        self._opener = urllib.request.build_opener(no_proxy_handler)
    
    def check_update(self) -> Optional[Dict[str, Any]]:
        """
        检查是否有新版本
        
        Returns:
            如果有新版本，返回包含以下字段的字典：
            - has_update: bool, 是否有更新
            - latest_version: str, 最新版本号
            - release_url: str, 发布页面链接
            - release_notes: str, 发布说明
            如果检查失败或无更新，返回 None
        """
        try:
            # 获取最新版本信息
            latest_info = self._fetch_latest_release()
            if not latest_info:
                return None
            
            latest_version = latest_info.get("tag_name", "").lstrip("v")
            if not latest_version:
                log("Failed to parse latest version from GitHub")
                return None
            
            # 比较版本号
            if self._is_newer_version(latest_version, self.current_version):
                return {
                    "has_update": True,
                    "latest_version": latest_version,
                    "current_version": self.current_version,
                    "release_url": latest_info.get("html_url", ""),
                    "release_notes": latest_info.get("body", "暂无发布说明")[:200]  # 限制长度
                }
            else:
                log(f"Already on latest version: {self.current_version}")
                return {
                    "has_update": False,
                    "latest_version": latest_version,
                    "current_version": self.current_version
                }
                
        except Exception as e:
            log(f"Version check failed: {e}")
            return None
    
    def _fetch_latest_release(self) -> Optional[Dict[str, Any]]:
        """
        从 GitHub API 获取最新 release 信息
        
        Returns:
            包含 release 信息的字典，失败返回 None
        """
        try:
            req = urllib.request.Request(
                self.GITHUB_API_URL,
                headers={"User-Agent": f"PasteMD/{self.current_version}",
                         "version": self.current_version}
            )
            
            with self._opener.open(req, timeout=self.TIMEOUT) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode("utf-8"))
                    return data
                else:
                    log(f"GitHub API returned status code: {response.status}")
                    return None
                    
        except urllib.error.URLError as e:
            log(f"Network error while checking version: {e}")
            return None
        except json.JSONDecodeError as e:
            log(f"Failed to parse GitHub API response: {e}")
            return None
        except Exception as e:
            log(f"Unexpected error while fetching release info: {e}")
            return None
    
    def _is_newer_version(self, latest: str, current: str) -> bool:
        """
        比较版本号
        
        Args:
            latest: 最新版本号
            current: 当前版本号
            
        Returns:
            如果最新版本更新，返回 True
        """
        try:
            return self._parse_version(latest) > self._parse_version(current)
        except Exception as e:
            log(f"Failed to compare versions: {e}")
            # 如果解析失败，使用字符串比较
            return latest > current
    
    @staticmethod
    def _parse_version(version_str: str) -> tuple:
        """
        将版本字符串解析为可比较的元组
        
        例如: "0.1.3.3" -> (0, 1, 3, 3)
        
        Args:
            version_str: 版本字符串
            
        Returns:
            版本号元组
        """
        try:
            parts = version_str.split(".")
            return tuple(int(p) for p in parts)
        except (ValueError, AttributeError):
            # 如果转换失败，返回空元组
            return ()
