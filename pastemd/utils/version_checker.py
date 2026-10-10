"""版本更新检查器"""

import json
import os
import re
import sys
import urllib.error
import urllib.request
from typing import Optional, Dict, Any, Tuple

from .logging import log
from .system_detect import is_windows
from .update_manifest import DEFAULT_MANIFEST_URL, DEFAULT_UPDATE_CHANNEL, parse_manifest, valid_version


class VersionChecker:
    """优先检查 R2 更新清单，失败时回退到 GitHub。"""
    
    GITHUB_API_URL = "https://api.github.richqaq.cn/repos/RICHQAQ/PasteMD/releases/latest"
    OFFICIAL_API_URL = "https://api.github.com/repos/RICHQAQ/PasteMD/releases/latest"
    TIMEOUT = 5  # 超时时间（秒）
    PRERANK = {"dev": 0, "alpha": 1, "beta": 2, "rc": 3, "final": 4}
    
    def __init__(self, current_version: str, manifest_url: str = DEFAULT_MANIFEST_URL,
                 channel: str = DEFAULT_UPDATE_CHANNEL):
        """
        初始化版本检查器
        
        Args:
            current_version: 当前应用版本号
        """
        self.current_version = current_version
        self.manifest_url = manifest_url
        self.channel = channel
    
    def check_update(self) -> Optional[Dict[str, Any]]:
        """
        检查是否有新版本
        
        Returns:
            如果有新版本，返回包含以下字段的字典：
            - has_update: bool, 是否有更新
            - latest_version: str, 最新版本号
            - release_url: str, 发布页面链接
            - release_notes: str, 发布说明
            如果检查失败，返回 None
        """
        try:
            # 获取最新版本信息
            latest_info = self._fetch_latest_release()
            if not latest_info:
                return None
            
            latest_version = latest_info.get("tag_name", "").lstrip("vV")
            if not valid_version(latest_version):
                log("Failed to parse latest version from GitHub")
                return None
            
            # 比较版本号
            if self._compare_versions(latest_version, self.current_version):
                return {
                    "has_update": True,
                    "latest_version": latest_version,
                    "current_version": self.current_version,
                    "release_url": latest_info.get("html_url", ""),
                    "release_notes": latest_info.get("body") or "",
                    "assets": latest_info.get("assets") or [],
                    "manifest_url": self.manifest_url,
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

        优先尝试直连（不使用任何代理），如果失败，再回退到使用系统代理。
        """
        self._prepare_ssl_environment()
        if self.manifest_url:
            from urllib.parse import urlparse
            if urlparse(self.manifest_url).scheme == "https":
                data = self._fetch_release_url(self.manifest_url)
                if data:
                    try:
                        result = parse_manifest(data, self.manifest_url, self.channel)
                        log(f"[update] Using R2 feed: {self.manifest_url}")
                        return result
                    except (ValueError, TypeError) as exc:
                        log(f"[update] Invalid R2 feed, falling back: {exc}")
            else:
                log("[update] Ignoring non-HTTPS update feed")
        if self.channel != "stable":
            log("[update] Preview feed unavailable; not switching to a production release")
            return None
        for url in (self.GITHUB_API_URL, self.OFFICIAL_API_URL):
            data = self._fetch_release_url(url)
            if data and data.get("tag_name") and not data.get("draft") and not data.get("prerelease"):
                return data
        return None

    def _fetch_release_url(self, url: str) -> Optional[Dict[str, Any]]:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": f"PasteMD/{self.current_version}",
                "version": self.current_version,
                "Accept": "application/vnd.github+json",
            },
        )

        # 依次尝试：先不使用代理，再使用系统代理
        for use_proxy in (False, True):
            try:
                if not use_proxy:
                    # 先不使用代理
                    log(f"[update] Checking {url} (direct)")
                    opener = urllib.request.build_opener(
                        urllib.request.ProxyHandler({})
                    )
                    response = opener.open(req, timeout=self.TIMEOUT)
                else:
                    # 回退：使用系统代理 / 环境变量配置的代理
                    log(f"[update] Retrying {url} with system proxy")
                    response = urllib.request.urlopen(req, timeout=self.TIMEOUT)

                with response:
                    if response.status == 200:
                        try:
                            if response.geturl().split(":", 1)[0] != "https":
                                raise ValueError("Non-HTTPS feed redirect")
                            payload = response.read(1024 * 1024 + 1)
                            if len(payload) > 1024 * 1024:
                                raise ValueError("Update metadata exceeds 1 MiB")
                            data = json.loads(payload.decode("utf-8"))
                            return data if isinstance(data, dict) else None
                        except json.JSONDecodeError as e:
                            log(f"Failed to parse GitHub API response: {e}")
                            return None

            except urllib.error.URLError as e:
                # 第一轮直连失败会进入这里，for 循环会继续第二轮使用代理
                # 第二轮再失败就直接退出循环
                mode = "no-proxy" if not use_proxy else "proxy"
                log(f"Network error while checking version ({mode}): {e}")
            except Exception as e:
                mode = "no-proxy" if not use_proxy else "proxy"
                log(f"Unexpected error while fetching release info ({mode}): {e}")

        # 两种方式都失败
        return None

    def _prepare_ssl_environment(self) -> None:
        """Windows: 尽量固定 OpenSSL DLL 来源，并记录诊断信息。"""
        if not is_windows():
            return

        base_dir = self._get_app_base_dir()
        log(f"SSL diag: base_dir={base_dir}")

        if base_dir:
            libssl_path = os.path.join(base_dir, "libssl-3-x64.dll")
            libcrypto_path = os.path.join(base_dir, "libcrypto-3-x64.dll")
            log(f"SSL diag: libssl exists={os.path.isfile(libssl_path)}")
            log(f"SSL diag: libcrypto exists={os.path.isfile(libcrypto_path)}")

            if hasattr(os, "add_dll_directory"):
                try:
                    os.add_dll_directory(base_dir)
                    log("SSL diag: added DLL search directory")
                except Exception as exc:
                    log(f"SSL diag: add_dll_directory failed: {exc}")

            self._preload_openssl_dlls(libcrypto_path, libssl_path)

        self._log_ssl_runtime_info()

    def _get_app_base_dir(self) -> str:
        if hasattr(sys, "_MEIPASS"):
            return str(sys._MEIPASS)
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        current_file = os.path.abspath(__file__)
        return os.path.dirname(os.path.dirname(os.path.dirname(current_file)))

    def _preload_openssl_dlls(self, libcrypto_path: str, libssl_path: str) -> None:
        try:
            import ctypes
        except Exception as exc:
            log(f"SSL diag: ctypes import failed: {exc}")
            return

        for path in (libcrypto_path, libssl_path):
            if not path or not os.path.isfile(path):
                continue
            try:
                ctypes.WinDLL(path)
                log(f"SSL diag: preloaded {os.path.basename(path)}")
            except Exception as exc:
                log(f"SSL diag: preload failed for {os.path.basename(path)}: {exc}")

    def _get_loaded_dll_path(self, dll_name: str) -> Optional[str]:
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
            kernel32.GetModuleHandleW.restype = wintypes.HMODULE
            kernel32.GetModuleFileNameW.argtypes = [
                wintypes.HMODULE,
                wintypes.LPWSTR,
                wintypes.DWORD,
            ]
            kernel32.GetModuleFileNameW.restype = wintypes.DWORD

            hmod = kernel32.GetModuleHandleW(dll_name)
            if not hmod:
                return None
            buf = ctypes.create_unicode_buffer(260)
            if kernel32.GetModuleFileNameW(hmod, buf, 260) == 0:
                return None
            return buf.value or None
        except Exception:
            return None

    def _log_ssl_runtime_info(self) -> None:
        try:
            import ssl

            log(f"SSL diag: OPENSSL_VERSION={ssl.OPENSSL_VERSION}")
            try:
                log(f"SSL diag: _ssl module={ssl._ssl.__file__}")
            except Exception:
                pass

            for dll_name in ("libssl-3-x64.dll", "libcrypto-3-x64.dll"):
                dll_path = self._get_loaded_dll_path(dll_name)
                if dll_path:
                    log(f"SSL diag: loaded {dll_name} from {dll_path}")
                else:
                    log(f"SSL diag: {dll_name} not loaded yet")
        except Exception as exc:
            log(f"SSL diag: failed to log ssl info: {exc}")
    
    def _compare_versions(self, latest: str, current: str) -> bool:
        ln, lrank, lpre = self._parse_version_parts(latest)
        cn, crank, cpre = self._parse_version_parts(current)

        n = max(len(ln), len(cn))
        ln = ln + (0,) * (n - len(ln))
        cn = cn + (0,) * (n - len(cn))

        return (ln, lrank, lpre) > (cn, crank, cpre)

    def _parse_version_parts(self, version_str: str) -> Tuple[Tuple[int, ...], int, int]:
        s = (version_str or "").strip().lstrip("vV").lower()
        if not s:
            return (), self.PRERANK["final"], 0

        s = s.split("+", 1)[0]

        pre_tag = None
        pre_num = 0
        m = re.search(r"(?:^|[.\-_])?(dev|alpha|beta|rc)\s*\.?\s*(\d*)\b", s)
        if m:
            pre_tag = m.group(1)
            pre_num = int(m.group(2) or 0)

        parts = re.split(r"[.\-_]", s)
        nums = []
        for part in parts:
            if not part:
                continue
            m2 = re.match(r"(\d+)", part)
            if not m2:
                break
            nums.append(int(m2.group(1)))

        if not nums:
            return (), self.PRERANK["final"], 0

        rank = self.PRERANK["final"] if pre_tag is None else self.PRERANK[pre_tag]
        return tuple(nums), rank, pre_num
