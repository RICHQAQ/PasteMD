"""Lightweight translation utilities for PasteMD."""

from __future__ import annotations

import json
import locale
import logging
import os
from typing import Dict, Iterator, Optional
from ..config.paths import resource_path
from ..utils.logging import log


FALLBACK_LANGUAGE = "en-US"

_loaded_translations: Dict[str, Dict[str, str]] = {}
_language_metadata: Dict[str, Dict[str, str]] = {}
_current_language = FALLBACK_LANGUAGE


def _get_locales_dir() -> str:
    """获取 locales 目录路径"""
    path = resource_path(os.path.join("i18n", "locales"))
    if os.path.isdir(path):
        return path
    fallback = resource_path(os.path.join("pastemd", "i18n", "locales"))
    if os.path.isdir(fallback):
        return fallback
    return path


def _normalize_to_bcp47(language: Optional[str]) -> Optional[str]:
    """
    将 locale 字符串转换为 BCP 47 格式。
    例如: 'zh_CN' -> 'zh-CN', 'en_US' -> 'en-US', 'zh' -> 'zh'
    """
    if not language:
        return None
    normalized = language.replace("_", "-")
    parts = normalized.split("-")
    if len(parts) >= 2:
        return f"{parts[0].lower()}-{parts[1].upper()}"
    return parts[0].lower()


def _load_translations(language: str) -> Dict[str, str]:
    """Load translation dictionary for a language (with caching) from file system."""
    if language in _loaded_translations:
        return _loaded_translations[language]

    data: Dict[str, str] = {}
    locales_dir = _get_locales_dir()
    json_path = os.path.join(locales_dir, f"{language}.json")

    if not os.path.isfile(json_path):
        base_lang = language.split("-")[0].lower()
        for fname in os.listdir(locales_dir):
            if fname.endswith(".json") and fname.lower().startswith(base_lang):
                json_path = os.path.join(locales_dir, fname)
                language = fname[:-5]
                break

    if os.path.isfile(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as fp:
                loaded = json.load(fp)

            if isinstance(loaded, dict):
                meta = loaded.get("_meta", {})
                if meta:
                    _language_metadata[language] = meta
                data = {str(k): str(v) for k, v in loaded.items() if k != "_meta"}
        except FileNotFoundError:
            log(f"Translation file missing for {language} at {json_path}", level=logging.WARNING)
        except Exception as exc:
            log(f"Failed to load translations for {language}: {exc}", level=logging.WARNING)

    _loaded_translations[language] = data or {}
    return _loaded_translations[language]


def _scan_available_languages() -> Dict[str, str]:
    """扫描 locales 目录，返回所有可用语言 {code: display_name}"""
    languages: Dict[str, str] = {}
    locales_dir = _get_locales_dir()
    
    if not os.path.isdir(locales_dir):
        return languages

    for fname in os.listdir(locales_dir):
        if not fname.endswith(".json") or fname.startswith("_"):
            continue
        code = fname[:-5]
        
        if code in _language_metadata:
            languages[code] = _language_metadata[code].get("name", code)
        else:
            json_path = os.path.join(locales_dir, fname)
            try:
                with open(json_path, "r", encoding="utf-8") as fp:
                    loaded = json.load(fp)
                meta = loaded.get("_meta", {})
                name = meta.get("name", code)
                _language_metadata[code] = meta
                languages[code] = name
            except Exception:
                languages[code] = code
    
    return languages


def is_supported_language(language: Optional[str]) -> bool:
    """Return True if the given language code is supported."""
    if not language:
        return False
    locales_dir = _get_locales_dir()
    json_path = os.path.join(locales_dir, f"{language}.json")
    if os.path.isfile(json_path):
        return True
    base_lang = language.split("-")[0].lower()
    for fname in os.listdir(locales_dir):
        if fname.endswith(".json") and fname.lower().startswith(base_lang):
            return True
    return False


def set_language(language: str) -> None:
    """Set the active language if supported, otherwise fall back to the default."""
    global _current_language
    if is_supported_language(language):
        _current_language = language
    else:
        _current_language = FALLBACK_LANGUAGE
    _load_translations(_current_language)


def get_language() -> str:
    """Return the current UI language code."""
    return _current_language


def get_language_label(language: str) -> str:
    """Return the human readable label for a language code."""
    if language in _language_metadata:
        return _language_metadata[language].get("name", language)
    _load_translations(language)
    if language in _language_metadata:
        return _language_metadata[language].get("name", language)
    # 兼容前版本
    if language == 'en':
        return 'English'
    if language == 'zh':
        return '简体中文'
    return language


def iter_languages() -> Iterator[tuple[str, str]]:
    """Yield supported languages (code, label) pairs."""
    languages = _scan_available_languages()
    for code, label in languages.items():
        yield code, label


def detect_system_language() -> Optional[str]:
    """
    Detect the system UI language and return a BCP 47 code if supported.

    Returns:
        Language code (e.g., 'zh-CN', 'en-US') if detection succeeds, otherwise None.
    """
    candidates: list[str | None] = []

    # Windows-specific detection (preferred if available).
    try:
        import ctypes
        from locale import windows_locale

        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        candidates.append(windows_locale.get(lang_id))
    except Exception:
        pass

    # Fallback to locale data provided by Python.
    try:
        lang, _encoding = locale.getdefaultlocale()
        candidates.append(lang)
    except (ValueError, TypeError):
        pass

    try:
        lang, _encoding = locale.getlocale()
        candidates.append(lang)
    except (ValueError, TypeError):
        pass

    for candidate in candidates:
        normalized = _normalize_to_bcp47(candidate)
        if normalized and is_supported_language(normalized):
            return normalized

    return None


def t(key: str, **kwargs) -> str:
    """
    Translate a key into the active language.

    Fallback chain: current language → en-US → raw key (with warning).

    Args:
        key: Translation key.
        **kwargs: Optional format arguments inserted via str.format.
    """
    translations = _load_translations(_current_language)
    text = translations.get(key)

    # Key missing in current language — fall back to en-US
    if text is None and _current_language != FALLBACK_LANGUAGE:
        fallback_text = _load_translations(FALLBACK_LANGUAGE).get(key)
        if fallback_text is not None:
            log(f"Translation key '{key}' missing in {_current_language}, falling back to {FALLBACK_LANGUAGE}", level=logging.WARNING)
            text = fallback_text

    # Key missing in fallback language — use raw key as last resort
    if text is None:
        log(f"Translation key '{key}' not found in {FALLBACK_LANGUAGE}, using raw key", level=logging.WARNING)
        text = key

    if kwargs:
        try:
            text = text.format(**kwargs)
        except (KeyError, IndexError, ValueError):
            pass

    return text


def get_no_app_action_map() -> Dict[str, str]:
    """获取动作值到显示文本的映射（用于 UI）"""
    from ..core.types import NoAppAction
    return {
        NoAppAction.OPEN.value: t("action.open"),
        NoAppAction.SAVE.value: t("action.save"),
        NoAppAction.CLIPBOARD.value: t("action.clipboard"),
        NoAppAction.NONE.value: t("action.none"),
    }
