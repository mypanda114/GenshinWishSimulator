#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
i18n 多语言支持模块
功能：加载指定语言的 JSON 文件，提供翻译函数 t(key, **kwargs)
用法：
    from i18n import init_i18n, t
    init_i18n("zh-CN")          # 初始化语言
    print(t("welcome.title"))   # 获取翻译
"""

import json
import os
from pathlib import Path

class I18n:
    """轻量级多语言支持类"""
    
    def __init__(self, locale="zh-CN", fallback_locale="zh-CN"):
        self.locale = locale
        self.fallback_locale = fallback_locale
        self.translations = {}
        self.fallback_translations = {}
        self._load_translations()
    
    def _load_translations(self):
        """加载语言文件"""
        locale_dir = Path(__file__).parent / "locales"
        
        # 加载当前语言
        current_file = locale_dir / f"{self.locale}.json"
        if current_file.exists():
            with open(current_file, 'r', encoding='utf-8') as f:
                self.translations = json.load(f)
        else:
            print(f"警告: 语言文件 {current_file} 不存在，使用空字典。")
        
        # 加载后备语言
        if self.fallback_locale != self.locale:
            fallback_file = locale_dir / f"{self.fallback_locale}.json"
            if fallback_file.exists():
                with open(fallback_file, 'r', encoding='utf-8') as f:
                    self.fallback_translations = json.load(f)
    
    def get(self, key, **kwargs):
        """
        获取翻译文本，支持层级访问（如 'menu.start'）
        支持变量替换：{name} 会被 kwargs 中的值替换
        """
        # 分割层级键
        keys = key.split('.')
        
        # 从当前语言获取
        value = self._get_nested(self.translations, keys)
        
        # 如果找不到，从后备语言获取
        if value is None and self.fallback_translations:
            value = self._get_nested(self.fallback_translations, keys)
        
        # 如果还是找不到，返回键名
        if value is None:
            return key
        
        # 替换变量
        if kwargs and isinstance(value, str):
            try:
                return value.format(**kwargs)
            except KeyError:
                # 如果变量缺失，保持原样
                return value
        return value
    
    def _get_nested(self, data, keys):
        """从嵌套字典中获取值"""
        current = data
        for k in keys:
            if isinstance(current, dict) and k in current:
                current = current[k]
            else:
                return None
        return current
    
    def set_locale(self, locale):
        """切换语言"""
        self.locale = locale
        self._load_translations()


# 全局单例
_i18n = None

def init_i18n(locale="zh-CN"):
    """初始化翻译模块（必须在主程序开头调用）"""
    global _i18n
    _i18n = I18n(locale)

def t(key, **kwargs):
    """翻译函数快捷方式，返回翻译后的字符串"""
    if _i18n is None:
        # 如果未初始化，使用默认语言初始化
        init_i18n()
    return _i18n.get(key, **kwargs)

def set_locale(locale):
    """切换当前语言"""
    if _i18n is not None:
        _i18n.set_locale(locale)