# -*- coding: utf-8 -*-
"""
配置管理器 - 管理应用配置包括API设置
"""

import json
import os
from typing import Dict, Any, Optional
import logging

from path_manager import path_manager

logger = logging.getLogger(__name__)


class ConfigManager:
    """配置管理器类"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化配置管理器
        
        参数:
            config_file: 配置文件名
        """
        self.config_file = path_manager.get_resource_path(config_file)
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """从文件加载配置"""
        default_config = {
            "openai": {
                "api_key": "sk-at-0b0e7e5415e7cb37a89700617f5e60d2b20893de093fbf632a750799477be62c",
                "base_url": "https://chat.appexnetworks.com/open-api/v1/relay/openai/v1",
                "model": "gpt-4.1"
            },
            "ui": {
                "theme": "dark",
                "auto_load_excel": True,
                "last_excel_dir": str(path_manager.get_resource_path("excel"))
            }
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # 合并默认配置和加载的配置
                    self._merge_config(default_config, loaded_config)
                    return default_config
            except Exception as e:
                logger.warning(f"加载配置文件失败，使用默认配置: {e}")
        
        return default_config
    
    def _merge_config(self, default: Dict[str, Any], loaded: Dict[str, Any]):
        """递归合并配置"""
        for key, value in loaded.items():
            if key in default:
                if isinstance(default[key], dict) and isinstance(value, dict):
                    self._merge_config(default[key], value)
                else:
                    default[key] = value
            else:
                default[key] = value
    
    def save_config(self) -> bool:
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
            return False
    
    def get_openai_config(self) -> Dict[str, str]:
        """获取OpenAI配置"""
        return self.config.get("openai", {})
    
    def set_openai_config(self, api_key: str, base_url: str, model: str):
        """设置OpenAI配置"""
        self.config["openai"]["api_key"] = api_key
        self.config["openai"]["base_url"] = base_url
        self.config["openai"]["model"] = model
    
    def get_ui_config(self) -> Dict[str, Any]:
        """获取UI配置"""
        return self.config.get("ui", {})
    
    def set_ui_config(self, **kwargs):
        """设置UI配置"""
        for key, value in kwargs.items():
            self.config["ui"][key] = value
    
    def apply_to_environment(self):
        """将OpenAI配置应用到环境变量"""
        openai_config = self.get_openai_config()
        if openai_config.get("api_key"):
            os.environ["OPENAI_API_KEY"] = openai_config["api_key"]
        if openai_config.get("base_url"):
            os.environ["OPENAI_BASE_URL"] = openai_config["base_url"]
        if openai_config.get("model"):
            os.environ["OPENAI_MODEL"] = openai_config["model"]
    
    def is_openai_configured(self) -> bool:
        """检查OpenAI配置是否完整"""
        openai_config = self.get_openai_config()
        return bool(openai_config.get("api_key") and 
                   openai_config.get("base_url") and 
                   openai_config.get("model"))

