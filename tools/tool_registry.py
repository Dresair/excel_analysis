# -*- coding: utf-8 -*-
"""
工具注册器 - 管理所有可用工具的注册和执行
"""

import json
import logging
from abc import ABC, abstractmethod
from typing import Dict, Any, Tuple, Optional, Type

logger = logging.getLogger(__name__)


class ToolHandler(ABC):
    """工具处理器抽象基类"""
    
    @property
    @abstractmethod
    def name(self) -> str:
        """工具名称"""
        pass
    
    @abstractmethod
    def execute(self, args: Dict[str, Any], context: Dict[str, Any]) -> Tuple[bool, str]:
        """
        执行工具
        
        参数:
            args: 工具参数
            context: 执行上下文（如excel_orchestrator等）
            
        返回:
            (成功标志, 结果字符串)
        """
        pass


class ToolRegistry:
    """工具注册器 - 管理所有可用的工具处理器"""
    
    def __init__(self):
        self._handlers: Dict[str, ToolHandler] = {}
        
    def register(self, handler: ToolHandler):
        """注册工具处理器"""
        self._handlers[handler.name] = handler
        logger.info(f"已注册工具: {handler.name}")
        
    def unregister(self, tool_name: str):
        """注销工具处理器"""
        if tool_name in self._handlers:
            del self._handlers[tool_name]
            logger.info(f"已注销工具: {tool_name}")
            
    def get_handler(self, tool_name: str) -> Optional[ToolHandler]:
        """获取工具处理器"""
        return self._handlers.get(tool_name)
        
    def list_tools(self) -> list[str]:
        """列出所有已注册的工具"""
        return list(self._handlers.keys())
        
    def execute_tool(self, tool_name: str, args: Dict[str, Any], context: Dict[str, Any]) -> Tuple[str, str]:
        """
        执行指定工具
        
        参数:
            tool_name: 工具名称
            args: 工具参数
            context: 执行上下文
            
        返回:
            (工具名称, 执行结果)
        """
        handler = self.get_handler(tool_name)
        if not handler:
            return tool_name, json.dumps({
                "error": f"未知的工具：{tool_name}"
            }, ensure_ascii=False)
            
        try:
            success, result = handler.execute(args, context)
            return tool_name, result
        except Exception as e:
            logger.error(f"工具 {tool_name} 执行失败: {e}")
            return tool_name, json.dumps({
                "error": f"工具执行失败：{str(e)}"
            }, ensure_ascii=False)
