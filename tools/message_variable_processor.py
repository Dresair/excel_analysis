# -*- coding: utf-8 -*-
"""
消息变量占位符处理器

职责：
- 在工具执行后保存大体量结果，生成可引用的变量名；
- 为LLM提供轻量级的引用提示payload；
- 在返回给用户前，对消息中的占位符 {"<tool>":"<var_name>"} 进行替换为真实数据；
- 控制内存占用与预览大小。
"""

from __future__ import annotations

import json
import logging
import re
import time
import uuid
from typing import Any, Dict, Optional, Tuple


logger = logging.getLogger(__name__)


class MessageVariableProcessor:
    """管理工具结果的变量绑定与占位符替换。"""

    def __init__(self, max_store_items: int = 50, preview_max_items: int = 50, preview_max_chars: int = 2000):
        # 存储结构：{ (tool_name, var_name): payload_dict }
        self._store: Dict[Tuple[str, str], Any] = {}
        self._order: list[Tuple[str, str]] = []  # 用于LRU裁剪
        self.max_store_items = max_store_items
        self.preview_max_items = preview_max_items
        self.preview_max_chars = preview_max_chars

        # 已注册的工具名，用于占位符解析中的白名单匹配（可选）
        self._known_tools: set[str] = set()

    def register_known_tool(self, tool_name: str):
        """登记一个可识别的工具名，提升占位符匹配的准确度。"""
        if tool_name:
            self._known_tools.add(tool_name)

    def register_binding(self, tool_name: str, value: Any, var_name: Optional[str] = None) -> str:
        """注册一个变量绑定，并返回变量名。"""
        if var_name is None:
            var_name = f"{tool_name}_{int(time.time()*1000)}_{uuid.uuid4().hex[:8]}"
        key = (tool_name, var_name)
        self._store[key] = value
        self._order.append(key)
        self._evict_if_necessary()
        logger.info(f"变量绑定已注册: tool={tool_name}, var={var_name}")
        return var_name

    def get_binding(self, tool_name: str, var_name: str) -> Optional[Any]:
        return self._store.get((tool_name, var_name))

    def build_lightweight_tool_payload(
        self,
        tool_name: str,
        var_name: str,
        original_value: Any,
        include_preview: bool = True,
    ) -> str:
        """
        构造传回给LLM的轻量级payload，避免暴露完整数据。
        """
        payload: Dict[str, Any] = {
            "variable_binding": {
                "tool": tool_name,
                "name": var_name,
                "usage": f"在最终回答中使用 {{\"{tool_name}\":\"{var_name}\"}} 引用完整数据；如需进一步计算，请描述需要的聚合/筛选，不要展开原始数据。"
            }
        }
        if include_preview:
            payload["variable_binding"]["preview"] = self._make_preview(original_value)
        payload["variable_binding"]["size_hint"] = self._size_hint(original_value)
        return json.dumps(payload, ensure_ascii=False)

    def resolve_placeholders_in_text(self, text: Optional[str]) -> str:
        """将文本中的占位符 {"<tool>":"<var_name>"} 替换为绑定的真实数据(JSON字符串)。"""
        if not text:
            return text or ""

        # 正则匹配形如 {"execute_sql":"var_123"} 的最小对象片段
        pattern = re.compile(r"\{\s*\"([a-zA-Z0-9_]+)\"\s*:\s*\"([a-zA-Z0-9_\-:.]+)\"\s*\}")

        def replacer(match: re.Match) -> str:
            tool = match.group(1)
            var = match.group(2)
            # 若登记了已知工具，则仅对已知工具进行替换
            if self._known_tools and tool not in self._known_tools:
                return match.group(0)
            value = self.get_binding(tool, var)
            if value is None:
                return match.group(0)
            try:
                # 对于SQL查询结果，生成HTML表格格式
                if tool == "execute_sql":
                    return self._format_sql_result_as_html_table(value)
                else:
                    return json.dumps(value, ensure_ascii=False)
            except Exception:
                # 如果无法序列化，退回原样
                return match.group(0)

        return pattern.sub(replacer, text)
    
    def _format_sql_result_as_html_table(self, value: Any) -> str:
        """将SQL查询结果格式化为HTML表格"""
        try:
            # 处理多查询结果
            if isinstance(value, dict) and value.get("multiple_queries"):
                html_parts = []
                for i, result_item in enumerate(value.get("results", [])):
                    if "error" in result_item:
                        html_parts.append(f"""
                        <div class="query-result-section">
                            <h4>查询 {result_item.get('query_index', i+1)}</h4>
                            <div class="error-message">❌ {result_item['error']}</div>
                        </div>
                        """)
                    else:
                        sql_preview = result_item.get("sql", "")
                        if len(sql_preview) > 50:
                            sql_preview = sql_preview[:50] + "..."
                        
                        table_html = self._create_html_table(result_item.get("result", []))
                        html_parts.append(f"""
                        <div class="query-result-section">
                            <h4>查询 {result_item.get('query_index', i+1)}: {sql_preview}</h4>
                            <div class="result-stats">行数: {result_item.get('row_count', 0)} | 列数: {result_item.get('column_count', 0)}</div>
                            {table_html}
                        </div>
                        """)
                return f'<div class="sql-results-container">{"".join(html_parts)}</div>'
            
            # 处理单查询结果（数组格式）
            elif isinstance(value, list):
                return self._create_html_table(value)
            
            # 处理错误情况
            elif isinstance(value, dict) and "error" in value:
                return f'<div class="error-message">❌ {value["error"]}</div>'
            
            # 其他情况回退到JSON
            else:
                return json.dumps(value, ensure_ascii=False)
                
        except Exception as e:
            logger.error(f"格式化SQL结果为表格失败: {e}")
            return json.dumps(value, ensure_ascii=False)
    
    def _create_html_table(self, data: list) -> str:
        """创建HTML表格"""
        if not data or not isinstance(data, list):
            return '<div class="no-data">暂无数据</div>'
        
        # 获取表头（假设所有行都有相同的键）
        if len(data) == 0:
            return '<div class="no-data">暂无数据</div>'
        
        headers = list(data[0].keys()) if isinstance(data[0], dict) else []
        if not headers:
            return '<div class="no-data">数据格式错误</div>'
        
        # 限制显示行数，避免页面过长
        display_data = data[:100]  # 最多显示100行
        show_more = len(data) > 100
        
        # 生成唯一的表格ID
        table_id = f"table_{int(time.time()*1000)}"
        
        # 生成表格容器HTML
        html_parts = ['<div class="table-container">']
        
        # 添加表格操作栏
        html_parts.append(f'''
        <div class="table-toolbar">
            <div class="table-info">
                数据行数: {len(display_data)} / {len(data)}
            </div>
            <div class="table-actions">
                <button class="copy-table-btn" onclick="copyTableData('{table_id}')" title="复制表格数据">
                    <i class="fas fa-copy"></i> 复制数据
                </button>
                <button class="copy-csv-btn" onclick="copyTableAsCSV('{table_id}')" title="复制为CSV格式">
                    <i class="fas fa-file-csv"></i> 复制CSV
                </button>
            </div>
        </div>
        ''')
        
        # 生成表格HTML
        html_parts.append(f'<table class="data-table" id="{table_id}">')
        
        # 表头
        html_parts.append('<thead><tr>')
        for header in headers:
            html_parts.append(f'<th>{str(header)}</th>')
        html_parts.append('</tr></thead>')
        
        # 表体
        html_parts.append('<tbody>')
        for row in display_data:
            html_parts.append('<tr>')
            for header in headers:
                cell_value = row.get(header, '')
                display_value = ''  # 初始化display_value

                # 处理特殊值
                if cell_value is None:
                    cell_value = ''
                    display_value = ''
                elif isinstance(cell_value, (int, float)):
                    cell_value = str(cell_value)
                    display_value = cell_value
                else:
                    cell_value = str(cell_value)
                    # 限制单元格内容长度（用于显示，但保存原始值到data属性）
                    display_value = cell_value
                    if len(display_value) > 100:
                        display_value = display_value[:97] + "..."
                escaped_value = str(cell_value).replace('"', '&quot;')
                html_parts.append(f'<td data-value="{escaped_value}">{display_value}</td>')
            html_parts.append('</tr>')
        html_parts.append('</tbody>')
        html_parts.append('</table>')
        
        if show_more:
            html_parts.append(f'<div class="table-more-info">显示前100行，共{len(data)}行数据。复制功能将包含所有显示的数据。</div>')
        
        html_parts.append('</div>')  # 结束table-container
        
        return ''.join(html_parts)

    # ---------------------------- 内部工具函数 ---------------------------- #

    def _evict_if_necessary(self):
        while len(self._order) > self.max_store_items:
            old_key = self._order.pop(0)
            self._store.pop(old_key, None)
            logger.info(f"变量绑定被回收: tool={old_key[0]}, var={old_key[1]}")

    def _make_preview(self, value: Any) -> Any:
        """生成轻量预览：数组/列表仅取前N条；字符串裁剪到最大长度。"""
        try:
            if isinstance(value, str):
                return value[: self.preview_max_chars]
            if isinstance(value, list):
                return value[: self.preview_max_items]
            if isinstance(value, dict):
                # 对典型结构 {"multiple_queries": True, "results": [...]} 做浅裁剪
                preview = {}
                for k, v in list(value.items())[:20]:
                    if isinstance(v, list):
                        preview[k] = v[: min(len(v), self.preview_max_items)]
                    elif isinstance(v, str):
                        preview[k] = v[: self.preview_max_chars]
                    else:
                        preview[k] = v
                return preview
            # 其它类型直接返回
            return value
        except Exception:
            return None

    def _size_hint(self, value: Any) -> Dict[str, Any]:
        try:
            if isinstance(value, list):
                return {"type": "list", "items": len(value)}
            if isinstance(value, dict):
                return {"type": "dict", "keys": len(value)}
            if isinstance(value, str):
                return {"type": "string", "chars": len(value)}
            return {"type": type(value).__name__}
        except Exception:
            return {"type": "unknown"}


