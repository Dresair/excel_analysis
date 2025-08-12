# llm_client.py
import openai
import json
import logging
from typing import Dict, Any, Optional, List, Tuple
import threading
import os
import dotenv
from datetime import datetime
from tools.message_variable_processor import MessageVariableProcessor
dotenv.load_dotenv()

# 设置日志
logger = logging.getLogger(__name__)

class OpenAIConnector:
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls, api_key: Optional[str] = None, **kwargs):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialize(api_key, **kwargs)
        return cls._instance
    
    def _initialize(self, api_key: Optional[str], **kwargs):
        """初始化 OpenAI 客户端"""
        # 尝试获取API密钥，优先级：参数 > 环境变量
        api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError(
                "OpenAI API key is required. "
                "Please provide it either as an argument or set the OPENAI_API_KEY environment variable."
            )
            
        self.client = openai.OpenAI(api_key=api_key,base_url=os.getenv("OPENAI_BASE_URL"), **kwargs)
        # 从环境变量读取默认模型，如果没有设置则使用 gpt-4o-mini
        self.default_model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        
        # 工具注册和执行相关
        self.tool_registry = None
        self.tool_context = {}

        # 消息变量占位符处理器
        self.message_var_processor = MessageVariableProcessor()
        
        # LLM交互日志
        self.llm_logs: List[Dict[str, Any]] = []
    
    @classmethod
    def get_instance(cls) -> 'OpenAIConnector':
        """获取单例实例"""
        if cls._instance is None:
            # 尝试使用环境变量中的API密钥初始化
            if os.getenv("OPENAI_API_KEY"):
                return cls()
            raise ValueError(
                "OpenAIConnector has not been initialized yet and no OPENAI_API_KEY environment variable found. "
                "Please initialize with an API key first."
            )
        return cls._instance
    
    def chat_completion(
        self,
        messages: List[Dict[str, str]],
        model: Optional[str] = None,
        tools: Optional[List[Dict[str, Any]]] = None,
        tool_choice: Optional[str] = None,
        auto_execute_tools: bool = True,
        max_tool_rounds: int = 10,
        **kwargs
    ) -> Dict[str, Any]:
        """
        与 AI 进行对话，支持自动工具调用
        
        参数:
            messages: 对话消息列表
            model: 使用的模型
            tools: 可用工具列表（如果未提供且有tool_registry，将自动获取）
            tool_choice: 工具选择策略
            auto_execute_tools: 是否自动执行工具调用
            max_tool_rounds: 最大工具调用轮数
            **kwargs: 其他OpenAI参数
            
        返回:
            包含最终响应的字典
        """
        model = model or self.default_model
        
        # 如果没有提供tools，尝试从tool_registry获取
        if tools is None and self.tool_registry:
            tools = self._get_tools_from_registry()
        tools = tools or []
        
        try:
            # 如果不支持工具调用或没有工具，直接调用
            if not auto_execute_tools or not tools:
                response = self.client.chat.completions.create(
                    model=model,
                    messages=messages,
                    tools=tools,
                    tool_choice=tool_choice,
                    **kwargs
                )
                return response
            
            # 启用自动工具调用
            return self._handle_chat_with_tools(
                messages, model, tools, tool_choice, max_tool_rounds, **kwargs
            )
            
        except Exception as e:
            raise RuntimeError(f"Failed to get chat completion: {str(e)}")
    
    def _handle_chat_with_tools(
        self, 
        messages: List[Dict[str, str]], 
        model: str, 
        tools: List[Dict[str, Any]], 
        tool_choice: Optional[str],
        max_tool_rounds: int,
        **kwargs
    ) -> Dict[str, Any]:
        """
        处理带工具调用的对话，自动执行多轮工具调用
        
        参数:
            messages: 对话消息列表
            model: 使用的模型
            tools: 可用工具列表
            tool_choice: 工具选择策略
            max_tool_rounds: 最大工具调用轮数
            **kwargs: 其他OpenAI参数
            
        返回:
            最终的LLM响应
        """
        current_messages = messages.copy()
        current_round = 0
        
        while current_round < max_tool_rounds:
            # 记录LLM请求
            self._log_llm_interaction(f"chat_round_{current_round}", current_messages, None)
            
            # 调用LLM
            response = self.client.chat.completions.create(
                model=model,
                messages=current_messages,
                tools=tools,
                tool_choice=tool_choice,
                **kwargs
            )
            
            message = response.choices[0].message
            
            # 记录LLM响应
            self._log_llm_interaction(f"chat_round_{current_round}", None, message.content)
            
            # 如果没有工具调用，返回最终结果
            if not message.tool_calls:
                return response
            
            # 添加包含tool_calls的assistant消息
            current_messages.append({
                "role": "assistant",
                "content": None,  # tool_calls消息的content必须是None
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {
                            "name": tc.function.name,
                            "arguments": tc.function.arguments
                        }
                    } for tc in message.tool_calls
                ]
            })
            
            # 执行每个工具调用并添加对应的tool消息
            for tool_call in message.tool_calls:
                try:
                    tool_name, result = self._execute_tool_call(tool_call)
                    logger.debug(f"工具 {tool_name} 执行完成，结果长度: {len(result)}")
                    
                    # 添加tool响应消息
                    current_messages.append({
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "content": result
                    })
                except Exception as e:
                    logger.error(f"工具调用失败: {e}")
                    # 即使失败也要添加tool消息，否则会导致格式错误
                    current_messages.append({
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "content": json.dumps({"error": str(e)}, ensure_ascii=False)
                    })
            
            current_round += 1
        
        # 达到最大轮数，强制要求LLM给出最终答案
        current_messages.append({
            "role": "user",
            "content": "请基于以上工具调用的结果，直接给出最终答案，不要再调用任何工具。"
        })
        
        try:
            final_response = self.client.chat.completions.create(
                model=model,
                messages=current_messages,
                tools=None,  # 禁用工具调用
                tool_choice=None,
                **kwargs
            )
            
            logger.warning(f"达到最大工具调用轮数 {max_tool_rounds}，已获取最终答案")
            return final_response
            
        except Exception as e:
            logger.error(f"获取最终LLM响应失败: {e}")
            # 创建一个模拟的响应对象
            class MockResponse:
                def __init__(self, content):
                    self.choices = [MockChoice(content)]
            
            class MockChoice:
                def __init__(self, content):
                    self.message = MockMessage(content)
            
            class MockMessage:
                def __init__(self, content):
                    self.content = content
                    self.tool_calls = None
            
            return MockResponse(f"达到最大工具调用轮数，且获取最终响应时出错: {str(e)}")
    
    def _execute_tool_call(self, tool_call) -> Tuple[str, str]:
        """
        执行工具调用
        
        参数:
            tool_call: OpenAI返回的tool_call对象
            
        返回:
            (工具名称, 执行结果)
        """
        if not self.tool_registry:
            return tool_call.function.name, json.dumps({
                "error": "工具注册器未初始化"
            }, ensure_ascii=False)
        
        function_name = tool_call.function.name
        function_args = json.loads(tool_call.function.arguments)
        
        # 使用工具注册器执行工具
        tool_name, result_str = self.tool_registry.execute_tool(function_name, function_args, self.tool_context)

        # 将结果注册为变量绑定，并返回轻量payload供LLM消费
        try:
            # 尝试解析为JSON对象，失败则按原字符串存储
            parsed: Any
            try:
                parsed = json.loads(result_str)
            except Exception:
                parsed = result_str

            var_name = self.message_var_processor.register_binding(tool_name, parsed)
            lightweight = self.message_var_processor.build_lightweight_tool_payload(
                tool_name=tool_name,
                var_name=var_name,
                original_value=parsed,
                include_preview=True,
            )
            return tool_name, lightweight
        except Exception:
            # 退化：返回原始结果
            return tool_name, result_str
    
    def _get_tools_from_registry(self) -> List[Dict[str, Any]]:
        """
        从tool_registry获取工具定义列表
        
        返回:
            工具定义列表，如果没有tool_registry则返回空列表
        """
        if not self.tool_registry:
            return []
        
        # 尝试加载工具定义文件
        try:
            tools_file_path = 'tools/pptx_json.json'
            if os.path.exists(tools_file_path):
                with open(tools_file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                logger.warning(f"工具定义文件不存在: {tools_file_path}")
                return []
        except Exception as e:
            logger.error(f"加载工具定义失败: {e}")
            return []
    
    def _log_llm_interaction(self, context: str, request: Any, response: Any):
        """
        记录LLM交互日志
        
        参数:
            context: 上下文标识
            request: 请求内容
            response: 响应内容
        """
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "context": context,
            "request": request,
            "response": response
        }
        self.llm_logs.append(log_entry)
        
        # 同时写入文件
        try:
            os.makedirs('log', exist_ok=True)
            with open('log/llm_interactions.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps(log_entry, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.warning(f"写入LLM交互日志失败: {e}")
    
    def set_default_model(self, model: str):
        """设置默认模型"""
        self.default_model = model
    
    def set_tool_registry(self, tool_registry):
        """设置工具注册器"""
        self.tool_registry = tool_registry
        # 同步登记已知工具名到变量处理器，提升占位符识别精准度
        try:
            for name in self.tool_registry.list_tools():
                self.message_var_processor.register_known_tool(name)
        except Exception:
            pass
    
    def set_tool_context(self, context: Dict[str, Any]):
        """设置工具执行上下文"""
        self.tool_context = context
    
    def update_tool_context(self, updates: Dict[str, Any]):
        """更新工具执行上下文"""
        self.tool_context.update(updates)

# 创建全局实例（不自动初始化）
global_openai_connector: Optional[OpenAIConnector] = None
