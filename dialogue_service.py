# -*- coding: utf-8 -*-
"""
对话服务系统 - 支持Excel分析和PPT生成
"""

import json
import os
import asyncio
import logging
from typing import Dict, List, Any, Optional
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import traceback

from llm_client import OpenAIConnector
from tools.db import ExcelAnalysisOrchestrator
from tools.create_ppt_simplified import create_pptx_from_json
from tools.tool_registry import ToolRegistry
from tools.db import SqlExecutionTool
from tools.create_ppt_simplified import PptCreationTool
from tools.message_variable_processor import MessageVariableProcessor
from path_manager import path_manager

# 配置日志
logging.basicConfig(
    level=logging.INFO,  # 改为INFO级别，减少日志噪音
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(path_manager.get_log_path('dialogue_service.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DialogueService:
    """
    对话服务主类
    提供Excel分析和PPT生成功能
    """
    
    def _parse_llm_json_response(self, content: str, context: str = "") -> Dict[str, Any]:
        """
        统一的LLM JSON响应解析器
        
        参数:
            content: LLM返回的原始内容
            context: 上下文信息，用于错误日志
            
        返回:
            解析后的字典对象
            
        抛出:
            ValueError: JSON解析失败时抛出，调用方需要处理
        """
        # 清理markdown格式
        cleaned = content.strip()
        if cleaned.startswith("```json"):
            cleaned = cleaned[7:]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()
        
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError as e:
            # 记录详细错误信息
            logger.error(f"JSON解析失败 [{context}]: {e}")
            logger.error(f"原始内容: {content[:200]}...")
            raise ValueError(f"LLM返回的JSON格式无效: {e}") from e
    
    def __init__(self):
        """
        初始化对话服务
        """
        # 初始化LLM客户端（模型配置在llm_client中通过环境变量处理）
        self.llm_client = OpenAIConnector()
        
        # 初始化工具注册器并设置到LLM客户端
        self.tool_registry = ToolRegistry()
        self._setup_tools()
        self.llm_client.set_tool_registry(self.tool_registry)
        
        # 存储当前的Excel分析器
        self.excel_orchestrator: Optional[ExcelAnalysisOrchestrator] = None
        
        # 对话历史
        self.conversation_history: List[Dict[str, str]] = []
        
        # 线程池用于并行处理
        self.executor = ThreadPoolExecutor(max_workers=10)

        # 消息变量占位符处理器（与llm_client保持同一个实例）
        self.message_var_processor: MessageVariableProcessor = self.llm_client.message_var_processor

        # 全局系统提示：指导LLM如何使用变量占位符，避免回传海量数据
        self.conversation_history.append({
            "role": "system",
            "content": (
                "当你调用工具并收到包含 variable_binding 的结果时，不要在回复中直接展开原始数据。"
                "如需引用完整数据，请在最终回答中使用占位符形式 {\"<工具名>\":\"<变量名>\"}。"
                "若需要对大数据进行进一步聚合/筛选/排序，请描述操作或再次调用工具，而不是复制粘贴原始数据。"
            )
        })
        
    def _setup_tools(self):
        """设置LLM可用的工具"""
        # 注册工具处理器（工具定义由llm_client自动从pptx_json.json加载）
        self.tool_registry.register(SqlExecutionTool())
        self.tool_registry.register(PptCreationTool())
    
    def load_excel(self, excel_path: str) -> str:
        """
        加载Excel文件并生成数据上下文
        
        参数:
            excel_path: Excel文件路径
            
        返回:
            数据上下文描述
        """
        try:
            self.excel_orchestrator = ExcelAnalysisOrchestrator(excel_path)
            context = self.excel_orchestrator.get_llm_context()
            
            # 更新LLM客户端的工具上下文
            self.llm_client.update_tool_context({
                "excel_orchestrator": self.excel_orchestrator
            })
            
            # 添加到对话历史作为系统消息
            self.conversation_history.append({
                "role": "system",
                "content": f"用户已上传Excel文件，数据上下文如下：\n{context}"
            })
            
            return f"✅ Excel文件加载成功！\n{context}"
        except Exception as e:
            return f"❌ Excel文件加载失败：{str(e)}"
    

    
    def _generate_ppt_outline(self, user_requirement: str, data_context: str) -> Dict[str, Any]:
        """
        根据用户需求和数据生成PPT大纲
        
        参数:
            user_requirement: 用户的需求描述
            data_context: Excel数据上下文
            
        返回:
            PPT大纲结构
        """
        logger.info("开始生成PPT大纲")
        
        # 读取结构化大纲prompt模板
        prompt_path = path_manager.get_resource_path('prompts/prompt_outline.txt')
        with open(prompt_path, 'r', encoding='utf-8') as f:
            prompt_template = f.read()
        
        # 构建完整的prompt
        messages = [
            {"role": "system", "content": prompt_template},
            {"role": "user", "content": f"""
基于以下Excel数据，生成一份PPT报告大纲：

数据上下文：
{data_context}

用户需求：
{user_requirement}

请按照系统要求的JSON格式生成结构化的PPT大纲。
"""}
        ]
        
        # 调用LLM生成大纲
        try:
            response = self.llm_client.chat_completion(
                messages=messages,
                temperature=0.3,  # 降低温度以获得更稳定的JSON输出
                max_tokens=4096,
                auto_execute_tools=False  # 大纲生成不需要工具调用
            )
            
            outline_text = response.choices[0].message.content

            # 在解析JSON前，先替换占位符（理论上大纲不含，但保持一致性）
            outline_text = self.message_var_processor.resolve_placeholders_in_text(outline_text)
            
            # 解析JSON格式的大纲
            try:
                outline = self._parse_llm_json_response(outline_text, "outline_generation")
                logger.info(f"成功解析大纲，包含 {len(outline.get('sections', []))} 个章节")
                return outline
            except ValueError as e:
                # JSON解析失败，直接抛出异常让上层处理
                raise RuntimeError(f"大纲生成失败：{e}") from e
                
        except Exception as e:
            logger.error(f"生成大纲失败: {traceback.format_exc()}")
            raise
    
    def _generate_slide_content(self, section_title: str, subsection: Dict[str, Any], 
                               data_context: str, main_objective: str) -> Dict[str, Any]:
        """
        为单个幻灯片生成内容
        
        参数:
            section_title: 章节标题
            subsection: 子章节信息
            data_context: 数据上下文
            main_objective: 主要任务目标
            
        返回:
            幻灯片内容
        """
        slide_title = subsection.get('subsection_title', 'Unknown')
        logger.info(f"开始生成幻灯片内容: {slide_title}")
        # 构建针对该页面的prompt
        key_points = "\n".join(subsection.get("key_points", []))
        analysis_type = subsection.get("analysis_type", "summary")
        chart_type = subsection.get("chart_type", "none")
        data_query = subsection.get("data_query", "")
        
        # 构建页面生成prompt
        page_prompt = f"""
任务：为PPT页面生成具体内容

主要目标：{main_objective}
章节：{section_title}
页面标题：{subsection.get('subsection_title', '')}
分析类型：{analysis_type}
图表类型：{chart_type}
数据需求：{data_query}
关键要点：
{key_points}

数据上下文：
{data_context}

要求：
1. 如果需要数据，首先使用execute_sql工具查询所需数据，调用工具最大轮次为10次
2. 基于查询结果生成内容，内容要有洞察力和价值
3. 必须返回JSON格式，格式如下：

```json
{{
    "text": "这里是对数据的分析说明，要有具体的数据支撑，1-2段话",
    "bullet_points": [
        "第一个关键发现或洞察",
        "第二个关键发现或洞察",
        "第三个关键发现或洞察"
    ],
    "chart": {{
        "type": "{chart_type}",
        "title": "图表标题",
        "data": {{
            "categories": ["类别1", "类别2", "类别3"],
            "series": {{
                "系列名称1": [数值1, 数值2, 数值3],
                "系列名称2": [数值1, 数值2, 数值3]
            }}
        }}
    }}
}}
```

注意：
- text字段必须包含具体的分析内容，不能为空
- bullet_points必须是3-5个要点的数组
- 如果chart_type是"none"，则不需要chart字段
- 如果需要图表，确保data中的categories数量与每个series的数值数量一致
- 返回的必须是纯JSON，不要包含```json标记

请确保内容与分析类型和数据相符，提供有价值的洞察。
"""
        
        messages = [
            {"role": "system", "content": "你是一个数据分析和PPT内容生成专家。请根据要求生成准确、有洞察力的内容。重要：你必须返回纯JSON格式，不要包含任何额外的文字说明或markdown标记。"},
            {"role": "user", "content": page_prompt}
        ]
        
        try:
            # 允许工具调用来查询数据，由LLM客户端自动处理
            response = self.llm_client.chat_completion(
                messages=messages,
                tools=None,  # 让llm_client自动从tool_registry获取工具定义
                tool_choice="auto",
                temperature=0.5,
                auto_execute_tools=True  # 启用自动工具调用
            )
            
            # 获取最终响应内容
            content = response.choices[0].message.content

            # 在解析JSON前，将占位符替换为真实数据
            content = self.message_var_processor.resolve_placeholders_in_text(content)
            
            # 记录详细的LLM响应内容用于调试
            logger.info(f"LLM响应内容（前500字符）: {content[:500] if content else 'None'}")
            
            # 解析LLM返回的JSON内容
            try:
                llm_data = self._parse_llm_json_response(content, f"slide_content_{subsection.get('subsection_title', '')}")
                
                # 直接构建幻灯片数据结构
                slide_data = {
                    "type": "content",
                    "title": subsection.get('subsection_title', ''),
                    "contents": []
                }
                
                # 添加文本内容
                text_content = llm_data.get("text", "")
                bullet_points = llm_data.get("bullet_points", [])
                
                # 确保bullet_points是列表
                if isinstance(bullet_points, str):
                    bullet_points = [bullet_points]
                
                # 如果有文本内容或要点，添加文本部分
                if text_content or bullet_points:
                    slide_data["contents"].append({
                        "type": "text",
                        "text": text_content,
                        "bullet_points": bullet_points
                    })
                
                # 添加图表内容
                chart_info = llm_data.get("chart")
                if chart_info and isinstance(chart_info, dict):
                    chart_data = chart_info.get("data", {})
                    if chart_data.get("categories") and chart_data.get("series"):
                        slide_data["contents"].append({
                            "type": "chart",
                            "chart_type": chart_info.get("type", "column"),
                            "chart_title": chart_info.get("title", ""),
                            "data": chart_data
                        })
                
                logger.info(f"成功生成幻灯片内容: {subsection.get('subsection_title', '')}, 包含 {len(slide_data['contents'])} 个内容块")
                return slide_data
                
            except ValueError as e:
                # JSON解析失败，返回错误页面
                logger.error(f"幻灯片内容JSON解析失败: {e}")
                return {
                    "type": "content",
                    "title": subsection.get('subsection_title', '错误页面'),
                    "contents": [{
                        "type": "text",
                        "text": f"内容生成失败：{e}",
                        "bullet_points": [
                            "LLM返回的JSON格式无效",
                            "请检查prompt设置"
                        ]
                    }]
                }
            
        except Exception as e:
            logger.error(f"生成幻灯片内容失败: {e}")
            logger.error(traceback.format_exc())
            # 返回一个基本的错误页面
            return {
                "type": "content",
                "title": subsection.get('subsection_title', '错误页面'),
                "contents": [{
                    "type": "text",
                    "text": f"生成内容时出错: {str(e)}",
                    "bullet_points": []
                }]
            }
    
    async def generate_ppt_async(self, user_requirement: str, output_filename: str = "report") -> str:
        """
        异步生成PPT（支持并行生成内容）
        
        参数:
            user_requirement: 用户需求描述
            output_filename: 输出文件名
            
        返回:
            生成结果消息
        """
        if not self.excel_orchestrator:
            return "❌ 请先上传Excel文件"
        
        try:
            # 获取数据上下文
            data_context = self.excel_orchestrator.get_llm_context()
            
            print("📋 正在生成PPT大纲...")
            # 生成大纲
            outline = self._generate_ppt_outline(user_requirement, data_context)
            
            # 构建PPT结构
            ppt_data = {
                "slides": []
            }
            
            # 添加封面页
            ppt_data["slides"].append({
                "type": "cover",
                "title": outline["title"],
                "subtitle": f"生成时间：{datetime.now().strftime('%Y年%m月%d日')}"
            })
            
            print(f"🚀 开始并行生成 {sum(len(s['subsections']) for s in outline['sections'])} 个内容页...")
            
            # 准备所有内容页生成任务（先并行生成所有内容）
            loop = asyncio.get_event_loop()
            futures = []
            content_map = {}  # 用于存储内容页和其对应的位置
            
            for section_idx, section in enumerate(outline["sections"]):
                for subsection_idx, subsection in enumerate(section.get("subsections", [])):
                    # 创建唯一键来标识每个内容页的位置
                    content_key = f"{section_idx}_{subsection_idx}"
                    
                    future = loop.run_in_executor(
                        self.executor,
                        self._generate_slide_content,
                        section.get('section_title', ''),
                        subsection,
                        data_context,
                        user_requirement  # 传递主要目标
                    )
                    futures.append((content_key, future))
            
            # 等待所有内容页生成完成
            content_results = []
            for content_key, future in futures:
                result = await future
                content_results.append((content_key, result))
            
            # 将结果存储到map中
            for content_key, result in content_results:
                content_map[content_key] = result
            
            # 按正确的顺序组装PPT
            for section_idx, section in enumerate(outline["sections"]):
                # 添加章节页
                section_slide = {
                    "type": "section",
                    "title": f"{section.get('section_number', '')} {section.get('section_title', '')}"
                }
                ppt_data["slides"].append(section_slide)
                
                # 添加该章节的所有内容页
                for subsection_idx, subsection in enumerate(section.get("subsections", [])):
                    content_key = f"{section_idx}_{subsection_idx}"
                    if content_key in content_map:
                        content_slide = content_map[content_key]
                        ppt_data["slides"].append(content_slide)

            # 在末尾追加总结页（复用页面生成prompt，并把已生成的PPT数据摘要放入数据上下文）
            try:
                slide_titles = [s.get("title", "") for s in ppt_data["slides"] if isinstance(s, dict) and s.get("title")]
                titles_block = "\n".join(f"- {t}" for t in slide_titles)
                summary_context = f"{data_context}\n\n已生成PPT页面标题列表：\n{titles_block}"
                summary_subsection = {
                    "subsection_title": "总结",
                    "analysis_type": "summary",
                    "chart_type": "none",
                    "data_query": "",
                    "key_points": []
                }
                summary_slide = self._generate_slide_content(
                    "总结与展望",
                    summary_subsection,
                    summary_context,
                    user_requirement
                )
                ppt_data["slides"].append(summary_slide)
            except Exception as e:
                logger.warning(f"追加总结页失败，将继续生成PPT：{e}")
            
            print("📝 正在创建PPT文件...")
            # 生成PPT文件
            output_file_path = path_manager.get_output_path(output_filename)
            file_path = create_pptx_from_json(ppt_data, str(output_file_path))
            
            return f"✅ PPT生成成功！\n文件路径：{file_path}\n总页数：{len(ppt_data['slides'])} 页"
            
        except Exception as e:
            logger.error(f"PPT生成失败：{e}")
            logger.error(traceback.format_exc())
            return f"❌ PPT生成失败：{str(e)}"
    
    def process_message(self, user_message: str, generate_ppt: bool = False) -> str:
        """
        处理用户消息
        
        参数:
            user_message: 用户输入的消息
            generate_ppt: 是否生成PPT
            
        返回:
            系统响应
        """
        # 添加用户消息到历史
        self.conversation_history.append({
            "role": "user",
            "content": user_message
        })
        
        if generate_ppt:
            # 同步调用异步方法 - 在FastAPI环境中正确处理
            try:
                # 检查是否已在事件循环中
                asyncio.get_running_loop()
                # 如果已在事件循环中，创建新的事件循环在线程中运行
                loop = asyncio.new_event_loop()
                try:
                    result = loop.run_until_complete(self.generate_ppt_async(user_message))
                finally:
                    loop.close()
            except RuntimeError:
                # 没有运行的事件循环，可以直接使用asyncio.run
                result = asyncio.run(self.generate_ppt_async(user_message))
            return result
        else:
            # 普通对话模式，支持Excel分析
            try:
                response = self.llm_client.chat_completion(
                    messages=self.conversation_history,
                    tools=None if self.excel_orchestrator else [],  # 有Excel时让llm_client自动获取工具，否则不使用工具
                    tool_choice="auto" if self.excel_orchestrator else None,
                    auto_execute_tools=True  # 启用自动工具调用
                )
                
                # 获取最终响应内容
                assistant_message = response.choices[0].message.content

                # 返回给用户前：替换占位符为真实数据
                rendered_message = self.message_var_processor.resolve_placeholders_in_text(assistant_message)
                
                # 添加助手消息到历史
                self.conversation_history.append({
                    "role": "assistant",
                    # 历史中保存未展开版本，以节省后续token
                    "content": assistant_message
                })
                
                return rendered_message
                
            except Exception as e:
                return f"❌ 处理消息时出错：{str(e)}"
    
    def clear_history(self):
        """清空对话历史"""
        self.conversation_history = []
        if self.excel_orchestrator:
            # 重新添加数据上下文
            context = self.excel_orchestrator.get_llm_context()
            self.conversation_history.append({
                "role": "system",
                "content": f"用户已上传Excel文件，数据上下文如下：\n{context}"
            })

    def __del__(self):
        """清理资源"""
        if hasattr(self, 'executor') and self.executor:
            self.executor.shutdown(wait=False)


def main():
    """
    主函数 - 命令行交互界面
    """
    print("="*60)
    print("🤖 Excel分析与PPT生成对话服务")
    print("1. 输入 'load <文件路径>' 加载Excel文件")
    print("2. 输入 'ppt <需求描述>' 生成PPT")
    print("3. 直接输入问题进行Excel数据分析")
    print("4. 输入 'clear' 清空对话历史")
    print("5. 输入 'exit' 退出程序")
    print("="*60)
    
    # 创建服务实例
    service = DialogueService()
    
    while True:
        user_input = input("\n👤 用户: ").strip()
        
        if user_input.lower() == 'exit':
            break
        
        elif user_input.lower() == 'clear':
            service.clear_history()
            print("✅ 对话历史已清空")
        
        elif user_input.lower().startswith('load '):
            file_path = user_input[5:].strip()
            result = service.load_excel(file_path)
            print(f"\n🤖 系统: {result}")
        
        elif user_input.lower().startswith('ppt '):
            requirement = user_input[4:].strip()
            print("\n🤖 系统: 正在生成PPT，请稍候...")
            result = service.process_message(requirement, generate_ppt=True)
            print(f"\n🤖 系统: {result}")
        
        else:
            print("\n🤖 系统: 正在处理您的请求...")
            result = service.process_message(user_input)
            print(f"\n🤖 系统: {result}")


if __name__ == "__main__":
    main()
