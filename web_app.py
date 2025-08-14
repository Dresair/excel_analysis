# -*- coding: utf-8 -*-
"""
FastAPI Web应用 - PPT生成工具的Web版本
"""

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
import uuid
import asyncio
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional
import json
import shutil
from pathlib import Path
import webbrowser
import threading

# 导入现有的服务
from dialogue_service import DialogueService
from config_manager import ConfigManager
from path_manager import path_manager

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 创建FastAPI应用
app = FastAPI(
    title="PPT生成工具",
    description="基于AI的Excel数据分析与PPT自动生成系统",
    version="2.0.0"
)

# 启用CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
async def on_startup():
    """
    应用启动时执行的事件，用于自动打开浏览器。
    """
    def open_browser():
        webbrowser.open("http://localhost:8000")

    # 延迟1秒后打开浏览器，确保Uvicorn服务已准备就绪
    threading.Timer(1, open_browser).start()


# 注意：不需要手动创建目录，path_manager的get_*_path方法会自动创建可写目录


# 配置管理器
config_manager = ConfigManager()
config_manager.apply_to_environment()

# 会话管理
sessions: Dict[str, DialogueService] = {}

# 数据模型
class ChatMessage(BaseModel):
    message: str
    session_id: str

class PPTGenerateRequest(BaseModel):
    message: str
    session_id: str

class ConfigUpdate(BaseModel):
    api_key: str
    base_url: str
    model: str

# ======================== API 路由 ========================

@app.get("/", response_class=HTMLResponse)
async def read_root():
    """返回主页面"""
    index_path = path_manager.get_resource_path("static/index.html")
    if not index_path.exists():
        return HTMLResponse(content="<html><body><h1>欢迎使用PPT生成工具</h1><p>主页面文件丢失。</p></body></html>", status_code=404)
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    """上传Excel文件"""
    try:
        # 验证文件类型
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="请上传Excel文件(.xlsx或.xls)")
        
        # 生成会话ID
        session_id = str(uuid.uuid4())
        
        # 保存文件到临时目录
        temp_path = path_manager.get_temp_path(f"{session_id}_{file.filename}")
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 创建对话服务并加载Excel
        dialogue_service = DialogueService()
        result = dialogue_service.load_excel(temp_path)
        
        # 保存会话
        sessions[session_id] = dialogue_service
        
        logger.info(f"Excel文件上传成功: {file.filename}, 会话ID: {session_id}")
        
        return {
            "success": True,
            "session_id": session_id,
            "filename": file.filename,
            "message": result
        }
        
    except Exception as e:
        logger.error(f"Excel上传失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/chat")
async def chat(request: ChatMessage):
    """处理聊天消息"""
    try:
        session_id = request.session_id
        if session_id not in sessions:
            raise HTTPException(status_code=404, detail="会话不存在，请先上传Excel文件")
        
        dialogue_service = sessions[session_id]
        response = dialogue_service.process_message(request.message)
        
        return {
            "success": True,
            "response": response,
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"聊天处理失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-ppt")
async def generate_ppt(request: PPTGenerateRequest, background_tasks: BackgroundTasks):
    """生成PPT（异步）"""
    try:
        session_id = request.session_id
        if session_id not in sessions:
            raise HTTPException(status_code=404, detail="会话不存在，请先上传Excel文件")
        
        # 生成任务ID
        task_id = str(uuid.uuid4())
        
        # 在后台执行PPT生成
        background_tasks.add_task(
            generate_ppt_task, 
            task_id, 
            session_id, 
            request.message
        )
        
        return {
            "success": True,
            "task_id": task_id,
            "message": "PPT生成任务已启动，请稍后查看进度"
        }
        
    except Exception as e:
        logger.error(f"PPT生成启动失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

# 任务状态存储
task_status: Dict[str, Dict] = {}

async def generate_ppt_task(task_id: str, session_id: str, message: str):
    """后台PPT生成任务"""
    try:
        task_status[task_id] = {
            "status": "processing",
            "progress": 0,
            "message": "正在生成PPT..."
        }
        
        dialogue_service = sessions[session_id]
        
        # 更新进度
        task_status[task_id]["progress"] = 30
        task_status[task_id]["message"] = "正在分析数据..."
        
        # 生成PPT
        result = dialogue_service.process_message(message, generate_ppt=True)
        
        # 更新进度
        task_status[task_id]["progress"] = 100
        task_status[task_id]["status"] = "completed"
        task_status[task_id]["message"] = result
        
        # 提取文件路径
        if "文件路径：" in result:
            file_path = result.split("文件路径：")[1].split("\n")[0].strip()
            task_status[task_id]["file_path"] = file_path
        
    except Exception as e:
        task_status[task_id] = {
            "status": "failed",
            "progress": 0,
            "message": f"PPT生成失败: {str(e)}"
        }
        logger.error(f"PPT生成任务失败: {e}")

@app.get("/api/task-status/{task_id}")
async def get_task_status(task_id: str):
    """获取任务状态"""
    if task_id not in task_status:
        raise HTTPException(status_code=404, detail="任务不存在")
    
    return task_status[task_id]

@app.get("/api/output-files")
async def list_output_files():
    """获取输出文件列表"""
    try:
        files = []
        output_dir = path_manager.get_resource_path("output")
        
        if output_dir.exists():
            for file_path in output_dir.glob("*.pptx"):
                stat = file_path.stat()
                files.append({
                    "filename": file_path.name,
                    "size": stat.st_size,
                    "created_time": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "download_url": f"/api/download/{file_path.name}"
                })
        
        # 按创建时间降序排列
        files.sort(key=lambda x: x["created_time"], reverse=True)
        
        return {"files": files}
        
    except Exception as e:
        logger.error(f"获取文件列表失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """下载PPT文件"""
    file_path = path_manager.get_output_path(filename)
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

@app.post("/api/config")
async def update_config(config: ConfigUpdate):
    """更新配置"""
    try:
        config_manager.set_openai_config(
            config.api_key,
            config.base_url,
            config.model
        )
        
        if config_manager.save_config():
            config_manager.apply_to_environment()
            
            # 重新初始化所有现有会话的LLM客户端
            from llm_client import OpenAIConnector
            llm_client = OpenAIConnector.get_instance()
            llm_client.reinitialize_client()
            
            return {"success": True, "message": "配置更新成功"}
        else:
            raise HTTPException(status_code=500, detail="配置保存失败")
            
    except Exception as e:
        logger.error(f"配置更新失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/config")
async def get_config():
    """获取当前配置"""
    try:
        openai_config = config_manager.get_openai_config()
        ui_config = config_manager.get_ui_config()
        
        # 隐藏API密钥的敏感信息
        api_key = openai_config.get("api_key", "")
        masked_api_key = ""
        if api_key:
            if len(api_key) > 10:
                masked_api_key = api_key[:8] + "..." + api_key[-4:]
            else:
                masked_api_key = "***已设置***"
        
        return {
            "api_key": api_key,  # 返回完整key用于编辑
            "api_key_masked": masked_api_key,  # 用于显示
            "base_url": openai_config.get("base_url", ""),
            "model": openai_config.get("model", ""),
            "is_configured": config_manager.is_openai_configured(),
            "config_file_path": str(config_manager.config_file_writable),
            "ui": ui_config
        }
    except Exception as e:
        logger.error(f"获取配置失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/logs/llm")
async def get_llm_logs(limit: int = 50):
    """获取LLM交互日志"""
    try:
        log_file_path = path_manager.get_log_path('llm_interactions.log')
        
        if not log_file_path.exists():
            return {"logs": [], "total_count": 0, "has_content": False}
        
        logs = []
        with open(log_file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            # 获取最新的limit条日志
            recent_lines = lines[-limit:] if len(lines) > limit else lines
            
            for line in recent_lines:
                try:
                    log_entry = json.loads(line.strip())
                    logs.append(log_entry)
                except json.JSONDecodeError:
                    continue
        
        return {
            "logs": logs,
            "total_count": len(logs),
            "has_content": len(logs) > 0,
            "file_path": str(log_file_path)
        }
        
    except Exception as e:
        logger.error(f"获取LLM日志失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/session/{session_id}")
async def clear_session(session_id: str):
    """清除会话"""
    if session_id in sessions:
        del sessions[session_id]
        return {"success": True, "message": "会话已清除"}
    else:
        raise HTTPException(status_code=404, detail="会话不存在")

@app.get("/api/table-data/{data_id}")
async def get_table_data(data_id: str):
    """获取完整表格数据供复制功能使用"""
    try:
        # 尝试从所有会话中查找表格数据
        table_data = None
        for session_id, service in sessions.items():
            if hasattr(service, 'message_var_processor'):
                data = service.message_var_processor.get_table_data_for_copy(data_id)
                if data:
                    table_data = data
                    break
        
        if not table_data:
            raise HTTPException(status_code=404, detail="表格数据不存在或已过期")
        
        return {
            "success": True,
            "data": table_data['data'],
            "headers": table_data['headers']
        }
        
    except Exception as e:
        logger.error(f"获取表格数据失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))



# 挂载静态文件
static_dir = path_manager.get_resource_path("static")
app.mount("/static", StaticFiles(directory=static_dir), name="static")

if __name__ == "__main__":
    import uvicorn
    
    # 检查配置
    if not config_manager.is_openai_configured():
        print("⚠️  警告：OpenAI配置未完成，请在Web界面中完成配置")
    
    print("🚀 启动PPT生成工具Web服务...")
    print("📱 请在浏览器中访问: http://localhost:8000")
    
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info",
        log_config=None
    )


