# -*- coding: utf-8 -*-
"""
路径管理器 - 统一处理开发环境和打包环境中的资源路径
"""
import sys
import os
from pathlib import Path

class PathManager:
    """
    路径管理器，用于获取不同环境下的正确资源路径。
    """
    def __init__(self):
        """
        初始化路径管理器，确定应用根目录。
        - 在PyInstaller打包环境中，根目录是 _MEIPASS 临时目录。
        - 在开发环境中，根目录是项目主文件所在的目录。
        """
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # 打包后的可执行文件
            # 资源文件（只读）在 _MEIPASS 临时目录
            self.resource_base_path = Path(sys._MEIPASS)
            # 可写文件（日志、输出）应在可执行文件旁边
            self.writable_base_path = Path(sys.executable).parent
        else:
            # 正常Python环境
            # 我们假设 path_manager.py 在项目根目录或者一个子目录中
            # 通过追溯到包含 'requirements.txt' 或 '.git' 的目录来确定项目根目录
            current_path = Path(__file__).resolve()
            project_root = current_path.parent
            # 如果需要更可靠的根目录查找，可以向上遍历
            while not (project_root / 'requirements.txt').exists() and not (project_root / '.git').exists():
                if project_root.parent == project_root:
                    # 到达文件系统根目录，使用初始路径
                    project_root = current_path.parent
                    break
                project_root = project_root.parent
            self.resource_base_path = project_root
            self.writable_base_path = project_root


    def get_resource_path(self, relative_path: str) -> Path:
        """
        获取捆绑的只读资源的绝对路径。
        例如: config.json, static/, prompts/
        
        参数:
            relative_path: 相对于项目根目录的资源路径。

        返回:
            资源的绝对路径 (pathlib.Path对象)。
        """
        return self.resource_base_path / relative_path

    def get_output_path(self, filename: str) -> Path:
        """
        获取输出文件的路径，并确保目录存在。

        参数:
            filename: 输出文件名。

        返回:
            输出文件的绝对路径。
        """
        output_dir = self.writable_base_path / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / filename

    def get_log_path(self, filename: str) -> Path:
        """
        获取日志文件的路径，并确保目录存在。

        参数:
            filename: 日志文件名。

        返回:
            日志文件的绝对路径。
        """
        log_dir = self.writable_base_path / "log"
        log_dir.mkdir(parents=True, exist_ok=True)
        return log_dir / filename

    def get_temp_path(self, filename: str) -> Path:
        """
        获取临时文件的路径，并确保目录存在。

        参数:
            filename: 临时文件名。

        返回:
            临时文件的绝对路径。
        """
        temp_dir = self.writable_base_path / "temp"
        temp_dir.mkdir(parents=True, exist_ok=True)
        return temp_dir / filename

# 创建一个全局实例供项目其他模块使用
path_manager = PathManager()
