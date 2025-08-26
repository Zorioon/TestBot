import sys
from pathlib import Path
from loguru import logger
from typing import Optional

class ProjectLogger:
    def __init__(self, project_name: str = "TESTBOT", log_level: str = "INFO"):
        self.project_name = project_name
        self.log_dir = Path("logs")
        self.log_level = log_level  # 默认从环境变量获取
        self._setup_logger()

    def _setup_logger(self):
        """配置logger，使用动态级别"""
        self.log_dir.mkdir(exist_ok=True)
        logger.remove()

        logger.add(
            sys.stdout,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
            level=self.log_level,  # 使用动态级别
            colorize=True
        )

        # 文件日志保持原样
        logger.add(
            self.log_dir / f"{self.project_name}.log",
            # serialize=True,  
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            level="DEBUG",
            rotation="10 MB",
            retention="30 days",
            compression="zip",
            enqueue=True,
            backtrace=True,
            diagnose=True
        )

        logger.add(
            self.log_dir / f"{self.project_name}.error.log",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            # serialize=True,  
            level="ERROR",
            rotation="10 MB",
            retention="60 days",
            compression="zip",
            enqueue=True
        )

    def get_logger(self, module_name: Optional[str] = None):
        if module_name:
            return logger.bind(module=module_name)
        return logger

    def set_level(self, level: str):
        """动态修改日志级别"""
        self.log_level = level
        # 重新配置 logger
        self._setup_logger()

# 全局实例
project_logger = ProjectLogger()
log = project_logger.get_logger(__name__)

def get_logger(module_name: str = None):
    return project_logger.get_logger(module_name)
