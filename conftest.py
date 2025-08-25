import pytest

from utils.log_tools.logger_utils import ProjectLogger, get_logger
from utils.yaml_tools.yaml_utils import YAMLUtil


@pytest.hookimpl(tryfirst=True)
def pytest_configure(config):
    # 从 pytest.ini 配置文件读取日志级别
    log_level = config.getini('log_level')  # 获取配置项 log_level
    # 根据日志级别重新初始化 ProjectLogger
    project_logger = ProjectLogger(log_level=log_level)
    # 将 log 绑定到全局
    global log
    log = project_logger.get_logger()

@pytest.fixture(scope='session', autouse=True)
def load_config():
    config = YAMLUtil.read_yaml('./common/config.yaml')
    return config

@pytest.fixture(scope='session', autouse=True)
def sc_config(load_config):
    return load_config.get('sc', {})

@pytest.fixture(scope='session', autouse=True)
def database_config(load_config):
    return load_config.get('database', {})

@pytest.fixture(scope='session', autouse=True)
def notice_config(load_config):
    return load_config.get('notice', {})

    
