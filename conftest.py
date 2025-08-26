import asyncio
from http import client
import pytest
import pytest_asyncio

from utils.auth_tools.auth_utils import AuthUtils
from utils.log_tools.logger_utils import ProjectLogger
from utils.request_tools.async_http_client import AsyncHttpClient
from utils.yaml_tools.yaml_utils import YAMLUtil
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)

@pytest.hookimpl(tryfirst=True)
def pytest_configure(config):
    # 从 pytest.ini 配置文件读取日志级别
    log_level = config.getini('log_level')  # 获取配置项 log_level
    # 根据日志级别重新初始化 ProjectLogger
    project_logger = ProjectLogger(log_level=log_level)
    # 将 log 绑定到全局
    global log
    log = project_logger.get_logger()


@pytest.fixture(scope="session")
def event_loop():
    """创建一个会话级别的事件循环"""
    policy = asyncio.get_event_loop_policy()
    loop = policy.new_event_loop()
    yield loop
    loop.close()


# @pytest.fixture(scope='session')
# def http_req():
#     client = AsyncHttpClient("https://192.192.101.175")
#     return client

@pytest_asyncio.fixture(scope='session')
async def https_req(sc_config):
    """返回一个配置好的 httpx 客户端"""
    client = AsyncHttpClient(f'https://{sc_config["sc_ip"]}')
    try:
        auth_token = await AuthUtils.login(https_req, sc_config["username"], sc_config["password"])
        https_req.set_token(auth_token)
        log.info("成功初始化会话级 https_req fixture")
        yield client
    except Exception as e:
        log.exception("初始化 https_req fixture 失败")
        raise RuntimeError("初始化 https_req fixture 失败") from e
    finally:
        await client.close()


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

    
