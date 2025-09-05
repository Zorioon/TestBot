import asyncio
import pytest
import pytest_asyncio

from utils.auth_tools.auth_utils import AuthUtils
from utils.log_tools.logger_utils import ProjectLogger
from utils.notice_tools.webcom_utils import WeComRobot
from utils.request_tools.async_http_client import AsyncHttpClient
from utils.ssh_tools.ssh_connect import AsyncSSHClient
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

@pytest_asyncio.fixture(scope='session')
async def http_req():
    """返回一个配置好的 http 客户端, 由于IP:PORT不一样, 支持set_url切换url"""
    try:
        client = AsyncHttpClient()
        log.success("成功初始化会话级 http_req fixture")
        yield client
    except Exception as e:
        log.exception("初始化 http_req fixture 失败")
        raise RuntimeError("初始化 http_req fixture 失败") from e
    finally:
        log.success("测试结束，释放 http_req fixture")
        await client.close()
    
@pytest_asyncio.fixture(scope='session')
async def https_req(sc_config):
    """返回一个配置好的 https 客户端"""
    try:
        client = AsyncHttpClient(f'https://{sc_config["sc_ip"]}')
        auth_token = await AuthUtils.login(client, sc_config["username"], sc_config["password"])
        client.set_token(auth_token)
        log.success("成功初始化会话级 https_req fixture")
        yield client
    except Exception as e:
        log.exception("初始化 https_req fixture 失败")
        raise RuntimeError("初始化 https_req fixture 失败") from e
    finally:
        log.success("测试结束，释放 https_req fixture")
        await client.close()

@pytest_asyncio.fixture(scope='session')
async def sc_ssh_client(sc_config):
    """返回一个配置好的 总控ssh 客户端"""
    try:
        ssh_client = AsyncSSHClient(host=sc_config["sc_ip"], username=sc_config["ssh_username"], password=sc_config["ssh_password"])
        await ssh_client.connect()
        log.success("成功初始化总控ssh连接 sc_ssh_client")
        yield ssh_client
    except Exception as e:
        log.exception("初始化 sc_ssh_client 失败")
        raise RuntimeError("初始化 sc_ssh_client 失败") from e
    finally:
        log.success("测试结束，释放 sc_ssh_client")
        await ssh_client.close()

@pytest_asyncio.fixture(scope='session')
async def wecom_robot(notice_config):
    """返回一个配置好的 总控ssh 客户端"""
    try:
        robot = WeComRobot(notice_config["webhook_key"])
        log.success("成功初始化企微通知机器人 wecom_robot")
        yield robot
    except Exception as e:
        log.exception("初始化 wecom_robot 失败")
        raise RuntimeError("初始化 wecom_robot 失败") from e
    finally:
        log.success("测试结束，释放 wecom_robot")
        await robot.close()


@pytest.fixture(scope='session', autouse=True)
def load_config():
    config = YAMLUtil.read_yaml('./common/config.yaml')
    return config

@pytest.fixture(scope='session')
def proxy_apps():
    return {
        "data_label": ("192.192.101.220:20010", "192.192.101.220:20011")
    }
@pytest.fixture(scope='session', autouse=True)
def sc_config(load_config):
    return load_config.get('sc', {})

@pytest.fixture(scope='session', autouse=True)
def database_config(load_config):
    return load_config.get('database', {})

@pytest.fixture(scope='session', autouse=True)
def businesses(load_config):
    return load_config.get('businesses', {})

@pytest.fixture(scope='session', autouse=True)
def notice_config(load_config):
    return load_config.get('notice', {})

    
