import asyncio
import functools
import inspect
from typing import Any, Callable, Coroutine, Optional, Type, TypeVar, Union
from utils.log_tools.logger_utils import get_logger
T = TypeVar('T')
log = get_logger(__name__)

def async_retry_on_empty(
    retries: int = 50,
    interval: int = 5,
    check: Optional[Callable[[Any], bool]] = None,
    target_param: str = "api"  # 新增：动态指定目标参数名
):
    def decorator(func: Callable[..., Coroutine]):
        @functools.wraps(func)
        async def wrapper(*args, **kwargs):
            # 动态获取目标参数值
            target_value = kwargs.get(target_param) or get_arg_by_name(func, args, target_param)
            identifier = target_value or "unknown"

            for attempt in range(retries + 1):
                result = await func(*args, **kwargs)
                if check is None or check(result):
                    return result
                if attempt < retries:
                    await asyncio.sleep(interval)
                    log.warning(f"{identifier} 未找到结果，重试{attempt + 1}次")
            log.warning(f"函数 {func.__name__} 重试 {retries} 次后仍没找到匹配数据")
            return result 
        return wrapper
    return decorator

def get_arg_by_name(func: Callable, args: tuple, param_name: str) -> Any:
    """通过函数签名从位置参数中获取指定参数名的值"""
    sig = inspect.signature(func)
    bound_args = sig.bind(*args)
    return bound_args.arguments.get(param_name)

def transform_to_data_class(data_class: Type[T]):
    """独立的数据转换装饰器函数"""
    def decorator(func):
        @functools.wraps(func)
        async def wrapper(*args, **kwargs):
            response = await func(*args, **kwargs)
            if not response:
                return None
            if isinstance(response, list):
                return [data_class(**item) for item in response if item]
            return data_class(**response)
        return wrapper
    return decorator