import asyncio
from functools import wraps
import json
import httpx
from typing import Callable, Dict, Any, Optional, Union, List
from enum import Enum
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)

class HttpMethod(Enum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    DELETE = "DELETE"
    PATCH = "PATCH"
    HEAD = "HEAD"
    OPTIONS = "OPTIONS"


class AsyncHttpClient:
    """异步HTTP客户端类"""
    
    def __init__(self, base_url: Optional[str] = None, default_headers: Dict[str, str] = {}, verify_ssl: bool = False):
        """
        初始化异步HTTP客户端
        
        Args:
            base_url: 基础URL，所有请求将基于此URL
            default_headers: 默认请求头
            verify_ssl: 是否验证SSL证书
        """
        self.base_url = base_url.rstrip('/') if base_url else None
        self.client: Optional[httpx.AsyncClient] = None
        self._auth_token: Optional[str] = None
        self.default_headers = {"content-type": "application/json;charset=utf-8"}
        self.verify_ssl = verify_ssl
        if default_headers:
            self.default_headers.update(default_headers)

    async def __aenter__(self):
        """异步上下文管理器入口"""
        await self.start()
        return self
        
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """异步上下文管理器退出"""
        await self.close()
        
    async def start(self):
        """启动客户端"""
        if self.client is None:
            headers = self.default_headers.copy()
            
            # 如果有token，添加到默认headers
            if self._auth_token:
                headers["token"] = self._auth_token
                
            self.client = httpx.AsyncClient(
                base_url=self.base_url,
                headers=headers,
                verify=self.verify_ssl  # 在这里设置SSL验证
            )
    
    async def close(self):
        """关闭客户端"""
        if self.client:
            await self.client.aclose()
            self.client = None
    
    def set_token(self, token: str):
        """
        设置认证token
        
        Args:
            token: 认证token字符串
        """
        self._auth_token = token
        
        # 如果客户端已经启动，更新headers
        if self.client:
            self.client.headers["token"] = token
    
    async def set_url(self, base_url: str):
        """
        设置基础URL
        
        Args:
            base_url: 新的基础URL
        """
        self.base_url = base_url.rstrip('/') if base_url else None
        # 如果客户端已经启动，重新启动客户端以应用新的基础URL
        if self.client:
            await self.close()
            await self.start()
    
    async def set_verify_ssl(self, verify_ssl: bool):
        """
        设置SSL验证
        
        Args:
            verify_ssl: 是否验证SSL证书
        """
        self.verify_ssl = verify_ssl
        # 如果客户端已经启动，重新启动客户端以应用新的SSL设置
        if self.client:
            await self.close()
            await self.start()
    
    def clear_token(self):
        """清除认证token"""
        self._auth_token = None
        if self.client and "token" in self.client.headers:
            del self.client.headers["token"]
    
    async def request(
        self,
        method: Union[str, HttpMethod],
        url: str,
        headers: Optional[Dict[str, str]] = None,
        params: Optional[Dict[str, Any]] = None,
        json: Optional[Dict[str, Any]] = None,
        data: Optional[Union[str, bytes, Dict[str, Any]]] = None,
        **kwargs
    ) -> httpx.Response:
        """
        发送HTTP请求
        
        Args:
            method: HTTP方法
            url: 请求URL
            headers: 请求头
            params: URL参数
            json: 请求的JSON体
            data: 请求的数据体
            **kwargs: 其他传递给httpx的参数
            
        Returns:
            httpx.Response: 响应对象
        """
        if self.client is None:
            await self.start()
            
        if isinstance(method, HttpMethod):
            method = method.value
            
        headers = headers or {}
        retries = 0
        
        while retries <= 3:  # 默认重试次数
            try:
                response = await self.client.request(
                    method=method,
                    url=url,
                    headers={**self.client.headers, **headers},
                    params=params,
                    json=json,
                    data=data,
                    **kwargs
                )
                response.raise_for_status()
                return response
            except (httpx.RequestError, httpx.HTTPStatusError) as e:
                retries += 1
                if retries > 3:
                    log.error(f"请求失败: {e}, URL: {url}, 方法: {method}")
                    raise e
                await asyncio.sleep(1 * retries)  # 重试延迟
                log.warning(f"第{retries}次重试: {url}, 错误信息: {str(e)}")

    def response_to_dict(func: Callable) -> Callable:
        """将响应转换为字典的装饰器"""
        @wraps(func)
        async def wrapper(*args, **kwargs) -> Dict[str, Any]:
            response = await func(*args, **kwargs)
            try:
                return response.json()
            except json.JSONDecodeError:
                return {"content": response.text, "status_code": response.status_code}
        return wrapper
    
    @response_to_dict  
    async def get(self, url: str, headers: Optional[Dict[str, str]] = None, params: Optional[Dict[str, Any]] = None, **kwargs) -> httpx.Response:
        """发送GET请求"""
        return await self.request(HttpMethod.GET, url, headers=headers, params=params, **kwargs)
    
    @response_to_dict 
    async def post(self, url: str, headers: Optional[Dict[str, str]] = None, data: Optional[Union[str, bytes, Dict[str, Any]]] = None, json: Optional[Dict[str, Any]] = None, **kwargs) -> httpx.Response:
        """发送POST请求"""
        return await self.request(HttpMethod.POST, url, headers=headers, data=data, json=json, **kwargs)
    
    @response_to_dict 
    async def put(self, url: str, headers: Optional[Dict[str, str]] = None, data: Optional[Union[str, bytes, Dict[str, Any]]] = None, json: Optional[Dict[str, Any]] = None, **kwargs) -> httpx.Response:
        """发送PUT请求"""
        return await self.request(HttpMethod.PUT, url, headers=headers, data=data, json=json, **kwargs)
    
    @response_to_dict 
    async def delete(self, url: str, headers: Optional[Dict[str, str]] = None, params: Optional[Dict[str, Any]] = None, **kwargs) -> httpx.Response:
        """发送DELETE请求"""
        return await self.request(HttpMethod.DELETE, url, headers=headers, params=params, **kwargs)
    
    @response_to_dict   
    async def patch(self, url: str, headers: Optional[Dict[str, str]] = None, data: Optional[Union[str, bytes, Dict[str, Any]]] = None, json: Optional[Dict[str, Any]] = None, **kwargs) -> httpx.Response:
        """发送PATCH请求"""
        return await self.request(HttpMethod.PATCH, url, headers=headers, data=data, json=json, **kwargs)
    
    @response_to_dict   
    async def head(self, url: str, headers: Optional[Dict[str, str]] = None, **kwargs) -> httpx.Response:
        """发送HEAD请求"""
        return await self.request(HttpMethod.HEAD, url, headers=headers, **kwargs)
    
    @response_to_dict  
    async def options(self, url: str, headers: Optional[Dict[str, str]] = None, **kwargs) -> httpx.Response:
        """发送OPTIONS请求"""
        return await self.request(HttpMethod.OPTIONS, url, headers=headers, **kwargs)
    
    async def download_file(self, url: str, file_path: str, headers: Optional[Dict[str, str]] = None, **kwargs) -> None:
        """
        下载文件到本地
        
        Args:
            url: 文件URL
            file_path: 本地文件路径
            headers: 请求头
            **kwargs: 其他参数
        """
        response = await self.get(url, headers=headers, **kwargs)
        
        with open(file_path, 'wb') as f:
            async for chunk in response.aiter_bytes():
                f.write(chunk)
    
    async def batch_request(
        self,
        requests: List[Dict[str, Any]],
        max_concurrent: int = 10
    ) -> List[Any]:
        """
        批量发送请求
        
        Args:
            requests: 请求列表，每个元素是包含method, url, config等参数的字典
            max_concurrent: 最大并发数
            
        Returns:
            List[Any]: 响应列表
        """
        semaphore = asyncio.Semaphore(max_concurrent)
        
        async def limited_request(request_args):
            async with semaphore:
                return await self.request(**request_args)
        
        tasks = [limited_request(req) for req in requests]
        return await asyncio.gather(*tasks, return_exceptions=True)