import asyncio
from functools import wraps
import json
import os
import urllib.request
import aiofiles
import httpx
from typing import Callable, Dict, Any, Optional, Union, List
from enum import Enum

import urllib

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

    def __init__(
        self,
        base_url: Optional[str] = None,
        default_headers: Dict[str, str] = {},
        verify_ssl: bool = False,
        timeout: Union[float, httpx.Timeout] = 60.0,
    ):
        """
        初始化异步HTTP客户端

        Args:
            base_url: 基础URL，所有请求将基于此URL
            default_headers: 默认请求头
            verify_ssl: 是否验证SSL证书
        """
        self.base_url = base_url.rstrip("/") if base_url else None
        self.client: Optional[httpx.AsyncClient] = None
        self._auth_token: Optional[str] = None
        self._timeout = timeout
        self.verify_ssl = verify_ssl
        self.default_headers = default_headers

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

            system_proxys = urllib.request.getproxies() or None
            format_systemt_proxy = (
                {
                    k + "://": v
                    for k, v in system_proxys.items()
                    if not k.endswith("://")
                }
                if system_proxys
                else None
            )
            self.client = httpx.AsyncClient(
                base_url=self.base_url,
                headers=headers,
                verify=self.verify_ssl,  # 在这里设置SSL验证
                proxies=format_systemt_proxy,
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
        self.base_url = base_url.rstrip("/") if base_url else None
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
        **kwargs,
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

        builtin_headers = {"content-type": "application/json;charset=utf-8"}
        headers = {**builtin_headers, **(headers or {})}
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
                    timeout=self._timeout,
                    **kwargs,
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
    async def get(
        self,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        params: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> httpx.Response:
        """发送GET请求"""
        return await self.request(
            HttpMethod.GET, url, headers=headers, params=params, **kwargs
        )

    @response_to_dict
    async def post(
        self,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        data: Optional[Union[str, bytes, Dict[str, Any]]] = None,
        json: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> httpx.Response:
        """发送POST请求"""
        return await self.request(
            HttpMethod.POST, url, headers=headers, data=data, json=json, **kwargs
        )

    @response_to_dict
    async def put(
        self,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        data: Optional[Union[str, bytes, Dict[str, Any]]] = None,
        json: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> httpx.Response:
        """发送PUT请求"""
        return await self.request(
            HttpMethod.PUT, url, headers=headers, data=data, json=json, **kwargs
        )

    @response_to_dict
    async def delete(
        self,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        params: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> httpx.Response:
        """发送DELETE请求"""
        return await self.request(
            HttpMethod.DELETE, url, headers=headers, params=params, **kwargs
        )

    @response_to_dict
    async def patch(
        self,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        data: Optional[Union[str, bytes, Dict[str, Any]]] = None,
        json: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> httpx.Response:
        """发送PATCH请求"""
        return await self.request(
            HttpMethod.PATCH, url, headers=headers, data=data, json=json, **kwargs
        )

    @response_to_dict
    async def head(
        self, url: str, headers: Optional[Dict[str, str]] = None, **kwargs
    ) -> httpx.Response:
        """发送HEAD请求"""
        return await self.request(HttpMethod.HEAD, url, headers=headers, **kwargs)

    @response_to_dict
    async def options(
        self, url: str, headers: Optional[Dict[str, str]] = None, **kwargs
    ) -> httpx.Response:
        """发送OPTIONS请求"""
        return await self.request(HttpMethod.OPTIONS, url, headers=headers, **kwargs)

    def get_mime_type(self, file_path: str) -> str:
        """根据文件后缀返回 MIME 类型"""
        ext = os.path.splitext(file_path)[1].lower()
        mime_types = {
            ".zip": "application/zip",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc": "application/msword",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls": "application/vnd.ms-excel",
            ".pdf": "application/pdf",
            ".csv": "text/csv; charset=utf-8",
            ".txt": "text/plain; charset=utf-8",
            ".gif": "image/gif",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
        }
        return mime_types.get(ext, "application/octet-stream")

    @response_to_dict
    async def upload_file(
        self,
        file_path: str,
        url: str,
        headers: Optional[Dict[str, str]] = None,
        use_multipart: bool = False,
        **kwargs,
    ) -> httpx.Response:
        """
        异步上传文件，并自动根据文件后缀设置 Content-Type
        """
        if self.client is None:
            await self.start()

        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        headers = headers or {}
        mime_type = self.get_mime_type(file_path)

        if use_multipart:
            async with aiofiles.open(file_path, "rb") as f:
                content = await f.read()
            filename = os.path.basename(file_path)
            files = {
                "file": (filename, content, mime_type),
                "path": (None, "/yzm"),  # 额外字段
            }

            response = await self.client.post(
                url, headers={**self.client.headers, **headers}, files=files, **kwargs
            )
            response.raise_for_status()
            return response
        else:
            # PUT 上传
            async with aiofiles.open(file_path, "rb") as f:
                content = await f.read()
            headers.update({"Content-Type": mime_type})
            url = url + "/" + os.path.basename(file_path)
            response = await self.request(
                method="PUT",
                url=url,
                headers={**self.client.headers, **headers},
                data=content,
                **kwargs,
            )
            response.raise_for_status()
            return response

    async def upload_files(
        self,
        files: Optional[List[str]] = None,
        folder: Optional[str] = None,
        url: str = "",
        use_multipart: bool = False,
        max_retries: int = 3,
        interval: float = 0.2,
        headers: Optional[Dict[str, str]] = None,
    ):
        """
        批量上传文件到 DUFS 或任意 HTTP 上传接口

        Args:
            files: 指定文件列表
            folder: 指定文件夹，将上传该目录下所有文件
            url: 上传目标 URL 目录（每个文件名会拼接到 url）
            use_multipart: 是否使用 multipart/form-data
            max_retries: 单个文件最大重试次数
            interval: 文件间延迟
            headers: 额外请求头
        """
        headers = headers or {}
        files_to_upload = []

        if self.client is None:
            await self.start()

        # 1. 文件列表
        if files:
            files_to_upload.extend([f for f in files if os.path.isfile(f)])

        # 2. 文件夹扫描
        if folder and os.path.isdir(folder):
            for root, _, filenames in os.walk(folder):
                for filename in filenames:
                    files_to_upload.append(os.path.join(root, filename))

        if not files_to_upload:
            raise ValueError("没有找到可上传的文件")

        for file_path in files_to_upload:
            filename = os.path.basename(file_path)
            # url = f"{url.rstrip('/')}/{filename}"  # 拼接完整 URL
            url = f"{url.rstrip('/')}"

            for attempt in range(1, max_retries + 1):
                try:
                    await self.upload_file(
                        file_path=file_path,
                        url=url,
                        headers=headers,
                        use_multipart=use_multipart,
                    )
                    log.success(f"{filename} 上传成功")
                    break
                except Exception as e:
                    log.error(f"{filename} 上传失败，第 {attempt} 次重试")
                    if attempt == max_retries:
                        log.error(f"{filename} 最终上传失败")
                    else:
                        await asyncio.sleep(interval)
            await asyncio.sleep(interval)

    async def download_file(
        self,
        url: str,
        file_path: str,
        headers: Optional[Dict[str, str]] = None,
        **kwargs,
    ) -> None:
        """
        下载文件到本地

        Args:
            url: 文件URL
            file_path: 本地文件路径
            headers: 请求头
            **kwargs: 其他参数
        """
        response = await self.get(url, headers=headers, **kwargs)

        with open(file_path, "wb") as f:
            async for chunk in response.aiter_bytes():
                f.write(chunk)

    async def batch_request(
        self,
        requests: List[Dict[str, Any]],
        max_concurrent: int = 10,
        interval: float = 0.2,
    ) -> List[Any]:
        semaphore = asyncio.Semaphore(max_concurrent)

        async def limited_request(request_args):
            url = request_args.get("url")
            try:
                async with semaphore:
                    if interval > 0:
                        await asyncio.sleep(interval)
                    result = await self.request(**request_args)
                    log.success(f"{url} 请求成功")
                    return result
            except Exception as e:
                log.error(f"{url} 请求失败")
                raise

        tasks = [limited_request(req) for req in requests]
        return await asyncio.gather(*tasks, return_exceptions=True)
