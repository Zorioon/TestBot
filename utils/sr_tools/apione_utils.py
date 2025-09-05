import asyncio
import json
import re
from typing import Any, Dict, List, Optional, Union
from xmlrpc.client import Boolean

import test
from entity.api_asset.api_asset import ApiAssetLabelDetail, ApiAssetRecord
from entity.file_asset.file_asset import FileAssetRecord
from utils.decorator_tools.decorator_utils import async_retry_on_empty, transform_to_data_class
from utils.request_tools.async_http_client import AsyncHttpClient
from utils.log_tools.logger_utils import get_logger
log = get_logger(__name__)

class ApioneUtils:
    """Apione 通用方法"""

    @staticmethod
    async def initial_rule(
        https_req: AsyncHttpClient, specification_id: int, timeout: int = 15
    ):
        """初始化系统规则（风险、弱点、数据标签规则）

        Args:
            https_req (AsyncHttpClient): _description_
            specification_id (int): _description_
        """
        init_resp = await https_req.post(
            "/apione/v2/initial/rules",
            json={"specification_name": "", "specification": specification_id},
        )
        if init_resp.get("code", -1) != 200:
            raise RuntimeError(f"初始化失败: {init_resp.get('message', '未知错误')}")

        async def listen_init_progress():
            while True:
                pro_resp = await https_req.get("/apione/v2/initial/progress")
                if pro_resp["code"] == 200 and pro_resp.get("data", {}).get(
                    "finish_tag"
                ):
                    return True
                await asyncio.sleep(0.5)

        return await asyncio.wait_for(listen_init_progress(), timeout=timeout)
    
    @staticmethod
    async def update_auto_merge_config(https_req: AsyncHttpClient, turn_on: Boolean = False) -> None:
        """修改自动合并状态

        Args:
            turn_on (Boolean, optional): False 代表禁用， True代表启用. Defaults to False.

        Raises:
            RuntimeError: _description_
            RuntimeError: _description_
            RuntimeError: _description_

        Returns:
            _type_: _description_
        """
        response = await https_req.put("/apione/v2/merger/auto-merge-config/update", json={
            "threshold": 0,
            "turn_on": turn_on
        })
        if response.get("code") != 200:
            raise RuntimeError(f"{'开启' if turn_on else '关闭'} 自动合并状态失败")
        log.success(f"{'开启' if turn_on else '关闭'} 自动合并状态成功")
    
    @staticmethod
    @async_retry_on_empty(check=lambda x: x)
    async def is_file_asset_count_equal_expected(https_req: AsyncHttpClient, expected_file_asset_count: int):
        """验证入库的文件资产是否符合预期

        Args:
            file_asset_expected_count (int): _description_

        Raises:
            RuntimeError: _description_
            RuntimeError: _description_
            RuntimeError: _description_
            RuntimeError: _description_
            RuntimeError: _description_

        Returns:
            _type_: _description_
        """
        response = await https_req.post("/apione/v2/file-assets", json={"time_layout":"2006-01-02 15:04:05","page_num":1,"page_size":10})
        if response["code"] != 200:
            raise RuntimeError("获取文件资产记录失败") 
        real_file_asset_count = response["data"]["row_count"]
        return real_file_asset_count == expected_file_asset_count
    
    
    @staticmethod
    @async_retry_on_empty(check=lambda api_asset: api_asset)
    @transform_to_data_class(ApiAssetRecord)
    async def get_api_asset_record(https_req: AsyncHttpClient, api: str) -> Optional[ApiAssetRecord]:
        """获取API记录

        Args:
            https_req (AsyncHttpClient): _description_
            api (str): _description_
        """
        response = await https_req.post(
            "/apione/v2/assets/list",
            json={
                "api": api,
                "page_num": 1,
                "page_size": 10,
                "time_layout": "2006-01-02 15:04:05",
            },
        )
        if response.get("code") != 200:
            raise RuntimeError("获取API资产记录失败")
        api_asset = response.get("data", {}).get("results")
        return api_asset and api_asset[0]
    
    def convert_to_dict_list_old(contents: List[str]) -> List[Dict[str, Any]]:
        """
        将 ["\"key\": \"value\"", ...] 转换为 [ {"key": "value"}, {"key": "value"}, ... ]
        - 每条都保留，不做 key 合并
        - 其他内容（如 IPv6 地址）放成 ["原始内容"]
        """
        result = []

        # 严格匹配 "key": "value"
        pattern = re.compile(r'^"\s*([^"]+)\s*"\s*:\s*"\s*([^"]+)\s*"$')

        for item in contents:
            item = item.strip()
            if not item:
                continue

            match = pattern.match(item)
            if match:
                key, value = match.groups()
                result.append({key: value})
            else:
                # 非 key:value 格式，比如 IPv6
                result.append(item)

        return result

    def convert_to_dict_list(contents: List[str]) -> List[Dict[str, Any]]:
        """
        将混合格式数据转换为字典列表
        - key:value 格式转换为 {key: value}
        - 其他（如 IPv6）保留原始字符串
        """
        result = []

        # 匹配 key:value，key 可以是字母、下划线、中文，value 任意内容
        kv_pattern = re.compile(r'^\s*([\w\u4e00-\u9fff_]+)"?\s*:\s*"?\s*(.+?)\s*"?$')

        for item in contents:
            item = item.strip()
            # 去掉外层单/双引号
            if (item.startswith("'") and item.endswith("'")) or (item.startswith('"') and item.endswith('"')):
                item = item[1:-1]

            if not item:
                continue

            # 判断冒号数量，如果多于1个，认为是 IPv6 或非 key:value
            if item.count(':') > 1:
                result.append(item)
                continue

            # 尝试匹配 key:value
            match = kv_pattern.match(item)
            if match:
                key, value = match.groups()
                result.append({key: value})
            else:
                result.append(item)

        return result
    
    @staticmethod
    @transform_to_data_class(ApiAssetLabelDetail)
    async def get_api_asset_label_detail(https_req: AsyncHttpClient, api_id: int) -> Dict[str, Any]:
        """获取 API 资产详情 + 原始请求响应细节"""

        # 1. 获取 API 资产详情
        detail_resp = await https_req.get(f"/apione/v2/assets/{api_id}/detail")
        if detail_resp.get("code") != 200:
            raise RuntimeError("获取API资产详情失败")

        latest_request_id = detail_resp.get("data", {}).get("latest_request_id")
        latest_storage_key = detail_resp.get("data", {}).get("latest_storage_key")

        # 2. 获取原始调用详情
        raw_detail_resp = await https_req.get(
            f"/apione/v2/call/records/{latest_request_id}/unmask?storage_key={latest_storage_key}"
        )
        if raw_detail_resp.get("code") != 200:
            raise RuntimeError("获取API接口详情失败")

        raw_data = raw_detail_resp.get("data", {})

        # 3. 处理 request/response label
        request_label_resp = raw_data.get("request", {}).get("label", {})
        response_label_resp = raw_data.get("response", {}).get("label", {})
        request_label = {
            "start_line": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in request_label_resp.get("start_line")
            }
            if request_label_resp.get("start_line")
            else None,
            "headers": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in request_label_resp.get("headers")
            }
            if request_label_resp.get("headers")
            else None,
            "body": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in request_label_resp.get("body")
            }
            if request_label_resp.get("body")
            else None,
        }

        response_label = {
            "start_line": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in response_label_resp.get("start_line")
            }
            if response_label_resp.get("start_line")
            else None,
            "headers": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in response_label_resp.get("headers")
            }
            if response_label_resp.get("headers")
            else None,
            "body": {
                item.get("name"): {"count": item.get("count"), "contents": ApioneUtils.convert_to_dict_list(item.get("contents"))}
                for item in response_label_resp.get("body")
            }
            if response_label_resp.get("body")
            else None,
        }

        # 4. 返回结果
        return {
            "storage_state": raw_data.get("storage_state"),
            "request": request_label,
            "response": response_label,
        }

    @staticmethod
    @async_retry_on_empty(check=lambda file_asset: file_asset, target_param="file_name")
    @transform_to_data_class(FileAssetRecord)
    async def get_file_asset_record(https_req: AsyncHttpClient, file_name: str, file_md5: Optional[str] = None) -> Optional[FileAssetRecord]:
        """获取文件记录

        Args:
            https_req (AsyncHttpClient): _description_
            api (str): _description_
        """
        response = await https_req.post(
            "/apione/v2/file-assets",
            json={
                "time_layout": "2006-01-02 15:04:05",
                "name": file_name,
                "md5": file_md5,
                "page_num": 1,
                "page_size": 10
            },
        )
        if response.get("code") != 200:
            raise RuntimeError("获取文件资产记录失败")
        file_asset = response.get("data", {}).get("results")
        return file_asset and file_asset[0]
    
    @staticmethod
    # @async_retry_on_empty(retries=4, interval=0.2, check=lambda file_data_label_detail: file_data_label_detail)
    async def get_file_asset_label_detail(https_req: AsyncHttpClient, file_id: int):
        """获取文件匹配数据标签详情

        Args:
            https_req (AsyncHttpClient): _description_
            file_id (int): _description_
        """
        # 1. 获取 文件 资产详情
        detail_resp = await https_req.get(f"apione/v2/file-assets/{file_id}/data-count/rank?period=7d&interval=1d&page_size=10&page_num=1")
        if detail_resp.get("code") != 200:
            raise RuntimeError("获取文件资产详情失败")
        data_labels = detail_resp.get("data", {}).get("results")
        file_data_label_detail = {
            data_label.get("data_label"): data_label.get("data_count")
            for data_label in data_labels
        }
        return file_data_label_detail
        