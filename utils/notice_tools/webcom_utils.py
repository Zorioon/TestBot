import httpx
import asyncio
from typing import List, Dict, Optional
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)

class WeComRobot:
    def __init__(self, webhook_key: str):
        """
        初始化企业微信机器人
        
        :param webhook_key: 机器人的webhook key
        """
        self.webhook_key = webhook_key
        self.send_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={webhook_key}"
        self.upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={webhook_key}&type=file"
        self.client = httpx.AsyncClient()

    async def send_message(self, message_type: str, content: Dict) -> Dict:
        """
        发送消息（通用方法）
        
        :param message_type: 消息类型 (text, markdown, image, file, news)
        :param content: 消息内容
        :return: 发送结果
        """
        payload = {"msgtype": message_type, message_type: content}
        
        try:
            response = await self.client.post(self.send_url, json=payload)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            return {"errcode": -1, "errmsg": str(e)}

    async def send_text(self, content: str, mentioned_list: List[str] = None) -> Dict:
        """发送文本消息"""
        data = {"content": content}
        if mentioned_list:
            data["mentioned_list"] = mentioned_list
        return await self.send_message("text", data)

    async def send_markdown(self, content: str) -> Dict:
        """发送markdown消息"""
        return await self.send_message("markdown", {"content": content})

    async def upload_file(self, file_path: str, filename: Optional[str] = None) -> Optional[str]:
        """上传文件并返回media_id"""
        if filename is None:
            filename = file_path.split("/")[-1]
        
        try:
            with open(file_path, "rb") as f:
                files = {"file": (filename, f, "application/octet-stream")}
                response = await self.client.post(self.upload_url, files=files)
                response.raise_for_status()
                result = response.json()
                
                if result["errcode"] == 0:
                    return result["media_id"]
                else:
                    print(f"文件上传失败: {result['errmsg']}")
                    return None
        except Exception as e:
            print(f"上传文件时出错: {e}")
            return None

    async def send_file(self, file_path: str) -> Dict:
        """发送文件消息"""
        media_id = await self.upload_file(file_path)
        if not media_id:
            return {"errcode": -1, "errmsg": "文件上传失败"}
        
        return await self.send_message("file", {"media_id": media_id})

    async def send_multiple_files(self, file_paths: List[str]) -> List[Dict]:
        """发送多个文件"""
        results = []
        for file_path in file_paths:
            result = await self.send_file(file_path)
            results.append({"file": file_path, "result": result})
            await asyncio.sleep(0.5)  # 避免频繁请求
        return results

    async def close(self):
        """关闭HTTP客户端"""
        await self.client.aclose()