import asyncssh

class AsyncSSHClient:
    def __init__(self, host, username, password):
        self.host = host
        self.username = username
        self.password = password
        self.client = None

    async def connect(self):
        """异步连接SSH"""
        self.client = await asyncssh.connect(
            host=self.host,
            username=self.username,
            password=self.password,
            known_hosts=None 
        )
        return self.client

    async def close(self):
        """关闭连接"""
        if self.client:
            self.client.close()

    async def __aenter__(self):
        """支持 async with"""
        await self.connect()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """退出时自动关闭"""
        await self.close()

