import asyncio
from utils.ssh_tools.ssh_connect import AsyncSSHClient
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)

class SSHOperation:
    """远程连接操作
    """
    
    @staticmethod
    async def exec_single_command(ssh_client: AsyncSSHClient, command: str) -> str:
        """执行单条命令

        Args:
            ssh_client (AsyncSSHClient): 异步ssh对象
            command (str): 命令
        """
        result = await ssh_client.client.run(command)
        if result.exit_status != 0:
            raise RuntimeError(
                f"Command failed: {command}\n"
                f"stderr: {result.stderr}\n"
                f"exit_code: {result.exit_status}"
            )
        return True
    
    @staticmethod
    async def check_process_log(ssh_client: AsyncSSHClient, process_name: str, keyword: str, timeout: int = 180):
        """检查进程是否有指定日志输出，表示进程已正式启动

        Args:
            ssh_client (AsyncSSHClient): _description_
            process_name (str): _description_
            keyword (_type_, optional): _description_. Defaults to " http server listening at address=127.0.0.1:29300".
            timeout (int, optional): _description_. Defaults to 60.
        """
        async def watch_process_log():
            cmd = f"tail -F /data/{process_name}/log/{process_name}.log"
            # cmd = f"journalctl -fu {process_name}"
            async with ssh_client.client.create_process(cmd) as process:
                while True:
                    line = await process.stdout.readline()
                    if not line:
                        # 防止空转
                        await asyncio.sleep(0.1)
                        continue
                    if keyword in line:
                        return True
        return await asyncio.wait_for(watch_process_log(), timeout=timeout)
    
    
    