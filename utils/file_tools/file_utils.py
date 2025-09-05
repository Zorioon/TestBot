import hashlib
import os
from pathlib import Path
import shutil
from typing import List, Optional
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)


class FileUtils:
    """文件通用方法"""

    @staticmethod
    def find_file_from_root(
        sub_path: str, root: Optional[str] = None, create_if_not_exists: bool = False
    ) -> Optional[str]:
        """
        在项目根目录下查找文件，如果存在返回绝对路径，否则根据参数决定是否创建或返回None

        Args:
            sub_path (str): 相对路径，例如 'data/specifications.json'
            root (str, optional): 根路径，可选，默认自动查找项目根目录
            create_if_not_exists (bool): 如果路径不存在，是否创建该路径。默认为False

        Returns:
            str | None: 文件绝对路径或 None
        """
        # 1. 确定项目根目录
        if root is None:
            cur_dir = Path(__file__).resolve()
            for parent in cur_dir.parents:
                # 判断条件可根据项目实际情况改，例如找 .git 或 pytest.ini
                if (parent / ".git").exists() or (parent / "pytest.ini").exists():
                    root = parent
                    break
            else:
                root = Path.cwd()  # 找不到就用当前工作目录

        root = Path(root)
        target_path = root / sub_path

        if target_path.exists():
            return str(target_path.resolve())

        # 如果路径不存在且允许创建
        if create_if_not_exists:
            try:
                # 如果是文件路径，创建父目录
                if not target_path.suffix:  # 没有后缀，可能是目录
                    target_path.mkdir(parents=True, exist_ok=True)
                else:  # 有后缀，可能是文件
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    # 如果是文件且需要创建空文件，可以取消下面的注释
                    # target_path.touch(exist_ok=True)

                return str(target_path.resolve())
            except Exception as e:
                raise ValueError(f"创建路径失败: {sub_path}, 错误: {e}")

        raise ValueError(f"未找到路径 {sub_path}")

    @staticmethod
    def remove_all_file_in_folder(file_path: str):
        """删除文件夹下的所有内容(文件、符号链接、文件夹)

        Args:
            file_path (str): _description_
        """
        if os.path.exists(file_path):  # 首先检查目录是否存在
            try:
                # 遍历目录中的所有内容
                for filename in os.listdir(file_path):
                    # 构建完整的文件/目录路径
                    file_path_to_remove = os.path.join(file_path, filename)

                    try:
                        # 如果是文件或符号链接，直接删除
                        if os.path.isfile(file_path_to_remove) or os.path.islink(
                            file_path_to_remove
                        ):
                            os.unlink(file_path_to_remove)  # 删除文件或链接

                        # 如果是目录，递归删除整个目录树
                        elif os.path.isdir(file_path_to_remove):
                            shutil.rmtree(file_path_to_remove)  # 删除整个目录

                    # 处理单个文件/目录删除时的异常
                    except Exception as e:
                        log.error(f"删除 {file_path_to_remove} 失败")
                        raise ValueError(f"删除 {file_path_to_remove} 失败") from e

                # 所有内容删除成功
                log.success(f"已清空目录: {file_path}")

            # 处理整个清空过程的异常
            except Exception as e:
                log.error(f"清空目录 {file_path} 失败")
                raise

    @staticmethod
    def calculate_file_md5(file_path: str, chunk_size: int = 8192) -> str:
        """计算文件的 MD5 值"""
        md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            while chunk := f.read(chunk_size):
                md5.update(chunk)
        return md5.hexdigest()

    @staticmethod
    def get_all_files(
        folder_path: str,
        file_name: Optional[str] = None,
        recursive: bool = False,
    ) -> List[str]:
        """
        获取指定文件夹下的所有文件路径

        :param folder_path: 文件夹路径
        :param recursive: 是否递归获取子文件夹中的文件
        :param file_name: 可选文件名（支持通配符），用于过滤结果
        :return: 文件路径列表
        """
        if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
            raise ValueError(f"路径无效或不是文件夹: {folder_path}")

        file_paths = []

        if recursive:
            # 递归遍历
            for root, _, files in os.walk(folder_path):
                file_paths.extend(os.path.join(root, f) for f in files)
        else:
            # 非递归遍历
            file_paths = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
            ]

        if file_name:
            file_paths = [
                path for path in file_paths if file_name in os.path.basename(path)
            ]

        return file_paths
