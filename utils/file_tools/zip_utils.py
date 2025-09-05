import os
import zipfile
import tempfile
import shutil
from typing import List, Optional, Union, Tuple
from pathlib import Path

class ZipUtils:
    """ZIP压缩解压工具类"""
    
    @staticmethod
    def zip_files(
        files: Union[List[str], str],
        output_path: str,
        *,
        base_dir: Optional[str] = None,
        compression: int = zipfile.ZIP_DEFLATED,
        exclude_extensions: Optional[List[str]] = None,
        exclude_dirs: Optional[List[str]] = None
    ) -> str:
        """
        压缩文件/目录
        
        Args:
            files: 文件路径列表或目录路径
            output_path: 输出的ZIP文件路径
            base_dir: 基础目录（用于保留相对路径结构）
            compression: 压缩算法 (zipfile.ZIP_STORED 或 zipfile.ZIP_DEFLATED)
            exclude_extensions: 要排除的文件扩展名 ['.log', '.tmp']
            exclude_dirs: 要排除的目录名 ['.git', '__pycache__']
            
        Returns:
            生成的ZIP文件路径
            
        Raises:
            ValueError: 输入路径无效时抛出
        """
        exclude_extensions = exclude_extensions or []
        exclude_dirs = exclude_dirs or []
        
        if isinstance(files, str):
            if os.path.isdir(files):
                return ZipUtils._zip_dir(
                    files, output_path, base_dir, compression, 
                    exclude_extensions, exclude_dirs
                )
            files = [files]
        
        if not files:
            raise ValueError("文件列表不能为空")
            
        with zipfile.ZipFile(output_path, 'w', compression) as zipf:
            for file_path in files:
                if not os.path.exists(file_path):
                    continue
                    
                if os.path.isdir(file_path):
                    raise ValueError("请使用目录模式压缩文件夹")
                
                if any(file_path.endswith(ext) for ext in exclude_extensions):
                    continue
                    
                arcname = (
                    os.path.relpath(file_path, base_dir) 
                    if base_dir 
                    else os.path.basename(file_path)
                )
                zipf.write(file_path, arcname=arcname)
        
        return output_path

    @staticmethod
    def _zip_dir(
        dir_path: str,
        output_path: str,
        base_dir: Optional[str],
        compression: int,
        exclude_extensions: List[str],
        exclude_dirs: List[str]
    ) -> str:
        """内部方法：处理目录压缩"""
        if not os.path.isdir(dir_path):
            raise ValueError(f"目录不存在: {dir_path}")
            
        base_dir = base_dir or dir_path
        
        with zipfile.ZipFile(output_path, 'w', compression) as zipf:
            for root, dirs, files in os.walk(dir_path):
                # 过滤排除目录
                dirs[:] = [d for d in dirs if d not in exclude_dirs]
                
                for file in files:
                    if any(file.endswith(ext) for ext in exclude_extensions):
                        continue
                        
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, base_dir)
                    zipf.write(file_path, arcname=rel_path)
        
        return output_path

    @staticmethod
    def create_temp_zip(
        source: Union[str, List[str]],
        prefix: str = "temp_zip_",
        **kwargs
    ) -> Tuple[str, str]:
        """
        创建临时ZIP文件（需手动清理）
        
        Args:
            source: 文件路径列表或目录路径
            prefix: 临时文件前缀
            kwargs: 传递给zip_files的参数
            
        Returns:
            (zip_path, temp_dir) 临时ZIP路径和临时目录（需要手动清理）
        """
        temp_dir = tempfile.mkdtemp(prefix="zip_temp_")
        zip_path = os.path.join(temp_dir, f"{prefix}{os.urandom(4).hex()}.zip")
        
        try:
            ZipUtils.zip_files(source, zip_path, **kwargs)
            return zip_path, temp_dir
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise

    @staticmethod
    def unzip(
        zip_path: str,
        output_dir: str,
        *,
        password: Optional[str] = None,
        overwrite: bool = True,
        preserve_permissions: bool = False
    ) -> List[str]:
        """
        解压ZIP文件
        
        Args:
            zip_path: ZIP文件路径
            output_dir: 解压目录
            password: 解压密码
            overwrite: 是否覆盖已存在文件
            preserve_permissions: 是否保留文件权限（Unix系统）
            
        Returns:
            解压出的文件路径列表
        """
        if not os.path.isfile(zip_path):
            raise ValueError(f"ZIP文件不存在: {zip_path}")
            
        os.makedirs(output_dir, exist_ok=True)
        extracted_files = []
        
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            if password:
                zipf.setpassword(password.encode('utf-8'))
                
            for file in zipf.namelist():
                try:
                    # 安全校验：防止ZIP炸弹或路径穿越攻击
                    target_path = os.path.join(output_dir, file)
                    if not target_path.startswith(os.path.abspath(output_dir) + os.sep):
                        continue
                        
                    if overwrite or not os.path.exists(target_path):
                        zipf.extract(file, output_dir)
                        extracted_files.append(target_path)
                        
                        if preserve_permissions and os.name == 'posix':
                            info = zipf.getinfo(file)
                            if info.external_attr > 0:
                                os.chmod(target_path, info.external_attr >> 16)
                except Exception as e:
                    continue
        
        return extracted_files

    @staticmethod
    def get_zip_contents(zip_path: str) -> List[str]:
        """获取ZIP文件内容列表"""
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            return zipf.namelist()