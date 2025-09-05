
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class FileAssetRecord:
    """文件资产类"""
    id: int
    name: str
    format: str
    md5: str
    app_id: int
    app_name: str
    icon_file_path: str
    size: int
    sens_level_ids: int
    upload_count: int
    download_count: int
    data_labels: List[str]
    create_time: str  # 可以改为 datetime 类型，如果需要进行时间操作
    latest_access_time: str  
    file_type: Optional[str] = None