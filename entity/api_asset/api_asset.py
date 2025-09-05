from dataclasses import dataclass
from typing import Any, Dict, List, Optional

@dataclass
class ApiAssetRecord:
    id: int
    http_authority: str
    http_path: str
    http_request_method_id: int
    api_protocol_id: int
    merger_rule_id: int
    offline_sign: int
    version: str
    app_name: str
    app_id: int
    address: str
    app_icon_file_path: str
    call_count: int
    today_call_count: int
    risk_level_id: int
    vul_level_id: int
    api_sens_level_id: int
    created_at: str 
    latest_access_time: str 
    src_ip_config_ids: List[int]
    api_labels: List[Any]  
    asset_label_names: List[Any]  
    data_labels: Optional[Any]  
    request_data_assets: List[Any]  
    response_data_assets: List[Any]  
    asset_source_name: str
    data_source_type: int
    active_id: int
    # risk_count: int
    # vul_count: int
    

@dataclass
class ApiAssetLabelDetail:
    request: Dict[str, Any]
    response: Dict[str, Any]
    storage_state: int
