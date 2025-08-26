import os
import sys
import base64
from ddddocr import DdddOcr
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)


def recognize_captcha_from_base64(base64_data):
    """
    识别base64，转换成具体验证码
    
    参数:
        base64_data: 包含或纯base64编码的图片数据
    """
    pure_base64 = base64_data.split(",")[1] if "," in base64_data else base64_data
    
    try:
        image_data = base64.b64decode(pure_base64)
        
        with open(os.devnull, 'w') as f:
            sys.stdout = f
            ocr = DdddOcr()
            sys.stdout = sys.__stdout__
        
        result = ocr.classification(image_data)
        return result
        
    except Exception:
        log.exception(f"识别验证码失败")
        raise