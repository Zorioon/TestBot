import json
from multiprocessing import Value
from typing_extensions import runtime
from urllib import response
from utils.crypto_tools.crypto_utils import md5enc, rsa_encrypt
from utils.ocr_tools.ocr_utils import recognize_captcha_from_base64
from utils.request_tools.async_http_client import AsyncHttpClient
from utils.log_tools.logger_utils import get_logger

log = get_logger(__name__)


class AuthUtils:
    """
    登录认证方法
    """

    @staticmethod
    async def login(
        https_req: AsyncHttpClient,
        username: str = "admin",
        password: str = "root123.",
        max_retries: int = 5,
    ):
        """
        获取 token
        """
        public_key = """-----BEGIN PUBLIC KEY-----
        MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDcu1sLod7mIz0EaYW7iM/glFNL
        kTFI5n87pFW/0Xv2UFUiPoFKiBagZ0NsBtPTKzFFimmqEdbj0W0O7wwoQ1bupTo8
        1qYm1EJ+Qc3REzmPyEJn9wof7vHvSlNdcIff6wJOOZ+Vqq08qK4p9HG73/8oKgVx
        Nw4cEJUnmqUqtAP31wIDAQAB
        -----END PUBLIC KEY-----"""  # 公钥

        for attempt in range(max_retries):
            try:
                # 1、获取登录随机字符串
                rand_resp = await https_req.get("/api/v1.2/randString")
                rand_str = rand_resp["data"]["rand"]

                # 2、获取验证码信息
                captcha_resp = await https_req.get("/api/v1.2/captcha")
                captcha_id = captcha_resp["data"]["id"]
                captcha_code = recognize_captcha_from_base64(
                    captcha_resp["data"]["captcha"]
                )

                # 3、拼装登录信息
                encrypted_data = rsa_encrypt(
                    json.dumps(
                        {
                            "username": username,
                            "password": md5enc(password),
                            "uuid": rand_str,
                            "captcha": captcha_code,
                            "captchaId": captcha_id,
                        }
                    ),
                    public_key,
                )

                login_resp = await https_req.post(
                    "/api/v1.2/login", json={"info": encrypted_data}
                )

                if login_resp["code"] == 200:
                    log.info(f"登录成功")
                    return login_resp["data"]["token"]

                raise ValueError(f"登录失败: {response}")

            except (KeyError, ValueError) as parse_err:
                # 数据解析/接口返回异常 → warning
                log.warning(f"业务逻辑错误 (第 {attempt} 次): {parse_err}")

            except Exception as net_err:
                # 网络错误/请求异常等 → error
                log.error(f"系统异常 (第 {attempt} 次)")

        raise RuntimeError(f"登录失败，已尝试 {max_retries} 次")
