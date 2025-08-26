import base64
from Crypto.PublicKey import RSA
from Crypto.Cipher import PKCS1_v1_5 as PKCS1_cipher
import hashlib

def md5enc(in_str):
    """
    字符串MD5加密
    """
    md5 = hashlib.md5()
    md5.update(in_str.encode("utf8"))
    return md5.hexdigest()

def rsa_encrypt(msg: str, publickey, max_length=100):
    """校验RSA加密 使用公钥进行加密"""
    cipher = PKCS1_cipher.new(RSA.importKey(publickey))
    res_byte = bytes()
    for i in range(0, len(msg), max_length):
        res_byte += cipher.encrypt(msg[i : i + max_length].encode("utf-8"))
    # cipher_text = base64.b64encode(cipher.encrypt(password.encode())).decode()
    return base64.b64encode(res_byte).decode("utf-8")