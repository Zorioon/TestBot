# !/usr/bin/env python
# -*- coding: utf-8 -*-

"""
-------------------------------------------------
   File Name：     yamlutil
   Description :   
   Author :       崔术森
   date：          2024/10/15
-------------------------------------------------
   Change Activity:
                   2024/10/15 9:09: 
-------------------------------------------------
"""
__author__ = '崔术森'

import os
from ruamel.yaml import YAML

class YAMLUtil:
    # 共享同一个 YAML 实例（确保配置一致）
    _yaml = YAML()
    _yaml.preserve_quotes = True  # 保留引号
    _yaml.width = 4096           # 避免长行换行
    _yaml.indent(mapping=2, sequence=4, offset=2)  # 设置缩进

    @staticmethod
    def read_yaml(file_path):
        """
        读取 YAML 文件并保留注释。
        :param file_path: 文件路径
        :return: 带注释的字典（CommentedMap）
        :raises FileNotFoundError: 文件不存在
        :raises Exception: YAML 解析错误
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件未找到: {file_path}")

        with open(file_path, 'r', encoding='utf-8') as file:
            try:
                return YAMLUtil._yaml.load(file)  # 使用共享实例
            except Exception as exc:
                raise Exception(f"YAML 解析出错: {exc}")

    @staticmethod
    def write_yaml(data, file_path):
        """
        写入 YAML 文件并保留注释。
        :param data: 带注释的数据（CommentedMap）
        :param file_path: 输出文件路径
        """
        with open(file_path, 'w', encoding='utf-8') as file:
            YAMLUtil._yaml.dump(data, file)  # 使用共享实例


# 示例使用
if __name__ == "__main__":
    config = YAMLUtil.read_yaml("../../common/config.yaml")
    print(config.get('sc'))
