# -*- coding: utf-8 -*-
"""
路径工具 - Mac 优化
统一处理输出目录创建与文件写入，避免 FileNotFoundError。
"""

import os


def ensure_dir(path):
    """
    确保目录存在，若不存在则创建（含父目录）。
    使用 os.makedirs 保证在 Mac 上可靠。
    """
    abs_path = os.path.abspath(os.path.normpath(path))
    os.makedirs(abs_path, exist_ok=True)
    return abs_path


def ensure_parent_and_open(filepath, mode='w', encoding=None, **kwargs):
    """
    确保文件父目录存在后打开文件。
    返回 (abs_path, file_handle)。
    """
    abs_path = os.path.abspath(os.path.normpath(filepath))
    parent = os.path.dirname(abs_path)
    os.makedirs(parent, exist_ok=True)
    if encoding:
        kwargs['encoding'] = encoding
    f = open(abs_path, mode, **kwargs)
    return abs_path, f


def resolve_path(path):
    """返回绝对路径字符串。"""
    return os.path.abspath(os.path.normpath(path))
