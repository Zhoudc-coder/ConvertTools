# utils.py
import re
import os
from pathlib import Path

def sanitize_filename(name: str, max_length: int = 200) -> str:
    """
    清洗文件名中的非法字符（Windows/macOS 通用），并截断长度。
    """
    # 替换非法字符
    illegal_chars = r'[\\/:*?"<>|]'
    name = re.sub(illegal_chars, '_', name)
    # 去除首尾空格和点号（Windows 不允许结尾点号）
    name = name.strip().rstrip('.')
    # 空字符串则使用默认名
    if not name:
        name = "unnamed"
    # 处理 Windows 保留名称
    reserved = {'CON', 'PRN', 'AUX', 'NUL',
                'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
                'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'}
    if name.upper() in reserved:
        name = f"_{name}"
    # 截断长度，保留扩展名（调用方会加 .xlsx）
    if len(name) > max_length:
        name = name[:max_length]
    return name

def detect_csv_encoding(file_path: Path):
    """
    尝试常用编码读取 CSV 文件，返回第一个成功的编码。
    """
    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030', 'latin-1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                # 读取一行测试
                f.readline()
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    # 如果全部失败，抛出异常
    raise UnicodeDecodeError(f"无法识别 CSV 文件编码，请手动指定。尝试过的编码：{encodings}")