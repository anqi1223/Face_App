"""图片 I/O 工具 - 解决 Windows 上 OpenCV 不支持中文/Unicode 路径的问题"""
import cv2
import numpy as np
from pathlib import Path


def imread_unicode(filepath):
    """使用 np.fromfile + cv2.imdecode 读取图片，支持中文/Unicode 路径"""
    try:
        data = np.fromfile(filepath, dtype=np.uint8)
        img = cv2.imdecode(data, cv2.IMREAD_COLOR)
        return img
    except Exception:
        return None


def imwrite_unicode(filepath, img):
    """使用 cv2.imencode + Python 原生写入，支持中文/Unicode 路径"""
    ext = Path(filepath).suffix
    success, encoded = cv2.imencode(ext, img)
    if success:
        with open(filepath, "wb") as f:
            f.write(encoded.tobytes())
        return True
    return False
