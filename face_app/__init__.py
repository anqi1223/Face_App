"""人脸识别引擎包

采用延迟加载（PEP 562 __getattr__）：
- 表生成模块（generate_table）只依赖 pandas/openpyxl，不应被 insightface 等重依赖拖累；
- 未安装 insightface / rapidocr 时，仅当真正用到对应函数才报错。
"""

__all__ = [
    "init_face_app",
    "enroll_faces",
    "recognize_faces",
    "save_results_to_excel",
    "init_ocr_engine",
    "extract_text_from_directory",
    "get_table1",
    "get_table2",
    "get_table3",
    "get_table4",
    "get_table5",
    "get_table6",
    "get_table7",
    "get_table8_9_10",
]

# 名称 → 所在子模块
_LAZY_MODULES = {
    "init_face_app": "face_engine",
    "enroll_faces": "face_engine",
    "recognize_faces": "face_engine",
    "save_results_to_excel": "face_engine",
    "init_ocr_engine": "ocr",
    "extract_text_from_directory": "ocr",
    "get_table1": "generate_table",
    "get_table2": "generate_table",
    "get_table3": "generate_table",
    "get_table4": "generate_table",
    "get_table5": "generate_table",
    "get_table6": "generate_table",
    "get_table7": "generate_table",
    "get_table8_9_10": "generate_table",
}


def __getattr__(name):
    """按需导入子模块并返回对应属性。"""
    if name in _LAZY_MODULES:
        import importlib

        mod = importlib.import_module(f"{__name__}.{_LAZY_MODULES[name]}")
        return getattr(mod, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
