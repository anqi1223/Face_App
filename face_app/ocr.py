"""OCR 文字提取 - 基于 RapidOCR (ONNX Runtime)"""
from pathlib import Path
from rapidocr_onnxruntime import RapidOCR

from .utils import imread_unicode


def _get_elapse(elapse):
    """RapidOCR 返回的 elapse 可能是列表/None，统一转为总耗时（秒）"""
    if elapse is None:
        return 0.0
    if isinstance(elapse, (list, tuple)):
        return float(elapse[-1]) if len(elapse) > 0 else 0.0
    return float(elapse)


def init_ocr_engine():
    """初始化 OCR 引擎"""
    print("正在加载 RapidOCR 模型...")
    engine = RapidOCR()
    print("OCR 模型加载完成！")
    return engine


def extract_text_from_image(engine, img_path):
    """
    从单张图片提取文字。
    返回 (text_lines, raw_result, elapse)

    - text_lines: 提取到的文字列表（按从上到下排列）
    - raw_result: RapidOCR 原始输出
    - elapse: 耗时（秒）
    """
    img = imread_unicode(str(img_path))
    if img is None:
        return [], None, 0

    result, elapse = engine(img)
    elapse = _get_elapse(elapse)

    if result is None:
        return [], None, elapse

    # result: list of [box, text, confidence]
    # 按 y 坐标从上到下排序
    result = sorted(result, key=lambda r: r[0][0][1])

    text_lines = []
    for box, text, conf in result:
        text_lines.append({"文字": text, "置信度": round(conf, 4)})

    return text_lines, result, elapse


def extract_text_from_directory(engine, directory: Path):
    """
    对目录下所有图片进行 OCR 文字提取。
    返回结果列表。
    """
    if not directory.exists():
        print(f"❌ 目录不存在: {directory}")
        return []

    image_files = sorted([
        f for f in directory.glob("*.*")
        if f.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
    ])

    if not image_files:
        print(f"⚠️ {directory} 中没有图片文件")
        return []

    print(f"\n{'=' * 50}")
    print(f"文字提取中... 共 {len(image_files)} 张图片")
    print(f"{'=' * 50}")

    results = []

    for img_path in image_files:
        text_lines, _, elapse = extract_text_from_image(engine, img_path)

        if not text_lines:
            print(f"⚠️ 未检测到文字: {img_path.name}  (耗时: {elapse:.2f}s)")
            results.append({
                "图片": img_path.name,
                "提取文字": "",
                "置信度": 0.0,
                "耗时(秒)": round(elapse, 3),
            })
            continue

        combined_text = " | ".join(t["文字"] for t in text_lines)
        avg_conf = round(sum(t["置信度"] for t in text_lines) / len(text_lines), 4)

        print(f"✅ {img_path.name}: {combined_text}  (置信度: {avg_conf}, 耗时: {elapse:.2f}s)")

        for t in text_lines:
            results.append({
                "图片": img_path.name,
                "提取文字": t["文字"],
                "置信度": t["置信度"],
                "耗时(秒)": round(elapse, 3),
            })

    print(f"{'=' * 50}\n")
    return results
