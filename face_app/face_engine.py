"""人脸识别引擎 - InsightFace 封装：初始化、录入、匹配、识别"""
import hashlib
import os
import pickle
import re

import numpy as np
import pandas as pd
from datetime import datetime
from pathlib import Path
from insightface.app import FaceAnalysis

from .utils import imread_unicode, imwrite_unicode
from .visualize import draw_face_annotations


# ============================================================
# 初始化
# ============================================================
# InsightFace 模型名（缓存校验依赖：更换模型后自动重新录入）
MODEL_NAME = "buffalo_l"

# 模型根目录：优先读环境变量 INSIGHTFACE_HOME（离线打包版用 runtime/insightface_home），
# 未设置时回退到默认 ~/.insightface（与旧行为一致）
MODEL_ROOT = os.environ.get("INSIGHTFACE_HOME") or "~/.insightface"


def init_face_app(providers=None):
    """初始化 InsightFace 人脸分析模型。providers 默认 CPU。"""
    if providers is None:
        providers = ["CPUExecutionProvider"]
    print(f"正在加载 InsightFace 模型... (providers: {providers})")
    app = FaceAnalysis(name=MODEL_NAME, providers=providers, root=MODEL_ROOT)
    app.prepare(ctx_id=0, det_size=DEFAULT_DET_SIZE)
    print(f"模型加载完成！检测阈值: {app.det_thresh}")
    return app


# ============================================================
# 稳健人脸检测
# ============================================================
# 默认检测分辨率 640（单人/小图快）；大图多人时把长边压到 640 会漏检远/小脸，
# 若一张脸都没检测到，就逐级提高到 1280/1920 重新检测（多人合照小脸补漏）。
DEFAULT_DET_SIZE = (640, 640)
ESCALATION_DET_SIZES = (1280, 1920)


def _set_det_size(app, det_size):
    """调整共享人脸模型的检测分辨率（仅改检测器输入尺寸，不重载权重）。"""
    app.prepare(
        ctx_id=0,
        det_thresh=getattr(app, "det_thresh", 0.5),
        det_size=det_size,
    )


def _bbox_iou(a, b):
    """两个检测框的 IoU（a/b 为 [x1,y1,x2,y2] 数组）。"""
    a = np.asarray(a, dtype=float)
    b = np.asarray(b, dtype=float)
    ax1, ay1, ax2, ay2 = a[:4]
    bx1, by1, bx2, by2 = b[:4]
    iw = min(ax2, bx2) - max(ax1, bx1)
    ih = min(ay2, by2) - max(ay1, by1)
    if iw <= 0 or ih <= 0:
        return 0.0
    inter = iw * ih
    area_a = (ax2 - ax1) * (ay2 - ay1)
    area_b = (bx2 - bx1) * (by2 - by1)
    return inter / (area_a + area_b - inter)


def _merge_faces(acc, new_faces, iou_thresh=0.4):
    """合并两次（不同分辨率）检测的结果：与已有框重叠的只保留，新框追加。"""
    merged = list(acc)
    for nf in new_faces:
        if not any(_bbox_iou(nf.bbox, mf.bbox) > iou_thresh for mf in merged):
            merged.append(nf)
    return merged


def detect_faces(app, img, escalation=True):
    """
    稳健人脸检测：先按当前 det_size 检测；若一张脸都没检测到且原图较大
    （说明被默认 640 分辨率压缩后远/小脸漏检），逐级提高分辨率重检，
    并把各分辨率检测到的人脸合并去重——不同分辨率能检出的脸可能不同。

    返回人脸列表（bbox 均在原图坐标系）。调用方应持有模型锁，避免与其他
    推理线程并发变更共享模型。结束后始终把检测分辨率还原为默认值。
    """
    faces = app.get(img)
    if faces or not escalation:
        return faces
    h, w = img.shape[:2]
    if max(h, w) <= DEFAULT_DET_SIZE[0]:
        return faces  # 原图本就不大，提高分辨率无意义
    try:
        for ds in ESCALATION_DET_SIZES:
            if ds <= max(h, w):
                _set_det_size(app, (ds, ds))
                batch = app.get(img)
                if batch:
                    faces = _merge_faces(faces, batch)
    finally:
        _set_det_size(app, DEFAULT_DET_SIZE)
    return faces


# ============================================================
# 人脸录入
# ============================================================
def _extract_name(img_path: Path) -> str:
    """从文件名提取人名：'张三_01.jpg' -> '张三'，'张三.jpg' -> '张三'"""
    name = img_path.stem
    if "_" in name:
        parts = name.rsplit("_", 1)
        if parts[-1].isdigit():
            name = parts[0]
    return name


def _collect_images(directory: Path):
    """收集目录下所有支持的图片文件"""
    if not directory.exists():
        return []
    return sorted([
        f for f in directory.glob("*.*")
        if f.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
    ])


# ============================================================
# 人脸库向量缓存
# ============================================================
def _dir_signature(directory: Path) -> str | None:
    """计算人脸库目录签名：文件名 + 文件内容哈希。目录不存在时返回 None。"""
    if not directory.exists():
        return None
    image_files = _collect_images(directory)
    h = hashlib.sha256()
    h.update(str(len(image_files)).encode("utf-8"))
    for f in image_files:
        h.update(f.name.encode("utf-8"))
        try:
            h.update(f.read_bytes())
        except OSError:
            continue
    return h.hexdigest()


def _load_face_db_cache(cache_file: Path, signature: str, model: str):
    """读取人脸库缓存；签名与模型都匹配才返回 face_db，否则返回 None。"""
    if not cache_file.exists():
        return None
    try:
        with open(cache_file, "rb") as f:
            cached = pickle.load(f)
    except Exception as e:
        print(f"⚠️ 人脸库缓存读取失败（{e}），将重新录入")
        return None
    if cached.get("signature") != signature or cached.get("model") != model:
        print("ℹ️ 人脸库图片有增删改或模型已更换，将重新录入")
        return None
    return cached.get("face_db")


def _save_face_db_cache(cache_file: Path, signature: str, face_db, model: str):
    """将人脸库向量写入缓存文件。"""
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        with open(cache_file, "wb") as f:
            pickle.dump(
                {"signature": signature, "model": model, "face_db": face_db}, f
            )
        print(f"✅ 人脸向量已缓存: {cache_file}")
    except Exception as e:
        print(f"⚠️ 人脸库缓存保存失败（{e}），本次不缓存")


def _print_face_db_summary(face_db):
    """打印人脸库汇总信息。"""
    print(f"\n人脸库共 {len(face_db)} 人:")
    for n, embs in face_db.items():
        print(f"   - {n}: {len(embs)} 张")
    print(f"{'=' * 50}\n")


def enroll_faces(
    app: FaceAnalysis,
    ref_dir: Path,
    save_annotated: bool = True,
    output_dir: Path = None,
    cache_file: Path = None,
):
    """
    从 ref_dir 录入人脸，返回 face_db: {姓名: [embedding1, embedding2, ...]}。

    cache_file: 传入时启用向量缓存——
      若 Ref_Figure 图片没有增删改，直接复用缓存向量（跳过模型推理）；
      若图片有变化，则重新录入并更新缓存。None 则不缓存、每次都重新计算。
    save_annotated: 是否保存检测可视化图到 output_dir/Ref_Figure/
    """
    # --- 0. 尝试命中缓存 ---
    signature = _dir_signature(ref_dir)
    if cache_file is not None and signature is not None:
        cached_db = _load_face_db_cache(cache_file, signature, MODEL_NAME)
        if cached_db is not None:
            print("✅ 人脸库未变化，使用缓存向量（跳过重新计算）")
            _print_face_db_summary(cached_db)
            return cached_db

    # --- 1. 正常录入 ---
    face_db = {}

    if not ref_dir.exists():
        print(f"❌ 目录不存在，已自动创建: {ref_dir}")
        ref_dir.mkdir(parents=True, exist_ok=True)
        return face_db

    image_files = _collect_images(ref_dir)
    if not image_files:
        print("⚠️ 没有图片，请放入参考人脸图片。命名规则：'张三.jpg' 或 '张三_01.jpg'")
        return face_db

    print(f"\n{'=' * 50}")
    print(f"人脸录入中... 共发现 {len(image_files)} 张图片")
    print(f"{'=' * 50}")

    for img_path in image_files:
        name = _extract_name(img_path)

        img = imread_unicode(str(img_path))
        if img is None:
            print(f"⚠️ 无法读取: {img_path.name}，跳过")
            continue

        faces = app.get(img)
        if len(faces) == 0:
            print(f"⚠️ 未检测到人脸: {img_path.name}，跳过")
            continue
        if len(faces) > 1:
            print(f"⚠️ {img_path.name} 检测到 {len(faces)} 张人脸，仅取第一张")

        face_db.setdefault(name, []).append(faces[0].normed_embedding)
        print(f"✅ 已录入: {img_path.name} -> {name}")

        if save_annotated and output_dir:
            out_path = output_dir / "Ref_Figure" / f"{img_path.stem}_detected.jpg"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            annotated = draw_face_annotations(img, faces)
            imwrite_unicode(str(out_path), annotated)
            print(f"   📷 已保存: {out_path}")

    _print_face_db_summary(face_db)

    # --- 2. 写入缓存 ---
    if cache_file is not None and signature is not None:
        _save_face_db_cache(cache_file, signature, face_db, MODEL_NAME)

    return face_db


# ============================================================
# 匹配
# ============================================================
def cosine_similarity(emb1, emb2):
    """余弦相似度"""
    return np.dot(emb1, emb2) / (np.linalg.norm(emb1) * np.linalg.norm(emb2))


def match_face(embedding, face_db: dict, threshold: float):
    """
    将 embedding 与人脸库比对，取最高分。
    返回 (姓名, 最高相似度)，低于阈值返回 ("Unknown", 最高相似度)。
    """
    best_name, best_score = "Unknown", 0.0
    for name, embeddings in face_db.items():
        for ref_emb in embeddings:
            sim = cosine_similarity(embedding, ref_emb)
            if sim > best_score:
                best_score, best_name = sim, name
    if best_score < threshold:
        return "Unknown", best_score
    return best_name, best_score


# ============================================================
# 人脸识别
# ============================================================
def recognize_faces(
    app: FaceAnalysis,
    face_db: dict,
    target_dir: Path,
    threshold: float,
    save_annotated: bool = True,
    output_dir: Path = None,
):
    """
    对 target_dir 下所有图片进行人脸识别。
    返回识别结果列表（可直接存为 CSV）。
    """
    if not face_db:
        print("❌ 人脸库为空，无法识别。请先在 Ref_Figure 中放入参考图片。")
        return []

    if not target_dir.exists():
        print(f"❌ 目录不存在，已自动创建: {target_dir}")
        target_dir.mkdir(parents=True, exist_ok=True)
        return []

    image_files = _collect_images(target_dir)
    if not image_files:
        print("⚠️ 没有图片，请放入待识别的图片。")
        return []

    print(f"\n{'=' * 50}")
    print(f"人脸识别中... 共 {len(image_files)} 张待识别图片")
    print(f"{'=' * 50}")

    results = []

    for img_path in image_files:
        img = imread_unicode(str(img_path))
        if img is None:
            print(f"⚠️ 无法读取: {img_path.name}，跳过")
            continue

        faces = detect_faces(app, img)
        if len(faces) == 0:
            print(f"⚠️ 未检测到人脸: {img_path.name}")
            results.append({
                "图片": img_path.name,
                "识别结果": "未检测到人脸",
                "相似度": 0.0,
                "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            continue

        face_names, face_scores = [], []
        for i, face in enumerate(faces):
            name, score = match_face(face.normed_embedding, face_db, threshold)
            face_names.append(name)
            face_scores.append(round(score, 4))

            label = f"{img_path.name} (人脸{i + 1})" if len(faces) > 1 else img_path.name
            status = "✅ 已识别" if name != "Unknown" else "❌ 未知人员"
            print(f"{status}: {label} -> {name} (相似度: {score:.4f})")

            results.append({
                "图片": label,
                "识别结果": name,
                "相似度": round(score, 4),
                "状态": "已识别" if name != "Unknown" else "未知",
                "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

        if save_annotated and output_dir:
            out_path = output_dir / "Target_Figure" / f"{img_path.stem}_recognized.jpg"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            annotated = draw_face_annotations(img, faces, face_names, face_scores)
            imwrite_unicode(str(out_path), annotated)
            print(f"   📷 已保存: {out_path}")

    print(f"{'=' * 50}\n")
    return results


# ============================================================
# Excel 结果输出
# ============================================================
def save_results_to_excel(results, output_path):
    """
    将人脸识别结果保存为 Excel (.xlsx) 文件。

    输出列：被识别图像名称 | 人脸 | 识别人名 | 相似度 | 状态 | 时间
    兼容 recognize_faces 返回的"图片"字段（单张人脸时为文件名，
    多张人脸时为 "文件名 (人脸N)"），自动拆分为图像名称和人脸序号两列。
    """
    if not results:
        print("⚠️ 无识别结果，跳过 Excel 保存")
        return False

    rows = []
    for r in results:
        img_name, face_seq = _split_image_and_face(r.get("图片", ""))
        rows.append({
            "被识别图像名称": img_name,
            "人脸": face_seq,
            "识别人名": r.get("识别结果", ""),
            "相似度": r.get("相似度", ""),
            "状态": r.get("状态", ""),
            "时间": r.get("时间", ""),
        })

    columns = ["被识别图像名称", "人脸", "识别人名", "相似度", "状态", "时间"]
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(output_path, index=False)
    print(f"📄 识别结果已保存到 Excel: {output_path}")
    return True


def _split_image_and_face(label) -> tuple[str, str]:
    """将 '文件名 (人脸N)' 拆分为 (图像名称, 人脸N)；无人脸编号时返回 (原字符串, '')。"""
    m = re.search(r"^(.*?)\s*\(人脸(\d+)\)$", str(label).strip())
    if m:
        return m.group(1), f"人脸{m.group(2)}"
    return str(label), ""
