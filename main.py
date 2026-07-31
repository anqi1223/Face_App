"""人脸识别打卡系统 - 入口"""

import os
from pathlib import Path
import pandas as pd

from face_app import init_face_app, enroll_faces, recognize_faces, save_results_to_excel
from face_app import init_ocr_engine, extract_text_from_directory
from face_app import (
    get_table1,
    get_table2,
    get_table3,
    get_table4,
    get_table5,
    get_table6,
    get_table7,
    get_table8_9_10,
)

os.system("cls" if os.name == "nt" else "clear")  # 清空终端

# ============================================================
# 配置
# ============================================================
BASE_DIR = Path(__file__).parent
REF_DIR = BASE_DIR / "Ref_Figure"
TARGET_DIR = BASE_DIR / "Target_Figure"
OUTPUT_DIR = BASE_DIR / "output"
RESULT_FILE = OUTPUT_DIR / "00_程序人脸识别结果.xlsx"

# 人脸库向量缓存：Ref_Figure 图片未变化时复用缓存，跳过向量计算
FACE_DB_CACHE_FILE = OUTPUT_DIR / "face_db_cache.pkl"

# 识别阈值：余弦相似度低于此值视为未知人员（0~1，越高越严格）
THRESHOLD = 0.45

# 推理后端：["CPUExecutionProvider"] 纯CPU，["CUDAExecutionProvider", "CPUExecutionProvider"] GPU+CPU
PROVIDERS = ["CUDAExecutionProvider"]

# 是否保存标注后的可视化图片到 Output/
SAVE_ANNOTATED = False

# OCR 文字提取：是否启用，以及从哪个目录提取
OCR_ENABLED = False
OCR_DIR = TARGET_DIR  # 从待识别图片目录提取文字
OCR_RESULT_FILE = OUTPUT_DIR / "ocr_result.csv"


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 50)
    print("       人脸识别打卡系统 (InsightFace)")
    print(f"       阈值: {THRESHOLD}")
    print("=" * 50)

    # 1. 初始化模型
    app = init_face_app(PROVIDERS)

    # 2. 录入人脸库（图片未变化时直接复用缓存向量）
    face_db = enroll_faces(app, REF_DIR, SAVE_ANNOTATED, OUTPUT_DIR, FACE_DB_CACHE_FILE)

    # 3. 识别
    results = recognize_faces(
        app, face_db, TARGET_DIR, THRESHOLD, SAVE_ANNOTATED, OUTPUT_DIR
    )

    # 4. 保存结果（Excel 版；CSV 已废弃不再生成）
    if results:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        save_results_to_excel(results, RESULT_FILE)

        # 控制台预览
        df = pd.DataFrame(results)
        print(df.to_string(index=False))

    # 5. 表生成（表1~表10，一次性跑通全部，便于调试）
    table_tasks = [
        ("表1今日相机出工信息提取表", get_table1),
        ("表2工作安排信息提取表", get_table2),
        ("表3出工照片人脸识别结果表", get_table3),
        ("表4出工照片识别出工人表", get_table4),
        ("表5出工信息错报漏报检验", get_table5),
        ("表6全体人员出工情况表", get_table6),
        ("07_表7出工地点及时长统计表", get_table7),
        ("表8/9/10 最终考勤表", get_table8_9_10),
    ]
    for label, func in table_tasks:
        try:
            func()
        except FileNotFoundError as e:
            print(f"⚠️ 跳过【{label}】：{e}")
        except Exception as e:
            print(f"❌【{label}】失败：{e}")


if __name__ == "__main__":
    main()
