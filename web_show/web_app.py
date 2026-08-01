"""
人脸识别考勤系统 - Web 交互界面（苹果风格 · 白蓝主题 · 分步向导）

流程：
  第1步 人脸识别：上传 Ref_Figure / Target_Figure 图片 → 点按钮识别 → 生成 output/00_程序人脸识别结果.xlsx
  第2步 上传输入表：5 个 input/ 输入表，全部就绪后进入下一步
  第3步 生成表格：表1~表10 依次生成，生成 05/06 时弹窗人工确认（可回传修正表）
  第4步 完成下载：下载 08/09/10 单表 + 一键 ZIP

启动（必须用 AI_PY312 环境，该环境有 insightface/cv2）：
  C:/Users/29037/.conda/envs/AI_PY312/python.exe web_show/web_app.py
"""

import io
import os
import sys
import zipfile
import contextlib
import threading
import traceback
from pathlib import Path
from datetime import datetime

# 注入项目根目录，便于 import face_app
BASE_DIR = Path(__file__).resolve().parents[1]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

WEB_SHOW_DIR = Path(__file__).resolve().parent

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from flask import Flask, request, jsonify, send_file, send_from_directory

from face_app import (  # 仅依赖 pandas/openpyxl，无需 insightface（懒加载）
    get_table1, get_table2, get_table3, get_table4,
    get_table5, get_table6, get_table7, get_table8_9_10,
)
from face_app.generate_table import (  # 表格输出路径的单一数据源
    TABLE1_FILE, TABLE2_FILE, TABLE3_FILE, TABLE4_FILE,
    TABLE5_FILE, TABLE5_ERROR_FILE, TABLE6_FILE, TABLE6_ERROR_FILE,
    TABLE7_FILE,
)

# ============ 配置（与 main.py 保持同步） ============
THRESHOLD = 0.45
SAVE_ANNOTATED = False
REF_DIR = BASE_DIR / "Ref_Figure"
TARGET_DIR = BASE_DIR / "Target_Figure"
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
RESULT_FILE = OUTPUT_DIR / "00_程序人脸识别结果.xlsx"
FACE_DB_CACHE_FILE = REF_DIR / "face_db_cache.pkl"  # 随人脸库存放
# ====================================================

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

# 输入表：web key → input/ 文件名
INPUT_FILES = {
    "photo_ledger": "01_照片台账表.xlsx",
    "work_plan": "02_工作安排表.xlsx",
    "person_class": "03_人员分类表.xlsx",
    "project_info": "04_项目信息表.xlsx",
    "attendance_template": "05_工程与外协考勤表模板.xlsx",
}

# 表格函数分发表
TABLE_FUNCS = {
    "table1": get_table1,
    "table2": get_table2,
    "table3": get_table3,
    "table4": get_table4,
    "table5": get_table5,
    "table6": get_table6,
    "table7": get_table7,
    "table8_9_10": get_table8_9_10,
}
TABLE_LABELS = {
    "table1": "表1 今日相机出工信息提取表",
    "table2": "表2 工作安排_提取表",
    "table3": "表3 出工照片人脸识别结果表",
    "table4": "表4 出工照片识别出工人表",
    "table5": "表5 该日出工人员表",
    "table6": "表6 全体人员出工情况表",
    "table7": "表7 出工地点及时长统计表",
    "table8_9_10": "表8/9/10 最终考勤表",
}
TABLE_OUTPUTS = {
    "table1": TABLE1_FILE,
    "table2": TABLE2_FILE,
    "table3": TABLE3_FILE,
    "table4": TABLE4_FILE,
    "table5": TABLE5_FILE,
    "table6": TABLE6_FILE,
    "table7": TABLE7_FILE,
}
FINAL_FILES = [
    "08_表8工程考勤表.xlsx",
    "09_表9外协考勤表1.xlsx",
    "10_表10外协考勤表2.xlsx",
]

# 需要人工确认的表格（每次生成都弹窗）
CONFIRM_TABLES = {"table5": "05", "table6": "06"}

# 输出文件 → 显示名
OUTPUT_FILES_LABELS = {
    "00_程序人脸识别结果.xlsx": "人脸识别结果",
    TABLE1_FILE.name: "表1",
    TABLE2_FILE.name: "表2",
    TABLE3_FILE.name: "表3",
    TABLE4_FILE.name: "表4",
    TABLE5_FILE.name: "表5",
    TABLE6_FILE.name: "表6",
    TABLE5_ERROR_FILE.name: "核对信息错误文档1",
    TABLE6_ERROR_FILE.name: "核对信息错误文档2",
    TABLE7_FILE.name: "表7",
    "08_表8工程考勤表.xlsx": "表8",
    "09_表9外协考勤表1.xlsx": "表9",
    "10_表10外协考勤表2.xlsx": "表10",
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB

# 防止识别与表格生成并发写 output/
RUN_LOCK = threading.Lock()


# ============================================================
# 工具函数
# ============================================================
def _image_files(d):
    if not d.is_dir():
        return []
    return [p for p in d.iterdir() if p.suffix.lower() in IMAGE_EXTS]


def _unique_path(d, name):
    """同名文件自动加序号：stem_1.ext。"""
    p = d / name
    if not p.exists():
        return p
    stem, suffix = Path(name).stem, Path(name).suffix
    i = 1
    while (d / f"{stem}_{i}{suffix}").exists():
        i += 1
    return d / f"{stem}_{i}{suffix}"


def _list_outputs():
    outputs = []
    for filename, label in OUTPUT_FILES_LABELS.items():
        path = OUTPUT_DIR / filename
        if path.exists():
            outputs.append({
                "name": filename,
                "label": label,
                "size": path.stat().st_size,
                "mtime": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
    return outputs


def _clear_dir(d):
    """删除目录内的所有文件与子目录（保留目录本身），返回删除项数。"""
    if not d.is_dir():
        return 0
    n = 0
    for item in d.iterdir():
        try:
            if item.is_dir():
                import shutil

                shutil.rmtree(item)
            else:
                item.unlink()
            n += 1
        except Exception:
            pass
    return n


def _resolve_providers():
    """按可用性选择推理后端：CUDA → DirectML → CPU（本机无 CUDA 时回退 CPU）。"""
    try:
        import onnxruntime as ort
        avail = set(ort.get_available_providers())
        for p in ("CUDAExecutionProvider", "DmlExecutionProvider", "CPUExecutionProvider"):
            if p in avail:
                return [p, "CPUExecutionProvider"] if p != "CPUExecutionProvider" else [p]
    except Exception:
        pass
    return ["CPUExecutionProvider"]


def _run_func(func):
    """直接调用表格函数，捕获 stdout 与返回值。返回 (success, output, result)。"""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            result = func()
        return True, buf.getvalue(), result
    except Exception:
        return False, buf.getvalue() + "\n" + traceback.format_exc(), None


# ============================================================
# 人脸识别后台线程（实时日志）
# ============================================================
class _LiveStream(io.StringIO):
    """把 stdout 按行实时转发给回调，同时保留完整文本。"""

    def __init__(self, on_line):
        super().__init__()
        self._on_line = on_line
        self._buf = ""

    def write(self, s):
        super().write(s)
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line:
                self._on_line(line)
        return len(s)

    def drain(self):
        if self._buf:
            self._on_line(self._buf)
            self._buf = ""


class RecognitionRunner:
    """后台运行人脸识别，前端轮询 /api/progress 获取实时日志。"""

    def __init__(self):
        self._lock = threading.Lock()
        self.running = False
        self.done = False
        self.success = False
        self.message = ""
        self.log_lines = []
        self.result_file = None

    def snapshot(self):
        with self._lock:
            return {
                "running": self.running,
                "done": self.done,
                "success": self.success,
                "message": self.message,
                "logs": list(self.log_lines),
                "result_file": str(self.result_file) if self.result_file else None,
            }

    def reset(self):
        """清空识别状态（用于清空文件夹后，让界面回到未识别状态）。"""
        with self._lock:
            self.done = False
            self.success = False
            self.message = ""
            self.log_lines = []
            self.result_file = None

    def start(self):
        if self.running:
            return False
        with self._lock:
            self.done = False
            self.success = False
            self.message = ""
            self.log_lines = []
            self.result_file = None
        RUN_LOCK.acquire()  # 整个识别期间独占，防止与表格生成冲突
        with self._lock:
            self.running = True
        threading.Thread(target=self._run, daemon=True).start()
        return True

    def _run(self):
        stream = _LiveStream(lambda line: self._add_log([line]))
        try:
            with contextlib.redirect_stdout(stream):
                # 惰性导入：仅此线程需要 insightface / cv2
                from face_app import (
                    init_face_app, enroll_faces, recognize_faces, save_results_to_excel,
                )

                app_ = init_face_app(_resolve_providers())
                face_db = enroll_faces(app_, REF_DIR, SAVE_ANNOTATED, OUTPUT_DIR, FACE_DB_CACHE_FILE)
                results = recognize_faces(
                    app_, face_db, TARGET_DIR, THRESHOLD, SAVE_ANNOTATED, OUTPUT_DIR
                )
                if results:
                    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                    save_results_to_excel(results, RESULT_FILE)
                    self.result_file = RESULT_FILE
                    self.message = f"识别完成，共处理 {len(results)} 张照片"
                else:
                    self.message = "未检测到可识别照片（Target_Figure 中可能没有检测到人脸）"
            self.success = True
        except ModuleNotFoundError:
            self.message = (
                "未安装 insightface / opencv，无法执行人脸识别。\n"
                "请使用 AI_PY312 环境启动本程序（C:/Users/29037/.conda/envs/AI_PY312/python.exe web_show/web_app.py）。"
            )
            self.success = False
        except Exception:
            self.message = traceback.format_exc()
            self.success = False
        finally:
            stream.drain()
            with self._lock:
                self.running = False
                self.done = True
            RUN_LOCK.release()

    def _add_log(self, lines):
        with self._lock:
            self.log_lines.extend(lines)


runner = RecognitionRunner()


# ============================================================
# 页面 / 静态资源
# ============================================================
@app.route("/")
def index():
    return send_file(WEB_SHOW_DIR / "index.html")


@app.route("/style.css")
def style_css():
    return send_file(WEB_SHOW_DIR / "style.css")


@app.route("/logo.png")
def logo_png():
    return send_file(WEB_SHOW_DIR / "logo.png", mimetype="image/png")


@app.route("/favicon.ico")
def favicon():
    return send_file(WEB_SHOW_DIR / "logo.png", mimetype="image/png")


@app.route("/app.js")
def app_js():
    return send_file(WEB_SHOW_DIR / "app.js")


# ============================================================
# 状态
# ============================================================
@app.route("/api/status")
def api_status():
    uploads = {key: (INPUT_DIR / fname).exists() for key, fname in INPUT_FILES.items()}
    snap = runner.snapshot()
    tables = {key: path.exists() for key, path in TABLE_OUTPUTS.items()}
    tables["table8_9_10"] = all((OUTPUT_DIR / n).exists() for n in FINAL_FILES)
    return jsonify({
        "uploads": uploads,
        "image_counts": {
            "ref": len(_image_files(REF_DIR)),
            "target": len(_image_files(TARGET_DIR)),
        },
        "outputs": _list_outputs(),
        "tables": tables,
        "recognition_running": snap["running"],
        "recognition_done": (snap["done"] and snap["success"]) or RESULT_FILE.exists(),
    })


# ============================================================
# 上传
# ============================================================
@app.route("/api/upload_images/<zone>", methods=["POST"])
def upload_images(zone):
    if zone not in ("ref", "target"):
        return jsonify({"success": False, "message": "未知上传区"}), 400
    target_dir = REF_DIR if zone == "ref" else TARGET_DIR
    target_dir.mkdir(parents=True, exist_ok=True)

    files = request.files.getlist("files")
    if not files:
        return jsonify({"success": False, "message": "未收到文件"}), 400

    saved, rejected = [], []
    for f in files:
        name = Path(f.filename or "").name
        if not name or "/" in name or "\\" in name or ".." in name:
            rejected.append(f.filename)
            continue
        if Path(name).suffix.lower() not in IMAGE_EXTS:
            rejected.append(name)
            continue
        try:
            f.save(_unique_path(target_dir, name))
            saved.append(name)
        except Exception as e:
            rejected.append(f"{name}({e})")

    return jsonify({
        "success": True,
        "saved": saved,
        "rejected": rejected,
        "image_counts": {
            "ref": len(_image_files(REF_DIR)),
            "target": len(_image_files(TARGET_DIR)),
        },
    })


@app.route("/api/upload_input/<file_key>", methods=["POST"])
def upload_input(file_key):
    if file_key not in INPUT_FILES:
        return jsonify({"success": False, "message": "未知文件类型"}), 400
    fname = INPUT_FILES[file_key]
    if "file" not in request.files:
        return jsonify({"success": False, "message": "未收到文件"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"success": False, "message": "文件名为空"}), 400
    if Path(f.filename).suffix.lower() != ".xlsx":
        return jsonify({"success": False, "message": "请输入 .xlsx 文件"}), 400
    save_path = INPUT_DIR / fname
    try:
        save_path.parent.mkdir(parents=True, exist_ok=True)
        f.save(save_path)
        return jsonify({"success": True, "message": f"已保存：input/{fname}"})
    except Exception as e:
        return jsonify({"success": False, "message": f"保存失败: {e}"}), 500


# ============================================================
# 人脸识别
# ============================================================
@app.route("/api/recognize", methods=["POST"])
def recognize():
    snap = runner.snapshot()
    if snap["running"]:
        return jsonify({
            "success": False, "started": False, "running": True,
            "message": "人脸识别已在进行中，请稍候",
        }), 409
    if not _image_files(REF_DIR):
        return jsonify({
            "success": False, "started": False, "running": False,
            "message": "Ref_Figure（员工人脸库）内没有图片，请先上传",
        }), 400
    if not _image_files(TARGET_DIR):
        return jsonify({
            "success": False, "started": False, "running": False,
            "message": "Target_Figure（出工照片）内没有图片，请先上传",
        }), 400
    started = runner.start()
    return jsonify({
        "success": started,
        "started": started,
        "running": started,
        "message": "人脸识别已启动" if started else "启动失败",
    })


@app.route("/api/progress")
def progress():
    return jsonify(runner.snapshot())


# ============================================================
# 表格生成
# ============================================================
@app.route("/api/run_table/<key>", methods=["POST"])
def run_table(key):
    if key not in TABLE_FUNCS:
        return jsonify({"success": False, "message": f"未知表格: {key}"}), 404
    if runner.snapshot()["running"]:
        return jsonify({"success": False, "message": "人脸识别正在运行，请稍后再试"}), 409

    with RUN_LOCK:
        success, output, result = _run_func(TABLE_FUNCS[key])

    info = None
    if key in CONFIRM_TABLES and success and isinstance(result, dict):
        info = {
            "normal": result.get("normal"),
            "abnormal": result.get("abnormal"),
            "error_doc": Path(result["error_doc"]).name if result.get("error_doc") else None,
            "table_file": Path(result["table_file"]).name if result.get("table_file") else None,
        }

    return jsonify({
        "success": success,
        "key": key,
        "label": TABLE_LABELS.get(key, key),
        "message": "执行成功" if success else "执行失败，请查看日志",
        "output": output,
        "confirm_required": key in CONFIRM_TABLES and success,
        "table_info": info,
        "outputs": _list_outputs(),
    })


@app.route("/api/confirm/<target>", methods=["POST"])
def confirm(target):
    mapping = {"05": TABLE5_FILE, "06": TABLE6_FILE}
    if target not in mapping:
        return jsonify({"success": False, "message": "未知确认目标"}), 400
    dest = mapping[target]
    f = request.files.get("file")
    if f and f.filename:
        if Path(f.filename).suffix.lower() != ".xlsx":
            return jsonify({"success": False, "message": "回传修正表必须是 .xlsx 文件"}), 400
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            f.save(dest)
        except Exception as e:
            return jsonify({"success": False, "message": f"保存失败: {e}"}), 500
    return jsonify({
        "success": True,
        "message": "已确认，继续",
        "outputs": _list_outputs(),
    })


# ============================================================
# 清空文件夹（只清 output/ 和 Target_Figure/）
# ============================================================
@app.route("/api/clear_folders", methods=["POST"])
def clear_folders():
    if runner.snapshot()["running"]:
        return jsonify({
            "success": False,
            "message": "人脸识别正在运行，请先停止后再清空",
        }), 409
    with RUN_LOCK:
        n_out = _clear_dir(OUTPUT_DIR)
        n_tgt = _clear_dir(TARGET_DIR)
    runner.reset()
    return jsonify({
        "success": True,
        "message": f"已清空：output/（{n_out} 项）、Target_Figure/（{n_tgt} 项）。人脸库与输入表不受影响。",
        "outputs": _list_outputs(),
    })


# ============================================================
# 下载
# ============================================================
@app.route("/api/download/<path:filename>")
def download(filename):
    name = Path(filename).name  # 防止路径逃逸
    if not (OUTPUT_DIR / name).exists():
        return jsonify({"success": False, "message": f"文件不存在: {name}"}), 404
    return send_from_directory(str(OUTPUT_DIR), name, as_attachment=True)


@app.route("/api/download_zip")
def download_zip():
    existing = [n for n in FINAL_FILES if (OUTPUT_DIR / n).exists()]
    if not existing:
        return jsonify({"success": False, "message": "尚未生成 8/9/10 表"}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in existing:
            zf.write(str(OUTPUT_DIR / name), name)  # 中文名自动置 UTF-8 flag
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"最终考勤表_{datetime.now():%Y-%m-%d}.zip",
    )


# ============================================================
# 错误处理
# ============================================================
@app.errorhandler(413)
def too_large(_e):
    return jsonify({"success": False, "message": "上传文件过大（超过 100MB）"}), 413


if __name__ == "__main__":
    print("=" * 50)
    print(" 人脸识别考勤系统 Web 界面")
    print(" 浏览器打开: http://localhost:5000")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, threaded=True, debug=False)
