"""
人脸考勤处理流程 - Web交互界面

按用户流程图逐步执行：
  步骤1: 上传基础表 → 人脸识别 → 生成表1/2/3/4 → 生成表5（05_该日出工人员表 + 核对信息错误文档1）
  步骤2: 生成表6（06_全体人员出工情况表 + 核对信息错误文档2）
  步骤3: 生成表7
  步骤4: 上传考勤模板 → 生成表8/9/10
"""

import os
import sys
import io
import json
import contextlib
import subprocess
import traceback
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_from_directory

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

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"            # 输入表目录
OUTPUT_DIR = BASE_DIR / "output"          # 统一输出目录
LEGACY_OUTPUT_DIR = BASE_DIR / "Output"   # 兼容旧的大写 Output/
UPLOAD_MAP = {
    "001人员分类表": "input/03_人员分类表.xlsx",
    "002项目信息表": "input/04_项目信息表.xlsx",
    "工作安排表": "input/02_工作安排表.xlsx",
    "照片台账表": "input/01_照片台账表.xlsx",
    "工程与外协考勤表模板": "input/05_工程与外协考勤表模板.xlsx",
}

# 按步骤分组脚本
STEPS = {
    "step1": {
        "name": "基础信息处理",
        "uploads": ["001人员分类表", "002项目信息表", "工作安排表", "照片台账表"],
        "scripts": [
            {"key": "face_recognition", "label": "人脸识别", "script": "main.py", "needs": []},
            {"key": "table1", "label": "生成表1", "func": get_table1, "needs": ["照片台账表"]},
            {"key": "table2", "label": "生成表2", "func": get_table2, "needs": ["工作安排表"]},
            {"key": "table3", "label": "生成表3", "func": get_table3, "needs": []},
            {"key": "table4", "label": "生成表4", "func": get_table4, "needs": []},
            {"key": "table5", "label": "生成表5/异常检验1", "func": get_table5, "needs": ["01_今日相机出工信息提取表.xlsx", "04_出工照片识别出工人表.xlsx"]},
        ],
    },
    "step2": {
        "name": "生成表6/异常检验2",
        "uploads": [],
        "scripts": [
            {"key": "table6", "label": "生成表6/异常检验2", "func": get_table6, "needs": ["05_该日出工人员表.xlsx", "02_工作安排_提取表.xlsx"]},
        ],
    },
    "step3": {
        "name": "生成表7",
        "uploads": [],
        "scripts": [
            {"key": "table7", "label": "生成表7", "func": get_table7, "needs": ["06_全体人员出工情况表.xlsx", "001人员分类表"]},
        ],
    },
    "step4": {
        "name": "生成表8/9/10",
        "uploads": ["工程与外协考勤表模板"],
        "scripts": [
            {"key": "table8_9_10", "label": "生成表8/9/10", "func": get_table8_9_10, "needs": ["工程与外协考勤表模板", "07_表7出工地点及时长统计表.xlsx"]},
        ],
    },
}

# 最终输出的关键文件
OUTPUT_FILES = {
    "00_程序人脸识别结果.xlsx": "人脸识别结果",
    "01_今日相机出工信息提取表.xlsx": "表1",
    "02_工作安排_提取表.xlsx": "表2",
    "03_出工照片人脸识别结果表.xlsx": "表3",
    "04_出工照片识别出工人表.xlsx": "表4",
    "05_该日出工人员表.xlsx": "表5",
    "06_全体人员出工情况表.xlsx": "表6",
    "核对信息错误文档1.xlsx": "核对信息错误文档1",
    "核对信息错误文档2.xlsx": "核对信息错误文档2",
    "07_表7出工地点及时长统计表.xlsx": "表7",
    "08_表8工程考勤表.xlsx": "表8",
    "09_表9外协考勤表1.xlsx": "表9",
    "10_表10外协考勤表2.xlsx": "表10",
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB


def _save_upload(file_type, file):
    """保存上传文件到项目根目录，使用脚本期望的文件名"""
    if file_type not in UPLOAD_MAP:
        return False, f"未知文件类型: {file_type}"
    save_path = BASE_DIR / UPLOAD_MAP[file_type]
    try:
        save_path.parent.mkdir(parents=True, exist_ok=True)
        file.save(save_path)
        return True, str(save_path)
    except Exception as e:
        return False, f"保存失败: {e}"


def _file_exists(file_type):
    """检查上传的基础文件是否已存在"""
    if file_type not in UPLOAD_MAP:
        return False
    return (BASE_DIR / UPLOAD_MAP[file_type]).exists()


def _check_script_needs(script_info):
    """检查脚本依赖文件是否齐全（生成类表格在 output/ 下查找）"""
    missing = []
    for need in script_info.get("needs", []):
        if need.endswith(".xlsx"):
            if not (BASE_DIR / need).exists() and not (OUTPUT_DIR / need).exists() and not (LEGACY_OUTPUT_DIR / need).exists():
                missing.append(need)
        else:
            if not _file_exists(need):
                missing.append(UPLOAD_MAP.get(need, need))
    return missing


def _run_script(script_name):
    """运行指定Python脚本，返回(成功标志, 输出文本)"""
    script_path = BASE_DIR / script_name
    if not script_path.exists():
        return False, f"脚本不存在: {script_path}"

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(BASE_DIR),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        success = result.returncode == 0
        return success, result.stdout
    except Exception as e:
        return False, f"执行异常: {e}\n{traceback.format_exc()}"


def _run_func(func):
    """直接调用函数并捕获 stdout，返回(成功标志, 输出文本)"""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            func()
        return True, buf.getvalue()
    except Exception as e:
        return False, buf.getvalue() + "\n" + traceback.format_exc()


def _run_script_info(script_info):
    """执行任务：有 func 直接调用函数，否则用 subprocess 跑脚本"""
    if "func" in script_info:
        return _run_func(script_info["func"])
    return _run_script(script_info["script"])


def _list_outputs():
    """列出当前已生成的输出文件（依次查 output/、Output/、根目录）"""
    outputs = []
    for filename, label in OUTPUT_FILES.items():
        path = OUTPUT_DIR / filename
        if not path.exists():
            path = LEGACY_OUTPUT_DIR / filename
        if not path.exists():
            path = BASE_DIR / filename
        if path.exists():
            outputs.append({
                "name": filename,
                "label": label,
                "size": path.stat().st_size,
                "mtime": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
    return outputs


# ============================================================
# 路由
# ============================================================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/status")
def api_status():
    """返回各步骤上传状态及输出文件列表"""
    uploads = {}
    for key, name in UPLOAD_MAP.items():
        uploads[key] = (BASE_DIR / name).exists()
    return jsonify({
        "uploads": uploads,
        "outputs": _list_outputs(),
    })


@app.route("/upload/<file_type>", methods=["POST"])
def upload_file(file_type):
    if "file" not in request.files:
        return jsonify({"success": False, "message": "未收到文件"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "message": "文件名为空"}), 400

    success, msg = _save_upload(file_type, file)
    if not success:
        return jsonify({"success": False, "message": msg}), 500

    return jsonify({"success": True, "message": f"已保存: {UPLOAD_MAP[file_type]}"})


@app.route("/run/<script_key>", methods=["POST"])
def run_script(script_key):
    """根据脚本key查找配置并执行"""
    script_info = None
    for step in STEPS.values():
        for s in step["scripts"]:
            if s["key"] == script_key:
                script_info = s
                break
        if script_info:
            break

    if not script_info:
        return jsonify({"success": False, "message": f"未知脚本: {script_key}"}), 404

    missing = _check_script_needs(script_info)
    if missing:
        return jsonify({
            "success": False,
            "message": f"缺少依赖文件: {', '.join(missing)}",
        }), 400

    success, output = _run_script_info(script_info)
    return jsonify({
        "success": success,
        "message": "执行成功" if success else "执行失败，请查看日志",
        "output": output,
        "outputs": _list_outputs(),
    })


@app.route("/files/<path:filename>")
def download_file(filename):
    """下载文件：依次查 output/、Output/、根目录"""
    for folder in (OUTPUT_DIR, LEGACY_OUTPUT_DIR, BASE_DIR):
        if (folder / filename).exists():
            return send_from_directory(str(folder), filename, as_attachment=True)
    return jsonify({"success": False, "message": f"文件不存在: {filename}"}), 404


@app.route("/api/run_all/<step_key>", methods=["POST"])
def run_all_in_step(step_key):
    """一键执行某一步骤的所有脚本"""
    if step_key not in STEPS:
        return jsonify({"success": False, "message": "未知步骤"}), 404

    step = STEPS[step_key]
    results = []
    overall_success = True

    for script_info in step["scripts"]:
        missing = _check_script_needs(script_info)
        if missing:
            results.append({
                "key": script_info["key"],
                "label": script_info["label"],
                "success": False,
                "message": f"缺少依赖文件: {', '.join(missing)}",
                "output": "",
            })
            overall_success = False
            continue

        success, output = _run_script_info(script_info)
        results.append({
            "key": script_info["key"],
            "label": script_info["label"],
            "success": success,
            "message": "成功" if success else "失败",
            "output": output,
        })
        if not success:
            overall_success = False

    return jsonify({
        "success": overall_success,
        "results": results,
        "outputs": _list_outputs(),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
