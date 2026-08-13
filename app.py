# -*- coding: utf-8 -*-
"""PPTX Studio —— 高度定制化 PPTX 生成 API

基于 Flask + python-pptx，提供丰富 UI 组件与多套主题，可生成精美 PPTX。

接口列表：
  GET    /api/health              健康检查
  POST   /api/pptx/render         渲染 PPTX（高度定制）
  GET    /api/themes              列出可用主题
  GET    /api/components          列出可用组件
  GET    /api/templates           列出示例模板
  POST   /api/templates/<name>    用内置示例模板快速渲染
  GET    /api/download/<filename> 下载文件
  GET    /api/files               列出已生成文件
  DELETE /api/files/<filename>    删除文件

启动：
  python app.py            # 默认 0.0.0.0:5002
  PORT=5002 python app.py  # 指定端口
"""
import os
import json
import logging

from flask import Flask, request, jsonify, send_from_directory, url_for
from werkzeug.utils import secure_filename

from core import engine
from core.themes import list_themes
from core.components import component_list

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated")
TEMPLATES_DIR = os.path.join(BASE_DIR, "examples", "templates")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("pptx_studio_api")


# ------------------------------------------------------------------ 工具
def error_response(message, code=400):
    return jsonify({"success": False, "error": message}), code


def build_download_info(filename):
    return {
        "filename": filename,
        "download_url": url_for("download_file", filename=filename, _external=True),
    }


# ------------------------------------------------------------------ 路由
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"success": True, "status": "ok",
                    "time": __import__("datetime").datetime.now().isoformat()})


@app.route("/api/themes", methods=["GET"])
def themes():
    return jsonify({"success": True, "count": len(list_themes()),
                    "themes": list_themes()})


@app.route("/api/components", methods=["GET"])
def components():
    return jsonify({"success": True, "count": len(component_list()),
                    "components": component_list()})


@app.route("/api/pptx/render", methods=["POST"])
def pptx_render():
    data = request.get_json(silent=True)
    if data is None:
        return error_response("请求体必须是合法 JSON（Content-Type: application/json）")
    if not isinstance(data.get("pages"), list) or not data["pages"]:
        return error_response("pages 字段必须是非空数组")
    try:
        filename = engine.render(data, OUTPUT_DIR)
    except Exception as e:
        logger.exception("渲染 PPTX 失败")
        return error_response(f"渲染失败: {str(e)}", 500)
    info = build_download_info(filename)
    return jsonify({"success": True, **info}), 201


@app.route("/api/templates", methods=["GET"])
def templates():
    names = []
    if os.path.isdir(TEMPLATES_DIR):
        for fn in sorted(os.listdir(TEMPLATES_DIR)):
            if fn.endswith(".json"):
                names.append(fn[:-5])
    return jsonify({"success": True, "count": len(names), "templates": names})


@app.route("/api/templates/<name>", methods=["POST"])
def template_render(name):
    safe = secure_filename(name)
    path = os.path.join(TEMPLATES_DIR, safe + ".json")
    if not os.path.isfile(path):
        return error_response("模板不存在", 404)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        return error_response(f"模板解析失败: {str(e)}", 500)
    # 允许覆盖主题与文件名前缀
    body = request.get_json(silent=True) or {}
    if body.get("theme"):
        data["theme"] = body["theme"]
    if body.get("filename_prefix"):
        data["filename_prefix"] = body["filename_prefix"]
    try:
        filename = engine.render(data, OUTPUT_DIR)
    except Exception as e:
        logger.exception("模板渲染失败")
        return error_response(f"渲染失败: {str(e)}", 500)
    info = build_download_info(filename)
    return jsonify({"success": True, "template": safe, **info}), 201


@app.route("/api/download/<path:filename>", methods=["GET"])
def download_file(filename):
    safe_name = secure_filename(filename)
    if not os.path.isfile(os.path.join(OUTPUT_DIR, safe_name)):
        return error_response("文件不存在", 404)
    return send_from_directory(OUTPUT_DIR, safe_name, as_attachment=True)


@app.route("/api/files", methods=["GET"])
def list_files():
    files = []
    for name in sorted(os.listdir(OUTPUT_DIR)):
        path = os.path.join(OUTPUT_DIR, name)
        if os.path.isfile(path):
            files.append({
                "filename": name,
                "size_bytes": os.path.getsize(path),
                "modified": __import__("datetime").datetime.fromtimestamp(
                    os.path.getmtime(path)).isoformat(),
                "download_url": url_for("download_file", filename=name, _external=True),
            })
    return jsonify({"success": True, "count": len(files), "files": files})


@app.route("/api/files/<path:filename>", methods=["DELETE"])
def delete_file(filename):
    safe_name = secure_filename(filename)
    path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.isfile(path):
        return error_response("文件不存在", 404)
    os.remove(path)
    return jsonify({"success": True, "message": f"{safe_name} 已删除"})


@app.errorhandler(413)
def too_large(e):
    return error_response("请求体过大", 413)


@app.errorhandler(404)
def not_found(e):
    return error_response("接口不存在", 404)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5002"))
    app.run(host="0.0.0.0", port=port, debug=True)
