from __future__ import annotations

import copy
import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Lock
from uuid import uuid4

from flask import Flask, jsonify, render_template, request, send_from_directory, url_for

from desensitizer.excel_preview import (
    generate_watermark_preview,
    inspect_excel_upload,
    options_from_form,
    process_excel_preview,
    process_excel_preview_from_upload,
)
from desensitizer.processor import (
    OUTPUT_ROOT,
    ProcessingError,
    add_manual_redaction,
    clear_manual_redactions,
    export_job_output,
    process_uploaded_file,
    reprocess_existing_job,
    undo_last_manual_redaction,
)


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024
TASK_EXECUTOR = ThreadPoolExecutor(max_workers=2)
TASKS: dict[str, dict] = {}
TASK_LOCK = Lock()


def with_urls(report: dict) -> dict:
    decorated = copy.deepcopy(report)
    job_id = decorated["job_id"]
    outputs = decorated.get("outputs", {})
    decorated["download_urls"] = {
        key: url_for("output_file", name=f"{job_id}/{value}")
        for key, value in outputs.items()
        if key in {"report"} and value
    }
    for page in decorated.get("pages", []):
        page["image_url"] = url_for("output_file", name=f"{job_id}/{page['image']}")
    return decorated


def with_excel_urls(report: dict) -> dict:
    decorated = copy.deepcopy(report)
    job_id = decorated["job_id"]
    outputs = decorated.get("outputs", {})
    decorated["download_urls"] = {
        key: url_for("output_file", name=f"{job_id}/{value}")
        for key, value in outputs.items()
        if key in {"excel", "report"} and value
    }
    return decorated


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/favicon.ico")
def favicon():
    return send_from_directory(Path(__file__).parent, "app.ico")


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/api/process")
def process_files():
    files = [file for file in request.files.getlist("files") if file.filename]
    if not files:
        return jsonify({"error": "请选择至少一个 PDF、DOCX 或 DOC 文件。"}), 400
    image_method = request.form.get("image_method", "mosaic")
    expected_pages = parse_optional_int(request.form.get("expected_pages"))

    jobs = []
    errors = []
    for file in files:
        try:
            report = process_uploaded_file(
                file.filename,
                file.save,
                OUTPUT_ROOT,
                image_method=image_method,
                expected_pages=expected_pages,
            )
            jobs.append(with_urls(report))
        except ProcessingError as exc:
            errors.append({"file": file.filename, "error": str(exc)})
        except Exception as exc:  # Keep batch processing moving for other files.
            errors.append({"file": file.filename, "error": f"处理失败：{exc}"})

    status = 207 if errors and jobs else 200
    if errors and not jobs:
        status = 400
    return jsonify({"jobs": jobs, "errors": errors}), status


@app.post("/excel/preview")
def excel_preview():
    upload_id = (request.form.get("upload_id") or "").strip()
    file = request.files.get("file")
    if not upload_id and (not file or not file.filename):
        return jsonify({"error": "请选择一个 Excel 文件。"}), 400
    try:
        options = options_from_form(request.form)
        if upload_id:
            report = process_excel_preview_from_upload(upload_id, OUTPUT_ROOT, options=options)
        else:
            report = process_excel_preview(file.filename, file.save, OUTPUT_ROOT, options=options)
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Excel 预览版生成失败：{exc}"}), 500
    return jsonify({"job": with_excel_urls(report)})


@app.post("/excel/export")
def start_excel_export():
    upload_id = (request.form.get("upload_id") or "").strip()
    if not upload_id:
        return jsonify({"error": "当前文件不存在，请重新上传 Excel。"}), 400
    try:
        options = options_from_form(request.form)
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400

    task_id = uuid4().hex
    task = {
        "task_id": task_id,
        "state": "queued",
        "message": "准备导出 Excel...",
        "current": 0,
        "total": None,
        "kind": "excel",
    }
    with TASK_LOCK:
        TASKS[task_id] = task
    TASK_EXECUTOR.submit(run_excel_export_task, task_id, upload_id, options)
    return jsonify({"task_id": task_id})


@app.get("/excel/export-status/<task_id>")
def excel_export_status(task_id: str):
    with TASK_LOCK:
        task = copy.deepcopy(TASKS.get(task_id))
    if not task:
        return jsonify({"error": "任务不存在。"}), 404
    if task.get("state") == "done" and task.get("report"):
        task["job"] = with_excel_urls(task["report"])
        task.pop("report", None)
    return jsonify(task)


def run_excel_export_task(task_id: str, upload_id: str, options):
    def progress(message: str, current: int | None = None, total: int | None = None):
        with TASK_LOCK:
            task = TASKS.get(task_id)
            if not task:
                return
            task["state"] = "running"
            task["message"] = message
            if current is not None:
                task["current"] = current
            if total is not None:
                task["total"] = total

    try:
        progress("正在后台导出 Excel...", 0, None)
        report = process_excel_preview_from_upload(upload_id, OUTPUT_ROOT, options=options, progress_callback=progress)
        with TASK_LOCK:
            TASKS[task_id].update({"state": "done", "message": "导出完成", "report": report})
    except Exception as exc:
        with TASK_LOCK:
            TASKS[task_id].update({"state": "error", "message": f"导出失败：{exc}"})


@app.post("/excel/inspect")
def excel_inspect():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "请选择一个 Excel 文件。"}), 400
    try:
        upload = inspect_excel_upload(file.filename, file.save, OUTPUT_ROOT)
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"读取 Excel 工作表失败：{exc}"}), 500
    return jsonify({"upload": upload})


@app.post("/excel/watermark-preview")
def excel_watermark_preview():
    upload_id = (request.form.get("upload_id") or "").strip()
    sheet_name = (request.form.get("sheet_name") or "").strip()
    range_type = (request.form.get("range_type") or "used_range").strip() or "used_range"
    if not upload_id:
        return jsonify({"error": "当前文件不存在，请重新上传 Excel。"}), 400
    if not sheet_name:
        return jsonify({"error": "请选择一个工作表。"}), 400
    try:
        options = options_from_form(request.form)
        result = generate_watermark_preview(upload_id, sheet_name, range_type, options, OUTPUT_ROOT)
        result["preview_url"] = url_for("output_file", name=result["preview_file"])
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"生成水印预览图失败：{exc}"}), 500
    return jsonify({"preview": result})


@app.get("/excel/report/<job_id>")
def excel_report(job_id: str):
    if "/" in job_id or "\\" in job_id or not job_id.strip():
        return jsonify({"error": "任务 ID 无效。"}), 400
    root = OUTPUT_ROOT.resolve()
    job_dir = (root / job_id).resolve()
    if root not in job_dir.parents:
        return jsonify({"error": "任务路径无效。"}), 400
    if not job_dir.exists():
        return jsonify({"error": "任务不存在。"}), 404
    reports = sorted(job_dir.glob("*_处理报告.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not reports:
        return jsonify({"error": "没有找到处理报告。"}), 404
    return send_from_directory(job_dir, reports[0].name, as_attachment=True)


def parse_optional_int(value: str | None) -> int | None:
    if not value:
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


@app.post("/api/manual-redaction")
def manual_redaction():
    data = request.get_json(silent=True) or {}
    required = {"job_id", "page", "x", "y", "width", "height", "image_width", "image_height"}
    missing = sorted(required - set(data))
    if missing:
        return jsonify({"error": f"缺少参数：{', '.join(missing)}"}), 400

    try:
        report = add_manual_redaction(
            job_id=str(data["job_id"]),
            page_number=int(data["page"]),
            image_rect=(
                float(data["x"]),
                float(data["y"]),
                float(data["width"]),
                float(data["height"]),
            ),
            image_size=(float(data["image_width"]), float(data["image_height"])),
            method=str(data.get("method", "mosaic")),
            output_root=OUTPUT_ROOT,
        )
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400

    return jsonify({"job": with_urls(report)})


@app.post("/api/manual-redaction/undo")
def undo_manual_redaction():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get("job_id", ""))
    if not job_id:
        return jsonify({"error": "缺少任务 ID。"}), 400
    try:
        page = parse_optional_int(data.get("page"))
        report = undo_last_manual_redaction(job_id=job_id, page_number=page, output_root=OUTPUT_ROOT)
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"job": with_urls(report)})


@app.post("/api/manual-redactions/clear")
def clear_manual_redaction_list():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get("job_id", ""))
    if not job_id:
        return jsonify({"error": "缺少任务 ID。"}), 400
    try:
        page = parse_optional_int(data.get("page"))
        report = clear_manual_redactions(job_id=job_id, page_number=page, output_root=OUTPUT_ROOT)
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"job": with_urls(report)})


@app.post("/api/export")
def start_export():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get("job_id", ""))
    kind = str(data.get("kind", ""))
    if not job_id:
        return jsonify({"error": "缺少任务 ID。"}), 400
    if kind not in {"pdf", "zip"}:
        return jsonify({"error": "导出类型无效。"}), 400

    task_id = uuid4().hex
    task = {
        "task_id": task_id,
        "state": "queued",
        "message": "准备生成...",
        "current": 0,
        "total": None,
        "job_id": job_id,
        "kind": kind,
    }
    with TASK_LOCK:
        TASKS[task_id] = task
    TASK_EXECUTOR.submit(run_export_task, task_id, job_id, kind)
    return jsonify({"task_id": task_id})


@app.get("/api/export-status/<task_id>")
def export_status(task_id: str):
    with TASK_LOCK:
        task = copy.deepcopy(TASKS.get(task_id))
    if not task:
        return jsonify({"error": "任务不存在。"}), 404
    if task.get("state") == "done" and task.get("file_name"):
        task["download_url"] = url_for("output_file", name=f"{task['job_id']}/{task['file_name']}")
    if task.get("state") == "done" and task.get("report"):
        task["job"] = with_urls(task["report"])
        task.pop("report", None)
    return jsonify(task)


def run_export_task(task_id: str, job_id: str, kind: str):
    def progress(message: str, current: int | None = None, total: int | None = None):
        with TASK_LOCK:
            task = TASKS.get(task_id)
            if not task:
                return
            task["state"] = "running"
            task["message"] = message
            if current is not None:
                task["current"] = current
            if total is not None:
                task["total"] = total

    try:
        progress("正在生成分页详情图", 0, None)
        report, file_name = export_job_output(job_id=job_id, kind=kind, output_root=OUTPUT_ROOT, progress_callback=progress)
        with TASK_LOCK:
            TASKS[task_id].update(
                {
                    "state": "done",
                    "message": "生成完成",
                    "file_name": file_name,
                    "report": report,
                }
            )
    except Exception as exc:
        with TASK_LOCK:
            TASKS[task_id].update({"state": "error", "message": f"生成失败：{exc}"})


@app.post("/api/reprocess")
def reprocess_job():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get("job_id", ""))
    if not job_id:
        return jsonify({"error": "缺少任务 ID。"}), 400
    try:
        report = reprocess_existing_job(
            job_id=job_id,
            image_method=str(data.get("image_method", "")) or None,
            output_root=OUTPUT_ROOT,
        )
    except ProcessingError as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"job": with_urls(report)})


@app.get("/outputs/<path:name>")
def output_file(name: str):
    return send_from_directory(OUTPUT_ROOT, name)


if __name__ == "__main__":
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get("APP_PORT", "5000"))
    debug = os.environ.get("APP_DEBUG", "0") == "1"
    app.run(host="127.0.0.1", port=port, debug=debug)
