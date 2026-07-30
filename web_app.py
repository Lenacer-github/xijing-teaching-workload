"""Bootstrap frontend with a small Flask API backend."""

from __future__ import annotations

import io
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file, send_from_directory

from excel_export import build_workload_excel
from workload_domain import (
    ACADEMIC_TERM_OPTIONS,
    APP_VERSION,
    DEFAULT_ACADEMIC_TERM,
    DEPARTMENT_OPTIONS,
    make_excel_filename,
    normalize_payload,
    prepare_export_data,
    public_catalog,
)


BASE_DIR = Path(__file__).resolve().parent
app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024
app.json.ensure_ascii = False
app.json.sort_keys = False


@app.get("/")
def index():
    catalog = public_catalog()
    return render_template(
        "index.html",
        app_config={
            "version": APP_VERSION,
            "departments": DEPARTMENT_OPTIONS,
            "terms": ACADEMIC_TERM_OPTIONS,
            "defaultTerm": DEFAULT_ACADEMIC_TERM,
            "categories": list(catalog),
            "catalog": catalog,
        },
    )


@app.get("/assets/<path:filename>")
def assets(filename: str):
    return send_from_directory(BASE_DIR / "assets", filename)


@app.get("/health")
def health():
    return jsonify({"status": "ok", "version": APP_VERSION})


@app.post("/api/restore")
def restore_progress():
    try:
        normalized = normalize_payload(
            request.get_json(force=True), require_complete=False
        )
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    return jsonify({"ok": True, "data": normalized})


@app.post("/api/export")
def export_excel():
    try:
        export_data = prepare_export_data(request.get_json(force=True))
        excel_bytes = build_workload_excel(export_data)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400
    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=make_excel_filename(export_data),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@app.errorhandler(413)
def payload_too_large(_error):
    return jsonify({"ok": False, "message": "暂存文件不能超过 2MB"}), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8601, debug=True)
