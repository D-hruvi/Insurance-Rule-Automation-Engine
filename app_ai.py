"""
Digit 2W Converter — AI Edition — Flask API
Endpoints:
  POST /api/process        - Upload Excel + params → AI-generates output files + validation report
  GET  /api/states          - List available states from an uploaded file
  GET  /api/download/<session_id>/<filename> - Download a generated file
  GET  /api/files/<session_id> - List all generated files in a session
  GET  /api/validate/<session_id>/<filename>  - Re-run validator on a generated file, get report
"""

import os
import json
import zipfile
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from ai_processor import (
    load_raw_sheets, build_rto_lookup, get_all_states, process_all_ai,
)
from validate_output import validate_file

app = Flask(__name__)
CORS(app)

UPLOAD_DIR = "/tmp/digit_ai_uploads"
OUTPUT_DIR = "/tmp/digit_ai_outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/api/health", methods=["GET"])
def health():
    groq_key_set = bool(os.environ.get("GROQ_API_KEY"))
    return jsonify({
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "groq_api_key_configured": groq_key_set,
    })


@app.route("/api/states", methods=["POST"])
def get_states():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported"}), 400

    session_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}_{f.filename}")
    f.save(save_path)

    try:
        raw = load_raw_sheets(save_path)
        _, rto_to_state = build_rto_lookup(raw["rto_rows"])
        states = get_all_states(rto_to_state)
        return jsonify({
            "session_id": session_id,
            "file_path": save_path,
            "states": states,
            "total_states": len(states),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/process", methods=["POST"])
def process():
    """
    Form fields:
      file          - input .xlsx
      effect_start  - e.g. "2026-02-01"
      effect_end    - e.g. "2026-02-28"
      states        - JSON array of state names (optional; omit for all)
    """
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    if not os.environ.get("GROQ_API_KEY"):
        return jsonify({
            "error": "GROQ_API_KEY not configured on server. "
                     "Get a free key at https://console.groq.com/keys"
        }), 500

    f = request.files["file"]
    effect_start = request.form.get("effect_start", "")
    effect_end = request.form.get("effect_end", "")
    if not effect_start or not effect_end:
        return jsonify({"error": "effect_start and effect_end are required"}), 400

    states_raw = request.form.get("states", None)
    states = json.loads(states_raw) if states_raw else None

    session_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}_{f.filename}")
    f.save(save_path)

    out_dir = os.path.join(OUTPUT_DIR, session_id)
    os.makedirs(out_dir, exist_ok=True)

    progress_log = []

    def progress_cb(msg, current, total):
        progress_log.append({"message": msg, "current": current, "total": total})

    try:
        generated = process_all_ai(
            save_path, out_dir, effect_start, effect_end,
            states=states, progress_callback=progress_cb,
        )

        zip_path = os.path.join(out_dir, f"Digit_2W_AI_{session_id}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in generated:
                zf.write(fp, os.path.basename(fp))

        file_info = []
        validation_summary = {}
        for fp in generated:
            fname = os.path.basename(fp)
            size = os.path.getsize(fp)
            violations = validate_file(fp)
            validation_summary[fname] = {
                "violation_count": len(violations),
                "violations": violations[:50],  # cap to keep response sane
            }
            file_info.append({
                "filename": fname,
                "size_bytes": size,
                "download_url": f"/api/download/{session_id}/{fname}",
                "violation_count": len(violations),
            })

        zip_info = {
            "filename": os.path.basename(zip_path),
            "size_bytes": os.path.getsize(zip_path),
            "download_url": f"/api/download/{session_id}/{os.path.basename(zip_path)}",
        }

        total_violations = sum(v["violation_count"] for v in validation_summary.values())

        return jsonify({
            "session_id": session_id,
            "files_generated": len(generated),
            "files": file_info,
            "zip": zip_info,
            "progress": progress_log,
            "validation": {
                "total_violations": total_violations,
                "per_file": validation_summary,
            },
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/validate/<session_id>/<filename>", methods=["GET"])
def validate(session_id, filename):
    file_path = os.path.join(OUTPUT_DIR, session_id, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    violations = validate_file(file_path)
    return jsonify({
        "filename": filename,
        "violation_count": len(violations),
        "violations": violations,
    })


@app.route("/api/download/<session_id>/<filename>", methods=["GET"])
def download(session_id, filename):
    file_path = os.path.join(OUTPUT_DIR, session_id, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)


@app.route("/api/files/<session_id>", methods=["GET"])
def list_files(session_id):
    out_dir = os.path.join(OUTPUT_DIR, session_id)
    if not os.path.exists(out_dir):
        return jsonify({"error": "Session not found"}), 404
    files = []
    for fname in sorted(os.listdir(out_dir)):
        fp = os.path.join(out_dir, fname)
        files.append({
            "filename": fname,
            "size_bytes": os.path.getsize(fp),
            "download_url": f"/api/download/{session_id}/{fname}",
        })
    return jsonify({"session_id": session_id, "files": files})


if __name__ == "__main__":
    print("Starting Digit 2W AI Converter API on port 5060...")
    if not os.environ.get("GROQ_API_KEY"):
        print("WARNING: GROQ_API_KEY not set. Set it before calling /api/process.")
    app.run(debug=True, host="0.0.0.0", port=5060)
