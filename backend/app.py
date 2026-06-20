"""
Digit 2W Converter - Flask API
Endpoints:
  POST /api/process   - Upload Excel + params → generate output files
  GET  /api/states    - List available states from an uploaded file
  GET  /api/download/<filename> - Download a generated file
  GET  /api/files     - List all generated files in current session
"""

import os
import json
import tempfile
import zipfile
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from processor import load_input_data, get_all_states, process_all, generate_for_state, write_output_excel

app = Flask(__name__)
CORS(app)

UPLOAD_DIR = "/tmp/digit_uploads"
OUTPUT_DIR = "/tmp/digit_outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ──────────────────────────────────────────────────────────────
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "timestamp": datetime.now().isoformat()})


# ──────────────────────────────────────────────────────────────
@app.route("/api/states", methods=["POST"])
def get_states():
    """Upload an Excel file and get all available states."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported"}), 400

    session_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}_{f.filename}")
    f.save(save_path)

    try:
        d = load_input_data(save_path)
        states = get_all_states(d)
        return jsonify({
            "session_id": session_id,
            "file_path": save_path,
            "states": states,
            "total_states": len(states)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ──────────────────────────────────────────────────────────────
@app.route("/api/process", methods=["POST"])
def process():
    """
    Process Excel and generate output files.
    Form fields:
      file          - The input .xlsx file
      effect_start  - e.g. "2026-02-01"
      effect_end    - e.g. "2026-02-28"
      states        - JSON array of state names (optional; omit for all)
    """
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

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
        generated = process_all(
            save_path, out_dir,
            effect_start, effect_end,
            states=states,
            progress_callback=progress_cb
        )

        # Create a zip of all output files
        zip_path = os.path.join(out_dir, f"Digit_2W_{session_id}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in generated:
                zf.write(fp, os.path.basename(fp))

        file_info = []
        for fp in generated:
            fname = os.path.basename(fp)
            size = os.path.getsize(fp)
            file_info.append({
                "filename": fname,
                "size_bytes": size,
                "download_url": f"/api/download/{session_id}/{fname}"
            })

        zip_info = {
            "filename": os.path.basename(zip_path),
            "size_bytes": os.path.getsize(zip_path),
            "download_url": f"/api/download/{session_id}/{os.path.basename(zip_path)}"
        }

        return jsonify({
            "session_id": session_id,
            "files_generated": len(generated),
            "files": file_info,
            "zip": zip_info,
            "progress": progress_log
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ──────────────────────────────────────────────────────────────
@app.route("/api/download/<session_id>/<filename>", methods=["GET"])
def download(session_id, filename):
    file_path = os.path.join(OUTPUT_DIR, session_id, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)


# ──────────────────────────────────────────────────────────────
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
            "download_url": f"/api/download/{session_id}/{fname}"
        })
    return jsonify({"session_id": session_id, "files": files})


if __name__ == "__main__":
    print("Starting Digit 2W Converter API on port 5050...")
    app.run(debug=True, host="0.0.0.0", port=5050)
