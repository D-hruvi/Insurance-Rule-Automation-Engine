"""
Digit 2W Converter — AI Edition — Flask API
Endpoints:
  POST /api/process              - Upload Excel + params → kicks off a background job, returns session_id immediately
  GET  /api/status/<session_id>  - Poll job status/progress; once done, includes the same payload /api/process used to return
  GET  /api/states                - List available states from an uploaded file
  GET  /api/download/<session_id>/<filename> - Download a generated file
  GET  /api/files/<session_id>   - List all generated files in a session
  GET  /api/validate/<session_id>/<filename>  - Re-run validator on a generated file, get report

IMPORTANT — deployment note:
Processing now runs in a background thread instead of inside the request
handler, because a full run can mean hundreds of sequential Groq calls and
easily takes 10-60+ minutes. Job state lives in an in-memory dict, so this
only works correctly with a SINGLE worker PROCESS (threads are fine, extra
gunicorn *workers* are not — they don't share memory). On Render, set the
start command to something like:
    gunicorn app_ai:app --bind 0.0.0.0:$PORT --workers 1 --worker-class gthread --threads 4 --timeout 120
"""

import os
import json
import zipfile
import uuid
import threading
import traceback
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

# session_id -> {"status": "queued"|"running"|"done"|"error",
#                "progress": [...], "result": {...}|None, "error": str|None}
JOBS = {}
JOBS_LOCK = threading.Lock()


@app.route("/", methods=["GET"])
def index():
    return jsonify({
        "service": "Digit 2W Converter — AI Edition API",
        "status": "running",
        "note": "This is an API-only service, no HTML frontend is served here. "
                 "Point your separately-deployed React frontend's API base URL "
                 "at this host.",
        "endpoints": [
            "POST /api/states",
            "POST /api/process",
            "GET /api/status/<session_id>",
            "GET /api/files/<session_id>",
            "GET /api/download/<session_id>/<filename>",
            "GET /api/validate/<session_id>/<filename>",
            "GET /api/health",
        ],
    })


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

    Returns immediately with a session_id + status_url. The actual AI
    processing happens in a background thread — poll status_url for
    progress and the final result (this can take many minutes for a
    full multi-state run, so don't block on this request).
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

    with JOBS_LOCK:
        JOBS[session_id] = {"status": "queued", "progress": [], "result": None, "error": None}

    thread = threading.Thread(
        target=_run_job,
        args=(session_id, save_path, out_dir, effect_start, effect_end, states),
        daemon=True,
    )
    thread.start()

    return jsonify({
        "session_id": session_id,
        "status": "queued",
        "status_url": f"/api/status/{session_id}",
    }), 202


def _run_job(session_id, save_path, out_dir, effect_start, effect_end, states):
    """Runs the actual AI pipeline off the request thread. Updates JOBS[session_id]
    as it goes so /api/status can report live progress."""

    def progress_cb(msg, current, total):
        with JOBS_LOCK:
            JOBS[session_id]["progress"].append({"message": msg, "current": current, "total": total})

    with JOBS_LOCK:
        JOBS[session_id]["status"] = "running"

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

        result = {
            "session_id": session_id,
            "files_generated": len(generated),
            "files": file_info,
            "zip": zip_info,
            "validation": {
                "total_violations": total_violations,
                "per_file": validation_summary,
            },
        }
        with JOBS_LOCK:
            JOBS[session_id]["status"] = "done"
            JOBS[session_id]["result"] = result

    except Exception as e:
        with JOBS_LOCK:
            JOBS[session_id]["status"] = "error"
            JOBS[session_id]["error"] = str(e)
            JOBS[session_id]["trace"] = traceback.format_exc()


@app.route("/api/status/<session_id>", methods=["GET"])
def job_status(session_id):
    """Poll this while a job runs. Recommended: every 3-5s — frequent
    polling also keeps a Render free-tier instance from spinning down
    mid-job, since spin-down is based on inbound request inactivity."""
    with JOBS_LOCK:
        job = JOBS.get(session_id)
        if job is None:
            return jsonify({"error": "Unknown session_id"}), 404
        recent_progress = job["progress"][-20:]
        return jsonify({
            "session_id": session_id,
            "status": job["status"],
            "latest_progress": recent_progress[-1] if recent_progress else None,
            "recent_progress": recent_progress,
            "progress_count": len(job["progress"]),
            "result": job["result"] if job["status"] == "done" else None,
            "error": job["error"] if job["status"] == "error" else None,
        })


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
    # Local dev only. In production (Render), run via gunicorn with a single
    # worker process + threads, e.g.:
    #   gunicorn app_ai:app --bind 0.0.0.0:$PORT --workers 1 --worker-class gthread --threads 4 --timeout 120
    app.run(debug=True, host="0.0.0.0", port=5060, threaded=True)
