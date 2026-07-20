"""
Digit 2W Converter - Flask API
Endpoints:
  POST /api/process     - Upload Excel + params -> kicks off a background job, returns job_id
  GET  /api/job/<job_id> - Poll job status/progress; once done, returns the same
                            payload /api/process used to return synchronously
  GET  /api/states      - List available states from an uploaded file
  GET  /api/download/<filename> - Download a generated file
  GET  /api/files       - List all generated files in current session

Processing now runs in a background thread instead of inside the request
handler. Render's proxy enforces a hard ~100s timeout on any single HTTP
request, and a synchronous /api/process call for "all states" easily blows
past that (plus cold-start time on a sleeping free-tier instance). The
frontend was also just showing a fake local timer as "progress" while it
waited, capped at 0% forever, since it had no way to see real progress from
a request that hadn't returned yet. This version returns a job_id almost
immediately and the frontend polls /api/job/<job_id> for real progress.
"""

import os
import json
import threading
import uuid
import time
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from processor import load_input_data, get_all_states, process_all, generate_for_state, write_output_excel
from processor import get_all_tata_states, process_all_tata, build_tata_rto_map

app = Flask(__name__)
CORS(app)

UPLOAD_DIR = "/tmp/digit_uploads"
OUTPUT_DIR = "/tmp/digit_outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ──────────────────────────────────────────────────────────────
# In-memory job store. Fine for a single-instance deployment; if this
# ever runs on more than one worker/dyno, swap this for something shared
# (redis, a DB row, etc) since each process would have its own JOBS dict.
JOBS = {}
JOBS_LOCK = threading.Lock()

JOB_TTL_SECONDS = 60 * 60  # prune finished jobs after an hour


def _set_job(job_id, **fields):
    with JOBS_LOCK:
        JOBS[job_id].update(fields)


def _get_job(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        return dict(job) if job else None


def _prune_old_jobs():
    cutoff = time.time() - JOB_TTL_SECONDS
    with JOBS_LOCK:
        stale = [jid for jid, j in JOBS.items() if j.get("created_at", 0) < cutoff]
        for jid in stale:
            JOBS.pop(jid, None)


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

    lc = request.form.get("lc", "digit")

    session_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}_{f.filename}")
    f.save(save_path)

    try:
        if lc == "tata":
            # Confirm the uploaded grid has the TW sheet the parser needs.
            from openpyxl import load_workbook as _lw
            wb = _lw(save_path, read_only=True)
            if "TW" not in wb.sheetnames:
                return jsonify({"error": "Uploaded file has no 'TW' sheet"}), 400
            wb.close()

            # States come from the RTO-TW Mapper master table, uploaded fresh
            # each month rather than bundled with the code.
            if "rto_master_file" not in request.files:
                return jsonify({"error": "TATA AIG requires the monthly RTO master file (rto_master_file)"}), 400
            mf = request.files["rto_master_file"]
            if not mf.filename.endswith(".xlsx"):
                return jsonify({"error": "RTO master file must be .xlsx"}), 400
            master_path = os.path.join(UPLOAD_DIR, f"{session_id}_master_{mf.filename}")
            mf.save(master_path)

            rto_map = build_tata_rto_map(master_path)
            states = get_all_tata_states(rto_map)
        else:
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
def _run_process_job(job_id, lc, save_path, master_path, effect_start, effect_end,
                      states, output_mode, combined_filename, session_id):
    """Runs in a background thread. Mirrors what /api/process used to do
    inline, but writes progress/results into the JOBS dict instead of
    returning them straight to a waiting HTTP response."""

    def progress_cb(msg, current, total):
        job = _get_job(job_id)
        log = (job or {}).get("progress", [])
        log = log + [{"message": msg, "current": current, "total": total}]
        _set_job(job_id, progress=log)

    out_dir = os.path.join(OUTPUT_DIR, session_id)
    os.makedirs(out_dir, exist_ok=True)

    try:
        if lc == "tata":
            generated = process_all_tata(
                save_path, out_dir,
                effect_start, effect_end,
                master_path,
                states=states,
                progress_callback=progress_cb,
                output_mode=output_mode,
                combined_filename=combined_filename
            )
        else:
            generated = process_all(
                save_path, out_dir,
                effect_start, effect_end,
                states=states,
                progress_callback=progress_cb,
                output_mode=output_mode,
                combined_filename=combined_filename
            )

        import zipfile
        zip_prefix = "TATA_2W" if lc == "tata" else "Digit_2W"
        zip_path = os.path.join(out_dir, f"{zip_prefix}_{session_id}.zip")
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

        job = _get_job(job_id) or {}
        result = {
            "session_id": session_id,
            "files_generated": len(generated),
            "files": file_info,
            "zip": zip_info,
            "progress": job.get("progress", [])
        }
        _set_job(job_id, status="done", result=result)

    except Exception as e:
        import traceback
        _set_job(job_id, status="error", error=str(e), trace=traceback.format_exc())


@app.route("/api/process", methods=["POST"])
def process():
    """
    Kicks off Excel generation in a background thread and returns a job_id
    immediately. Poll GET /api/job/<job_id> for progress and the final result.

    Form fields:
      file             - The input .xlsx file
      effect_start     - e.g. "2026-02-01"
      effect_end       - e.g. "2026-02-28"
      states           - JSON array of state names (optional; omit for all)
      output_mode      - "per_state" (default), "combined", or "both"
      combined_filename- optional custom filename for the combined workbook
    """
    _prune_old_jobs()

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    effect_start = request.form.get("effect_start", "")
    effect_end = request.form.get("effect_end", "")

    if not effect_start or not effect_end:
        return jsonify({"error": "effect_start and effect_end are required"}), 400

    states_raw = request.form.get("states", None)
    states = json.loads(states_raw) if states_raw else None

    output_mode = request.form.get("output_mode", "per_state")
    if output_mode not in ("per_state", "combined", "both"):
        return jsonify({"error": "output_mode must be one of: per_state, combined, both"}), 400
    combined_filename = request.form.get("combined_filename") or None

    lc = request.form.get("lc", "digit")

    session_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}_{f.filename}")
    f.save(save_path)

    master_path = None
    if lc == "tata":
        if "rto_master_file" not in request.files:
            return jsonify({"error": "TATA AIG requires the monthly RTO master file (rto_master_file)"}), 400
        mf = request.files["rto_master_file"]
        if not mf.filename.endswith(".xlsx"):
            return jsonify({"error": "RTO master file must be .xlsx"}), 400
        master_path = os.path.join(UPLOAD_DIR, f"{session_id}_master_{mf.filename}")
        mf.save(master_path)

    job_id = str(uuid.uuid4())[:12]
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "running",
            "progress": [],
            "result": None,
            "error": None,
            "created_at": time.time(),
        }

    thread = threading.Thread(
        target=_run_process_job,
        args=(job_id, lc, save_path, master_path, effect_start, effect_end,
              states, output_mode, combined_filename, session_id),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id, "session_id": session_id})


# ──────────────────────────────────────────────────────────────
@app.route("/api/job/<job_id>", methods=["GET"])
def job_status(job_id):
    job = _get_job(job_id)
    if job is None:
        return jsonify({"error": "Job not found (it may have expired)"}), 404

    if job["status"] == "error":
        return jsonify({
            "status": "error",
            "error": job.get("error", "Unknown error"),
            "trace": job.get("trace"),
            "progress": job.get("progress", []),
        }), 200

    if job["status"] == "done":
        return jsonify({"status": "done", **job["result"]})

    return jsonify({"status": "running", "progress": job.get("progress", [])})


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
