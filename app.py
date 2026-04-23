#!/usr/bin/env python3
"""
Web interface for the Newspaper Analyzer tool.
Run with: python3 app.py
Then open http://localhost:5000 in your browser.
"""

import json
import os
import queue
import threading
import uuid

from flask import Flask, render_template, request, send_from_directory, Response, jsonify

from newspaper_analyzer import (
    pdf_to_images, extract_metadata, analyze_page,
    export_results, parse_page_selection,
)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "uploads")
app.config["OUTPUT_FOLDER"] = os.path.join(os.path.dirname(__file__), "outputs")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

jobs = {}


def run_analysis(job_id: str, pdf_path: str, pages_arg: str, fmt: str):
    """Run the full analysis in a background thread, posting progress updates."""
    q = jobs[job_id]["queue"]

    def log(msg):
        q.put({"type": "log", "message": msg})

    try:
        log("Converting PDF to images...")
        images = pdf_to_images(pdf_path)
        log(f"Found {len(images)} page(s).")

        page_indices = parse_page_selection(pages_arg, len(images))
        if not page_indices:
            raise ValueError("No valid pages selected.")

        log("Extracting newspaper name and date...")
        metadata = extract_metadata(images[page_indices[0]])
        log(f"Newspaper: {metadata['newspaper_name']}  |  Date: {metadata['date']}")

        all_rows = []
        for idx in page_indices:
            page_num = idx + 1
            log(f"Analysing page {page_num} of {len(images)}...")
            rows = analyze_page(images[idx], page_num, metadata)
            all_rows.extend(rows)
            log(f"  → {len(rows)} element(s) found on page {page_num}.")

        output_base = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
        export_results(all_rows, output_base, fmt)

        files = []
        if fmt in ("csv", "both") and os.path.exists(output_base + ".csv"):
            files.append({"name": "results.csv", "url": f"/download/{job_id}/results.csv"})
        if fmt in ("xlsx", "both") and os.path.exists(output_base + ".xlsx"):
            files.append({"name": "results.xlsx", "url": f"/download/{job_id}/results.xlsx"})

        log(f"Done! {len(all_rows)} total elements extracted.")
        q.put({"type": "done", "files": files, "total": len(all_rows),
               "newspaper": metadata["newspaper_name"], "date": metadata["date"]})

    except Exception as exc:
        q.put({"type": "error", "message": str(exc)})
    finally:
        jobs[job_id]["finished"] = True


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    if "pdf" not in request.files or request.files["pdf"].filename == "":
        return jsonify({"error": "No PDF file uploaded."}), 400

    pdf_file = request.files["pdf"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400

    pages_arg = request.form.get("pages", "").strip()
    fmt = request.form.get("format", "both")

    job_id = str(uuid.uuid4())
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

    pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{job_id}.pdf")
    pdf_file.save(pdf_path)

    jobs[job_id] = {"queue": queue.Queue(), "finished": False}

    thread = threading.Thread(
        target=run_analysis,
        args=(job_id, pdf_path, pages_arg, fmt),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/stream/<job_id>")
def stream(job_id):
    """Server-Sent Events endpoint — streams progress logs to the browser."""
    if job_id not in jobs:
        return "Job not found", 404

    def generate():
        q = jobs[job_id]["queue"]
        while True:
            try:
                event = q.get(timeout=30)
                yield f"data: {json.dumps(event)}\n\n"
                if event["type"] in ("done", "error"):
                    break
            except queue.Empty:
                yield "data: {\"type\": \"ping\"}\n\n"
                if jobs[job_id]["finished"]:
                    break

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    safe_name = os.path.basename(filename)
    actual = safe_name.replace("results.", f"{job_id}.")
    return send_from_directory(app.config["OUTPUT_FOLDER"], actual, as_attachment=True,
                               download_name=safe_name)


if __name__ == "__main__":
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting Newspaper Analyzer at http://localhost:{port}")
    app.run(debug=False, host="0.0.0.0", port=port)
