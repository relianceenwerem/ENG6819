#!/usr/bin/env python3
"""
Web interface for the Newspaper Analyzer tool.
Run with: python3 app.py
Then open http://localhost:5000 in your browser.
Requires ANTHROPIC_API_KEY environment variable to be set.
"""

import json
import os
import queue
import threading
import uuid

import anthropic
from flask import Flask, render_template, request, send_from_directory, Response, jsonify

from newspaper_analyzer import (
    pdf_to_images, analyze_page, export_results, parse_page_selection,
)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "uploads")
app.config["OUTPUT_FOLDER"] = os.path.join(os.path.dirname(__file__), "outputs")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

jobs = {}


def run_analysis(job_id: str, pdf_entries: list, pages_arg: str, fmt: str):
    """
    Run analysis on one or more PDFs in a background thread.
    pdf_entries: list of (pdf_path, original_filename)
    """
    q = jobs[job_id]["queue"]

    def log(msg):
        q.put({"type": "log", "message": msg})

    try:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY is not set on the server.")
        client = anthropic.Anthropic(api_key=api_key)

        all_rows        = []
        newspapers_seen = []
        total_cost      = 0.0
        total_in_tokens = 0
        total_out_tokens = 0

        for file_num, (pdf_path, original_filename) in enumerate(pdf_entries, 1):
            filename = os.path.splitext(original_filename)[0] if original_filename else job_id
            log(f"── File {file_num}/{len(pdf_entries)}: {original_filename}")

            log("  Converting PDF to images...")
            images = pdf_to_images(pdf_path)
            log(f"  {len(images)} page(s) found.")

            page_indices = parse_page_selection(pages_arg, len(images))
            if not page_indices:
                log(f"  Warning: no valid pages for {original_filename}, skipping.")
                continue

            for idx in page_indices:
                page_num = idx + 1
                log(f"  Analysing page {page_num}/{len(images)} with Claude Vision...")
                rows, usage = analyze_page(client, images[idx], page_num, filename)
                all_rows.extend(rows)
                total_cost       += usage["cost_usd"]
                total_in_tokens  += usage["input_tokens"]
                total_out_tokens += usage["output_tokens"]
                if rows:
                    newspapers_seen.append(rows[0]["newspaper_name"])
                img_flag = rows[0]["image_detected"] if rows else "No"
                log(f"    → {len(rows)} element(s) found. Image detected: {img_flag}. "
                    f"Page cost: ${usage['cost_usd']:.4f}")

        if not all_rows:
            raise ValueError("No elements could be extracted from the uploaded files.")

        output_base = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
        export_results(all_rows, output_base, fmt)

        files = []
        if fmt in ("csv", "both") and os.path.exists(output_base + ".csv"):
            files.append({"name": "results.csv", "url": f"/download/{job_id}/results.csv"})
        if fmt in ("xlsx", "both") and os.path.exists(output_base + ".xlsx"):
            files.append({"name": "results.xlsx", "url": f"/download/{job_id}/results.xlsx"})

        summary = ", ".join(dict.fromkeys(newspapers_seen)) or "Unknown"
        log(f"Done! {len(all_rows)} total elements from {len(pdf_entries)} file(s). "
            f"Total cost: ${total_cost:.4f} ({total_in_tokens} in / {total_out_tokens} out tokens)")
        q.put({"type": "done", "files": files, "total": len(all_rows),
               "newspaper": summary, "date": f"{len(pdf_entries)} file(s) processed",
               "cost": f"${total_cost:.4f}",
               "tokens": f"{total_in_tokens:,} in / {total_out_tokens:,} out"})

    except Exception as exc:
        q.put({"type": "error", "message": str(exc)})
    finally:
        jobs[job_id]["finished"] = True


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    uploaded_files = request.files.getlist("pdf")
    uploaded_files = [f for f in uploaded_files if f.filename]

    if not uploaded_files:
        return jsonify({"error": "No PDF files uploaded."}), 400

    invalid = [f.filename for f in uploaded_files if not f.filename.lower().endswith(".pdf")]
    if invalid:
        return jsonify({"error": f"Only PDF files are supported. Invalid: {', '.join(invalid)}"}), 400

    pages_arg = request.form.get("pages", "").strip()
    fmt = request.form.get("format", "both")

    job_id = str(uuid.uuid4())
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

    pdf_entries = []
    for f in uploaded_files:
        safe_name = f"{job_id}_{len(pdf_entries)}.pdf"
        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
        f.save(pdf_path)
        pdf_entries.append((pdf_path, f.filename))

    jobs[job_id] = {"queue": queue.Queue(), "finished": False}

    thread = threading.Thread(
        target=run_analysis,
        args=(job_id, pdf_entries, pages_arg, fmt),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id, "file_count": len(pdf_entries)})


@app.route("/stream/<job_id>")
def stream(job_id):
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
