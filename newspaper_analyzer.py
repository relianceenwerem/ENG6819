#!/usr/bin/env python3
"""
Newspaper Analyzer — reads newspapers.com PDF clippings using Claude Vision API.
Extracts headlines and subheadlines, detects photographs, and exports to CSV/Excel.
"""

import argparse
import base64
import csv
import io
import json
import os
import sys

import anthropic
from pdf2image import convert_from_path
from PIL import Image

try:
    import openpyxl
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

MODEL = "claude-haiku-4-5-20251001"

ANALYSIS_PROMPT = """\
Analyze this newspaper page image carefully.

Return ONLY a valid JSON object — no explanation, no markdown fences:
{
  "newspaper_name": "full publication title as printed (e.g. Chicago Tribune)",
  "date": "publication date exactly as printed on the page",
  "image_detected": true,
  "elements": [
    {"type": "headline",    "text": "exact headline text"},
    {"type": "subheadline", "text": "exact subheadline text"}
  ]
}

Rules:
- newspaper_name: the masthead / publication title
- date: the publication date (e.g. "Monday, March 30, 2020")
- image_detected: true if any photograph or illustration appears on this page, false otherwise
- elements: ONLY headlines (largest bold text introducing an article) and subheadlines \
(smaller text immediately below a headline summarising the article). \
Exclude body text, captions, bylines, page numbers, and advertisements.
- Preserve exact spelling and punctuation of every text element.
"""


def pdf_to_images(pdf_path: str) -> list:
    print(f"  Converting PDF to images: {pdf_path}")
    images = convert_from_path(pdf_path, dpi=200)
    print(f"  {len(images)} page(s) found.")
    return images


def image_to_base64(image: Image.Image) -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=85)
    return base64.standard_b64encode(buffer.getvalue()).decode("utf-8")


def analyze_page(client: anthropic.Anthropic, image: Image.Image,
                 page_num: int, filename: str) -> list:
    """Send one page to Claude Vision and return structured rows."""
    print(f"  Analysing page {page_num} with Claude Vision...")
    image_b64 = image_to_base64(image)

    response = client.messages.create(
        model=MODEL,
        max_tokens=2048,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": image_b64,
                    },
                },
                {"type": "text", "text": ANALYSIS_PROMPT},
            ],
        }],
    )

    raw = response.content[0].text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        print(f"  Warning: could not parse Claude response for page {page_num} — skipping.")
        return []

    newspaper_name = data.get("newspaper_name", "Unknown")
    date           = data.get("date", "Unknown")
    image_detected = "Yes" if data.get("image_detected", False) else "No"
    elements       = data.get("elements", [])

    rows = []
    for el in elements:
        if not isinstance(el, dict):
            continue
        content_type = el.get("type", "").strip().lower()
        if content_type not in ("headline", "subheadline"):
            continue
        text = el.get("text", "").strip()
        if not text:
            continue
        rows.append({
            "filename":       filename,
            "newspaper_name": newspaper_name,
            "date":           date,
            "content_type":   content_type,
            "text":           text,
            "image_detected": image_detected,
        })

    print(f"    {len(rows)} element(s) found. Image detected: {image_detected}")
    return rows


def export_csv(rows: list, path: str) -> None:
    fieldnames = ["filename", "newspaper_name", "date", "content_type", "text", "image_detected"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"CSV saved: {path}")


def export_xlsx(rows: list, path: str) -> None:
    if not XLSX_AVAILABLE:
        print("openpyxl not installed — skipping .xlsx export.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Newspaper Analysis"
    headers = ["filename", "newspaper_name", "date", "content_type", "text", "image_detected"]
    ws.append(headers)
    for row in rows:
        ws.append([row[h] for h in headers])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 80)
    wb.save(path)
    print(f"Excel saved: {path}")


def export_results(rows: list, output_base: str, fmt: str) -> None:
    if fmt in ("csv", "both"):
        export_csv(rows, output_base + ".csv")
    if fmt in ("xlsx", "both"):
        export_xlsx(rows, output_base + ".xlsx")


def parse_page_selection(pages_arg: str, total: int) -> list:
    if not pages_arg:
        return list(range(total))
    indices = []
    for part in pages_arg.split(","):
        try:
            n = int(part.strip())
            if 1 <= n <= total:
                indices.append(n - 1)
            else:
                print(f"Warning: page {n} out of range (1–{total}), skipping.")
        except ValueError:
            print(f"Warning: invalid page number '{part}', skipping.")
    return indices


def main():
    parser = argparse.ArgumentParser(
        description="Analyze a newspapers.com PDF using Claude Vision API."
    )
    parser.add_argument("--pdf",    required=True, help="Path to the newspaper PDF.")
    parser.add_argument("--output", default="results",
                        help="Base output path (default: results → results.csv / results.xlsx).")
    parser.add_argument("--pages",  default="",
                        help="Comma-separated page numbers (default: all).")
    parser.add_argument("--format", choices=["csv", "xlsx", "both"], default="both",
                        help="Export format (default: both).")
    args = parser.parse_args()

    if not os.path.isfile(args.pdf):
        sys.exit(f"Error: file not found: {args.pdf}")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("Error: ANTHROPIC_API_KEY environment variable is not set.")

    client   = anthropic.Anthropic(api_key=api_key)
    filename = os.path.splitext(os.path.basename(args.pdf))[0]
    images   = pdf_to_images(args.pdf)
    indices  = parse_page_selection(args.pages, len(images))

    if not indices:
        sys.exit("No valid pages to analyze.")

    all_rows = []
    for idx in indices:
        rows = analyze_page(client, images[idx], idx + 1, filename)
        all_rows.extend(rows)

    print(f"\nTotal elements extracted: {len(all_rows)}")
    export_results(all_rows, args.output, args.format)
    print("Done.")


if __name__ == "__main__":
    main()
