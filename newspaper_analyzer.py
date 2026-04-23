#!/usr/bin/env python3
"""
Newspaper Analyzer — reads a PDF newspaper and classifies text elements
(headline, subheadline, main_text, image_caption) using Claude Vision API.
Also extracts newspaper name and publication date.
Exports results as CSV and/or Excel spreadsheet.
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

MODEL = "claude-opus-4-7"

METADATA_PROMPT = (
    "Look at this newspaper page. Return ONLY a JSON object with two keys: "
    "'newspaper_name' (the publication title, e.g. 'The New York Times') and "
    "'date' (the publication date as YYYY-MM-DD, or the text as printed if the "
    "exact format is unclear). No explanation, no markdown — just the JSON object."
)

CLASSIFICATION_PROMPT = (
    "Analyze this newspaper page image. Identify and extract every piece of text "
    "and classify each one as exactly one of: headline, subheadline, main_text, "
    "image_caption. Return ONLY a JSON array — no explanation, no markdown. "
    "Each element must have this shape: "
    "{\"type\": \"<category>\", \"text\": \"<exact text as it appears on the page>\"}."
)


def pdf_to_images(pdf_path: str) -> list:
    """Convert each PDF page to a PIL Image."""
    print(f"Converting PDF to images: {pdf_path}")
    images = convert_from_path(pdf_path, dpi=200)
    print(f"  {len(images)} page(s) found.")
    return images


def image_to_base64(image: Image.Image) -> str:
    """Encode a PIL Image as a base64 JPEG string."""
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=90)
    return base64.standard_b64encode(buffer.getvalue()).decode("utf-8")


def call_claude(client: anthropic.Anthropic, image_b64: str, prompt: str) -> str:
    """Send an image + prompt to Claude and return the raw text response."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        messages=[
            {
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
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    )
    return response.content[0].text.strip()


def extract_metadata(client: anthropic.Anthropic, image: Image.Image) -> dict:
    """Extract newspaper name and date from the first page."""
    print("  Extracting newspaper name and date...")
    image_b64 = image_to_base64(image)
    raw = call_claude(client, image_b64, METADATA_PROMPT)
    try:
        raw = raw.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        metadata = json.loads(raw)
        newspaper_name = metadata.get("newspaper_name", "Unknown")
        date = metadata.get("date", "Unknown")
    except (json.JSONDecodeError, KeyError):
        print(f"  Warning: could not parse metadata response — {raw[:120]}")
        newspaper_name, date = "Unknown", "Unknown"
    print(f"  Newspaper: {newspaper_name} | Date: {date}")
    return {"newspaper_name": newspaper_name, "date": date}


def analyze_page(
    client: anthropic.Anthropic,
    image: Image.Image,
    page_num: int,
    metadata: dict,
) -> list:
    """Classify all text elements on a single page. Returns list of row dicts."""
    print(f"  Classifying text elements on page {page_num}...")
    image_b64 = image_to_base64(image)
    raw = call_claude(client, image_b64, CLASSIFICATION_PROMPT)
    try:
        raw = raw.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        elements = json.loads(raw)
        if not isinstance(elements, list):
            raise ValueError("Expected a JSON array")
    except (json.JSONDecodeError, ValueError) as exc:
        print(f"  Warning: could not parse page {page_num} response ({exc}) — skipping page.")
        return []

    rows = []
    for el in elements:
        if not isinstance(el, dict):
            continue
        rows.append(
            {
                "newspaper_name": metadata["newspaper_name"],
                "date": metadata["date"],
                "page": page_num,
                "type": el.get("type", "unknown"),
                "text": el.get("text", ""),
            }
        )
    print(f"    {len(rows)} element(s) extracted.")
    return rows


def export_csv(rows: list, path: str) -> None:
    """Write rows to a CSV file."""
    fieldnames = ["newspaper_name", "date", "page", "type", "text"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"CSV saved: {path}")


def export_xlsx(rows: list, path: str) -> None:
    """Write rows to an Excel .xlsx file."""
    if not XLSX_AVAILABLE:
        print("openpyxl not installed — skipping .xlsx export.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Newspaper Analysis"
    headers = ["newspaper_name", "date", "page", "type", "text"]
    ws.append(headers)
    for row in rows:
        ws.append([row[h] for h in headers])
    # Auto-width for readability
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
    """Parse a comma-separated page list like '1,3,5' into 0-based indices."""
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
        description="Analyze a PDF newspaper with Claude Vision and export categorized text."
    )
    parser.add_argument("--pdf", required=True, help="Path to the newspaper PDF file.")
    parser.add_argument(
        "--output",
        default="results",
        help="Base path for output files (default: results → results.csv / results.xlsx).",
    )
    parser.add_argument(
        "--pages",
        default="",
        help="Comma-separated page numbers to analyze, e.g. '1,2,3' (default: all pages).",
    )
    parser.add_argument(
        "--format",
        choices=["csv", "xlsx", "both"],
        default="both",
        help="Export format (default: both).",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.pdf):
        sys.exit(f"Error: file not found: {args.pdf}")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("Error: ANTHROPIC_API_KEY environment variable is not set.")

    client = anthropic.Anthropic(api_key=api_key)

    images = pdf_to_images(args.pdf)
    page_indices = parse_page_selection(args.pages, len(images))

    if not page_indices:
        sys.exit("No valid pages to analyze.")

    # Extract metadata from the first selected page (usually the front page)
    metadata = extract_metadata(client, images[page_indices[0]])

    all_rows = []
    for idx in page_indices:
        page_num = idx + 1
        print(f"\nPage {page_num}/{len(images)}:")
        rows = analyze_page(client, images[idx], page_num, metadata)
        all_rows.extend(rows)

    print(f"\nTotal elements extracted: {len(all_rows)}")
    export_results(all_rows, args.output, args.format)
    print("Done.")


if __name__ == "__main__":
    main()
