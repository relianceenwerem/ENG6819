#!/usr/bin/env python3
"""
Newspaper Analyzer — reads a PDF newspaper and classifies text elements
(headline, subheadline, main_text, image_caption) using Tesseract OCR.
Extracts newspaper name and publication date. Exports as CSV and/or Excel.
Free — no API key required.
"""

import argparse
import csv
import os
import sys

import numpy as np
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

try:
    import openpyxl
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False


def pdf_to_images(pdf_path: str) -> list:
    print(f"Converting PDF to images: {pdf_path}")
    images = convert_from_path(pdf_path, dpi=200)
    print(f"  {len(images)} page(s) found.")
    return images


def _get_blocks(image: Image.Image) -> list:
    """
    Run Tesseract and return a list of text blocks.
    Each block: {"text": str, "avg_height": float, "word_count": int,
                  "top": int, "left": int, "page_h": int}
    """
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    page_h = image.height

    # Group words into lines keyed by (block_num, par_num, line_num)
    lines: dict = {}
    for i, level in enumerate(data["level"]):
        if level != 5:  # word level
            continue
        text = data["text"][i].strip()
        if not text or data["conf"][i] < 10:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append({
            "text": text,
            "height": data["height"][i],
            "top": data["top"][i],
            "left": data["left"][i],
        })

    # Group lines into paragraphs (same block_num + par_num)
    paras: dict = {}
    for (bn, pn, ln), words in lines.items():
        key = (bn, pn)
        paras.setdefault(key, []).append({
            "text": " ".join(w["text"] for w in words),
            "avg_height": float(np.mean([w["height"] for w in words])),
            "top": min(w["top"] for w in words),
            "left": min(w["left"] for w in words),
        })

    blocks = []
    for (bn, pn), para_lines in paras.items():
        full_text = " ".join(l["text"] for l in para_lines).strip()
        if not full_text:
            continue
        avg_h = float(np.mean([l["avg_height"] for l in para_lines]))
        top = min(l["top"] for l in para_lines)
        left = min(l["left"] for l in para_lines)
        blocks.append({
            "text": full_text,
            "avg_height": avg_h,
            "word_count": len(full_text.split()),
            "top": top,
            "left": left,
            "page_h": page_h,
        })

    return blocks


def _classify_blocks(blocks: list) -> list:
    """
    Assign a type to each block based on font-size heuristics.
    Returns list of (type, text) tuples.
    """
    if not blocks:
        return []

    heights = [b["avg_height"] for b in blocks]
    p75 = float(np.percentile(heights, 75))
    p50 = float(np.percentile(heights, 50))
    p25 = float(np.percentile(heights, 25))

    results = []
    for b in blocks:
        h = b["avg_height"]
        wc = b["word_count"]
        relative_top = b["top"] / b["page_h"] if b["page_h"] else 0

        if h >= p75 and wc <= 30:
            category = "headline"
        elif h >= p50 and wc <= 60:
            category = "subheadline"
        elif wc <= 25 and h <= p25 and relative_top > 0.3:
            # Short, small text in the lower portion of the page → likely a caption
            category = "image_caption"
        else:
            category = "main_text"

        results.append((category, b["text"]))

    return results


def extract_metadata(image: Image.Image) -> dict:
    """
    Heuristically extract newspaper name and date from the first page.
    The masthead (name) is usually the largest text at the very top.
    The date is usually a short line near the top containing digits.
    """
    print("  Extracting newspaper name and date...")
    blocks = _get_blocks(image)
    if not blocks:
        return {"newspaper_name": "Unknown", "date": "Unknown"}

    # Sort by position — top of page first
    top_blocks = sorted(blocks, key=lambda b: b["top"])

    # Newspaper name: largest font-size block in the top 20% of the page
    top_zone = [b for b in top_blocks if b["top"] / image.height < 0.20]
    if top_zone:
        name_block = max(top_zone, key=lambda b: b["avg_height"])
        newspaper_name = name_block["text"]
    else:
        newspaper_name = top_blocks[0]["text"] if top_blocks else "Unknown"

    # Date: short block in the top 30% containing a 4-digit year or date pattern
    import re
    date = "Unknown"
    date_zone = [b for b in top_blocks if b["top"] / image.height < 0.30]
    for b in date_zone:
        if re.search(r"\b(19|20)\d{2}\b", b["text"]) and b["word_count"] <= 15:
            date = b["text"]
            break

    print(f"  Newspaper: {newspaper_name} | Date: {date}")
    return {"newspaper_name": newspaper_name, "date": date}


def analyze_page(image: Image.Image, page_num: int, metadata: dict) -> list:
    """Classify all text elements on a single page. Returns list of row dicts."""
    print(f"  Classifying text elements on page {page_num}...")
    blocks = _get_blocks(image)
    classified = _classify_blocks(blocks)

    rows = []
    for category, text in classified:
        rows.append({
            "newspaper_name": metadata["newspaper_name"],
            "date": metadata["date"],
            "page": page_num,
            "type": category,
            "text": text,
        })
    print(f"    {len(rows)} element(s) extracted.")
    return rows


def export_csv(rows: list, path: str) -> None:
    fieldnames = ["newspaper_name", "date", "page", "type", "text"]
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
    headers = ["newspaper_name", "date", "page", "type", "text"]
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
        description="Analyze a PDF newspaper with Tesseract OCR and export categorized text."
    )
    parser.add_argument("--pdf", required=True, help="Path to the newspaper PDF file.")
    parser.add_argument(
        "--output", default="results",
        help="Base path for output files (default: results → results.csv / results.xlsx).",
    )
    parser.add_argument(
        "--pages", default="",
        help="Comma-separated page numbers to analyze, e.g. '1,2,3' (default: all pages).",
    )
    parser.add_argument(
        "--format", choices=["csv", "xlsx", "both"], default="both",
        help="Export format (default: both).",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.pdf):
        sys.exit(f"Error: file not found: {args.pdf}")

    images = pdf_to_images(args.pdf)
    page_indices = parse_page_selection(args.pages, len(images))

    if not page_indices:
        sys.exit("No valid pages to analyze.")

    metadata = extract_metadata(images[page_indices[0]])

    all_rows = []
    for idx in page_indices:
        page_num = idx + 1
        print(f"\nPage {page_num}/{len(images)}:")
        rows = analyze_page(images[idx], page_num, metadata)
        all_rows.extend(rows)

    print(f"\nTotal elements extracted: {len(all_rows)}")
    export_results(all_rows, args.output, args.format)
    print("Done.")


if __name__ == "__main__":
    main()
