#!/usr/bin/env python3
"""
Newspaper Analyzer — reads newspapers.com PDF clippings and extracts:
  - newspaper name and publication date (from the top metadata strip)
  - headlines, subheadlines, and image captions (via Tesseract OCR)
Exports results as CSV and/or Excel. Free — no API key required.
"""

import argparse
import csv
import os
import re
import subprocess
import sys
from collections import defaultdict

import numpy as np
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter

try:
    import openpyxl
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

# newspapers.com overlay heights at 200 DPI (pixels to skip)
TOP_STRIP_PX    = 160
BOTTOM_STRIP_PX = 300

# Overlay text patterns to discard from OCR output
OVERLAY_PATTERNS = re.compile(
    r"(newspapers?\.com|copyright|downloaded\s+on|clipped\s+by|all\s+rights\s+reserved"
    r"|https?://|www\.|cvd_\d+)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Metadata extraction — uses pdftotext (reliable on the vector overlay layer)
# ---------------------------------------------------------------------------

def extract_metadata_from_pdf(pdf_path: str) -> dict:
    """
    Pull newspaper name and date from the newspapers.com metadata strip
    embedded as a vector text layer in the PDF.
    """
    try:
        result = subprocess.run(
            ["pdftotext", pdf_path, "-"],
            capture_output=True, text=True, timeout=30,
        )
        text = result.stdout
    except Exception:
        return {"newspaper_name": "Unknown", "date": "Unknown"}

    # The newspapers.com format puts the title on its own line:
    # "Chicago Tribune (Chicago, Illinois) · Sun, Mar 29, 2020 · Page 1-29"
    name, date = "Unknown", "Unknown"
    for line in text.splitlines():
        line = line.strip()
        m = re.match(
            r"^((?:The\s+)?[A-Z][A-Za-z\s,()\']+?)\s*·\s*"
            r"(\w+,\s+\w+\s+\d{1,2},\s+\d{4})",
            line,
        )
        if m and "newspapers.com" not in m.group(1).lower():
            name = m.group(1).strip()
            date = m.group(2).strip()
            break

    return {"newspaper_name": name, "date": date}


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------

def pdf_to_images(pdf_path: str) -> list:
    print(f"Converting PDF to images: {pdf_path}")
    images = convert_from_path(pdf_path, dpi=300)
    print(f"  {len(images)} page(s) found.")
    return images


def _preprocess(image: Image.Image) -> Image.Image:
    """Crop overlay strips then enhance contrast for Tesseract."""
    W, H = image.width, image.height
    # Scale strip sizes from 200 DPI baseline to actual DPI (300 DPI → 1.5×)
    scale = H / 2200
    top    = int(TOP_STRIP_PX    * scale * 1.5)
    bottom = int(BOTTOM_STRIP_PX * scale * 1.5)
    cropped = image.crop((0, top, W, H - bottom))
    gray = cropped.convert("L")
    sharpened = gray.filter(ImageFilter.SHARPEN)
    return ImageEnhance.Contrast(sharpened).enhance(2.0)


def _get_blocks(image: Image.Image) -> list:
    """
    Run Tesseract on the preprocessed image and group words into
    paragraph-level blocks.

    Returns list of dicts:
      text, avg_height, word_count, top, page_h
    """
    processed = _preprocess(image)
    page_h = processed.height

    data = pytesseract.image_to_data(
        processed, config="--oem 1 --psm 3",
        output_type=pytesseract.Output.DICT,
    )

    # Group words into lines keyed by (block_num, par_num, line_num)
    lines: dict = defaultdict(list)
    for i, level in enumerate(data["level"]):
        if level != 5:
            continue
        text = data["text"][i].strip()
        if not text or data["conf"][i] < 15:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines[key].append({
            "text":   text,
            "height": data["height"][i],
            "top":    data["top"][i],
            "left":   data["left"][i],
        })

    # Group lines into paragraphs
    paras: dict = defaultdict(list)
    for (bn, pn, ln), words in lines.items():
        paras[(bn, pn)].append({
            "text":       " ".join(w["text"] for w in words),
            "avg_height": float(np.mean([w["height"] for w in words])),
            "top":        min(w["top"] for w in words),
        })

    blocks = []
    for para_lines in paras.values():
        full_text = " ".join(l["text"] for l in para_lines).strip()
        if not full_text:
            continue
        # Skip newspapers.com overlay text that leaked through
        if OVERLAY_PATTERNS.search(full_text):
            continue
        # Skip garbled OCR: require at least 60% alphabetic characters
        # and at least 3 characters total
        if len(full_text) < 3:
            continue
        alpha_ratio = sum(c.isalpha() for c in full_text) / len(full_text)
        if alpha_ratio < 0.60:
            continue
        avg_h = float(np.mean([l["avg_height"] for l in para_lines]))
        top   = min(l["top"] for l in para_lines)
        blocks.append({
            "text":       full_text,
            "avg_height": avg_h,
            "word_count": len(full_text.split()),
            "top":        top,
            "page_h":     page_h,
        })

    return blocks


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def _classify_blocks(blocks: list) -> list:
    """
    Return list of (type, text) for headline / subheadline / image_caption.
    Body text is skipped.

    Headline     : tallest text on page (>= 50% of max height), short block
    Subheadline  : first block immediately below a headline (deck text)
                   OR ALL-CAPS medium-size block (section label / byline header)
    Image caption: short block (<= 20 words), small font, in lower half of page
    Everything else is treated as body text and skipped.
    """
    if not blocks:
        return []

    max_h = max(b["avg_height"] for b in blocks)
    if max_h == 0:
        return []

    HEADLINE_THRESH   = 0.50   # >= 50% of tallest text on page
    ALLCAPS_THRESH    = 0.22   # ALL-CAPS section headers (medium size)
    CAPTION_MAX_WORDS = 20
    CAPTION_H_MAX     = 0.25   # relative to max_h

    ordered = sorted(blocks, key=lambda b: b["top"])
    results = []
    last_headline_bottom = -1
    last_headline_h      = 0
    deck_captured        = False   # only first block after headline = deck

    for b in ordered:
        h       = b["avg_height"]
        wc      = b["word_count"]
        rel_h   = h / max_h
        rel_top = b["top"] / b["page_h"] if b["page_h"] else 0
        gap     = b["top"] - last_headline_bottom

        # ALL-CAPS test (ignore punctuation and digits)
        alpha_text = re.sub(r"[^A-Za-z]", "", b["text"])
        is_allcaps = bool(alpha_text) and alpha_text == alpha_text.upper()

        non_space = len(b["text"].replace(" ", ""))
        if rel_h >= HEADLINE_THRESH and wc <= 40 and non_space > 12:
            category = "headline"
            last_headline_bottom = b["top"] + int(h)
            last_headline_h      = h
            deck_captured        = False

        elif (last_headline_h > 0
              and not deck_captured
              and gap <= last_headline_h * 3
              and rel_h >= 0.12
              and wc <= 120):
            # First block directly below a headline → deck / subheadline
            category      = "subheadline"
            deck_captured = True

        elif is_allcaps and rel_h >= ALLCAPS_THRESH and 1 < wc <= 15 and non_space > 12:
            # ALL-CAPS medium block → section header / name label
            category = "subheadline"

        elif wc <= CAPTION_MAX_WORDS and rel_h <= CAPTION_H_MAX and rel_top > 0.30:
            category = "image_caption"

        else:
            continue   # body text — skip

        results.append((category, b["text"]))

    return results


# ---------------------------------------------------------------------------
# Page analysis
# ---------------------------------------------------------------------------

def analyze_page(image: Image.Image, page_num: int, metadata: dict) -> list:
    print(f"  Classifying page {page_num}...")
    blocks     = _get_blocks(image)
    classified = _classify_blocks(blocks)

    rows = []
    for category, text in classified:
        rows.append({
            "newspaper_name": metadata["newspaper_name"],
            "date":           metadata["date"],
            "type":           category,
            "text":           text,
        })
    print(f"    {len(rows)} element(s) found.")
    return rows


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_csv(rows: list, path: str) -> None:
    fieldnames = ["newspaper_name", "date", "type", "text"]
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
    headers = ["newspaper_name", "date", "type", "text"]
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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Analyze a newspapers.com PDF with Tesseract OCR."
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

    print("Extracting newspaper name and date...")
    metadata = extract_metadata_from_pdf(args.pdf)
    print(f"  Newspaper: {metadata['newspaper_name']} | Date: {metadata['date']}")

    images      = pdf_to_images(args.pdf)
    page_indices = parse_page_selection(args.pages, len(images))

    if not page_indices:
        sys.exit("No valid pages to analyze.")

    all_rows = []
    for idx in page_indices:
        rows = analyze_page(images[idx], idx + 1, metadata)
        all_rows.extend(rows)

    print(f"\nTotal elements extracted: {len(all_rows)}")
    export_results(all_rows, args.output, args.format)
    print("Done.")


if __name__ == "__main__":
    main()
