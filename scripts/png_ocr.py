#!/usr/bin/env python3
"""png_ocr.py
Extracts text from PNG images using OCR (pytesseract/Tesseract).
Useful for reading text from screenshot images in documentation/screenshots/.
Usage: python scripts/png_ocr.py [--input <IMAGE_FILE>] [--output <OUTPUT_FILE>] [--all]
"""

import argparse
import os
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "png_ocr.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Extract text from PNG images using OCR")
    parser.add_argument("--input", default=None, help="Path to PNG image file")
    parser.add_argument("--output", default=None, help="Output text file path")
    parser.add_argument("--all", action="store_true", help="Process all PNG files in documentation/screenshots/")
    parser.add_argument("--psm", type=int, default=6, help="Tesseract page segmentation mode (default: 6)")
    args = parser.parse_args()

    log("=== PNG OCR started ===")

    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        log("ERROR: Required packages not installed. Run: pip install pytesseract pillow")
        log(f"Details: {e}")
        sys.exit(1)

    # Set Tesseract path explicitly
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    # Check if tesseract executable is available
    try:
        import pytesseract as pt
        pt.get_tesseract_version()
    except Exception as e:
        log("ERROR: Tesseract OCR engine not found.")
        log("Install Tesseract OCR: https://github.com/tesseract-ocr/tesseract/wiki")
        log(f"Expected path: {pytesseract.pytesseract.tesseract_cmd}")
        sys.exit(1)

    screenshot_dir = os.path.join(project_root, "documentation", "screenshots")

    if args.all:
        png_files = [
            f for f in os.listdir(screenshot_dir)
            if f.lower().endswith(".png")
        ]
        if not png_files:
            log("No PNG files found in documentation/screenshots/")
            return
        for png_file in sorted(png_files):
            filepath = os.path.join(screenshot_dir, png_file)
            process_image(filepath, args.output, args.psm)
    elif args.input:
        if not os.path.exists(args.input):
            log(f"ERROR: Image file not found: {args.input}")
            sys.exit(1)
        process_image(args.input, args.output, args.psm)
    else:
        log("ERROR: Please specify --input <FILE> or --all")
        sys.exit(1)


def process_image(image_path, output_file, psm=6):
    """Process a single PNG image and extract text via OCR with layout preservation."""
    try:
        from PIL import Image
        import pytesseract

        log(f"Processing: {os.path.basename(image_path)}")

        with Image.open(image_path) as img:
            # Get OCR data with bounding boxes for layout preservation
            data = pytesseract.image_to_data(img, config=f"--psm {psm}", output_type=pytesseract.Output.DICT)

            # Group text by line number and block number
            lines = {}
            for i in range(len(data['text'])):
                text = data['text'][i].strip()
                if not text:
                    continue
                block_num = data['block_num'][i]
                line_num = data['line_num'][i]
                word_num = data['word_num'][i]
                left = data['left'][i]
                top = data['top'][i]
                conf = data['conf'][i]

                key = (block_num, line_num)
                if key not in lines:
                    lines[key] = []
                lines[key].append((left, text, conf, top))

            # Sort lines by top coordinate, then sort words within each line by left coordinate
            sorted_keys = sorted(lines.keys(), key=lambda k: lines[k][0][3])

            output_lines = []
            for key in sorted_keys:
                words = sorted(lines[key], key=lambda x: x[0])  # Sort by left position
                line_parts = [words[0][1]]  # Start with first word
                for i in range(1, len(words)):
                    prev_word = words[i - 1]
                    curr_word = words[i]
                    gap = curr_word[0] - prev_word[0] - len(prev_word[1]) * 6  # Approximate gap
                    # Add extra spaces for large gaps (column separators)
                    if gap > 60:  # Large gap = column separator
                        line_parts.append("    " + curr_word[1])
                    elif gap > 35:  # Medium gap = extra space
                        line_parts.append("  " + curr_word[1])
                    else:
                        line_parts.append(curr_word[1])
                line_text = " ".join(line_parts)
                if line_text.strip():
                    output_lines.append(line_text)

            full_text = "\n".join(output_lines)

        if full_text.strip():
            print(f"\n--- Text from {os.path.basename(image_path)} ---")
            print(full_text)
            print(f"--- End of {os.path.basename(image_path)} ---\n")

            if output_file:
                with open(output_file, "a", encoding="utf-8") as f:
                    f.write(f"\n--- {os.path.basename(image_path)} ---\n")
                    f.write(full_text)
                    f.write("\n")
                log(f"Appended text to: {output_file}")
            else:
                log(f"No output file specified. Text printed above only.")
        else:
            log(f"No text found in {os.path.basename(image_path)}")

    except Exception as e:
        log(f"ERROR processing {image_path}: {e}")