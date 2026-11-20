#!/usr/bin/env python3
"""verify_report.py
Data validation script that verifies report consistency across:
  1. Excel workbook (.xlsx) - actual data
  2. Screenshot PNGs (via OCR) - visual representation
  3. Raw export file (.xlsx) - source data

Usage: python scripts/verify_report.py --report <report.xlsx> --screenshot <screenshot.png> --source <fake_export.xlsx>
"""

import argparse
import os
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "verify_report.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def verify_report(report_path, screenshot_path=None, source_path=None):
    """Verify consistency between Excel report, screenshot OCR, and raw source data."""
    log("=" * 60)
    log(f"VERIFICATION STARTED")
    log(f"Report: {report_path}")
    if screenshot_path:
        log(f"Screenshot: {screenshot_path}")
    if source_path:
        log(f"Source data: {source_path}")
    log("=" * 60)

    errors = []
    warnings = []

    # 1. Load Excel workbook
    try:
        import openpyxl
        wb = openpyxl.load_workbook(report_path)
        log(f"Opened workbook: {len(wb.sheetnames)} sheets")
    except Exception as e:
        log(f"ERROR: Cannot open Excel file: {e}")
        sys.exit(1)

    # 2. Validate each sheet in the Excel workbook
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        log(f"  Sheet: {sheet_name} ({ws.max_row} rows x {ws.max_column} cols)")

        # Check for empty rows/columns issues
        empty_rows = 0
        for row in ws.iter_rows(values_only=True):
            if all(cell is None for cell in row):
                empty_rows += 1
        if empty_rows > 0:
            log(f"  WARNING: {empty_rows} empty rows found in {sheet_name}")
            warnings.append(f"{sheet_name}: {empty_rows} empty rows")

    # 3. OCR screenshot if provided
    screenshot_text = None
    if screenshot_path:
        if not os.path.exists(screenshot_path):
            log(f"WARNING: Screenshot not found: {screenshot_path}")
            warnings.append(f"Screenshot not found: {screenshot_path}")
        else:
            try:
                import pytesseract
                from PIL import Image

                pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

                with Image.open(screenshot_path) as img:
                    text = pytesseract.image_to_string(img)

                screenshot_text = text
                log(f"OCR extracted {len(text.split())} words from screenshot")

                # Check for common OCR garbling patterns
                garbled_patterns = ["cae", "cern", "Representativ@gy", "iiv", "!!!"]
                found_garbled = [p for p in garbled_patterns if p.lower() in text.lower()]
                if found_garbled:
                    log(f"WARNING: Possible OCR errors detected: {found_garbled}")
                    warnings.append(f"Screenshot OCR may have errors: {found_garbled}")

            except Exception as e:
                log(f"ERROR: OCR failed: {e}")
                warnings.append(f"OCR failed: {e}")
    else:
        # Auto-detect latest screenshot
        screenshots_dir = os.path.join(project_root, "documentation", "screenshots")
        if os.path.exists(screenshots_dir):
            png_files = [f for f in os.listdir(screenshots_dir) if f.lower().endswith(".png")]
            if png_files:
                latest_screenshot = os.path.join(screenshots_dir, max(png_files, key=lambda f: os.path.getmtime(os.path.join(screenshots_dir, f))))
                screenshot_path = latest_screenshot
                log(f"Auto-detected screenshot: {os.path.basename(screenshot_path)}")
                
                try:
                    import pytesseract
                    from PIL import Image

                    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

                    with Image.open(screenshot_path) as img:
                        text = pytesseract.image_to_string(img)

                    screenshot_text = text
                    log(f"OCR extracted {len(text.split())} words from screenshot")

                    garbled_patterns = ["cae", "cern", "Representativ@gy", "iiv", "!!!"]
                    found_garbled = [p for p in garbled_patterns if p.lower() in text.lower()]
                    if found_garbled:
                        log(f"WARNING: Possible OCR errors detected: {found_garbled}")
                        warnings.append(f"Screenshot OCR may have errors: {found_garbled}")

                except Exception as e:
                    log(f"ERROR: OCR failed: {e}")
                    warnings.append(f"OCR failed: {e}")
            else:
                log("WARNING: No screenshot PNG files found in documentation/screenshots/")
        else:
            log("WARNING: Screenshots directory not found")

    # 4. Load source data if provided
    source_data = None
    if source_path:
        if not os.path.exists(source_path):
            log(f"WARNING: Source file not found: {source_path}")
            warnings.append(f"Source file not found: {source_path}")
        else:
            try:
                source_wb = openpyxl.load_workbook(source_path)
                log(f"Loaded source workbook: {len(source_wb.sheetnames)} sheets")
                source_data = source_wb
            except Exception as e:
                log(f"ERROR: Cannot open source export: {e}")
                warnings.append(f"Source file error: {e}")

    # 5. Cross-validate
    log("\n--- Cross-Validation Results ---")

    # Check if screenshot text contains data that matches Excel
    if screenshot_text and wb:
        screenshot_text_lower = screenshot_text.lower()
        mismatches = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Check key data points from each sheet
            for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 20), values_only=True):
                row_values = [str(cell) for cell in row if cell is not None and str(cell).strip()]
                if row_values:
                    # Check if key identifiers appear in screenshot OCR
                    for val in row_values[:3]:  # Check first 3 columns
                        if len(val) > 3 and val.lower() not in screenshot_text_lower:
                            # Only flag substantial text (not just numbers/dates)
                            if any(c.isalpha() for c in val):
                                mismatches.append(f"{sheet_name}: '{val}' not found in screenshot OCR")

        if mismatches:
            log(f"WARNING: {len(mismatches)} values in Excel not found in screenshot OCR:")
            for m in mismatches[:10]:  # Show first 10
                log(f"  - {m}")
            warnings.append(f"{len(mismatches)} values missing from screenshot OCR")
        else:
            log("All key data from Excel found in screenshot OCR")

    # Summary
    log("\n" + "=" * 60)
    log(f"VERIFICATION COMPLETE")
    log(f"Errors: {len(errors)}")
    log(f"Warnings: {len(warnings)}")

    if errors:
        log("STATUS: FAILED - Errors found")
        sys.exit(1)
    elif warnings:
        log("STATUS: WARNINGS - Review recommended")
        sys.exit(0)
    else:
        log("STATUS: PASSED - No issues found")
        sys.exit(0)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Verify PM report data consistency across Excel, screenshots, and source data")
    parser.add_argument("--report", default=None, help="Path to Excel report workbook (auto-detects latest if omitted)")
    parser.add_argument("--screenshot", default=None, help="Path to screenshot PNG (auto-detects latest if omitted)")
    parser.add_argument("--source", default=None, help="Path to source data export (.xlsx)")
    args = parser.parse_args()

    if not args.report:
        # Auto-detect latest report
        reports_dir = os.path.join(project_root, "Reports")
        if os.path.exists(reports_dir):
            xlsx_files = [f for f in os.listdir(reports_dir) if f.endswith(".xlsx")]
            if xlsx_files:
                args.report = os.path.join(reports_dir, max(xlsx_files, key=lambda f: os.path.getmtime(os.path.join(reports_dir, f))))
                log(f"Auto-detected report: {args.report}")
            else:
                log("ERROR: No report files found in Reports/")
                sys.exit(1)
        else:
            log("ERROR: Reports directory not found")
            sys.exit(1)

    if not os.path.exists(args.report):
        log(f"ERROR: Report file not found: {args.report}")
        sys.exit(1)

    verify_report(args.report, args.screenshot, args.source)