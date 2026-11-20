#!/usr/bin/env python3
"""export_report.py
Exports PM report data from the generated report workbook to various formats
(CSV, JSON, Excel) using pandas.
Usage: python scripts/export_report.py [--report <REPORT_FILE>] [--format csv|json|excel] [--output <OUTPUT_FILE>]
"""

import argparse
import os
import subprocess
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "export_report.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def find_latest_report():
    """Find the most recent report in the Reports directory."""
    reports_dir = os.path.join(project_root, "Reports")
    if not os.path.exists(reports_dir):
        return None
    candidates = [
        f for f in os.listdir(reports_dir) if f.endswith(".xlsx")
    ]
    if not candidates:
        return None
    return os.path.join(reports_dir, max(candidates, key=lambda f: os.path.getmtime(os.path.join(reports_dir, f))))


def main():
    parser = argparse.ArgumentParser(description="Export PM report data using pandas")
    parser.add_argument("--report", default=None, help="Path to report workbook")
    parser.add_argument("--format", default="csv", choices=["csv", "json", "excel"], help="Output format")
    parser.add_argument("--output", default=None, help="Output file path")
    args = parser.parse_args()

    log("=== PM Report Export started ===")

    try:
        import pandas as pd
    except ImportError:
        log("ERROR: pandas is not installed. Run: pip install pandas")
        sys.exit(1)

    # Determine input report
    report_path = args.report
    if not report_path:
        report_path = find_latest_report()
    if not report_path:
        log("ERROR: No report file found. Use --report to specify one.")
        sys.exit(1)

    if not os.path.exists(report_path):
        log(f"ERROR: Report file not found: {report_path}")
        sys.exit(1)

    log(f"Reading report: {report_path}")

    # Determine output file
    if args.output:
        output_file = args.output
    else:
        base_name = os.path.splitext(os.path.basename(report_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = {"csv": "csv", "json": "json", "excel": "xlsx"}[args.format]
        output_file = os.path.join(project_root, "Reports", f"{base_name}_{timestamp}.{ext}")

    try:
        # Read all sheets from the report workbook
        xl = pd.ExcelFile(report_path)

        if args.format == "excel":
            # For Excel output, just copy the file
            import shutil
            shutil.copy2(report_path, output_file)
            log(f"Report copied to: {output_file}")
        else:
            # Read all sheets and export each one
            all_sheets = pd.read_excel(xl, sheet_name=None)
            for sheet_name, df in all_sheets.items():
                sheet_output = output_file
                if args.format == "csv":
                    sheet_output = output_file.replace(".csv", f"_{sheet_name}.csv")
                elif args.format == "json":
                    sheet_output = output_file.replace(".json", f"_{sheet_name}.json")

                if args.format == "csv":
                    df.to_csv(sheet_output, index=False)
                elif args.format == "json":
                    df.to_json(sheet_output, orient="records", indent=2)

                log(f"Exported sheet '{sheet_name}' to {sheet_output} ({len(df)} rows, {len(df.columns)} columns)")

        log(f"Export complete: {output_file}")

    except Exception as e:
        log(f"ERROR: Export failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()