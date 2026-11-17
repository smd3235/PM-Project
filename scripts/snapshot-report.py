# snapshot-report.py
# Exports sheets of a PM report as PNG images.
# Usage: python scripts/snapshot-report.py [--path PATH] [--sheets SHEET1,SHEET2] [--out-folder FOLDER]

import argparse
import os
import sys

import win32com.client

parser = argparse.ArgumentParser(description="Snapshot PM report sheets as PNG")
parser.add_argument("--path", default=None)
parser.add_argument("--sheets", default=None)
parser.add_argument("--out-folder", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

if args.path:
    path = args.path
else:
    reports_dir = os.path.join(project_root, "Reports")
    candidates = sorted(
        [f for f in os.listdir(reports_dir) if f.endswith(".xlsx")],
        key=lambda f: os.path.getmtime(os.path.join(reports_dir, f)),
        reverse=True,
    )
    if not candidates:
        print("No report found in Reports/", file=sys.stderr)
        sys.exit(2)
    path = os.path.join(reports_dir, candidates[0])

if not os.path.exists(path):
    print(f"Workbook not found: {path}", file=sys.stderr)
    sys.exit(2)

out_folder = args.out_folder or os.path.join(project_root, "documentation", "screenshots")
os.makedirs(out_folder, exist_ok=True)

# Clear existing PNGs
for f in os.listdir(out_folder):
    if f.endswith(".png"):
        os.remove(os.path.join(out_folder, f))

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
excel.AutomationSecurity = 3

wb = None
exported = 0
try:
    wb = excel.Workbooks.Open(path, False, True)
    targets = []
    if args.sheets:
        names = [s.strip() for s in args.sheets.split(",")]
        for name in names:
            for ws in wb.Worksheets:
                if ws.Name == name:
                    targets.append(ws)
    else:
        targets = list(wb.Worksheets)

    for ws in targets:
        used = ws.UsedRange
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1))
        ws.Activate()
        excel.ActiveWindow.Zoom = 100
        rng.CopyPicture(1, 2)  # xlScreen=1, xlBitmap=2

        import time
        time.sleep(0.25)
        safe = ws.Name.replace(" ", "-").lower()
        png_path = os.path.join(out_folder, f"report-{safe}.png")
        chart = wb.Charts.Add()
        try:
            chart.Paste()
            time.sleep(0.25)
            if os.path.exists(png_path):
                os.remove(png_path)
            chart.Export(png_path)
            print(f"Exported: {png_path}")
            exported += 1
        finally:
            chart.Delete()
finally:
    if wb:
        wb.Close(False)
    excel.Quit()

if exported == 0:
    sys.exit(3)
