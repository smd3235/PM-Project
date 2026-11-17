# update-docs.py
# Captures screenshots of template sheets, updates README, and pushes to git.
# Usage: python scripts/update-docs.py

import os
import subprocess
import sys
import time

import win32com.client

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = os.path.join(project_root, "PM_Requests_Master_Template.xlsm")
out_folder = os.path.join(project_root, "documentation", "screenshots")
os.makedirs(out_folder, exist_ok=True)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# Capture sample manager tab from latest report
import glob as _glob
report_files = sorted(_glob.glob(os.path.join(project_root, "Reports", "*.xlsx")), key=os.path.getmtime, reverse=True)
if report_files:
    try:
        wb2 = excel.Workbooks.Open(report_files[0], False, True)
        mgr_sheet = None
        for s in wb2.Worksheets:
            if s.Name not in ("Dashboard", "Raw Data", "START HERE"):
                mgr_sheet = s
                break
        if mgr_sheet:
            used = mgr_sheet.UsedRange
            rng = mgr_sheet.Range(mgr_sheet.Cells(1, 1), mgr_sheet.Cells(used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1))
            mgr_sheet.Activate()
            excel.ActiveWindow.Zoom = 100
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            rng.CopyPicture(1, 2)
            time.sleep(1)
            png_path = os.path.join(out_folder, "sample-manager-tab.png")
            chart = wb2.Charts.Add()
            try:
                chart.Paste()
                time.sleep(1)
                if os.path.exists(png_path):
                    os.remove(png_path)
                chart.Export(png_path)
                print("Exported: " + png_path)
            finally:
                chart.Delete()
        wb2.Close(False)
    except Exception as e:
        print("Report screenshot failed: " + str(e))


# Clear existing template screenshots
for f in os.listdir(out_folder):
    if f.startswith("template-") and f.endswith(".png"):
        os.remove(os.path.join(out_folder, f))


excel.DisplayAlerts = False
excel.AutomationSecurity = 3

wb = None
try:
    wb = excel.Workbooks.Open(template, False, True)
    targets = ["START HERE", "Manager Email Template", "Director Email Template"]
    for name in targets:
        ws = None
        for s in wb.Worksheets:
            if s.Name == name:
                ws = s
                break
        if ws:
            used = ws.UsedRange
            rng = ws.Range(ws.Cells(1, 1), ws.Cells(used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1))
            ws.Activate()
            excel.ActiveWindow.Zoom = 100
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            rng.CopyPicture(1, 2)
            time.sleep(0.25)
            safe = name.replace(" ", "-").lower()
            png_path = os.path.join(out_folder, f"template-{safe}.png")
            chart = wb.Charts.Add()
            try:
                chart.Paste()
                time.sleep(0.25)
                if os.path.exists(png_path):
                    os.remove(png_path)
                chart.Export(png_path)
                print(f"Exported: {png_path}")
            finally:
                chart.Delete()
finally:
    if wb:
        wb.Close(False)
    excel.Quit()

# Update README
readme_path = os.path.join(project_root, "README.md")
with open(readme_path, "r", encoding="utf-8") as f:
    readme = f.read()

# Ensure screenshot references exist
if "template-start-here.png" not in readme:
    # Add screenshots section before "Sample Manager Tab" or at end
    section = "\n\n## Template Screenshots\n\n### START HERE\n\n![START HERE](documentation/screenshots/template-start-here.png)\n\n### Manager Email Template\n\n![Manager Email Template](documentation/screenshots/template-manager-email-template.png)\n\n### Director Email Template\n\n![Director Email Template](documentation/screenshots/template-director-email-template.png)\n"
    # Insert before "## Sample" if exists, otherwise append
    idx = readme.find("## Sample")
    if idx >= 0:
        readme = readme[:idx] + section + "\n" + readme[idx:]
    else:
        readme = readme.rstrip() + section

with open(readme_path, "w", encoding="utf-8") as f:
    f.write(readme)
print("README updated")

# Git add, commit, push
subprocess.run(["git", "add", "-A"], cwd=project_root, check=True)
result = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=project_root)
if result.returncode != 0:
    subprocess.run(["git", "commit", "-m", "docs: auto-update screenshots and README"], cwd=project_root, check=True)
subprocess.run(["git", "push"], cwd=project_root, check=True)
print("Pushed to git")
