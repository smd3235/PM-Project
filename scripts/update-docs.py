# update-docs.py
# Captures screenshots of template sheets and updates README.
# Usage: python scripts/update-docs.py [--push]
# Headless by default: no visible Excel, no git push (push hangs without credentials).
import argparse
import glob as _glob
import os
import subprocess
import sys
import time

import win32com.client

parser = argparse.ArgumentParser(description="Update docs screenshots and README")
parser.add_argument("--push", action="store_true", help="also git push after local commit")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = os.path.join(project_root, "PM_Requests_Master_Template.xlsm")
out_folder = os.path.join(project_root, "documentation", "screenshots")
os.makedirs(out_folder, exist_ok=True)

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1


def capture_sheet(wb, sheet_name, png_path):
    ws = None
    for s in wb.Worksheets:
        if s.Name == sheet_name:
            ws = s
            break
    if ws is None:
        print(f"Sheet not found: {sheet_name}")
        return
    used = ws.UsedRange
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1))
    ws.Activate()
    excel.ActiveWindow.Zoom = 100
    excel.ActiveWindow.DisplayGridlines = False
    excel.ActiveWindow.DisplayHeadings = False
    rng.CopyPicture(1, 2)
    time.sleep(0.25)
    chart = wb.Charts.Add()
    try:
        chart.Paste()
        time.sleep(0.25)
        if os.path.exists(png_path):
            os.remove(png_path)
        chart.Export(png_path)
        print("Exported: " + png_path)
    finally:
        chart.Delete()


# Capture sample manager tab from latest report
report_files = sorted(
    _glob.glob(os.path.join(project_root, "Reports", "*.xlsx")),
    key=os.path.getmtime,
    reverse=True,
)
if report_files:
    wb2 = None
    try:
        wb2 = excel.Workbooks.Open(report_files[0], False, True)
        mgr_sheet = None
        for s in wb2.Worksheets:
            if s.Name not in ("Dashboard", "Raw Data", "START HERE"):
                mgr_sheet = s
                break
        if mgr_sheet:
            png_path = os.path.join(out_folder, "sample-manager-tab.png")
            capture_sheet(wb2, mgr_sheet.Name, png_path)
    except Exception as e:
        print("Report screenshot failed: " + str(e))
    finally:
        if wb2:
            try:
                wb2.Close(False)
            except Exception:
                pass

# Clear existing template screenshots
for f in os.listdir(out_folder):
    if f.startswith("template-") and f.endswith(".png"):
        os.remove(os.path.join(out_folder, f))

wb = None
try:
    wb = excel.Workbooks.Open(template, False, True)
    for name, safe in [
        ("START HERE", "start-here"),
        ("Manager Email Template", "manager-email-template"),
        ("Director Email Template", "director-email-template"),
    ]:
        png_path = os.path.join(out_folder, f"template-{safe}.png")
        capture_sheet(wb, name, png_path)
finally:
    if wb:
        try:
            wb.Close(False)
        except Exception:
            pass
    try:
        excel.Quit()
    except Exception:
        pass

# Update README
readme_path = os.path.join(project_root, "README.md")
with open(readme_path, "r", encoding="utf-8") as f:
    readme = f.read()

if "template-start-here.png" not in readme:
    section = (
        "\n\n## Template Screenshots\n\n### START HERE\n\n"
        "![START HERE](documentation/screenshots/template-start-here.png)\n\n"
        "### Manager Email Template\n\n"
        "![Manager Email Template](documentation/screenshots/template-manager-email-template.png)\n\n"
        "### Director Email Template\n\n"
        "![Director Email Template](documentation/screenshots/template-director-email-template.png)\n"
    )
    idx = readme.find("## Sample")
    if idx >= 0:
        readme = readme[:idx] + section + "\n" + readme[idx:]
    else:
        readme = readme.rstrip() + section

with open(readme_path, "w", encoding="utf-8") as f:
    f.write(readme)
print("README updated")

# Git add + local commit only; push only with --push (avoids hanging on credentials)
try:
    subprocess.run(["git", "add", "-A"], cwd=project_root, check=True)
    result = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=project_root)
    if result.returncode != 0:
        subprocess.run(
            ["git", "commit", "-m", "docs: auto-update screenshots and README"],
            cwd=project_root,
            check=True,
        )
        print("Committed docs changes locally")
    if args.push:
        subprocess.run(["git", "push"], cwd=project_root, check=True)
        print("Pushed to git")
except Exception as e:
    print("git step skipped: " + str(e))
