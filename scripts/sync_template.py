# sync_template.py
# Re-imports PM_Report_Macro.bas into the template's embedded VBA project
# and rebuilds the START HERE tab.
# Run this after every change to PM_Report_Macro.bas.

import argparse
import os
import sys
import zipfile

import win32com.client

parser = argparse.ArgumentParser(description="Sync VBA macro into PM template")
parser.add_argument("--template", default=None)
parser.add_argument("--bas", default=None)
parser.add_argument("--skip-package", action="store_true")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "PM_Requests_Master_Template.xlsm")
bas = args.bas or os.path.join(project_root, "PM_Report_Macro.bas")

if not os.path.exists(template):
    print(f"Template not found: {template}", file=sys.stderr)
    sys.exit(1)
if not os.path.exists(bas):
    print(f"Bas not found: {bas}", file=sys.stderr)
    sys.exit(1)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 3

try:
    wb = excel.Workbooks.Open(template)
    vbp = wb.VBProject

    # Remove existing module if present
    for c in list(vbp.VBComponents):
        if c.Name == "PM_Report_Macro":
            vbp.VBComponents.Remove(c)

    # Import the .bas module
    vbp.VBComponents.Import(bas)
    print(f"Imported PM_Report_Macro from {bas}")

    # Rebuild START HERE (calls RebuildStartSheet via macro)
    try:
        excel.Run("RebuildStartSheet")
        print("Rebuilt START HERE.")
    except Exception as e:
        print(f"RebuildStartSheet failed: {e}")

    wb.Save()
    wb.Close(True)
    print(f"Synced template: {template}")
finally:
    excel.Quit()

if not args.skip_package:
    no_macro_path = os.path.join(project_root, "PM_Requests_Master_Template_(No Macros).xlsx")
    instructions_path = os.path.join(project_root, "BUILD-INSTRUCTIONS.txt")
    zip_path = os.path.join(project_root, "Project-Local-Build.zip")

    excel2 = win32com.client.Dispatch("Excel.Application")
    excel2.Visible = False
    excel2.DisplayAlerts = False
    excel2.AutomationSecurity = 3
    try:
        wb2 = excel2.Workbooks.Open(template)
        if os.path.exists(no_macro_path):
            os.remove(no_macro_path)
        wb2.SaveAs(no_macro_path, FileFormat=51)
        wb2.Close(False)
        print(f"Created macro-free base: {no_macro_path}")
    finally:
        excel2.Quit()

    module_files = [bas]
    zip_items = [no_macro_path] + module_files
    if os.path.exists(instructions_path):
        zip_items.append(instructions_path)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in zip_items:
            zf.write(item, arcname=os.path.basename(item))
    print(f"Built local build package: {zip_path}")
