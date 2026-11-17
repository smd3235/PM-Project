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
parser.add_argument("--package", action="store_true", help="also rebuild macro-free base + deploy zip + docs")
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

excel = win32com.client.DispatchEx("Excel.Application")
try:
    excel.Visible = False
except Exception:
    pass

try:
    wb = excel.Workbooks.Open(template)
    vbp = wb.VBProject

    # Remove every PM_Report_* module if present, then import all modules.
    module_basenames = [
        "PM_Report_Macro.bas",
        "PM_Report_UI.bas",
        "PM_Report_Host.bas",
        "PM_Report_Setup.bas",
        "PM_Report_Data.bas",
        "PM_Report_Output.bas",
    ]
    module_names = [os.path.splitext(b)[0] for b in module_basenames]

    for c in list(vbp.VBComponents):
        if c.Name in module_names:
            vbp.VBComponents.Remove(c)

    bas_dir = os.path.dirname(os.path.abspath(bas))
    for b in module_basenames:
        module_path = os.path.join(bas_dir, b)
        if not os.path.exists(module_path):
            print(f"Module not found: {module_path}", file=sys.stderr)
            sys.exit(1)
        vbp.VBComponents.Import(module_path)
        print(f"Imported {b}")

    # Rebuild START HERE (calls RebuildStartSheet via macro)
    try:
        excel.Run("RebuildStartSheet")
        print("Rebuilt START HERE.")
    except Exception as e:
        print(f"RebuildStartSheet failed: {e}")

    # Format email template sheets with distinct colors
    for sheet_name, accent, accent_light in [("Manager Email Template", 0x8B4513, 0xE8F0FE), ("Director Email Template", 0x006400, 0xE8F5E9)]:
        try:
            ws = wb.Worksheets(sheet_name)
            ws.Unprotect()
            ws.Cells.Interior.Color = 0xFFFFFF
            ws.Cells.Borders.LineStyle = -4142
            ws.Range("B2:K2").Interior.Color = accent
            ws.Range("B2").Font.Color = 0xFFFFFF
            ws.Range("B3:K3").Interior.Color = accent_light
            ws.Range("B4:K4").Interior.Color = 0xFFF2CC
            for lr in [5, 7, 9, 11]:
                ws.Range("C" + str(lr)).Interior.Color = accent_light
                ws.Range("C" + str(lr)).Font.Color = accent
                ws.Range("C" + str(lr)).Font.Bold = True
            for vr in [5, 7, 9]:
                rng = ws.Range("D" + str(vr) + ":K" + str(vr))
                rng.Interior.Color = 0xFFFFFF
                rng.Borders.LineStyle = 1
                rng.Borders.Color = accent
                rng.Borders.Weight = 2
            rng = ws.Range("D11:K31")
            rng.Interior.Color = 0xFFFFFF
            rng.Borders.LineStyle = 1
            rng.Borders.Color = accent
            rng.Borders.Weight = 2
            ws.Protect()
            ws.EnableSelection = 2
            print("Formatted " + sheet_name)
        except Exception as e:
            print("Format " + sheet_name + " failed: " + str(e))

    wb.Save()
    wb.Close(True)
    print(f"Synced template: {template}")
finally:
    try:
        excel.Quit()
    except Exception:
        pass

if args.package:
    no_macro_path = os.path.join(project_root, "PM_Requests_Master_Template_(No Macros).xlsx")
    instructions_path = os.path.join(project_root, "BUILD-INSTRUCTIONS.txt")
    zip_path = os.path.join(project_root, "Project-Local-Build.zip")

    excel2 = win32com.client.DispatchEx("Excel.Application")
    try:
        excel2.Visible = False
    except Exception:
            pass
    excel2.AutomationSecurity = 1
    try:
        wb2 = excel2.Workbooks.Open(template)
        # Remove VBA project
        try:
            vbp2 = wb2.VBProject
            for c in list(vbp2.VBComponents):
                vbp2.VBComponents.Remove(c)
        except Exception:
            pass
        if os.path.exists(no_macro_path):
            os.remove(no_macro_path)
        try:
            wb2.SaveAs(no_macro_path, FileFormat=51)
            wb2.Close(False)
            print(f"Created macro-free base: {no_macro_path}")
        except Exception as e:
            print(f"Macro-free creation skipped: {e}")
            wb2.Close(False)
    finally:
        try:
            excel2.Quit()
        except Exception:
            pass

    module_files = [os.path.join(bas_dir, b) for b in module_basenames]
    zip_items = [no_macro_path] + module_files
    if os.path.exists(instructions_path):
        zip_items.append(instructions_path)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in zip_items:
            zf.write(item, arcname=os.path.basename(item))
    print(f"Built local build package: {zip_path}")

# Auto-update docs, screenshots, and push to git
try:
    import subprocess
    subprocess.run([sys.executable, os.path.join(script_dir, "update-docs.py")], check=True, timeout=120)
except Exception as e:
    print("update-docs failed: " + str(e))
