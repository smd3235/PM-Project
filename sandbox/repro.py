import os, sys, shutil
import win32com.client

ROOT = r"C:\Users\smd32\Downloads\PM Project\sandbox"
template = os.path.join(ROOT, "committed_template.xlsm")
fake = os.path.join(ROOT, "fake.xlsx")
out_folder = os.path.join(ROOT, "Reports")

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1

# Import the CURRENT working .bas into the committed template
wb = excel.Workbooks.Open(template)
vbp = wb.VBProject
for c in list(vbp.VBComponents):
    if c.Name == "PM_Report_Macro":
        vbp.VBComponents.Remove(c)
vbp.VBComponents.Import(r"C:\Users\smd32\Downloads\PM Project\PM_Report_Macro.bas")
print("imported current working .bas")
wb.Save()
wb.Close(False)

# Check the embedded module's line endings
from oletools.olevba import VBA_Parser
p = VBA_Parser(template)
for (fn, sn, vn, code) in p.extract_all_macros():
    if sn == "VBA/PM_Report_Macro":
        b = code.encode("utf-8", "replace")
        print("embedded: CRLF", b.count(b"\r\n"), "LF", b.count(b"\n"), "CR", b.count(b"\r"))
        open(os.path.join(ROOT, "embedded_after_import.bas"), "w", encoding="utf-8", newline="").write(code)
        break

# Run the report
wb = excel.Workbooks.Open(template)
try:
    res = excel.Run("RunReport", fake, out_folder, "")
    print("RESULT:", res)
except Exception as e:
    print("RunReport raised:", e)
    try:
        print("LogPath:", excel.Run("LogPath"))
    except Exception as e2:
        print("logpath err", e2)
wb.Close(False)
excel.Quit()
print("done")