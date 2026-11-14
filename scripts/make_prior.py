"""Build a prior report with rows removed from manager/pending tabs,
so the next run must flag new PMs and pending pickups."""
import glob
import os
import shutil

import openpyxl

OUT = r"C:\Users\smd32\AppData\Local\Temp\kilo\pmtest\xl_out"
SRC = sorted(glob.glob(os.path.join(OUT, "PM Requests 2026-11-13 (3).xlsx")), key=os.path.getmtime)[-1]
DST = os.path.join(OUT, "PRIOR_PM_Requests_2026-10-13.xlsx")

shutil.copyfile(SRC, DST)
wb = openpyxl.load_workbook(DST)

removed_mgr = 0
removed_pend = 0
for name in wb.sheetnames:
    if name in ("Dashboard", "Raw Data"):
        continue
    ws = wb[name]
    if ws.max_row < 2:
        continue
    if name == "Pending Pickup":
        ws.delete_rows(2, 4)          # drop 4 pending rows
        removed_pend = 4
    else:
        ws.delete_rows(2, 3)          # drop 3 manager rows
        removed_mgr += 3
wb.save(DST)
wb.close()
print("prior built:", DST, "removed manager rows:", removed_mgr, "pending rows:", removed_pend)
print("expect next run: new PMs>0, green rows>0, pending New>0")