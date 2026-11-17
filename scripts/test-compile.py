# test-compile.py
# Compile-test harness for PM_Report_Macro.bas
# Copies the template, injects a tiny probe macro, then runs it via COM.
# Running ANY macro forces VBA to compile the ENTIRE project.

import argparse
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

import psutil
import win32com.client

from excel_guard import DEFAULT_TIMEOUT, ExcelGuard, start_excel

parser = argparse.ArgumentParser(description="Compile-test PM_Report_Macro.bas")
parser.add_argument("--template", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "PM_Requests_Master_Template.xlsm")
test_copy = os.path.join(project_root, "_compile_test.xlsm")
marker = os.path.join(project_root, "probe_ran.txt")
watcher = os.path.join(script_dir, "watch-compile-error.py")
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
watch_log = os.path.join(log_dir, "watcher.log")
test_log = os.path.join(log_dir, "compile-test.log")
err_log = os.path.join(log_dir, "compile-errors.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(test_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)
if os.path.exists(marker):
    os.remove(marker)
log(f"=== Compile test started (template: {template}) ===")

watch_proc = subprocess.Popen(
    [sys.executable, watcher],
    creationflags=subprocess.CREATE_NO_WINDOW,
)
log(f"Dialog watcher started (pid {watch_proc.pid}).")
time.sleep(1)

excel, excel_pid = start_excel()
log(f"Excel COM instance pid: {excel_pid}")

log_mark = 0
if os.path.exists(watch_log):
    with open(watch_log, "r", encoding="utf-8") as f:
        log_mark = sum(1 for _ in f)

run_error = None
wb = None
try:
    with ExcelGuard(excel_pid, DEFAULT_TIMEOUT):
        wb = excel.Workbooks.Open(test_copy)
    vbp = wb.VBProject
    mod = vbp.VBComponents.Add(1)
    mod.Name = "zzCompileProbe"
    probe_code = (
        "Option Explicit\r\n"
        "Public Sub Probe()\r\n"
        f'    Open "{marker}" For Output As #1\r\n'
        '    Print #1, "ran"\r\n'
        "    Close #1\r\n"
        "End Sub"
    )
    mod.CodeModule.AddFromString(probe_code)
    log("Running Probe (forces full-project compile)...")
    with ExcelGuard(excel_pid, DEFAULT_TIMEOUT):
        excel.Run("Probe")
    log("Run call returned without COM error.")
except Exception as e:
    run_error = str(e)
    log(f"Run call raised: {run_error}")

time.sleep(1)
new_lines = []
if os.path.exists(watch_log):
    with open(watch_log, "r", encoding="utf-8") as f:
        all_lines = f.readlines()
    if len(all_lines) > log_mark:
        new_lines = all_lines[log_mark:]

compile_ok = os.path.exists(marker)
if compile_ok:
    log("RESULT: COMPILE OK - probe executed, project compiled cleanly.")
else:
    log("RESULT: COMPILE FAILED - probe did not execute.")
    where = ""
    if new_lines:
        where = " | dialog=" + " / ".join(l.rstrip() for l in new_lines)
        log("Captured dialog text (Logs\\watcher.log):")
        for l in new_lines:
            log("  " + l.rstrip())
    else:
        log(f"No dialog captured. Run error: {run_error}")
    err_line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | COMPILE FAILED - probe did not execute{where}"
    try:
        with open(err_log, "a", encoding="utf-8") as f:
            f.write(err_line + "\n")
    except Exception:
        pass
    log(f"Full error details written to Logs\\compile-errors.log")

try:
    if wb:
        wb.Close(False)
except Exception as e:
    log(f"Close failed: {e}")
try:
    if excel_pid:
        excel.Quit()
        try:
            psutil.Process(excel_pid).kill()
            log(f"Killed lingering Excel pid {excel_pid}.")
        except Exception:
            pass
except Exception as e:
    log(f"Quit failed: {e}")
try:
    watch_proc.kill()
except Exception:
    pass
try:
    os.remove(test_copy)
except OSError:
    pass
log("=== Compile test finished ===")
