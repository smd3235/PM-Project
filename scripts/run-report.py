# run-report.py
# Full end-to-end build against the fake export.
# Opens the master template, runs RunReport(source, outputFolder, priorPath),
# and prints the result string expected by the mandatory loop:
#   RESULT=OK|savePath|pmTotal|skipped|priorFileName
# or RESULT=ERROR|...
import argparse
import os
import sys
import threading
import time
from datetime import datetime

import psutil
import win32com.client

parser = argparse.ArgumentParser(description="Run the PM report end-to-end via RunReport")
parser.add_argument("--template", default=None)
parser.add_argument("--source", default=None)
parser.add_argument("--output-folder", default=None)
parser.add_argument("--prior", default="")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "PM_Requests_Master_Template.xlsm")
source = args.source or os.path.join(project_root, "AOTIOpsTreatmentPMDatesResults_FAKE.xlsx")
out_folder = args.output_folder or os.path.join(project_root, "Reports")
prior = args.prior

log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
os.makedirs(out_folder, exist_ok=True)
run_log = os.path.join(log_dir, "run-report.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(run_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


if not os.path.exists(template):
    log(f"Template not found: {template}")
    sys.exit(1)
if not os.path.exists(source):
    log(f"Source export not found: {source}")
    sys.exit(1)

log(f"=== run-report started (template={template})")
log(f"source={source}  output={out_folder}  prior={prior or '(auto)'}")

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1

watchdog_timer = None
excel_pid = None

try:
    before = {p.pid for p in psutil.process_iter(["name"])
              if p.info["name"] and p.info["name"].lower() == "excel.exe"}
except Exception:
    before = set()

wb = None
result = None
try:
    wb = excel.Workbooks.Open(template)

    def watchdog():
        time.sleep(10)
        try:
            proc = psutil.Process(excel_pid)
            proc.kill()
            log(f"Watchdog killed Excel pid {excel_pid}.")
        except Exception:
            pass

    try:
        after = {p.pid for p in psutil.process_iter(["name"])
                 if p.info["name"] and p.info["name"].lower() == "excel.exe"}
        new = after - before
        excel_pid = next(iter(new)) if new else None
    except Exception:
        excel_pid = None

    if excel_pid:
        watchdog_timer = threading.Thread(target=watchdog, daemon=True)
        watchdog_timer.start()

    result = excel.Run("RunReport", source, out_folder, prior)
    log(f"RunReport returned: {result}")

    if watchdog_timer and watchdog_timer.is_alive():
        watchdog_timer.join(timeout=1)

    ok = isinstance(result, str) and result.startswith("OK|")
    log(f"RESULT={'OK' if ok else 'FAILED'}|{result if result else '(no result)'}")
    sys.exit(0 if ok else 1)
except Exception as e:
    result = str(e)
    log(f"RunReport raised: {result}")
    sys.exit(1)
finally:
    if wb:
        try:
            wb.Close(False)
        except Exception as e:
            log(f"Close failed: {e}")
    try:
        excel.Quit()
    except Exception as e:
        log(f"Quit failed: {e}")
    if excel_pid:
        try:
            psutil.Process(excel_pid).kill()
            log(f"Killed lingering Excel pid {excel_pid}.")
        except Exception:
            pass
