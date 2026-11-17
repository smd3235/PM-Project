# excel_guard.py
# Shared helper: run blocking Excel COM calls with a hard timeout so nothing
# ever waits minutes for a stuck macro. Any operation that exceeds
# DEFAULT_TIMEOUT seconds kills the Excel process and logs the hang to
# Logs\excel-hangs.log.
import os
import sys
import threading
import time
from datetime import datetime

import psutil

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
hang_log = os.path.join(log_dir, "excel-hangs.log")

DEFAULT_TIMEOUT = 15  # seconds - nothing waits longer than this


def log_hang(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line, file=sys.stderr)
    with open(hang_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def excel_pids():
    pids = set()
    try:
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    pids.add(p.pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except ImportError:
        pass
    return pids


def start_excel():
    """Create a headless Excel instance and return (excel, pid)."""
    import win32com.client

    before = excel_pids()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1

    # Wait briefly for the process to register so we can track its PID.
    pid = None
    for _ in range(20):  # up to ~2s
        after = excel_pids()
        new = after - before
        if new:
            pid = next(iter(new))
            break
        time.sleep(0.1)
    return excel, pid


class ExcelGuard:
    """Guard a blocking Excel COM call with a hard timeout.

    Usage:
        excel, pid = start_excel()
        with ExcelGuard(pid):
            excel.Workbooks.Open(path)

    If the call exceeds timeout seconds, the Excel process is killed and the
    hang is logged.
    """

    def __init__(self, pid, timeout=DEFAULT_TIMEOUT):
        self.pid = pid
        self.timeout = timeout
        self.timer = None

    def _kill(self):
        try:
            proc = psutil.Process(self.pid)
            proc.kill()
        except Exception:
            pass
        log_hang(f"Excel pid {self.pid} hung (>{self.timeout}s), killed.")

    def __enter__(self):
        self.timer = threading.Timer(self.timeout, self._kill)
        self.timer.daemon = True
        self.timer.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.timer:
            self.timer.cancel()
        if exc_type is not None:
            log_hang(f"Excel operation raised: {exc}")
        return False  # don't suppress exceptions
