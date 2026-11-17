# excel_guard.py
# Shared helper: run blocking Excel COM calls with a hard timeout so nothing
# ever waits minutes for a stuck macro. Also runs a dialog watcher alongside
# every Excel session so modal popups (OK clicks, errors, file pickers) are
# captured with their full text and reported instead of silently hanging.
import os
import subprocess
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
watch_log = os.path.join(log_dir, "watcher.log")
pipeline_log = os.path.join(log_dir, "pipeline.log")
watcher_script = os.path.join(script_dir, "watch-compile-error.py")

DEFAULT_TIMEOUT = 15  # seconds - nothing waits longer than this


def log_pipeline(msg):
    """Write a timestamped line to Logs\pipeline.log."""
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(pipeline_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def log_hang(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line, file=sys.stderr)
    with open(hang_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def log_dialog(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(watch_log, "a", encoding="utf-8") as f:
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
    if pid:
        log_pipeline(f"Excel session started pid={pid}")
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
        else:
            log_pipeline(f"Excel call completed pid={self.pid}")
        return False  # don't suppress exceptions


class DialogWatcher:
    """Starts watch-compile-error.py as a background process that watches for
    Excel modal dialogs and logs their full text. On exit, terminates the
    watcher and returns any new dialog entries captured since start.

    Usage:
        with DialogWatcher() as dw:
            excel, pid = start_excel()
            with ExcelGuard(pid):
                excel.Run("SomeMacro")
        for d in dw.dialogs:
            print("Dialog during run:", d)
    """

    def __init__(self):
        self.proc = None
        self.start_lines = 0

    def __enter__(self):
        # Count existing lines so we only report new dialogs.
        if os.path.exists(watch_log):
            with open(watch_log, "r", encoding="utf-8") as f:
                self.start_lines = sum(1 for _ in f)
        else:
            self.start_lines = 0
        # Launch watcher with stdout captured so we can read its output.
        self.proc = subprocess.Popen(
            [sys.executable, watcher_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )
        log_pipeline("Dialog watcher started")
        return self

    def _read_new_dialogs(self):
        """Return new watcher.log lines captured since start, filtered to
        real dialog entries (DIALOG: or KILLED EXCEL)."""
        if not os.path.exists(watch_log):
            return []
        with open(watch_log, "r", encoding="utf-8") as f:
            lines = f.readlines()
        new = lines[self.start_lines:]
        dialogs = []
        for l in new:
            s = l.strip()
            if "DIALOG:" in s or "KILLED EXCEL" in s:
                dialogs.append(s)
        return dialogs

    def __exit__(self, exc_type, exc, tb):
        dialogs = []
        if self.proc:
            try:
                self.proc.terminate()
                self.proc.wait(timeout=3)
            except Exception:
                try:
                    self.proc.kill()
                except Exception:
                    pass
            # Drain any stdout the watcher produced.
            try:
                out, _ = self.proc.communicate(timeout=3)
                if out:
                    for line in out.splitlines():
                        if line.strip():
                            dialogs.append(line.strip())
            except Exception:
                pass
        # Also read watcher.log for any dialogs logged during the session.
        dialogs.extend(self._read_new_dialogs())
        # Deduplicate while preserving order.
        seen = set()
        unique = []
        for d in dialogs:
            if d not in seen:
                seen.add(d)
                unique.append(d)
        self.dialogs = unique
        for d in unique:
            log_dialog(f"DIALOG DURING SESSION: {d}")
        log_pipeline(f"Dialog watcher stopped ({len(unique)} dialog(s) captured)")
        return False
