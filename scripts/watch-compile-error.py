# watch-compile-error.py
# Watches Excel/VBA modal dialogs and logs them to Logs\watcher.log

import ctypes
import ctypes.wintypes as wintypes
import os
import sys
import time
from datetime import datetime

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
PostMessageW = user32.PostMessageW

WM_COMMAND = 0x0111
IDOK = 1

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "watcher.log")


def write_log(text):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return
    with open(log_path, "a", encoding="utf-8") as f:
        for ln in lines:
            f.write(f"{stamp} | {ln}\n")


def win_text(hwnd):
    buf = ctypes.create_unicode_buffer(1024)
    GetWindowTextW(hwnd, buf, 1024)
    return buf.value


def win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def dialog_text(hwnd):
    parts = [win_text(hwnd)]
    sb = []

    def child_cb(child, _):
        t = win_text(child)
        if t:
            sb.append(t)
        return True

    EnumChildWindows(hwnd, ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(child_cb), 0)
    parts.append("|" + "\n".join(sb))
    return "".join(parts)


def excel_pids():
    pids = set()
    try:
        import psutil
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    pids.add(p.pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except ImportError:
        pass
    return pids


class Watcher:
    def __init__(self):
        self.seen = set()
        self.stuck = {}

    def poll(self):
        result = []
        killed = []
        pids = excel_pids()
        if not pids:
            return result, killed

        def enum_cb(hwnd, _):
            if not IsWindowVisible(hwnd):
                return True
            pid = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids:
                return True
            if win_class(hwnd) != "#32770":
                return True
            text = dialog_text(hwnd)
            is_error = "error" in text.lower()
            if text not in self.seen:
                self.seen.add(text)
                result.append(text)
            if not is_error:
                return True
            c = self.stuck.get(hwnd, 0) + 1
            self.stuck[hwnd] = c
            if c >= 5:
                try:
                    proc = kernel32.OpenProcess(1, False, pid.value)
                    if proc:
                        kernel32.TerminateProcess(proc, 1)
                        kernel32.CloseHandle(proc)
                        killed.append(pid.value)
                except Exception:
                    pass
            else:
                PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            return True

        EnumWindows(ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(enum_cb), 0)
        return result, killed


def main():
    watcher = Watcher()
    write_log(f"watcher started (pid {os.getpid()})")
    print(f"Watching Excel for dialogs -> {log_path}")
    print("Run the macro in Excel now. Stop with Ctrl+C.")
    deadline = time.time() + 2 * 60 * 60
    try:
        while time.time() < deadline:
            time.sleep(0.6)
            hits, killed = watcher.poll()
            for hit in hits:
                write_log(f"DIALOG: {hit}")
                print(f"Captured: {hit}")
            for kp in killed:
                write_log(f"KILLED EXCEL pid {kp} (dialog would not close)")
                print(f"KILLED EXCEL pid {kp} (dialog would not close)")
        write_log("watcher stopped (2h limit reached)")
    finally:
        write_log("watcher stopped")


if __name__ == "__main__":
    main()
