# monitor-logs.py
# Realtime monitor for PM report run logs.
# Tails the newest Logs\PM_Report_*.log (and Logs\watcher.log) and flags:
#   * `!!! ERROR` lines (errors with stage breadcrumb)
#   * stall gaps (> stall-ms between consecutive lines; default 10000ms)
# Usage:
#   python scripts/monitor-logs.py              # one-shot summary of newest run
#   python scripts/monitor-logs.py --follow     # tail live during a run (Ctrl+C to stop)
#   python scripts/monitor-logs.py --fragment 20261117  # only matching logs
import argparse
import glob as _glob
import os
import re
import sys
import time

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")

LINE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s+\+(\d+)ms\s(.*)$")
ERR_RE = re.compile(r"!!! ERROR")


def parse_ms(line):
    m = LINE_RE.match(line.rstrip("\n"))
    if not m:
        return None, None
    return m.group(1), int(m.group(2))


def newest_logs(fragment=""):
    pattern = os.path.join(log_dir, "PM_Report_*.log")
    files = [f for f in _glob.glob(pattern) if fragment in os.path.basename(f)]
    files.sort(key=os.path.getmtime)
    return files


def follow(path, stall_ms=10000):
    """Tail a file live, printing lines and flagging errors/stalls."""
    print(f"Following {path} (stall threshold {stall_ms}ms). Ctrl+C to stop.")
    prev_ts = None
    with open(path, encoding="utf-8", errors="replace") as f:
        f.seek(0, os.SEEK_END)
        while True:
            line = f.readline()
            if not line:
                time.sleep(0.5)
                continue
            line = line.rstrip("\n")
            if not line.strip():
                continue
            ts, ms = parse_ms(line)
            is_err = bool(ERR_RE.search(line))
            stall = False
            if prev_ts is not None and ms is not None and ms > stall_ms:
                stall = True
            tag = ""
            if is_err:
                tag = " [ERROR]"
            if stall:
                tag += f" [STALL +{ms}ms]"
            if tag or is_err:
                print(f"ALERT: {line}{tag}")
            else:
                print(line)
            prev_ts = ts


def summarize(fragment=""):
    files = newest_logs(fragment)
    if not files:
        print(f"No PM_Report_*.log found in {log_dir}" + (f" matching {fragment!r}" if fragment else ""))
        return 0
    path = files[-1]
    print(f"Newest run log: {os.path.basename(path)}  ({os.path.getsize(path)} bytes)")
    with open(path, encoding="utf-8", errors="replace") as f:
        lines = [ln.rstrip("\n") for ln in f if ln.strip()]
    if not lines:
        print("  (empty)")
        return 0
    prev_ms = None
    prev_msg = None
    errors = 0
    stalls = 0
    for ln in lines:
        ts, ms = parse_ms(ln)
        msg = ln
        is_err = bool(ERR_RE.search(ln))
        if is_err:
            errors += 1
            print(f"  ERROR: {msg}")
        if prev_ms is not None and ms is not None and ms > 10000:
            stalls += 1
            print(f"  STALL +{ms}ms after: {prev_msg}")
        prev_ms = ms
        prev_msg = msg
    print(f"  lines={len(lines)} errors={errors} stalls(>{10000}ms)={stalls}")
    print(f"  first: {lines[0]}")
    print(f"  last : {lines[-1]}")
    # watcher.log dialogs
    watcher = os.path.join(log_dir, "watcher.log")
    if os.path.exists(watcher):
        with open(watcher, encoding="utf-8", errors="replace") as f:
            wl = [ln.rstrip("\n") for ln in f if ln.strip()]
        if wl:
            print(f"  watcher.log entries: {len(wl)}")
            for ln in wl[-5:]:
                print(f"    DIALOG: {ln}")
    return 1 if errors else 0


def main():
    parser = argparse.ArgumentParser(description="Monitor PM report run logs in realtime")
    parser.add_argument("--follow", "-f", action="store_true", help="tail the newest run log live")
    parser.add_argument("--fragment", default="", help="only consider logs whose name contains this")
    parser.add_argument("--stall-ms", type=int, default=10000, help="stall threshold in ms")
    args = parser.parse_args()

    if not os.path.isdir(log_dir):
        print(f"Logs dir not found: {log_dir}")
        return 1
    if args.follow:
        files = newest_logs(args.fragment)
        if not files:
            print(f"No PM_Report_*.log found in {log_dir}" + (f" matching {args.fragment!r}" if args.fragment else ""))
            return 1
        try:
            follow(files[-1], stall_ms=args.stall_ms)
        except KeyboardInterrupt:
            print("\nStopped.")
            return 0
    return summarize(args.fragment)


if __name__ == "__main__":
    sys.exit(main())
