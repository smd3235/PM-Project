#!/usr/bin/env python3
"""pyupgrade_check.py
Runs pyupgrade to modernize Python syntax (old-style formatting -> f-strings,
etc.). Reports what WOULD change without modifying files unless --fix is passed.
Usage: python scripts/pyupgrade_check.py [--py3-plus --fix]
"""

import argparse
import os
import subprocess
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "pyupgrade.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Check Python syntax modernization with pyupgrade")
    parser.add_argument("--py3-plus", action="store_true", help="Use py3-plus rules")
    parser.add_argument("--fix", action="store_true", help="Apply fixes automatically")
    args = parser.parse_args()

    log("=== PyUpgrade check started ===")

    # Check that pyupgrade is available
    result = subprocess.run(
        ["python", "-m", "pyupgrade", "--help"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: pyupgrade is not installed. Run: pip install pyupgrade")
        sys.exit(1)
    log("pyupgrade is available")

    # Build command - pass individual .py files, not the directory
    cmd = ["python", "-m", "pyupgrade"]
    # Collect all .py files in scripts/
    py_files = []
    for root, dirs, files in os.walk(os.path.join(project_root, "scripts")):
        for f in files:
            if f.endswith(".py"):
                cmd.append(os.path.join(root, f))
    if args.py3_plus:
        cmd.append("--py3-plus")
    if not args.fix:
        cmd.append("--exit-zero-even-if-changed")

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"pyupgrade found changes to make (exit code {result.returncode})")
        if not args.fix:
            log("Run with --fix to apply changes automatically")
        sys.exit(result.returncode)

    log("PyUpgrade check passed: all files are modernized")


if __name__ == "__main__":
    main()