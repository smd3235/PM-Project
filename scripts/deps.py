#!/usr/bin/env python3
"""deps.py
Runs deptry to find unused dependencies in Python projects.
Usage: python scripts/deps.py [--ignore-numpy --ignore-pandas]
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
log_path = os.path.join(log_dir, "deps.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Detect unused dependencies with deptry")
    parser.add_argument("--ignore-numpy", action="store_true", help="Ignore numpy (used indirectly)")
    parser.add_argument("--ignore-pandas", action="store_true", help="Ignore pandas (used indirectly)")
    args = parser.parse_args()

    log("=== Dependency check (deptry) started ===")

    result = subprocess.run(
        ["python", "-m", "deptry", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: deptry is not installed. Run: pip install deptry")
        sys.exit(1)
    log(f"deptry version: {result.stdout.strip()}")

    cmd = ["python", "-m", "deptry", "scripts/"]
    if args.ignore_numpy:
        cmd.extend(["--ignore", "numpy"])
    if args.ignore_pandas:
        cmd.extend(["--ignore", "pandas"])

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"Dependency check found unused dependencies (exit code {result.returncode})")
        sys.exit(result.returncode)

    log("Dependency check passed: no unused dependencies found")


if __name__ == "__main__":
    main()