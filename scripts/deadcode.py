#!/usr/bin/env python3
"""deadcode.py
Runs vulture to find dead code (unused code) in Python files.
Usage: python scripts/deadcode.py [--min-confidence 90]
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
log_path = os.path.join(log_dir, "deadcode.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Find dead code with vulture")
    parser.add_argument("--min-confidence", type=int, default=90, help="Minimum confidence to report (0-100)")
    args = parser.parse_args()

    log("=== Dead code detection (vulture) started ===")

    result = subprocess.run(
        ["python", "-m", "vulture", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: vulture is not installed. Run: pip install vulture")
        sys.exit(1)
    log(f"vulture version: {result.stdout.strip()}")

    cmd = ["python", "-m", "vulture", "scripts/", f"--min-confidence={args.min_confidence}"]

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"Dead code found (exit code {result.returncode})")
        sys.exit(result.returncode)

    log("Dead code check passed: no dead code found")


if __name__ == "__main__":
    main()