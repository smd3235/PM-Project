#!/usr/bin/env python3
"""complexity.py
Runs radon to analyze code complexity (cyclomatic complexity) of Python files.
Usage: python scripts/complexity.py [--min-rank C]
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
log_path = os.path.join(log_dir, "complexity.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Analyze code complexity with radon")
    parser.add_argument("--min-rank", default="C", help="Minimum rank to report (A=best, F=worst)")
    args = parser.parse_args()

    log("=== Complexity analysis (radon) started ===")

    result = subprocess.run(
        ["python", "-m", "radon", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: radon is not installed. Run: pip install radon")
        sys.exit(1)
    log(f"radon version: {result.stdout.strip()}")

    cmd = ["python", "-m", "radon", "cc", "scripts/", "-a", "--min", args.min_rank]

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"Complexity check found issues (exit code {result.returncode})")
        sys.exit(result.returncode)

    log("Complexity check passed: no high-complexity functions found")


if __name__ == "__main__":
    main()