#!/usr/bin/env python3
"""typecheck.py
Runs mypy type checker on all Python files in the project.
Usage: python scripts/typecheck.py [--strict]
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
log_path = os.path.join(log_dir, "typecheck.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Type-check Python files with mypy")
    parser.add_argument("--strict", action="store_true", help="Run with --strict flags")
    args = parser.parse_args()

    log("=== Type check (mypy) started ===")

    # Check that mypy is available
    result = subprocess.run(
        ["python", "-m", "mypy", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: mypy is not installed. Run: pip install mypy")
        sys.exit(1)
    log(f"mypy version: {result.stdout.strip()}")

    # Build mypy command
    cmd = ["python", "-m", "mypy", "scripts/", "--ignore-missing-imports"]
    if args.strict:
        cmd.append("--strict")

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"Type check FAILED with exit code {result.returncode}")
        sys.exit(result.returncode)

    log("Type check passed: no type errors found")


if __name__ == "__main__":
    main()