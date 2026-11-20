#!/usr/bin/env python3
"""pytest_run.py
Runs pytest with coverage on all Python tests.
Usage: python scripts/pytest_run.py [--cov]
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
log_path = os.path.join(log_dir, "pytest.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Run pytest with coverage")
    parser.add_argument("--cov", action="store_true", help="Run with coverage reporting")
    args = parser.parse_args()

    log("=== pytest run started ===")

    result = subprocess.run(
        ["python", "-m", "pytest", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: pytest is not installed. Run: pip install pytest pytest-cov")
        sys.exit(1)
    log(f"pytest version: {result.stdout.strip()}")

    cmd = ["python", "-m", "pytest", "scripts/", "-v"]
    if args.cov:
        cmd.extend(["--cov=scripts", "--cov-report=term-missing"])

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"pytest failed with exit code {result.returncode}")
        sys.exit(result.returncode)

    log("pytest passed: all tests passed")


if __name__ == "__main__":
    main()