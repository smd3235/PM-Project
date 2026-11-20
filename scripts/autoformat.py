#!/usr/bin/env python3
"""autoformat.py
Runs autoflake to remove unused imports and unused variables from Python files.
Usage: python scripts/autoformat.py [--fix]
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
log_path = os.path.join(log_dir, "autoformat.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Remove unused imports/variables with autoflake")
    parser.add_argument("--fix", action="store_true", help="Apply fixes automatically")
    args = parser.parse_args()

    log("=== AutoFormat (autoflake) started ===")

    result = subprocess.run(
        ["python", "-m", "autoflake", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: autoflake is not installed. Run: pip install autoflake")
        sys.exit(1)
    log(f"autoflake version: {result.stdout.strip()}")

    # Build command - pass individual .py files
    cmd = ["python", "-m", "autoflake"]
    for root, dirs, files in os.walk(os.path.join(project_root, "scripts")):
        for f in files:
            if f.endswith(".py"):
                cmd.append(os.path.join(root, f))
    if not args.fix:
        cmd.append("--check")

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"AutoFormat failed with exit code {result.returncode}")
        sys.exit(result.returncode)

    log("AutoFormat passed: no unused imports/variables found")


if __name__ == "__main__":
    main()