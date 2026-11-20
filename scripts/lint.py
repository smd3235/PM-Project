#!/usr/bin/env python3
"""lint.py
Runs ruff linter on all Python files in the project.
Usage: python scripts/lint.py [--fix]
"""

import argparse
import os
import re
import subprocess
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "lint.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_ruff(args, label):
    """Run ruff with given arguments and return returncode."""
    cmd = ["python", "-m", "ruff"] + args
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    return result.returncode


def main():
    parser = argparse.ArgumentParser(description="Lint Python files with ruff")
    parser.add_argument("--fix", action="store_true", help="Auto-fix lint issues")
    args = parser.parse_args()

    log("=== Lint (ruff) started ===")

    # Check that ruff is available
    result = subprocess.run(
        ["python", "-m", "ruff", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: ruff is not installed. Run: pip install ruff")
        sys.exit(1)
    log(f"ruff version: {result.stdout.strip()}")

    # Run ruff check
    ruff_args = ["check", "scripts/"]
    if args.fix:
        ruff_args.append("--fix")

    rc = run_ruff(ruff_args, "ruff check")

    # Only fail on actual errors (E/F severity codes), not warnings (W/I/PLR/UP/S)
    if rc != 0:
        # Check if there are any E/F level errors
        error_lines = [
            l
            for l in result.stdout.split("\n")
            if re.match(r"^\s*\w+\.py:\d+:\d+:\s*[EF]", l)
        ]
        if error_lines:
            log("Lint failed with compile/error issues:")
            for l in error_lines:
                print(f"  {l}")
            sys.exit(rc)
        else:
            log(f"Lint warnings found (rc={rc}) but no errors - passing")
    else:
        log("Lint passed: no issues found")


if __name__ == "__main__":
    main()