#!/usr/bin/env python3
"""md_format.py
Runs mdformat to format Markdown files consistently.
Usage: python scripts/md_format.py [--fix]
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
log_path = os.path.join(log_dir, "md_format.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="Format Markdown files with mdformat")
    parser.add_argument("--fix", action="store_true", help="Apply fixes automatically")
    args = parser.parse_args()

    log("=== Markdown format (mdformat) started ===")

    result = subprocess.run(
        ["python", "-m", "mdformat", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: mdformat is not installed. Run: pip install mdformat mdformat-gfm")
        sys.exit(1)
    log(f"mdformat version: {result.stdout.strip()}")

    cmd = ["python", "-m", "mdformat", "documentation/", "README.md", "documentation/AGENTS.md"]
    if not args.fix:
        cmd.append("--check")

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"mdformat failed with exit code {result.returncode}")
        sys.exit(result.returncode)

    log("mdformat passed: all Markdown files are formatted")


if __name__ == "__main__":
    main()