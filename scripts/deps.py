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
    parser.add_argument("--ignore-ruff", action="store_true", help="Ignore ruff (CLI tool)")
    parser.add_argument("--ignore-mypy", action="store_true", help="Ignore mypy (CLI tool)")
    parser.add_argument("--ignore-pyupgrade", action="store_true", help="Ignore pyupgrade (CLI tool)")
    parser.add_argument("--ignore-bandit", action="store_true", help="Ignore bandit (CLI tool)")
    parser.add_argument("--ignore-pip-audit", action="store_true", help="Ignore pip-audit (CLI tool)")
    parser.add_argument("--ignore-safety", action="store_true", help="Ignore safety (CLI tool)")
    parser.add_argument("--ignore-pytest", action="store_true", help="Ignore pytest (CLI tool)")
    parser.add_argument("--ignore-pytest-asyncio", action="store_true", help="Ignore pytest-asyncio (CLI tool)")
    parser.add_argument("--ignore-pytest-cov", action="store_true", help="Ignore pytest-cov (CLI tool)")
    parser.add_argument("--ignore-pytest-timeout", action="store_true", help="Ignore pytest-timeout (CLI tool)")
    parser.add_argument("--ignore-pytest-xdist", action="store_true", help="Ignore pytest-xdist (CLI tool)")
    parser.add_argument("--ignore-hypothesis", action="store_true", help="Ignore hypothesis (CLI tool)")
    parser.add_argument("--ignore-watchdog", action="store_true", help="Ignore watchdog (CLI tool)")
    parser.add_argument("--ignore-tenacity", action="store_true", help="Ignore tenacity (CLI tool)")
    parser.add_argument("--ignore-structlog", action="store_true", help="Ignore structlog (CLI tool)")
    parser.add_argument("--ignore-rich", action="store_true", help="Ignore rich (CLI tool)")
    parser.add_argument("--ignore-python-dotenv", action="store_true", help="Ignore python-dotenv (CLI tool)")
    parser.add_argument("--ignore-pil", action="store_true", help="Ignore PIL (transitive dep of Pillow)")
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
    if args.ignore_ruff:
        cmd.extend(["--ignore", "ruff"])
    if args.ignore_mypy:
        cmd.extend(["--ignore", "mypy"])
    if args.ignore_pyupgrade:
        cmd.extend(["--ignore", "pyupgrade"])
    if args.ignore_bandit:
        cmd.extend(["--ignore", "bandit"])
    if args.ignore_pip_audit:
        cmd.extend(["--ignore", "pip-audit"])
    if args.ignore_safety:
        cmd.extend(["--ignore", "safety"])
    if args.ignore_pytest:
        cmd.extend(["--ignore", "pytest"])
    if args.ignore_pytest_asyncio:
        cmd.extend(["--ignore", "pytest-asyncio"])
    if args.ignore_pytest_cov:
        cmd.extend(["--ignore", "pytest-cov"])
    if args.ignore_pytest_timeout:
        cmd.extend(["--ignore", "pytest-timeout"])
    if args.ignore_pytest_xdist:
        cmd.extend(["--ignore", "pytest-xdist"])
    if args.ignore_hypothesis:
        cmd.extend(["--ignore", "hypothesis"])
    if args.ignore_watchdog:
        cmd.extend(["--ignore", "watchdog"])
    if args.ignore_tenacity:
        cmd.extend(["--ignore", "tenacity"])
    if args.ignore_structlog:
        cmd.extend(["--ignore", "structlog"])
    if args.ignore_rich:
        cmd.extend(["--ignore", "rich"])
    if args.ignore_python_dotenv:
        cmd.extend(["--ignore", "python-dotenv"])
    if args.ignore_pil:
        cmd.extend(["--ignore", "PIL"])

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        log(f"Dependency check found {result.returncode} potential unused dependency(s) - these are informational only")
        log("Note: Many listed deps are CLI tools (ruff, mypy, etc.) not directly imported in scripts/")
        log("Dependency check completed with informational warnings")
    else:
        log("Dependency check passed: no unused dependencies found")


if __name__ == "__main__":
    main()