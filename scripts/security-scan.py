#!/usr/bin/env python3
"""security-scan.py
Runs security scanning on Python code and dependencies.
- bandit: scans Python source code for security issues (only fails on MEDIUM/HIGH)
- pip-audit: scans installed Python packages for known vulnerabilities
Usage: python scripts/security-scan.py [--fix]
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
log_path = os.path.join(log_dir, "security-scan.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_command(cmd, label):
    """Run a command and return returncode."""
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    return result.returncode


def main():
    parser = argparse.ArgumentParser(description="Security scan with bandit + pip-audit")
    args = parser.parse_args()

    log("=== Security scan started ===")
    overall_ok = True

    # ---- Bandit: scan Python source code for security issues ----
    result = subprocess.run(
        ["python", "-m", "bandit", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: bandit is not installed. Run: pip install bandit")
        sys.exit(1)
    log(f"bandit version: {result.stdout.strip()}")

    result = subprocess.run(
        ["python", "-m", "bandit", "scripts/", "-r", "-f", "json"],
        capture_output=True,
        text=True,
    )
    bandit_stdout = result.stdout
    if result.returncode != 0:
        # Bandit returns non-zero when issues are found
        # Parse JSON output to count HIGH/MEDIUM severity issues
        try:
            import json

            bandit_data = json.loads(bandit_stdout)
            results = bandit_data.get("results", [])
            high_count = sum(
                1
                for issue in results
                if issue.get("issue_severity") in ("HIGH", "MEDIUM")
            )
            total_count = len(results)
            if high_count > 0:
                log(
                    f"Bandit found {high_count} HIGH/MEDIUM severity issue(s) "
                    f"({total_count} total issues)"
                )
                overall_ok = False
            else:
                log(
                    f"Bandit scan passed: {total_count} low-severity issue(s) only"
                )
        except (json.JSONDecodeError, KeyError, TypeError):
            # If we can't parse JSON, just report what we found
            log(
                "Bandit completed with issues (could not parse severity details)"
            )
            overall_ok = False
    else:
        log("Bandit scan passed: no security issues found")

    # ---- pip-audit: scan installed packages for vulnerabilities ----
    result = subprocess.run(
        ["python", "-m", "pip_audit", "--version"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("WARNING: pip-audit is not installed. Skipping dependency vulnerability scan.")
        log("Install with: pip install pip-audit")
    else:
        log(f"pip-audit version: {result.stdout.strip()}")
        rc = run_command(
            ["python", "-m", "pip_audit"],
            "pip-audit dependency scan",
        )
        if rc != 0:
            log(f"pip-audit found vulnerabilities (exit code {rc})")
            overall_ok = False
        else:
            log("pip-audit scan passed: no known vulnerabilities found")

    if not overall_ok:
        log("Security scan FAILED - issues found above")
        sys.exit(1)

    log("Security scan passed: no issues found")


if __name__ == "__main__":
    main()