#!/usr/bin/env python3
"""pipeline.py
Full reliable pipeline: compile gate -> lint -> typecheck -> security scan ->
sync modules -> run report -> screenshots -> docs -> git push.
Headless throughout; no modal dialogs.
"""

import argparse
import os
import subprocess
import sys
import time
from datetime import datetime

from excel_guard import log_pipeline

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

pipeline_log = os.path.join(project_root, "Logs", "pipeline.log")


def log_step(label, status, duration_ms=None):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dur = f" ({duration_ms}ms)" if duration_ms is not None else ""
    line = f"{ts} | STEP {label}: {status}{dur}"
    print(line)
    with open(pipeline_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(cmd, label, timeout=60000):
    print(f"=== {label} ===")
    print(" ".join(cmd))
    start = time.time()
    log_step(label, "STARTED")
    try:
        p = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True, timeout=timeout)
        elapsed_ms = int((time.time() - start) * 1000)
        if p.stdout:
            print(p.stdout)
        if p.stderr:
            print(p.stderr, file=sys.stderr)
        if p.returncode != 0:
            log_step(label, "FAILED", elapsed_ms)
            print(f"FAILED: {label} (exit {p.returncode})", file=sys.stderr)
            sys.exit(p.returncode)
        log_step(label, "OK", elapsed_ms)
        print(f"OK: {label}")
    except subprocess.TimeoutExpired:
        elapsed_ms = int((time.time() - start) * 1000)
        log_step(label, "TIMEOUT", elapsed_ms)
        print(f"FAILED: {label} timed out after {timeout}ms", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Full PM report pipeline")
    parser.add_argument("--skip-push", action="store_true", help="commit locally but do not push")
    args = parser.parse_args()

    python = sys.executable

    log_pipeline("=== PIPELINE START ===")

    # 1. Compile gate
    run([python, os.path.join(script_dir, "test-compile.py")], "test-compile (compile gate)", timeout=60000)

    # 2. Lint Python code with ruff
    run([python, os.path.join(script_dir, "lint.py")], "lint (ruff)", timeout=60000)

    # 3. Type-check Python code with mypy
    run([python, os.path.join(script_dir, "typecheck.py")], "typecheck (mypy)", timeout=60000)

    # 4. Security scan with bandit + pip-audit
    run([python, os.path.join(script_dir, "security-scan.py")], "security-scan (bandit + pip-audit)", timeout=120000)

    # 5. Sync modules into xlsm + build non-macro zip
    run([python, os.path.join(script_dir, "sync_template.py"), "--package"], "sync_template (sync modules + zip)", timeout=120000)

    # 6. End-to-end report build
    run([python, os.path.join(script_dir, "run-report.py")], "run-report (end-to-end build)", timeout=60000)

    # 7. Regenerate screenshots
    run([python, os.path.join(script_dir, "snapshot-report.py")], "snapshot-report (screenshots)", timeout=60000)

    # 8. Update docs + README photos + commit
    docs_cmd = [python, os.path.join(script_dir, "update-docs.py")]
    if not args.skip_push:
        docs_cmd.append("--push")
    run(docs_cmd, "update-docs (README photos + commit" + (" + push" if not args.skip_push else "") + ")", timeout=120000)

    log_pipeline("=== PIPELINE COMPLETE ===")


if __name__ == "__main__":
    main()