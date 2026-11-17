# pipeline.py
# Full reliable pipeline: compile gate -> sync modules -> run report ->
# screenshots -> docs -> git push. Headless throughout; no modal dialogs.
import argparse
import os
import subprocess
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)


def run(cmd, label, timeout=60000):
    print(f"=== {label} ===")
    print(" ".join(cmd))
    p = subprocess.run(cmd, cwd=project_root, capture_output=True, text=True, timeout=timeout)
    if p.stdout:
        print(p.stdout)
    if p.stderr:
        print(p.stderr, file=sys.stderr)
    if p.returncode != 0:
        print(f"FAILED: {label} (exit {p.returncode})", file=sys.stderr)
        sys.exit(p.returncode)
    print(f"OK: {label}")


def main():
    parser = argparse.ArgumentParser(description="Full PM report pipeline")
    parser.add_argument("--skip-push", action="store_true", help="commit locally but do not push")
    args = parser.parse_args()

    python = sys.executable

    # 1. Compile gate
    run([python, os.path.join(script_dir, "test-compile.py")], "test-compile (compile gate)", timeout=60000)

    # 2. Sync modules into xlsm + build non-macro zip
    run([python, os.path.join(script_dir, "sync_template.py"), "--package"], "sync_template (sync modules + zip)", timeout=120000)

    # 3. End-to-end report build
    run([python, os.path.join(script_dir, "run-report.py")], "run-report (end-to-end build)", timeout=60000)

    # 4. Regenerate screenshots
    run([python, os.path.join(script_dir, "snapshot-report.py")], "snapshot-report (screenshots)", timeout=60000)

    # 5. Update docs + README photos + commit
    docs_cmd = [python, os.path.join(script_dir, "update-docs.py")]
    if not args.skip_push:
        docs_cmd.append("--push")
    run(docs_cmd, "update-docs (README photos + commit" + (" + push" if not args.skip_push else "") + ")", timeout=120000)

    print("PIPELINE COMPLETE")


if __name__ == "__main__":
    main()
