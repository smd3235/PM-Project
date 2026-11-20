import re

content = open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'r', encoding='utf-8').read()

# Update section 3 - add lint, typecheck, security-scan steps
old_section3 = """| Step | Command | Purpose |
|------|---------|---------|
| 1 | `python scripts\\test-compile.py` | Compile gate \\u2014 injects a probe macro, forces VBA to compile every procedure in every module. Fast \u2014 seconds. |
| 2 | `python scripts\\sync_template.py` | Imports all six `.bas` modules into the master `.xlsm` and rebuilds `START HERE`. Use `--package` to also rebuild the deploy zip. |
| 3 | `python scripts\\run-report.py` | Full end-to-end build against the fake export. Must return `RESULT=OK\\|...\\|124\\|0\\|<prior>`. |
| 4 | `python scripts\\audit-vba.py` | Deep static inspection of all `.bas` modules. Read-only; a candidate list, not a gate. |
| 5 | Read `Logs\\PM_Report_*.log` | Every step is timestamped with milliseconds since the previous line. |
| 6 | Kill any stray `EXCEL` process before finishing. | \u201d |"""

new_section3 = """| Step | Command | Purpose |
|------|---------|---------|
| 1 | `python scripts\\test-compile.py` | Compile gate \\u2014 injects a probe macro, forces VBA to compile every procedure in every module. Fast \u2014 seconds. |
| 2 | `python scripts\\lint.py` | Lint Python code with ruff \u2014 catches style issues, unused imports, syntax problems. |
| 3 | `python scripts\\typecheck.py` | Type-check Python code with mypy \u2014 catches type mismatches and missing annotations. |
| 4 | `python scripts\\security-scan.py` | Security scan with bandit (source) + pip-audit (dependencies) \u2014 finds vulnerabilities. |
| 5 | `python scripts\\sync_template.py` | Imports all six `.bas` modules into the master `.xlsm` and rebuilds `START HERE`. Use `--package` to also rebuild the deploy zip. |
| 6 | `python scripts\\run-report.py` | Full end-to-end build against the fake export. Must return `RESULT=OK\\|...\\|124\\|0\\|<prior>`. |
| 7 | `python scripts\\audit-vba.py` | Deep static inspection of all `.bas` modules. Read-only; a candidate list, not a gate. |
| 8 | Read `Logs\\PM_Report_*.log` | Every step is timestamped with milliseconds since the previous line. |
| 9 | Kill any stray `EXCEL` process before finishing. | \u201d |"""

if old_section3 in content:
    content = content.replace(old_section3, new_section3)
    print("Section 3 updated successfully")
else:
    print("Section 3 pattern not found")
    lines = content.split('\n')
    for i, line in enumerate(lines):
        if 'test-compile.py' in line:
            print(f"Line {i+1}: {repr(line)}")
            for j in range(max(0,i-1), min(len(lines), i+10)):
                print(f"  {j+1}: {repr(lines[j])}")

# Update section 6 - file layout - add new scripts to scripts\ listing
old_scripts_line = "| `scripts\\` | All tooling: Python (`sync_template`, `test-compile`, `run-report`, `snapshot-report`, `update-docs`, `check-toc`, `audit-vba`, `monitor-logs`, `pipeline`, `make_fake_data`, `make_prior`). New test scripts go here. |"
new_scripts_line = "| `scripts\\` | All tooling: Python (`sync_template`, `test-compile`, `run-report`, `snapshot-report`, `update-docs`, `check-toc`, `audit-vba`, `lint`, `typecheck`, `security-scan`, `monitor-logs`, `pipeline`, `make_fake_data`, `make_prior`). New test scripts go here. |"

if old_scripts_line in content:
    content = content.replace(old_scripts_line, new_scripts_line)
    print("Section 6 updated successfully")
else:
    print("Section 6 pattern not found")
    lines = content.split('\n')
    for i, line in enumerate(lines):
        if 'scripts\\' in line and 'All tooling' in line:
            print(f"Line {i+1}: {repr(line)}")

# Update section 13 - pipeline reference
old_section13 = """- `scripts\\pipeline.py` runs the full mandatory loop (steps 1\u20135 of section 3) plus
  screenshots and docs, in order, headless, with 15-second timeouts on every Excel
  call. Use it as the single entry point for a verified change:
  `python scripts\\pipeline.py` (add `--skip-push` to commit locally without pushing)."""

new_section13 = """- `scripts\\pipeline.py` runs the full mandatory loop (steps 1\u20139 of section 3) plus
  screenshots and docs, in order, headless, with 15-second timeouts on every Excel
  call. Use it as the single entry point for a verified change:
  `python scripts\\pipeline.py` (add `--skip-push` to commit locally without pushing)."""

if old_section13 in content:
    content = content.replace(old_section13, new_section13)
    print("Section 13 updated successfully")
else:
    print("Section 13 pattern not found")
    lines = content.split('\n')
    for i, line in enumerate(lines):
        if 'pipeline.py' in line.lower():
            print(f"Line {i+1}: {repr(line)}")

open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'w', encoding='utf-8').write(content)
print("Done writing")