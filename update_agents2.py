content = open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'r', encoding='utf-8').read()

# Update section 3 - add lint, typecheck, security-scan steps
old_section3 = """| Step | Command | Purpose |
|------|---------|---------|
| 1 | `python scripts\\test-compile.py` | Compile gate \u2014 injects a probe macro, forces VBA to compile every procedure in every module. Fast \u2014 seconds. |
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
    print("Section 3 pattern still not found")
    lines = content.split('\n')
    for i, line in enumerate(lines):
        if 'test-compile.py' in line:
            print(f"Line {i+1}: {repr(line)}")

open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'w', encoding='utf-8').write(content)
print("Done")