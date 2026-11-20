content = open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'r', encoding='utf-8').read()
lines = content.split('\n')

# Find section 3 table boundaries
start_line = None
for i, line in enumerate(lines):
    if '| Step | Command | Purpose |' in line:
        start_line = i
        break

if start_line is None:
    print('Could not find section 3 table header')
else:
    # Find end of table (next blank line or ---)
    end_line = start_line + 1
    while end_line < len(lines) and lines[end_line].strip() != '':
        end_line += 1

    em_dash = '\u2014'
    new_table = [
        '| Step | Command | Purpose |',
        '|------|---------|---------|',
        f'| 1 | `python scripts\\\\test-compile.py` | Compile gate {em_dash} injects a probe macro, forces VBA to compile every procedure in every module. Fast {em_dash} seconds. |',
        '| 2 | `python scripts\\\\lint.py` | Lint Python code with ruff \u2014 catches style issues, unused imports, syntax problems. |',
        '| 3 | `python scripts\\\\typecheck.py` | Type-check Python code with mypy \u2014 catches type mismatches and missing annotations. |',
        '| 4 | `python scripts\\\\security-scan.py` | Security scan with bandit (source) + pip-audit (dependencies) \u2014 finds vulnerabilities. |',
        '| 5 | `python scripts\\\\sync_template.py` | Imports all six `.bas` modules into the master `.xlsm` and rebuilds `START HERE`. Use `--package` to also rebuild the deploy zip. |',
        '| 6 | `python scripts\\\\run-report.py` | Full end-to-end build against the fake export. Must return `RESULT=OK\\|...\\|124\\|0\\|<prior>`. |',
        '| 7 | `python scripts\\\\audit-vba.py` | Deep static inspection of all `.bas` modules. Read-only; a candidate list, not a gate. |',
        '| 8 | Read `Logs\\\\PM_Report_*.log` | Every step is timestamped with milliseconds since the previous line. |',
        '| 9 | Kill any stray `EXCEL` process before finishing. | \u201d |',
        ''
    ]

    new_lines_list = lines[:start_line] + new_table + lines[end_line:]
    content = '\n'.join(new_lines_list)
    open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'w', encoding='utf-8').write(content)
    print('Section 3 updated successfully')