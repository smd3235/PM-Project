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
    # Find end of table (next blank line)
    end_line = start_line + 1
    while end_line < len(lines) and lines[end_line].strip() != '':
        end_line += 1

    em_dash = '\u2014'  # Actual em-dash character
    
    new_table = [
        '| Step | Command | Purpose |',
        '|------|---------|---------|',
        '| 1 | `python scripts\\test-compile.py` | Compile gate ' + em_dash + ' injects a probe macro, forces VBA to compile every procedure in every module. Fast ' + em_dash + ' seconds. |',
        '| 2 | `python scripts\\lint.py` | Lint Python code with ruff \u2014 catches style issues, unused imports, syntax problems. |',
        '| 3 | `python scripts\\typecheck.py` | Type-check Python code with mypy \u2014 catches type mismatches and missing annotations. |',
        '| 4 | `python scripts\\security-scan.py` | Security scan with bandit (source) + pip-audit (dependencies) \u2014 finds vulnerabilities. |',
        '| 5 | `python scripts\\pyupgrade_check.py` | Modernize Python syntax with pyupgrade \u2014 upgrades old-style formatting to f-strings. |',
        '| 6 | `python scripts\\autoformat.py` | Remove unused imports/variables with autoflake \u2014 cleans up dead code in imports. |',
        '| 7 | `python scripts\\complexity.py` | Code complexity analysis with radon \u2014 finds overly complex functions. |',
        '| 8 | `python scripts\\deadcode.py` | Dead code detection with vulture \u2014 finds unused code and imports. |',
        '| 9 | `python scripts\\deps.py` | Unused dependency detection with deptry \u2014 finds unused Python packages. |',
        '| 10 | `python scripts\\sync_template.py` | Imports all six `.bas` modules into the master `.xlsm` and rebuilds `START HERE`. Use `--package` to also rebuild the deploy zip. |',
        '| 11 | `python scripts\\run-report.py` | Full end-to-end build against the fake export. Must return `RESULT=OK\\|...\\|124\\|0\\|<prior>`. |',
        '| 12 | `python scripts\\audit-vba.py` | Deep static inspection of all `.bas` modules. Read-only; a candidate list, not a gate. |',
        '| 13 | Read `Logs\\PM_Report_*.log` | Every step is timestamped with milliseconds since the previous line. |',
        '| 14 | Kill any stray `EXCEL` process before finishing. | \u201d |',
        ''
    ]

    new_lines_list = lines[:start_line] + new_table + lines[end_line:]
    content = '\n'.join(new_lines_list)
    open('C:/Users/smd32/Downloads/PM Project/documentation/AGENTS.md', 'w', encoding='utf-8').write(content)
    print('Section 3 updated successfully')
else:
    print('Could not find section 3 table header')