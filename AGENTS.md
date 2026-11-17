# Agent Rules

The canonical rule set for this project lives in **`documentation/AGENTS.md`**.
Read it before changing anything; it covers the mandatory loop
(compile gate -> sync -> run -> `/logs`), the Excel 10-second kill rule, the VBA
declaration rules, the log format, the file layout, and version control.

Short version, so this file is useful on its own:

- `PM_Report_Macro.bas` is the only source of truth for VBA; never edit the
  module inside the `.xlsm`.
- Every `.bas` change: `python scripts/test-compile.py`, then
  `python scripts/sync_template.py`, then `python scripts/run-report.py`.
- The build must return `RESULT=OK|...|106|18|0|<prior>`.
- New test scripts go in `scripts\`; Markdown and image assets go in
  `documentation\`; run logs go in `Logs\`; generated reports in `Reports\`.