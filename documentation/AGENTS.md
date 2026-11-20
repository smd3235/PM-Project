# Agent Rules — PM Requests Report Builder

These rules apply to every change, no exceptions.

______________________________________________________________________

## 1. Project overview

This is a weekly sales manager report builder. The macro runs inside Excel
(`PM_Requests_Master_Template.xlsm`) and builds one formatted tab per sales
manager from a synthetic export (`AOTIOpsTreatmentPMDatesResults_FAKE.xlsx`).
All VBA source lives in `.bas` files in the project root — never edit the
embedded module inside the `.xlsm` directly.

Every change flows through the **mandatory loop** (section 2) before it is
committed and pushed.

______________________________________________________________________

## 2. The one true source — VBA

- `PM_Report_Macro.bas` (project root) is the canonical source of the VBA.
  Never edit the module inside the `.xlsm`; edit the `.bas` and sync it.
- `scripts\sync_template.py` is the **only** supported way to push a `.bas`
  change into `PM_Requests_Master_Template.xlsm`.
- The embedded VBA module is named `PM_Report_Macro`. Module identifiers cannot
  contain spaces, so the filename and identifier both use underscores.
- The `START HERE` button runs `BuildManagerTabs` by procedure name; the module
  name is invisible to users.
- Five supporting modules live alongside it:
  `PM_Report_UI.bas`, `PM_Report_Host.bas`, `PM_Report_Setup.bas`,
  `PM_Report_Data.bas`, `PM_Report_Output.bas`.

______________________________________________________________________

## 3. Mandatory loop (run after every `.bas` change)

Run in order. Stop at the first failure. Do not push a broken build.

| Step | Command                                          | Purpose                                                                                                                           |
| ---- | ------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------- |
| 1    | `python scripts\\test-compile.py`                | Compile gate — injects a probe macro, forces VBA to compile every procedure in every module. Fast — seconds.                      |
| 2    | `python scripts\\lint.py`                        | Lint Python code with ruff — catches style issues, unused imports, syntax problems.                                               |
| 3    | `python scripts\\typecheck.py`                   | Type-check Python code with mypy — catches type mismatches and missing annotations.                                               |
| 4    | `python scripts\\security-scan.py`               | Security scan with bandit (source) + pip-audit (dependencies) — finds vulnerabilities.                                            |
| 5    | `python scripts\\sync_template.py`               | Imports all six `.bas` modules into the master `.xlsm` and rebuilds `START HERE`. Use `--package` to also rebuild the deploy zip. |
| 6    | `python scripts\\run-report.py`                  | Full end-to-end build against the fake export. Must return `RESULT=OK\|...\|124\|0\|<prior>`.                                     |
| 7    | `python scripts\\audit-vba.py`                   | Deep static inspection of all `.bas` modules. Read-only; a candidate list, not a gate.                                            |
| 8    | Read `Logs\\PM_Report_*.log`                     | Every step is timestamped with milliseconds since the previous line.                                                              |
| 9    | Kill any stray `EXCEL` process before finishing. | ”                                                                                                                                 |

**Expected result:** `RESULT=OK|...|124|0|<prior>` (124 PMs, 0 skipped, prior
comparison found). Anything else is a bug.

### Testing rules that bite

- **A pass needs three things, not one:** the compile command executed, no error
  dialog appeared, and the probe marker was written. "No dialog" alone is not a
  pass — a silently skipped compile looks identical.
- Compile errors surface as **modal dialogs**, never exceptions: `Application.Run`
  raises `0x800A03E8`/`0x800706BE` with no message. `Logs\watcher.log` has the
  dialog text.
- Never trust a clean log from a workbook that was built and closed without saving
  — it gives a false pass that looks exactly like a working gate.

______________________________________________________________________

## 4. Excel handling — hard rules

- **No blocking Excel call may wait longer than 15 seconds.** Every script wraps
  blocking COM calls (`Workbooks.Open`, `Application.Run`, `SaveAs`,
  `chart.Export`) in `scripts\excel_guard.py` (`ExcelGuard`), which kills the
  Excel process and logs the hang to `Logs\excel-hangs.log` if the timeout is
  exceeded. Nothing waits minutes for a stuck macro.
- **Every Excel session runs a dialog watcher** (`DialogWatcher` in
  `excel_guard.py` launching `scripts\watch-compile-error.py`) alongside it.
  Modal popups (OK clicks, errors, file pickers) are captured with their full
  text and printed in the script output, so a failure caused by a dialog is
  visible instead of silent. Dialogs are also logged to `Logs\watcher.log`.
- A run that stalls is killed after 15 seconds with no new log line. Excel dialogs
  are modal and will hang headless automation silently; this is the only reliable
  way to find out where it stopped.
- **Never leave Excel running** after a test. Kill `EXCEL` and any worker process.
- VBA compiles **one procedure at a time** (compile-on-demand), so running a
  procedure does not compile its neighbours: a typo in a procedure nobody calls
  survives every `Application.Run`. `test-compile.py` triggers the full-project
  compile; that is the gate.

______________________________________________________________________

## 5. VBA-specific rules

- **Module-level declarations (`Const`, `Dim`) belong in the declarations section
  at the top of the module.** A declaration placed between two procedures is not
  visible to other procedures and produces "Compile error: Variable not defined"
  at the use site.
- **Every procedure that can fail has error handling:** `On Error GoTo <label>`
  with a cleanup block, or `On Error Resume Next` for non-critical formatting work.
- **Log errors with procedure name, `Err.Number`, `Err.Description`, and the
  `stage` breadcrumb.**
- **Bounds-check row indices before indexing `data(...)`, and check objects for
  `Nothing` before use.**
- **Guard `CDbl`/`CLng` on cell values with `IsNumeric` or an error handler.**
- **Formatting, charts, and file I/O always run under `On Error Resume Next`.**

### 5a. Deep VBA inspection audit (`scripts\audit-vba.py`)

- Static deep inspection of all `PM_Report_*.bas` modules. Checks patterns the
  compile gate cannot catch (compile-on-demand means a typo in an uncalled
  procedure survives every `Application.Run`).
- **Run it before committing any `.bas` change:** `python scripts\audit-vba.py`.
- It is read-only (`--fix` is accepted for compatibility but does nothing).
- The audit is a candidate list, not a pass/fail gate: review each flag, fix real
  bugs, and treat style items as candidates for future refactoring.
- Severity levels: **CRITICAL** (likely runtime bugs), **HIGH** (error-handling
  gaps), **MEDIUM** (missing stage breadcrumbs), **INFO** (long procedures).
- What it checks:
  - `Err.Raise` without an `On Error GoTo` handler
  - `CDbl`/`CLng`/`CInt`/`CSng` on cell values without an `IsNumeric` guard
  - File I/O / workbook access with no error handling
  - `Fail` label with no `Exit Sub/Function` (falls through on success)
  - `Fail` label with no `LogErr` call (error swallowed)
  - `On Error Resume Next` without `On Error GoTo 0` to restore handling
  - Public procedures that assign `stage` but never call `LogStage`/`SetStatus`
  - Public procedures over 40 lines with no `LogStage`/`SetStatus`
  - Procedures over 80 lines (candidate for refactoring)
  - Module-level `Public`/`Private` declarations outside the declarations section
  - Array indexing `data(r, col)` without recent `UBound`/`LastUsedRow` bounds check

______________________________________________________________________

## 6. Script language — Python only

- All project tooling under `scripts\` is Python (`.py`). Never write or use
  PowerShell (`.ps1`) scripts for build, compile, sync, run-report, watcher,
  audit, or any automation task.
- If a `.ps1` file exists anywhere in the project, it is obsolete and should be
  removed or replaced with a `.py` equivalent.
- All Python scripts resolve paths from `__file__` / `os.path.dirname`, never
  from hardcoded user profiles.

______________________________________________________________________

## 7. Logging standard

- Every run writes `Logs\PM_Report_yyyyMMdd_hhmmss.log` next to the master
  workbook. Never write run logs into `Reports\` (that folder is user output).
- Line format: `yyyy-MM-dd hh:mm:ss  +<ms since previous line>ms  message`.
  The full date is mandatory so an old line can never be mistaken for a fresh one.
- Read logs oldest first, newest at the bottom.
- Stage breadcrumbs come from `RunReport`'s `stage` label; `SetStatus` logs too,
  so every user-visible step is traced automatically.
- Errors are logged as `!!! ERROR <n>: <description>   [stage: ...]`.

______________________________________________________________________

## 8. File layout

| Path                                                                                                              | Role                                                                                                                                                                                                                                                                                                                                                                                                                |
| ----------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PM_Report_Macro.bas`                                                                                             | Canonical VBA source.                                                                                                                                                                                                                                                                                                                                                                                               |
| `PM_Report_UI.bas` / `PM_Report_Host.bas` / `PM_Report_Setup.bas` / `PM_Report_Data.bas` / `PM_Report_Output.bas` | Supporting VBA modules.                                                                                                                                                                                                                                                                                                                                                                                             |
| `PM_Requests_Master_Template.xlsm`                                                                                | Master workbook (embedded module, `START HERE`, settings).                                                                                                                                                                                                                                                                                                                                                          |
| `AOTIOpsTreatmentPMDatesResults_FAKE.xlsx`                                                                        | Synthetic export for regression tests.                                                                                                                                                                                                                                                                                                                                                                              |
| `Reports/`                                                                                                        | Generated reports — build artifacts, never committed.                                                                                                                                                                                                                                                                                                                                                               |
| `Logs/`                                                                                                           | Run logs, watcher log, compile-test log. Safe to delete.                                                                                                                                                                                                                                                                                                                                                            |
| `backups/`                                                                                                        | At most two master snapshots. Safe to delete.                                                                                                                                                                                                                                                                                                                                                                       |
| `scripts\`                                                                                                        | All tooling: Python (`sync_template`, `test-compile`, `run-report`, `snapshot-report`, `update-docs`, `check-toc`, `audit-vba`, `lint`, `typecheck`, `security-scan`, `pyupgrade_check`, `autoformat`, `complexity`, `deadcode`, `deps`, `db_query`, `export_report`, `db_query`, `md_format`, `png_optimize`, `pytest_run`, `monitor-logs`, `pipeline`, `make_fake_data`, `make_prior`). New test scripts go here. |
| `scripts\temp\`                                                                                                   | Temp/scratch files. Safe to delete.                                                                                                                                                                                                                                                                                                                                                                                 |
| `documentation/`                                                                                                  | `AGENTS.md` (rules), `CHANGELOG.md`, `README.md`, `screenshots/`, logo art.                                                                                                                                                                                                                                                                                                                                         |
| `.kilo/command/`                                                                                                  | Kilo slash-command definitions (discovered by Kilo; do not move).                                                                                                                                                                                                                                                                                                                                                   |
| `README.md` (root)                                                                                                | Canonical GitHub-facing README.                                                                                                                                                                                                                                                                                                                                                                                     |

**Hard rule: every file stays inside the project folder.** Scripts resolve paths
from `__file__` / `os.path.dirname`, never from `%TEMP%`, `AppData`, or a
hardcoded user profile, and never write scratch files outside the project.
Scratch belongs in `Logs\` (gitignored), `scripts\temp\` (gitignored), build
output in `Reports\`, master snapshots in `backups\`.

______________________________________________________________________

## 9. Temp file paths

- **Never** use `C:\Users\smd32\AppData\Local\Temp\kilo` for any temp files,
  scratch work, or intermediate output. That path is prohibited.
- All temp/scratch files go under
  `C:\Users\smd32\Downloads\Re-Auth Project\scripts\temp`. Scripts resolve this
  path relative to `__file__`, never from `%TEMP%`, `AppData`, or a hardcoded
  user profile outside this project.
- All other files stay inside the project folder (see section 8 hard rule).

______________________________________________________________________

## 10. Documentation

- Update `documentation\CHANGELOG.md` with every change, newest entry first.
  Entry format: `- **YYYY-MM-DD HH:MM:SS** — what changed and why.`
  Every entry must include date/time, scope (VBA / Python / docs / deploy), and
  a one-line rationale.
- Update `README.md` (root) when paths, scripts, or behaviour change, and
  regenerate `documentation/screenshots/` after any layout change with
  `scripts\snapshot-report.py`.
- Everything else Markdown stays in `documentation/` (plus the root `AGENTS.md`
  stub and `.kilo/command/*.md`, which Kilo must find where it is).

______________________________________________________________________

## 11. Version control

- Always commit and push to git after making changes. If the change is complete
  and verified (compile gate + report build), commit it in the same session; do
  not leave finished work sitting in the working tree.
- Push when a remote and an upstream branch are configured. If there is no
  remote, commit locally and say so.
- One commit per logical change; keep the subject short and imperative.
- Stage only intended files. `git add -A` is acceptable only after `git status`
  confirms nothing unintended (logs, reports, backups) is included.
- Never commit `Reports\*.xlsx`, `Logs\*.log`, `scripts\temp\*`, or `backups\*`.
- Never commit secrets or keys.
- After committing, report the commit id and confirm the tree is clean.

______________________________________________________________________

## 12. Deployment

- After every verified change (mandatory loop passed), build **both** zips:
  1. `Project-Local-Build.zip` — non-macro `PM_Requests_Master_Template_(No Macros).xlsx`
     - all six `.bas` modules + `BUILD-INSTRUCTIONS.txt` (what work machines import
       into a fresh macro-free workbook).
  1. `PM_Requests_Master_Template_Macro_Archive.zip` — macro-enabled
     `PM_Requests_Master_Template.xlsm` + `BUILD-INSTRUCTIONS.txt` only (the `.bas`
     modules are already embedded; no separate `.bas` files).
- End users receive only the `.xlsm`. The `.bas`, `scripts\`, fake export, and logs
  are internal.
- The template must stay self-contained: module imported, logo embedded, button wired
  to `BuildManagerTabs`.

______________________________________________________________________

## 13. Pipeline (`scripts\pipeline.py`)

- `scripts\pipeline.py` runs the full mandatory loop (steps 1–17 of section 3) plus
  screenshots and docs, in order, headless, with 15-second timeouts on every Excel
  call. Use it as the single entry point for a verified change:
  `python scripts\pipeline.py` (add `--skip-push` to commit locally without pushing).
- Pipeline steps:
  1. test-compile (compile gate)
  1. lint (ruff)
  1. typecheck (mypy)
  1. security-scan (bandit + pip-audit)
  1. pyupgrade_check (syntax modernization)
  1. autoformat (autoflake)
  1. complexity (radon)
  1. deadcode (vulture)
  1. deps (deptry)
  1. db_query (ODBC database query)
  1. export_report (pandas export)
  1. sync_template (sync modules + zip)
  1. run-report (end-to-end build)
  1. snapshot-report (screenshots)
  1. png_optimize (Pillow)
  1. md_format (mdformat)
  1. update-docs (README photos + commit + push)
- If any step fails, fix it before committing. Do not push a broken build.
- Run it after every verified `.bas` change, every Python tooling change, and
  before every git push.

______________________________________________________________________

## 14. CHANGELOG format

- Update `documentation\CHANGELOG.md` with every change, newest entry first.

- Entry format: `- **YYYY-MM-DD HH:MM:SS** — what changed and why.`

- Every entry must include: date/time, scope (VBA / Python / docs / deploy), and
  a one-line rationale. Example:

  ```
  - **2026-11-19 18:30:00** — VBA: added LogErr to AddDashboardLogo Fail handler.
    Reason: shape delete on first run swallowed error #-2147024809.
  ```

- Never omit the rationale. A change without "why" is not reviewable.

______________________________________________________________________

## 15. Root AGENTS.md stub sync

- The root `AGENTS.md` (16 lines) is a stub that points to
  `documentation\AGENTS.md`. Kilo loads rules from the root stub.
- **All rule edits go in `documentation\AGENTS.md` only.** Never edit the root
  stub's rule text.
- If you add, remove, or renumber a section in `documentation\AGENTS.md`, update
  the root stub's cross-references to match.
- The root stub's short version (the 6-bullet summary) must stay consistent with
  the full rules; update both together.

______________________________________________________________________

## 16. Branching strategy

- All work happens on `main`. No feature branches unless explicitly approved.
- If a branch is needed, name it `feature/<short-description>` (lowercase,
  hyphens, no spaces).
- Delete feature branches after merging. Never leave stale branches.

______________________________________________________________________

## 17. Fixture / test data management

- `AOTIOpsTreatmentPMDatesResults_FAKE.xlsx` is synthetic test data, never real
  customer data. It is committed to git.
- Regenerate it only when the export schema changes (new headers, new columns,
  new row types). Use `scripts\make_fake_data.py`.
- When regenerating, bump the version comment inside the file and note it in
  `CHANGELOG.md`.
- Never hand-edit the fake export; regenerate it instead.

______________________________________________________________________

## 18. Screenshot naming and regeneration

- Screenshots live in `documentation\screenshots\`.
- Naming convention: `template-<sheet-name>.png` for template sheets,
  `report-<sheet-name>.png` for report sheets (sheet name lowercased, spaces
  replaced with `-`).
- Regenerate screenshots whenever the template layout, report layout, or any
  sheet geometry changes: `python scripts\snapshot-report.py`.
- After regenerating, commit the changed PNGs with the code change that
  motivated them.

______________________________________________________________________

## 19. check-toc.py vs audit-vba.py — which to run when

- `scripts\check-toc.py` — verifies module split integrity: `Option Explicit`,
  matching `VB_Name`, unique procedure names across modules. Run **after any
  structural change** (moving code between modules, adding/removing a procedure,
  changing a module filename).
- `scripts\audit-vba.py` — deep static inspection: error handling, `LogErr`,
  `IsNumeric` guards, stage breadcrumbs, long procedures. Run **before committing
  any `.bas` change**.
- Run both after any `.bas` edit; `check-toc.py` catches split drift,
  `audit-vba.py` catches runtime-risk patterns.

______________________________________________________________________

## 20. Log rotation and cleanup

- `Logs\` is safe to delete at any time; it contains only generated run logs.
- Keep the most recent 20 `PM_Report_*.log` files; older ones may be removed.
- `compile-test.log`, `watcher.log`, `excel-hangs.log`, `pipeline.log`, and
  `compile-errors.log` are append-only; truncate them manually when they grow
  beyond a few thousand lines.
- Never commit `.log` files.

______________________________________________________________________

## 21. Backup rotation

- `backups\` holds at most two master workbook snapshots.
- Create a new backup before any `.bas` change that touches `RunReport` or
  `BuildManagerTabs`.
- When creating a third snapshot, delete the oldest one. Never exceed two.
- Name backups `PM_Requests_Master_Template_<YYYYMMDD>_<HHMM>.xlsm`.

______________________________________________________________________

## 22. Rollback after a failed sync

- If `sync_template.py` fails mid-import (VBA import error, `.xlsm` corruption,
  Excel hang), restore `PM_Requests_Master_Template.xlsm` from the most recent
  snapshot in `backups\`.
- If no snapshot exists, recreate the `.xlsm` from the non-macro deploy zip
  (`Project-Local-Build.zip`) by importing all six `.bas` modules via
  `sync_template.py`, then re-run the mandatory loop.
- After any rollback, re-run `test-compile.py` to confirm the restored state
  compiles cleanly before attempting further changes.

______________________________________________________________________

## 23. Stray Excel cleanup procedure

- Before starting any Excel automation, check `Get-Process excel` and kill any
  stray `EXCEL` processes.
- If Excel is hung from a previous failed run (modal dialog not dismissed), kill
  the process — do not try to dismiss dialogs interactively.
- After any script ends (success or failure), verify `Get-Process excel` returns
  nothing. If it does, kill it.
- Never leave Excel running overnight or across sessions.

______________________________________________________________________

## 24. README dual maintenance

- `README.md` (root) is the canonical, GitHub-facing project README.
- `documentation\README.md` is a pointer to the root README; it contains only a
  one-line link and no independent content.
- When paths, scripts, or behavior change, update root `README.md` first, then
  regenerate screenshots if layout changed.
- Never edit `documentation\README.md` content beyond the link line.
