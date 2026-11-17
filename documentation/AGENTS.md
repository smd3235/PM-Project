# Agent Rules — PM Requests Report Builder

These rules apply to every change, no exceptions.

## 1. The one true source

- `PM_Report_Macro.bas` in the project root is the **canonical source** of the VBA.
  Never edit the module inside the `.xlsm`; edit the `.bas` and sync it.
- The embedded module is named **`PM_Report_Macro`**. VBA module identifiers cannot
  contain spaces (importing a module called `PM Report Macro` fails with "Not a
  legal object name"), which is why the file name and the identifier both use
  underscores. The `START HERE` button runs `BuildManagerTabs` by procedure name,
  so the module name is invisible to users.
- `scripts\sync_template.py` is the only supported way to push a `.bas` change
  into `PM_Requests_Master_Template.xlsm`.

## 2. Mandatory loop after any `.bas` change

1. `.\scripts\test-compile.py` — compile gate (fast, seconds). It compiles the
   whole project, not a call graph.
2. `.\scripts\sync_template.py` — import into the master and rebuild `START HERE`.
3. `.\scripts\run-report.py` — full end-to-end build against the fake export.
4. `/logs` — read the run log; every step is timestamped with milliseconds since
   the previous line.
5. Kill any stray `EXCEL` process before finishing.

Expected result of step 3: `RESULT=OK|...|124|0|<prior>` (124 PMs, 0 rows without a
manager). Anything else is a bug.

### Testing rules that bite

- **Prove your fixture when you test a detector.** A broken workbook that was built
  and closed without saving gives a false pass and looks exactly like a working
  gate. After editing the gate run `.\scripts\test-compile.py -SelfTest`: it
  injects a broken procedure, asserts the gate fails, removes it, and asserts the
  gate passes. If it prints `SELFTEST FAILED`, believe it over your reading of the
  output.
- **A pass needs three things, not one:** the compile command was available and
  executed, no error dialog appeared, and the probe marker was written. "No dialog"
  alone is not a pass - a silently skipped compile looks identical.
- Compile errors surface as modal dialogs, never as exceptions: `Application.Run`
  raises `0x800A03E8`/`0x800706BE` with no message. `Logs\watcher.log` has the text.

## 3. Excel handling — hard rules

- **A run that stalls is killed after 10 seconds** with no new log line. Excel
  dialogs are modal and will hang headless automation silently; this is the only
  reliable way to find out where it stopped.
- **Never leave Excel running** after a test. Kill `EXCEL` and any worker process.
- **Modal dialogs must be logged with their full text before being dismissed.**
  `scripts\watch-compile-error.py` does this into `Logs\watcher.log`.
- Compile errors surface through modal dialogs, not exceptions:
  `Application.Run` raises `0x800A03E8`/`0x800706BE` with no message. Read
  `Logs\watcher.log` for the dialog text.
- VBA compiles **one procedure at a time** (compile-on-demand), so running a
  procedure does not compile its neighbours: a typo in a procedure nobody calls
  survives every `Application.Run`. `scripts\test-compile.py` therefore triggers
  **Debug > Compile VBAProject** (VBE command bar, Debug menu, control id 578),
  which compiles every procedure in every module. That is the gate; run it before
  `sync_template.py`.

## 4. VBA-specific rules

- The module is one file with a **MAP block** at the top: section start lines and
  the procedures in each section. Each section also carries its own contract
  (what it owns, its entry points, and what must stay true).
  After moving code between sections, run `.\scripts\check-toc.py` and fix any
  drift with `-Fix`; a stale MAP is worse than none, because it sends the next
  reader to the wrong place.
- All module-level declarations (`Const`, `Dim`) belong in the **declarations
  section** at the top of the module. A `Private Const` placed between two
  procedures is not visible to other procedures and produces
  "Compile error: Variable not defined" at the use site.
- Every procedure that can fail has error handling: `On Error GoTo <label>` with a
  cleanup block, or `On Error Resume Next` for non-critical formatting work.
- Log errors with procedure name, `Err.Number`, `Err.Description`, and the
  `stage` breadcrumb.
- Bounds-check row indices before indexing `data(...)`, and check objects for
  `Nothing` before use.
- Guard `CDbl`/`CLng` on cell values with `IsNumeric` or an error handler.
- Formatting, charts, and file I/O always run under `On Error Resume Next`.

## 5. Logging standard

- Every run writes `Logs\PM_Report_yyyyMMdd_hhmmss.log` next to the master
  workbook. Never write run logs into `Reports\` (that folder is user output).
- Line format: `yyyy-MM-dd hh:mm:ss  +<ms since previous line>ms  message`.
  The full date is mandatory so an old line can never be mistaken for a fresh one.
- `/logs` prints logs **oldest first, newest at the bottom**.
- Stage breadcrumbs come from `RunReport`'s `stage` label; `SetStatus` logs too, so
  every user-visible step is traced automatically.
- Errors are logged as `!!! ERROR <n>: <description>   [stage: ...]`.

## 6. File layout

| Path | Role |
|---|---|
| `PM_Report_Macro.bas` | Canonical VBA source. |
| `PM_Requests_Master_Template.xlsm` | Master workbook (embedded module, `START HERE`, settings). |
| `AOTIOpsTreatmentPMDatesResults_FAKE.xlsx` | Synthetic export for regression tests. |
| `Reports/` | Generated reports — build artifacts, never committed. |
| `Logs/` | Run logs, watcher log, compile-test log. Safe to delete. |
| `backups/` | At most two master snapshots. Safe to delete. |
| `scripts/` | All tooling: PowerShell (`sync_template`, `test-compile`, `run-report`, `build-worker`, `watch-compile-error`, `answer-dialogs`, `audit-peer`, `snapshot-report`, `dump-sheet`, `check-toc`) and Python helpers (`make_fake_data`, `make_prior`). New test scripts go here. |
| `.kilo/command/logs.md` | The `/logs` slash command. Must stay here: Kilo only discovers commands under `.kilo\command\` (or legacy `.kilocode\`), walking up from the working directory. Moving it to `scripts\` silently removes `/logs` for everyone; `KILO_CONFIG_DIR` could redirect discovery but it is a machine env var, not something a repo can ship. |
| `.kilo/command/audit.md` | The `/audit` slash command: cross-audits the sibling `Pending Pickup Project` for portable scripts, VBA procedures and rules, and the reverse. Read-only; porting is a separate approved task. The same two files exist in the peer so the audit runs from either side. |
| `documentation/` | `AGENTS.md` (rules), `CHANGELOG.md`, `README.md` (pointer to the root README), `screenshots/`, the logo art. A stub `AGENTS.md` at the root points here so Kilo still loads the rules. |
| `README.md` (root) | The canonical, GitHub-facing project README. GitHub only renders a README at the repository root, so this is the one to edit; `documentation/README.md` just points here. |

Hard rule: **everything stays inside the project folder.** Scripts resolve paths from
`$PSScriptRoot`/`__file__`, never from `%TEMP%`, `AppData`, or a hardcoded user profile, and
never write scratch files outside the project. Scratch belongs in `Logs\` (gitignored); build
output in `Reports\`; master snapshots in `backups\`.

## 7. Documentation

- Update `CHANGELOG.md` with every change, newest entry first, format
  `- **HH:MM:SS** — what changed and why.`
- Update `README.md` (root) when paths, scripts, or behaviour change, and regenerate
  `documentation/screenshots/` after any layout change with `scripts\snapshot-report.py`.
- Everything else Markdown stays in `documentation/` (plus the root `AGENTS.md` stub and
  `.kilo/command/*.md`, which Kilo must find where it is).

## 8. Borrowing from the sibling project

The sibling **Pending Pickup** project (the other project folder next to this one) is a
separate report builder with the same harness design and different data.
`scripts\audit-peer.py` (via `/audit`) compares the two read-only and reports what each side
has that the other does not: files, VBA procedure names, script function names, and Markdown
sections. It finds the sibling by itself; no absolute path is written down.

- Use `/audit` before writing a new helper. Something similar often already exists
  on the other side.
- Copying is allowed, but the copy is a new file here: adapt its data source, paths,
  and logging to this project, then run the mandatory loop and update the changelog.
- Never change the peer project from here, and never let the audit write into it.
- If a file exists in both projects, treat this project's copy as authoritative for
  this project and note the divergence in the changelog.

## 9. Version control

- Always commit and push to git after making changes. If the change is
  complete and verified (compile gate + report build), commit it in the same
  session; do not leave finished work sitting in the working tree.
- Push when a remote and an upstream branch are configured. If there is no
  remote, commit locally and say so.
- One commit per logical change; keep the subject short and imperative.
- Stage only intended files. `git add -A` is acceptable only after
  `git status` confirms nothing unintended (logs, reports, backups) is
  included.
- Never commit `Reports\*.xlsx`, `Logs\*.log`, or `backups\*`.
- Never commit secrets or keys.
- After committing, report the commit id and confirm the tree is clean.

## 10. Deployment

- End users receive only the `.xlsm`. The `.bas`, `scripts\`, fake export, and logs
  are internal.
- The template must stay self-contained: module imported, logo embedded, button
  wired to `BuildManagerTabs`.