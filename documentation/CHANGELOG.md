# Changelog

All notable changes to the PM Requests Report Builder are logged here.
Timestamps use `YYYY-MM-DD HH:MM:SS` format.

## 1.4

### 2026-11-14 00:50:00
- Nothing in the project points outside the project folder any more.
  - `scripts\sync_template.ps1` saved the master through `%TEMP%` before swapping it in; it
    now saves to `Logs\_sync_<stamp>.xlsm` and deletes it after the swap. (A leftover
    `%TEMP%` copy of the master from an older run was removed.)
  - `scripts\make_fake_data.py` had the output path hardcoded to this project's folder
    under one user's `Downloads`; it now derives the project root from `__file__`.
  - `documentation\CHANGELOG.md` no longer quotes old temp paths or another machine's user
    profile, and `documentation\AGENTS.md` no longer spells out the sibling project's
    absolute path (the audit finds it by itself).
  - New AGENTS.md rule: scripts resolve paths from `$PSScriptRoot`/`__file__`; scratch goes in
    `Logs\`, output in `Reports\`, snapshots in `backups\`.

### 2026-11-14 00:35:00
- Pending Pickup removed from this project entirely. It was a data filter, not a feature:
  `Treatment Status` still has to exist in the export and is still carried into the report,
  but nothing reads it any more. **Behaviour change:** a row with a manager is now always a
  PM request, so the regression numbers move from `106|18|0` to `124|0`.
  - `PENDING_TEXT` and the status branch in the classification loop are gone; `pendingRows`
    became `excludedRows` and then went away with the filter.
  - Result string is now `OK|savePath|pmTotal|skipped|prior`; `BuildManagerTabs` reads the
    new indexes.
  - Log: `pending=` dropped from `classified` and `SUMMARY`; the `pending skipped` stage is
    gone (15 stages, not 16).
  - `IsManagerTab` no longer special-cases a Pending Pickup sheet; the module header no
    longer advertises a Pending Pickup tab or comparison card.
  - Dashboard footnote and the START HERE bullets no longer mention it.
- `scripts\make_fake_data.py`: generated rows are all `Active` (the 4 real rows are kept
  verbatim, so one historical `Pending Pickup` value remains in the data); the pending-count
  print became a "managers with no rows" check. Regenerated fixture: 124 rows, 12 managers.
- `scripts\make_prior.py` rewritten: it had hardcoded scratch paths outside the
  project and looked for a Pending Pickup sheet that no longer exists. It now
  takes the newest report in
  `Reports\`, drops N rows from each manager tab (header row 9, data from row 10), and names
  the output `PM Requests <date - 7 days>.xlsx` so auto-detection finds it next run.
- Docs: README data rule, result contract, log sample and stage list updated; AGENTS.md
  expected result is `RESULT=OK|...|124|0|<prior>`.

### 2026-11-14 00:20:00
- Cleaned the last of the Pending Pickup feature out of the code and the docs. The
  feature was removed earlier; what was left was stale copy and vocabulary:
  - Module header: version bumped to v1.4, result-string contract corrected to
    `OK|savePath|managerPms|skipped|noManager`, and the workflow/business-rule lists no
    longer claim a Pending Pickup tab or a prior/now/new/cleared card.
  - `pendingRows` renamed to `excludedRows`: these rows are excluded, not "pending".
    `PENDING_TEXT` stays - it is the status value the classifier matches - and now has a
    comment saying what it means.
  - Status text "Building manager and Pending Pickup tabs..." -> "Building manager tabs...".
  - Completion message "Pending Pickup: N" -> "Excluded by status: N".
  - Comments that still described removed behaviour: the tab-name collision note, the
    prior-report sheet list, the formatting helper's row-9 note.
  - Dashboard footnote no longer explains an untracked status; the START HERE bullet about
    skipped rows became a bullet about never overwriting a dated report.
  - Log field `pending=`, the `pending skipped` stage label, and the result-string token
    are unchanged: they are the contract `scripts\run-report.ps1` and `/logs` document, and
    old logs use them.
- README: the "What happens to Pending Pickup rows" section is gone. What remains is one
  data rule under Source file requirements (a row is a PM request when it has a manager and
  its status does not contain "Pending Pickup"), because that rule decides what the report
  contains. Screenshots regenerated.

### 2026-11-14 07:15:00
- Added a root `README.md`, because GitHub only renders a README that sits at the
  repository root: `documentation/README.md` was invisible on the repository page.
  The root file is now the canonical project README (contents, quick start, dashboard and
  manager-tab anatomy, Pending Pickup behaviour, logging, layout, architecture, the
  mandatory loop, the compile gate, testing, the cross-project audit, troubleshooting).
  `documentation/README.md` is now a short pointer to it, so there is still one source of
  truth. AGENTS.md section 6/7 updated to say so.
- Added `scripts\snapshot-report.ps1`: exports any sheet as a PNG (copy as picture ->
  throwaway chart -> `Chart.Export`), which captures fills, fonts, merges and shapes the
  way Excel renders them. Defaults to the newest report in `Reports\`, `-List` prints the
  sheet inventory.
- Added `documentation\screenshots\`: `master-start-here.png`, `report-dashboard.png`,
  `report-alice-fontaine.png`, generated from the current template and the current build.

### 2026-11-13 23:57:00
- `START HERE` redesigned to match the sibling Pending Pickup builder, with the
  PM-specific wording kept:
  - Column geometry is now the peer's (`B=5, C=28, D=28, E=28, F=2`) instead of
    `B=5.6, C=28.6, D=16, E=16, F=2.6`, so the sheet has the same width and the
    number gutter lines up with the step text.
  - Two-tone rule under the title: navy across `B4:D4`, gold `RGB(246,206,62)` in
    `E4:F4`. The row was empty before.
  - `B14:F14` gets the peer's light-blue CTA band `RGB(221,235,247)`, so the button
    sits on a filled strip.
  - "Ready to build this week's report?" is no longer italic, matching the peer.
  - Bullets are now `"•  "` with indent 1 and centred vertically (new
    `AddStartBullet`), instead of `"-  "` with no indent.
  - Logo is 150 pt wide (height keeps the aspect, 42.9 pt) instead of 105 x 30.
  - Button is 360 x 48 instead of 340 x 48.
  - Subtitle wraps, and the footer is `v1.4  •  PM Requests Report Builder  •  Output: ...`.
  - The sheet tab is navy `RGB(31,78,121)`.
  - Bullets are built with `ChrW(8226)` so this module stays ASCII-only and imports
    identically whatever the codepage of the `.bas` file is (the peer ships a UTF-8
    bullet literal instead).
- Fixed stale copy in `GOOD TO KNOW`: it still described a Pending Pickup tab, which
  the report no longer produces. It now says such rows are skipped, and that one tab
  per Assigned Sales Manager plus the Dashboard and Raw Data are produced.

## 1.2

## 1.3

### 2026-11-13 23:45:00
- Cross-project audit added, read-only, in both directions:
  - `scripts\audit-peer.ps1` compares this project with the sibling
    `Pending Pickup Project`: files only one side has, files that differ (with
    per-file line counts and the functions/sections each side adds), VBA procedure
    names (shared vs peer-only), PowerShell/Python functions, and Markdown section
    headings. `-Sections` narrows it to `scripts`, `documentation`, `commands`,
    `vba`, `rules`; `-Report <path>` saves Markdown; `-SymbolLimit` caps the listing.
  - The peer is auto-discovered (a sibling directory with a root `*.bas` and a
    `scripts\` folder); `-PeerPath` overrides it. The same two files exist in the
    peer, so the audit runs correctly from either side.
  - `.kilo\command\audit.md` adds `/audit`, which runs the audit and asks for a
    shortlist of port candidates with effort and risk. It never copies anything.
  - Rules: new AGENTS.md section 8 (Borrowing from the sibling project), and a
    `README.md` row for the new script.

### 2026-11-13 23:35:00
- Compile gate hardened, and made faster:
  - `-SelfTest` switch: injects a deliberately broken procedure, asserts the gate
    FAILS, verifies the fixture is really in the module (a fixture that was never
    saved is the classic false pass), removes it, then asserts the gate passes.
    This exists because a `BadProc` fixture that was closed without saving once
    produced a false pass here.
  - A pass now needs three things: the compile command was available and executed
    (it refuses to report "clean" if the control is disabled or missing), no error
    dialog appeared, and the probe marker was written. Previously "no dialog"
    alone counted as a pass, which a silently skipped compile also produces.
  - Polling is now 100 ms instead of 400 ms and the wait window is 1 s instead of
    3 s, so a clean gate takes ~2 s instead of ~5 s.
- Documented in `documentation/AGENTS.md`: `.kilo\command\logs.md` cannot move to
  `scripts\` - Kilo only discovers commands under `.kilo\command\` (or legacy
  `.kilocode\`), so moving it would silently delete `/logs` for anyone else, and
  `KILO_CONFIG_DIR` is a machine env var, not something the repo can ship.

### 2026-11-13 23:30:00
- Project layout aligned with the sibling Pending Pickup project:
  - Markdown and image assets moved to **`documentation/`** - `README.md`,
    `CHANGELOG.md`, `AGENTS.md`, `aoti-logo.png`, `aoti-logo-2023.svg`. A stub
    `AGENTS.md` stays at the root because Kilo loads the rules from there; it
    points at `documentation/AGENTS.md`, so there is still only one rule set.
  - Scratch and test tooling now lives in **`scripts/`** instead of
    a scratch folder outside the project. Two Python helpers were kept because they still earn
    their place (`make_fake_data.py` regenerates the fake export,
    `make_prior.py` builds a prior-week report for comparison tests); the rest of
    that folder was one-off investigation scratch whose functionality now lives in
    `test-compile.ps1` and `watch-compile-error.ps1`.
  - Logo path: `AddDashboardLogo` now resolves `documentation\aoti-logo.png`
    first, then falls back to the workbook root, so the art moves with the docs
    without breaking older folders.

### 2026-11-13 23:20:00
- Module renamed: the standard module is now **`PM_Report_Macro`**. VBA module
  identifiers cannot contain spaces (importing a module called `PM Report Macro`
  fails with "Not a legal object name"), so the file is `PM_Report_Macro.bas` -
  matching `Pending_Pickup_Macro.bas` in the sibling project. A spaced file name
  was tried first and rejected: it only ever causes quoting bugs (the worker
  launch in `run-report.ps1` failed on `-File 'C:\...\PM Report Macro.bas'` until
  the path was quoted). The button's `OnAction` still points at `BuildManagerTabs`,
  so nothing in the UI changes. `sync_template.ps1` removes the legacy
  `CreateManagerTabs` component on its next run so no stale copy survives.
- Run logging: every build writes `Logs\PM_Report_yyyyMMdd_hhmmss.log` next to the
  master workbook. Every line is `yyyy-MM-dd hh:mm:ss  +<ms since previous line>ms  message`,
  so a hang shows up as one line followed by silence and the line before it names
  the step that hung. Stage breadcrumbs come from `RunReport`'s `stage` label;
  `SetStatus` also logs, so each user-visible step is traced. Errors are logged as
  `!!! ERROR <n>: <description>   [stage: ...]`.
- New `/logs` slash command (`.kilo/command/logs.md`): prints the run logs
  oldest-first so the newest run is the last block, plus the slowest steps across
  the newest five logs.
- New `run-report.ps1` watchdog: builds one report headlessly and, if no new log
  line appears for 10 seconds (or no log file at all), kills Excel and the worker,
  prints the last 40 log lines with the slowest step, and exits `2`. `build-worker.ps1`
  is its headless worker.
- All PowerShell tooling moved into `scripts/` (`sync_template`, `run-report`,
  `build-worker`, plus the three below), following the sibling Pending Pickup
  project's layout. Each script resolves the project root from `$PSScriptRoot`.
- Ported from the Pending Pickup project, adapted to this one:
  - `scripts\test-compile.ps1` - compile gate that copies the master, starts the
    dialog watcher, appends a probe *inside* `PM_Report_Macro` (VBA compiles one
    procedure at a time, so a probe in its own module would prove nothing) and
    runs it. Verdict lands in `Logs\compile-test.log`.
  - `scripts\watch-compile-error.ps1` - stands outside Excel, logs every modal
    dialog to `Logs\watcher.log` with a timestamp, dismisses error dialogs, and
    kills the Excel process if one will not close. This replaces the manual
    "read the highlighted token from the VBE" procedure that found the
    `MANAGER_LAST_COL` bug.
  - `scripts\answer-dialogs.ps1` - answers the real UI path's dialogs (source
    picker, output folder, final MsgBox) so `BuildManagerTabs` can run unattended.
- New `AGENTS.md` at the project root: rules for agents - sync, compile-test, run,
  `/logs`, the 10-second kill rule, module naming, log line format, file layout.
  Sibling project keeps the same file, so the two projects now behave alike.
- `test-compile.ps1` upgraded from a call-graph probe to a **real whole-project
  compile**. VBA compiles one procedure at a time (compile-on-demand), so a probe
  only proves its own call graph: a negative test - a `BadProc` with an undefined
  name, never called - still reported OK. The gate now triggers
  **Debug > Compile VBAProject** through the VBE command bar (Debug menu, control
  id 578), which compiles every procedure in every module, then runs the execution
  probe. Verified both ways: clean master passes, and the `BadProc` copy now fails
  with `module=PM_Report_Macro line=64 token=[NoSuchNameHere =]` plus the dialog
  text from `Logs\watcher.log`. Gate time is ~5 s.
- The compile gate now **imports `PM_Report_Macro.bas` into the test copy before
  compiling**. It used to compile whatever was embedded in the master, so a stale
  module could pass the gate and then fail on the first real run - which is exactly
  what happened with the logo-path change below.
- Fixed "Compile error: Expected array" at `logoPath = LogoPath()`: the local
  `String logoPath` and the new function `LogoPath` are the same identifier to VBA,
  so `LogoPath()` parsed as an array index. Renamed the function to
  `ResolveLogoPath()`. Verified the logo still lands on the report Dashboard
  (`AOTI_Logo` picture present in the generated `.xlsx`).

### 2026-11-13 15:00:00
- Dashboard KPI cards no longer show nonsense without a prior report: with no
  comparison, "Completed" and "Prior PMs Still Open" are 0 instead of the full
  PM total (they are derived from prior minus retained, which is meaningless at
  prior = 0).
- XLIDE is no longer needed: the module is synced with `sync_template.ps1`, a
  plain PowerShell + Excel COM script (remove + import via `VBProject`,
  rebuild `START HERE`, save). It needs "Trust access to the VBA project object
  model" once, nothing else.
  - Idempotent: when the embedded module already matches the `.bas` file it does
    nothing at all - no backup, no write. That is why backups used to pile up:
    the old apply step copied the master on every invocation, changed or not.
  - A backup is taken only when the module actually changes, and only the two
    newest are kept.
  - Saves via `SaveAs` to a temp file + file copy, because `Save()` on the master
    can block on an invisible modal under headless automation.
- `CreateManagerTabs.bas` moved from `src/` to the project root (the sync script
  and the docs point there); `src/` is gone.
- `START HERE` redesigned to match the Pending Pickup builder: narrow gutters,
  navy title + "HOW IT WORKS" / "GOOD TO KNOW" label bars, four numbered step
  rows on alternating light fills, "Ready to build this week's report?" prompt,
  the CTA in a tall row, and a version footer. The logo sits at the right of the
  title rows.
- The CTA is now a rounded-rectangle AutoShape (was a form-control button), so
  the old VML `<x:FmlaMacro>` patching step is obsolete; `SetupReportButton`
  recreates it fresh every run, so type/style cannot drift.
- `RebuildStartSheet` is a public entry point so the sync script can rebuild the
  UI without running a report.

### 2026-11-13 14:45:00
- Weekly mode finds the prior report itself. With no explicit prior, the builder
  scans the output folder for `PM Requests *.xlsx`, scores the six newest by
  name overlap with the current report, and compares against the newest high
  scorer (a quarter of the names, minimum 5; stops at 90%). The prior-report
  dialog is gone, so the weekly run is one click; `RunReport` now returns the
  prior file name as a 6th field and the completion message shows it (or says
  no prior was found). Explicit `priorPath` (automation/tests) still wins.
- Report layout rebuilt to match the 2.0.18 work build:
  - Dashboard: white logo zone + navy title band, meta strip (current / prior /
    created), six KPI cards (PMs Requested, Prior PMs, Completed, New PMs,
    Pending Pickups, Prior PMs Still Open), manager table with Completed / New /
    per-manager Pending Pickup counts, and the three-block Sales Rep grid.
  - Manager tabs: KPI cards on rows 2-7, table header on row 9, data from row 10,
    "PMs by Rep" card. New PM rows stay light green + bold.
- Dashboard logo is placed from `aoti-logo.png` (works in automation, no
  clipboard) with a clipboard fallback for visible Excel; any existing logo is
  removed first so re-runs stay idempotent.
- Fixed four real defects found by the regression run:
  - `repCounts(idx) = ...` raised 424 "Object required" (Collection items cannot
    be reassigned); rep tallies now live in a Long array.
  - The rep-count array parameter was `ByVal`, which VBA rejects for arrays
    ("Array argument must be ByRef") - now ByRef.
  - `SortManagersByLastName` was called with a Collection (arrays only) -> type
    mismatch; the per-manager rep list now keeps first-appearance order.
  - The error handler read `Err` after cleanup statements, so failures reported
    `ERROR|0`; it now snapshots number/description/stage on entry.
- `RunReport` reports a `stage` marker with any failure to localize the step.
- Project folder pruned: 41 stale backups, the `.bak_button` snapshot, the stale
  generated report, and the local 2.0.18 `.bas` reference removed; README file
  inventory updated.

### 2026-11-13 13:55:00
- The report is weekly, not monthly: banner/title/comments/README updated to
  "WEEKLY PM REQUEST REPORT" and weekly wording throughout the UI and docs.
- Company logo added (`aoti-logo-2023.svg` supplied by the user; rasterized to
  `aoti-logo.png` with Edge headless since the toolchain has no SVG rasterizer):
  - START HERE: embedded picture (AOTI_Logo, 105x30) at the top-right of the
    navy banner, shipped inside the .xlsm.
  - Report Dashboard: logo at the top-right of the title band, copied from the
    embedded START HERE shape (the end user receives only the template, so the
    logo must live inside the .xlsm; no external png dependency). Skipped in
    invisible automation (clipboard copy can crash a hidden instance) - the
    earlier headless "crash" was actually Excel being closed mid-run, so the
    skip is a test-environment guard, not a workaround.
- `aoti-logo.png` is kept as source art only; regression passes
  (`OK|...|106|18|0`, runs 9/10).

### 2026-11-13 13:50:00
- Redesigned the START HERE tab (the template's first thing the user sees):
  navy 20pt banner, tagline, "HOW IT WORKS" 3-step list, light-blue button
  card with a centered 340x56 bold button, footer with the output filename,
  gridlines hidden, column grid normalized. `GetOrCreateStartSheet` rebuilds
  the canvas idempotently on every run, so existing installs upgrade
  themselves on the next build. The styled layout is already applied to the
  shipped template (COM restyled + SaveAs swap; VBA save via COM hangs
  headlessly). Regression after the redesign: `OK|...|106|18|0`.

### 2026-11-13 13:40:00
- Fixed the START HERE button error "Cannot run the macro": the interactive
  shell procedure was named `CreateManagerTabs`, the same as its module.
  Excel's macro engine resolved the name to the module (not a runnable
  procedure), so the button failed even though `SetupReportButton` and
  `RunReport` ran fine. Renamed the shell to `BuildManagerTabs` and rewired
  the button's stored reference (`xl/drawings/vmlDrawing1.vml`) to match.
  Verified via COM: the procedure now starts (reaches the source-file dialog)
  instead of erroring immediately. README and this changelog updated.

### 2026-11-13 13:20:00
- Post-port code review of the module (v1.2, 1202 lines) and fixes:
  1. Crash risk: `BuildDashboard` indexed the current data array with
     `nameCol` without checking it was found; a report without a `Name`
     column + a prior comparison would have raised "subscript out of range".
     Now guarded (`priorNames.Count > 0 And nameCol > 0`).
  2. Dashboard with no prior showed a misleading "New PMs: 0" card; it now
     shows a "Managers" count card instead.
  3. The status bar kept "Step 1 of 3..." text when the user cancelled the
     source/prior/folder dialogs; it is now reset on cancel.
- Output folder preset: the remembered folder in `PM Settings!B2` was the old
  work path on another machine; it is now `<project>\Reports`. Set by a direct XML edit of
  the sheet part (Excel's own save of the master hangs in headless COM), with
  VBA and all sheets verified intact afterwards (opens cleanly in Excel 16.0).
- Removed orphaned Excel save temp files (`5C553B00`, `E4C73B00`) left by the
  killed headless sessions.
- Regression re-run after the fixes: `OK|...|106|18|0` with and without prior.

### 2026-11-13 13:10:00
- Cleanup after the compare port: removed the stale LibreOffice lock file
  (`.~lock.AOTIOpsTreatmentPMDatesResults_FAKE.xlsx#`), the superseded
  `src/CreateManagerTabs_trace.bas` (LO-trace scaffolding), the original tiny
  sample `AOTIOpsTreatmentPMDatesResults820.xlsx` (replaced by
  `AOTIOpsTreatmentPMDatesResults_FAKE.xlsx` as the regression input), and the
  local reference copy `Create_Manager_Tabs_Monthly_v2.0.18.bas` (compare
  feature is ported and verified; the live 2.0.18 build at work is untouched).

### 2026-11-13 13:03:29
- Ported the prior-report comparison feature from the 2.0.18 work build into the
  1.x module (Excel-only runtime):
  - The interactive shell now asks for the prior PM Requests report after the
    source export (Cancel = no comparison). `RunReport` gained an optional
    `priorPath` parameter.
  - `LoadPriorComparison` reads the prior report read-only: every sheet except
    `Dashboard` / `Pending Pickup` / `Raw Data` / `START HERE` is a manager tab
    (Name header found within the first 30 rows, fallback column C); Pending
    Pickup tab feeds the prior-pending name set. An unreadable prior degrades to
    "no comparison" instead of failing the build.
  - Manager tabs: new-PM rows (names absent from the prior report) are light
    green + bold, overriding the banding (same color as 2.0.18).
  - Pending Pickup tab: comparison card with Prior / Now / New / Cleared counts.
  - New `Dashboard` tab (first in the report): KPI cards (Current PMs, Prior PMs,
    Net Change excluding new PMs, New PMs; Pending Now / Prior / New / Cleared)
    and a manager table (PMs Requested, Prior PMs, Change, New PMs) with
    hyperlinks to each manager tab.
  - `Raw Data` is copied into the finished report as the last tab (2.0.18 layout).
  - Verified in Excel 16.0 (real COM run, ~1.2 s): no-prior run returns
    `OK|...|106|18|0` with unchanged formatting; prior run with rows removed from
    the prior reports 36 new PMs (12 managers x 3), 36 green rows, pending card
    `Prior: 14 / Now: 18 / New: 4 / Cleared: 0`, dashboard change/new columns
    correct. Excel-only validation: headless LibreOffice is no longer used.

### 2026-11-13 12:30:00
- Removed all LibreOffice/UNO scaffolding from the module (headless LO could not
  open the report workbooks; Excel 16.0 is now available for testing): dropped
  the Trace/SelfTest/IsExcelHost helpers, `SetStatus` host guard,
  `FindOpenWorkbook`/open-source reuse, `PathSep`, clipboard helpers, and the
  `InputBox` fallbacks for the dialogs.
- Replaced clipboard-based format copying with property-based per-cell
  `CopyRowFormats` (NumberFormat + font name/size/bold/color) in
  `FormatManagerSheet` and `LoadSourceIntoRawData` (clipboard hung headless LO
  and failed under Excel COM automation).
- `FreezePanes` is now best-effort (`On Error Resume Next`): invisible automation
  has no window, which raised `1004 Unable to set the FreezePanes property`.
- Header font: `ApplyReaderFriendlyFormatting` copies font name/size from the
  first data row, fixing the "Aptos Narrow 11" header carried from the source.
- Runtime verification in Excel 16.0 via COM: `RunReport` green
  (`OK|...|106|18|0`, ~1.1 s); output verified with openpyxl (13 sheets, navy
  header Arial 10, freeze A2, autofilter, red pending text, PMs Needed box,
  date formats).

## 1.0

### 2026-11-13 11:45:18
- Second review pass found six issues in the optimized module; all fixed and
  re-applied to the master (backup in `backups/`, analysis 0 errors / 0
  warnings, oletools extraction matches the canonical source, workbook opens
  cleanly; XLIDE review pending in the tree):
  1. Pending Pickup tab with zero rows now gets the full header styling
     (borders, navy header, AutoFilter, frozen row) instead of a bare header.
  2. The master is now saved immediately after a new output folder is chosen,
     so the remembered folder survives closing without a save prompt.
  3. Removed the unused `lastSourceCol` parameter from `FormatManagerSheet`.
  4. PMs Needed box no longer overrides the 34pt header row height (rows 2:3
     only).
  5. Freeze-panes is guarded against a missing window (safe under automation).
  6. Trailing path separators are stripped from output-folder paths
     (new `NormalizeFolder`; prevents double-separator paths).

### 2026-11-13 11:42:02
- Applied the optimized module to the master `PM_Requests_Master_Template.xlsm`
  through the xlide-mcp server code path (pyopenvba; content-token guarded;
  backups in `backups/`).
- One fix made during verification: replaced the `sourceWasOpened` flag pattern
  with `If Not srcWb Is Nothing` so the analyzer can prove the object is set
  (the old pattern raised a run-time-91 finding).
- Verified: read-back token matches the canonical source; static analysis
  reports 0 errors / 0 warnings; oletools extraction matches
  `src/CreateManagerTabs.bas` exactly; workbook opens cleanly.
- XLIDE shows the change in its project tree with a before/after diff and
  pending review (Keep/Revert). Runtime testing in Excel is still pending
  (no Excel on this machine).

### 2026-11-13 11:38:43
- Toolchain established for working on the project without Excel on this machine:
  XLIDE (VS Code extension) for editing the VBA embedded in the `.xlsm` directly,
  Python/oletools for extraction and verification, openpyxl for report
  generation and testing, LibreOffice for file-validity checks. The optimized
  module is applied to the master via XLIDE (open module > replace with
  `src/CreateManagerTabs.bas` > save), then verified by extracting the embedded
  module with oletools.

### 2026-11-13 11:36:25
- Code review of the VBA module completed. Issues found:
  - **Null crash (error 94):** `CStr()` on empty cells in the classification
    loop and `FindHeaderColumn` — any blank manager/status cell or blank header
    cell aborted the whole run.
  - **Date/number formats lost:** only header-row formats were copied, so date
    columns rendered as serials (e.g. `46003` instead of `2025-12-12`).
  - **Inaccurate PM count** in the completion message (counted rows with no
    manager that appear on no tab).
  - **Sheet-name collisions:** managers like `A/B` vs `A\B` (both sanitize to
    `A-B`), or a manager named "Pending Pickup", caused a duplicate-name error.
  - **`Application.FileDialog`** folder picker is unsupported on macOS Excel
    (first-run setup would fail there).
  - Dead code: `FormatOutputSheet` was never called.
  - Unqualified button `OnAction` (could break after rename/copy).
  - Mojibake em-dashes (`â€"`) in status-bar strings (ANSI/UTF-8 mismatch).
- Optimized module v1.0 written to `src/CreateManagerTabs.bas` (canonical source
  from now on). Fixes: Null-safe `SafeStr` everywhere, full-range/column format
  copy (dates survive), accurate PM count + skipped-row count in the completion
  message, `UniqueSheetName` collision handling, macOS folder-path fallback,
  dead code removed, workbook-qualified `OnAction`, ASCII-only strings.
- The embedded copy inside `PM_Requests_Master_Template.xlsm` is still the old
  module — apply the update by importing `src/CreateManagerTabs.bas` (Alt+F11 >
  remove module > File > Import File), then save.

### 2026-11-13 11:33:34
- Established project version **1.0** (baseline carried over from the original v2.0.22 distribution; no version bumps planned — changes are tracked here).
- Overhauled the readme: replaced `README_PM_Requests_v2.0.22.txt` with `README.md` (overview, business rules, setup, monthly workflow, output spec, troubleshooting, architecture notes).
- Renamed the master workbook: `PM_Requests_Master_Template_v2.0.22_CLEAN.xlsm` → `PM_Requests_Master_Template.xlsm` (no code changes).
- Added this changelog.

### 2026-11-13 11:31:00
- Validated the master `.xlsm` package (opens cleanly).
- Verified the full report pipeline against the sample raw data (`AOTIOpsTreatmentPMDatesResults820.xlsx`, fake data) with a faithful algorithm simulation — all business rules pass (manager tabs sorted by last name, PMs Needed = row count, Pending Pickup tab last in red, frozen headers, filters, banding, unique output filename).
- Open issue: date columns render as Excel serial numbers (e.g. `46003` instead of `2025-12-12`) in generated reports, because only header-row formats are copied from the source. Fix pending.