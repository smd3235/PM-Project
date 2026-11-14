# PM Requests Report Builder

**Version:** 1.3
**Platform:** Desktop Excel (Windows or macOS) with VBA support

A one-click Excel macro workbook that turns a raw treatment export into a
formatted, ready-to-distribute **PM Requests** report: a summary **Dashboard**,
one tab per assigned Sales Manager, a **Pending Pickup** tab, and a copy of the
raw data — optionally highlighting what changed versus the prior report.

---

## What it does

1. You click **CREATE PM REQUEST REPORT** on the `START HERE` sheet.
2. You pick the current raw-data export (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`).
3. The previous week's report is found automatically in the remembered output
   folder and used for comparison (nothing to pick). Reports built by the 2.0.18
   work build are read fine — they are opened read-only.
4. The macro finds the source worksheet automatically by looking for the two
   required headers, validates the data, and loads it into the `Raw Data` sheet
   of the master workbook.
5. It builds a **Dashboard** summary tab: KPI cards, a hyperlinked manager table
   with Completed / New / Pending Pickup counts, and a Sales Rep grid.
6. It builds one tab per Sales Manager (sorted by manager last name), each with
   KPI summary cards and a "PMs by Rep" breakdown, reader-friendly formatting,
   filters, frozen headers — and, when comparing, new PM rows highlighted light
   green.
7. It builds a **Pending Pickup** tab (red text) for rows whose Treatment Status
   contains "Pending Pickup", with a prior/now/new/cleared comparison card when
   comparing.
8. It copies the loaded raw data in as the last tab and saves the finished report
   to the remembered output folder as `PM Requests YYYY-MM-DD.xlsx` (or a
   numbered name if that file already exists).

The source export is opened **read-only** and closed without saving; the master
workbook itself is never modified except for its remembered output-folder setting.

---

## Project files

| File | Purpose |
|---|---|
| `PM_Requests_Master_Template.xlsm` | The master macro workbook. Contains the embedded VBA, the `START HERE` UI (with the company logo), and the hidden settings sheet. This is the file you use every week. |
| `PM_Report_Macro.bas` | **Canonical source of the VBA module.** The embedded copy in the `.xlsm` is synced from this file with `scripts\sync_template.ps1` (see "Updating the macro code"). |
| `AOTIOpsTreatmentPMDatesResults_FAKE.xlsx` | Synthetic raw-data export (124 rows, 12 managers) used for regression testing. |
| `documentation\aoti-logo.png` | Company logo art (rasterized from `aoti-logo-2023.svg`); read at run time to place the logo on the report Dashboard. |
| `documentation\aoti-logo-2023.svg` | Original vector logo art, kept so the PNG can be regenerated. |
| `scripts\sync_template.ps1` | Imports `PM_Report_Macro.bas` into the master's VBA project and rebuilds `START HERE`. Run it after every `.bas` change. |
| `scripts\run-report.ps1` | Builds one report headlessly and watches the run log: if no new log line appears for 10 s it kills Excel and prints the step that hung. |
| `scripts\build-worker.ps1` | Headless worker used by `scripts\run-report.ps1` (one Excel instance, one `RunReport` call). |
| `scripts\test-compile.ps1` | Compile gate: imports the `.bas` into a copy of the master, triggers **Debug > Compile VBAProject** (compiles every procedure, not just the ones a probe calls), then runs an execution probe. `-SelfTest` proves the detector detects. Prints the exact module, line, and token on failure. |
| `scripts\watch-compile-error.ps1` | Stands outside Excel and records/dismisses any modal dialog to `Logs\watcher.log` (compile errors, runtime errors, security prompts). |
| `scripts\answer-dialogs.ps1` | Answers the real UI path's dialogs (source picker, output folder, final MsgBox) so `BuildManagerTabs` can run unattended. |
| `scripts\audit-peer.ps1` | Read-only cross audit against the sibling `Pending Pickup Project`: which files, VBA procedures, PowerShell/Python functions, and Markdown sections exist on one side only. `-Sections` narrows it, `-Report <path>` saves Markdown. Used by the `/audit` command. |
| `documentation\AGENTS.md` | Rules for any agent working on this project: sync, compile-test, run, log. |
| `Logs/` | One run log per build, `PM_Report_yyyyMMdd_hhmmss.log`. Written by the macro. Safe to delete. |
| `.kilo/command/logs.md` | The `/logs` slash command: shows the run logs and the slowest steps. |
| `backups/` | At most two snapshots of the master, taken only when the embedded module actually changes. Safe to delete. |
| `CHANGELOG.md` | Version history with timestamps. |

The VBA module is embedded in the `.xlsm` — there is no separate `.bas` file to
import in normal use. (The "import `Create_Manager_Tabs_Monthly_v2.0.22.bas`"
step from older readmes is legacy and no longer required.)

---

## Business rules

These are the rules the report enforces. Change them only by changing the code.

- Every data row assigned to a manager (non-empty **Assigned Sales Manager**)
  is a PM request.
- Rows whose **Treatment Status** contains `Pending Pickup` (case-insensitive)
  are excluded from manager tabs and manager PM counts, and appear **only** on
  the Pending Pickup tab.
- **PMs Needed** on a manager tab equals the number of data rows on that tab.
- **Treatment PM Date is NOT used** to decide whether a PM is needed.
- A manager with zero non-pending rows gets no tab (but will still appear on
  Pending Pickup if they have pending rows).

---

## One-time setup

1. Open `PM_Requests_Master_Template.xlsm` in desktop Excel.
2. If Excel blocks macros, allow them: **File > Options > Trust Center > Trust
   Center Settings > Macro Settings** (or the yellow security bar > Enable).
3. On the `START HERE` sheet, click **CREATE PM REQUEST REPORT** (or run
   `SetupReportButton` once from `Alt+F8` if the button is missing — it is
   re-installed automatically on every run anyway).
4. The first run asks where finished reports should be saved. That folder is
   remembered in the very-hidden `PM Settings` sheet and reused automatically
   on later runs. If the folder disappears, you are asked for a replacement.
   On macOS there is no folder picker, so the path is typed into an InputBox.

### Updating the macro code

`PM_Report_Macro.bas` in the project root is the canonical source; the copy
embedded in the `.xlsm` must be kept in sync. Use the sync script (it needs
desktop Excel and "Trust access to the VBA project object model" enabled once in
Trust Center > Macro Settings):

```powershell
.\scripts\sync_template.ps1
```

It compares the embedded module with the `.bas` file, does nothing when they
match, and otherwise takes one backup, re-imports the module, rebuilds `START
HERE`, and saves. Manual alternative: `Alt+F11` > remove the `PM_Report_Macro`
module > **File > Import File...** > `PM_Report_Macro.bas` > save.

---

## Logging and stuck runs

Every build writes one log file to `Logs\` in the project root (next to the master
workbook, not in the report output folder):

```
Logs\PM_Report_20261114_091530.log
```

Every line carries a full date, the time, and the milliseconds since the previous
line:

```
2026-11-14 09:15:33  +0000ms  [stage] fastmode
2026-11-14 09:15:33  +0004ms  [stage] source open
2026-11-14 09:15:33  +0122ms  source sheet 'Results': headerRow=1 lastRow=124 lastCol=12 managerCol=5 statusCol=9
2026-11-14 09:15:36  +0336ms  [stage] raw load
2026-11-14 09:15:36  +0000ms  [stage] data read
...
2026-11-14 09:15:40  +0000ms  SUMMARY: saved=...\Reports\PM Requests 2026-11-14.xlsx  pmTotal=106  managers=12  newPMs=0  pending=18  skipped=0
2026-11-14 09:15:40  +0000ms  === RUN OK in 4.8s
```

A hang is one line followed by silence: the line *before* the silence is the step
that hung. Errors are logged as `!!! ERROR <number>: <description>   [stage: ...]`.

`/logs` prints the run logs oldest-first (so the newest run is the last block) and
ranks the slowest steps. Ask for one run with `/logs 20261114`.

### The 10-second watchdog

```powershell
.\scripts\run-report.ps1
```

This starts a headless build, follows the log, and enforces a stall limit:

- no new log line for **10 seconds** (or no log file at all) means the run is stuck;
- Excel and the worker are killed, the last 40 log lines are printed with the
  slowest step highlighted, and the script exits `2`;
- a normal run prints `RESULT=OK|...`, the summary, and exits `0`.

Typical causes of a stall, in the order the log shows them: a modal dialog Excel
is waiting on (macro security prompt, file picker, a `MsgBox` left in the code), a
compile error dialog, or a single very large copy (`raw load`, `dashboard`).

---

## Weekly use

1. Open the saved `.xlsm` master workbook.
2. Click **CREATE PM REQUEST REPORT**.
3. Select the current raw-data export.
4. The prior report is picked automatically from the output folder (the newest
   saved report with enough name overlap). It is only read, never modified, and
   the completion message names it.
5. The macro finds the correct worksheet, loads its data into `Raw Data`, and
   builds Dashboard, manager tabs, and Pending Pickup automatically.
6. The report is saved automatically (with the raw data as the last tab).
7. A completion message shows the output path and row counts
   (manager PM requests + Pending Pickup).

---

## Source file requirements

The selected export must contain a worksheet with these exact headers
(case-insensitive, trimmed; the header row can be below row 1 — the macro scans
the first 30 rows of each sheet):

- **Assigned Sales Manager**
- **Treatment Status**

All other columns are carried into the generated report as supplied, in their
original order and with their original values and header formats.

---

## Output specification

The generated workbook contains (tab order):

- **Dashboard** (first tab): white logo zone plus navy title band, a meta strip
  (current file, prior file, creation time), six KPI cards (PMs Requested, Prior
  PMs, Completed, New PMs, Pending Pickups, Prior PMs Still Open), a manager
  table with hyperlinks to each manager tab, and a Sales Rep grid (name + PM
  count, three blocks). Without a prior report the prior-based cards show 0.
- **One tab per assigned Sales Manager**, in order of manager last name
  (handles both `Last, First` and `First Last` name formats).
- **`Pending Pickup`**, then **`Raw Data`** as the last tab.

Each manager tab:

- Excludes Pending Pickup rows.
- Has a banner row, four KPI cards (PMs Requested, Prior PMs, Completed, New PMs)
  and a "PMs by Rep" card above the table; the table header is on row 9 with
  data from row 10.
- Has a styled header row, alternating row banding, hairline borders, an
  AutoFilter, and a frozen header row.
- Is sorted by column A.
- When comparing, rows whose Name is absent from the prior report are light
  green + bold (new PMs).

The Pending Pickup tab:

- Contains every row whose Treatment Status contains "Pending Pickup".
- Shows all rows in red text, sorted by columns A then B.
- When comparing, carries a Prior / Now / New / Cleared card to the right of the
  data.

---

## Settings and file handling

- **Output folder memory:** stored in the very-hidden `PM Settings` sheet
  (cell `B2`, setting `OutputFolder`). Deleting or unhiding that sheet and
  clearing `B2` forces the folder prompt again.
- **Filename:** `PM Requests YYYY-MM-DD.xlsx` in the chosen folder. If the file
  exists, the macro writes `PM Requests YYYY-MM-DD (2).xlsx`, `(3)`, etc., so
  nothing is ever overwritten.
- **Button self-repair:** `SetupReportButton` runs at the start of every report
  build and re-creates the `btnCreatePMReport` form-control button if it is
  missing or broken.
- **Safety:** Excel calculation, screen updating, events, and display alerts are
  saved and restored on success, cancel, or error; the source workbook is
  always closed read-only without saving.
- **Status bar** progress messages show the current step during the build.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| "Could not find a worksheet containing both required headers" | The export lacks `Assigned Sales Manager` or `Treatment Status`, or the headers are below row 30, or have different spelling/spacing. Fix the export. |
| "The selected raw-data file has no data rows" | Header found but no rows below it. Check the export. |
| "The selected raw-data file is the PM Requests master workbook" | You selected the `.xlsm` itself instead of the raw export. |
| Macro/button does nothing | Macros are disabled. See One-time setup. On macOS, form-control button `OnAction` must point to `BuildManagerTabs` (re-run `SetupReportButton`). |
| Wrong output folder | Open the VBA editor and clear `PM Settings!B2`, then run again. |
| Extra columns/blank rows in output | The macro uses the source sheet's last used row/column. Remove stray formatting or trailing junk rows in the export. |

---

## Architecture notes (for maintainers)

All VBA lives in the single standard module **`PM_Report_Macro`**. The
canonical source is `PM_Report_Macro.bas` in the project root; the embedded copy
inside the `.xlsm` must be kept in sync with `scripts\sync_template.ps1`. Key entry points
and helpers:

| Member | Role |
|---|---|
| `BuildManagerTabs` | Main pipeline: pick source, auto-detect the prior report, load, split, build, save, report. |
| `SetupReportButton` / `RebuildStartSheet` / `GetOrCreateStartSheet` / `AddStartStep` / `PositionStartLogo` | `START HERE` canvas and the rounded-rectangle CTA button (`RebuildStartSheet` is what `scripts\sync_template.ps1` calls). |
| `DetectPriorReport` / `PriorNameOverlap` / `SortPathsNewestFirst` | Weekly mode: finds the previous week's report in the output folder by name overlap. |
| `LoadPriorComparison` / `FindHeaderRowForSingle` / `PriorNameExists` / `NormalizeName` / `CountDiff` / `PriorMgrCount` | Prior-report comparison: read-only open of the prior report, prior name/manager/pending sets. |
| `BuildDashboard` / `WriteTableHeader` / `AddPolishedCard` / `AddPolishedTextCard` / `AddDashboardLogo` / `PositionDashboardLogo` | Dashboard summary tab (meta strip, KPI cards, manager table, rep grid, logo). |
| `FindSourceWorksheet` / `FindHeaderRow` / `FindHeaderColumn` | Locate the source sheet and required header cells (first 30 rows, case-insensitive). |
| `LoadSourceIntoRawData` | Copies the source range into the `Raw Data` sheet of the master. |
| `SortManagersByLastName` / `LastNameKey` | Bubble-sort manager names by last-name key. |
| `BuildOutputArray` / `GetHeaderArray` | Build 2D arrays for fast bulk writes to output tabs. |
| `FormatManagerSheet` / `ApplyReaderFriendlyFormatting` / `CountInSheet` / `AddPolishedCard` / `AddPolishedTextCard` | Output formatting: banner, four uniform KPI cards, full-width rep card, per-employee row banding, new-PM green rows, table header style, filters. |
| `LogOpen` / `LogLine` / `LogStage` / `LogPath` / `SetStatus` | Run log: one timestamped file per run in `Logs\`, one line per step with milliseconds since the previous line. `SetStatus` also logs, so every user-visible step is traced. |
| `GetRememberedOutputFolder` / `SaveOutputFolder` / `GetOrCreateSettingsSheet` | Output-folder memory in the very-hidden `PM Settings` sheet. |
| `GetUniqueOutputPath` | Numbered filenames so reports are never overwritten. |
| `RestoreExcelState` | Restores calculation/events/screen updating/alerts. |

Workbook sheets:

| Sheet | Visibility | Role |
|---|---|---|
| `START HERE` | Visible | Instructions + report button. |
| `Raw Data` | Visible | Working copy of the current source export (rewritten each run). |
| `PM Settings` | Very hidden | Stores the remembered output folder. |

Conventions: `Option Explicit` throughout; constants at the top of the module
(`RAW_SHEET`, `SETTINGS_SHEET`, `DASHBOARD_SHEET`, `MANAGER_HEADER`,
`NAME_HEADER`, `STATUS_HEADER`, `PENDING_TEXT`, `BUTTON_NAME`, `BUTTON_TEXT`,
`SETTING_OUTPUT_FOLDER`) define the file names, headers, and UI text — change
those to rebrand.

## Changelog

See [CHANGELOG.md](CHANGELOG.md). Project version is 1.3; every change is logged
there with a `YYYY-MM-DD HH:mm:ss` timestamp.

## Known issues

- Comparison figures are name-based: a PM whose **Name** cell changed between
  weeks counts as one completed plus one new PM.
- Auto-detection only looks at `PM Requests *.xlsx` in the remembered output
  folder; a report stored elsewhere has to be picked explicitly (automation)
  or moved into that folder.