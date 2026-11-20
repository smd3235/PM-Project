# VBA Error Handling & Realtime Audit Logging System

A copy-pasteable error-handling and audit-logging framework for Excel/VBA desktop projects.
Originally built for the **PM Requests Report Builder** (`PM_Report_Macro.bas`) and adapted so
another project can adopt the same patterns.

> **Philosophy:** every error must leave a trace in a log file, not a modal dialog.
> Every run leaves a timestamped breadcrumb trail so a hang or crash can be diagnosed
> from the last line written, not from a screenshot.

______________________________________________________________________

## Contents

- [1. Design principles](#1-design-principles)
- [2. Module-level state](#2-module-level-state)
- [3. The `LocalFail` error-handling pattern](#3-the-localfail-error-handling-pattern)
- [4. The `LogErr` + `DebugWriteLine` helpers](#4-the-logerr--debugwriteline-helpers)
- [5. Wiring into your entry point](#5-wiring-into-your-entry-point)
- [6. Realtime stage tracking & breadcrumbs](#6-realtime-stage-tracking--breadcrumbs)
- [7. Audit counters (data-quality checks)](#7-audit-counters-data-quality-checks)
- [8. Compile gate enhancements](#8-compile-gate-enhancements)
- [9. Example log output](#9-example-log-output)
- [10. Copy-paste starter kit](#10-copy-paste-starter-kit)
- [11. Cross-project adaptation checklist](#11-cross-project-adaptation-checklist)

______________________________________________________________________

## 1. Design principles

| Principle                           | How it's implemented                                                                                                                             |
| ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| **No silent failures**              | Every procedure that touches the Excel object model has `On Error GoTo LocalFail` and calls `LogErr` before exiting on error.                    |
| **Errors logged to file**           | `LogErr` writes \`timestamp                                                                                                                      |
| **No modal dialogs in automation**  | Error handlers log instead of `MsgBox` (user-facing buttons are the exception, not the rule).                                                    |
| **Realtime breadcrumbs**            | Stage labels (`stage = "dashboard": LogStage "dashboard"`) plus per-item progress lines (`manager tab 9/12: Derek Owens (10 rows)`).             |
| **Whole-project compile gate**      | VBA compiles one procedure at a time; `test-compile.py` triggers real `Debug > Compile VBAProject` so uncalled procedures with typos are caught. |
| **Persistent compile-error record** | Failed gates write to `Logs\compile-errors.log` with module, line, token, and line text.                                                         |

______________________________________________________________________

## 2. Module-level state

Add these to the declarations section at the top of your standard module (with `Option Explicit`):

```vba
' Realtime error-logging state.
'
' mInReportRun    - true while your main build routine executes, so LogErr also
'                   writes to the per-run log. False means only the global debug log.
' mRunLogPath     - path to the current per-run log (mirrors your gLogPath).
' mDebugLogPath   - path to the global debug log (appended across every run).
Private mInReportRun As Boolean
Private mRunLogPath As String
Private mDebugLogPath As String
```

______________________________________________________________________

## 3. The `LocalFail` error-handling pattern

Apply this to **every** procedure that touches Worksheets, Workbooks, Ranges, Shapes, Charts, or the FileSystem.

### Anatomy

```vba
Private Sub FormatManagerSheet(ws As Worksheet, ByVal manager As String)
    Dim lastCol As Long, lastRow As Long

    On Error GoTo LocalFail          ' <-- 1. trap at the top

    ' ... procedure body ...
    lastCol = LastUsedColumn(ws)
    lastRow = LastUsedRow(ws)
    ' ...

    Exit Sub                         ' <-- 2. normal exit BEFORE the label
LocalFail:
    LogErr "FormatManagerSheet"      ' <-- 3. log and stop
End Sub
```

### Rules

| Rule                                                                | Reason                                                                                                                                  |
| ------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `Exit Sub` / `Exit Function` **before** `LocalFail:`                | The label only runs on errors; normal execution never falls into it.                                                                    |
| Call `LogErr "<ProcedureName>"` in `LocalFail`                      | Captures procedure name, `Err.Number`, `Err.Description`, and `Erl` (line number).                                                      |
| Do **not** route expected "not found" lookups to `LocalFail`        | `SheetExists`, `GetRememberedOutputFolder`, etc. keep `On Error Resume Next` around the lookup — that error is normal, not exceptional. |
| Capture `Err.Number`/`Err.Description`/`Erl` **before** any cleanup | Cleanup statements (`Close`, `Quit`, `Set ws = Nothing`) can clobber `Err`.                                                             |

### Which procedures get it (PM Project example)

```
Entry points          : RunReport, SetupReportButton
Workbook/sheet setup  : GetOrCreateStartSheet, AddStartStep, AddStartBullet,
                        PositionStartLogo, FindSourceWorksheet,
                        LoadSourceIntoRawData, GetOrCreateRawSheet,
                        GetOrCreateSettingsSheet, SaveOutputFolder
Path/state helpers    : FolderExists, NormalizeFolder, GetUniqueOutputPath,
                        RestoreExcelState
Data/format helpers   : BuildOutputArray, GetHeaderArray, CopyRowFormats,
                        SortManagersByLastName, FindHeaderRow, FindHeaderColumn,
                        LastUsedRow, LastUsedColumn, FindManager,
                        GetOrCreateOutputSheet, UniqueSheetName
Prior comparison      : DetectPriorReport, SortPathsNewestFirst
Dashboard             : BuildDashboard, AddManagerChart, WriteTableHeader,
                        AddPolishedCard, AddPolishedTextCard,
                        AddDashboardLogo, PositionDashboardLogo,
                        SortByColumnA
Manager formatting    : FormatManagerSheet, CountInSheet,
                        ApplyReaderFriendlyFormatting, ColorReportTabs
```

______________________________________________________________________

## 4. The `LogErr` + `DebugWriteLine` helpers

Copy these two procedures into your module:

```vba
' Write one line to BOTH logs:
'   * macro-debug.log   (global, appended across every run - never rotated)
'   * per-run log       (only when mInReportRun = True)
' Each line: wall-time | ERR | procedure | #Err.Number | Err.Description | line Erl
Private Sub LogErr(ByVal procName As String)
    Dim errNo As Long, errDesc As String, erl As Long
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    erl = Erl
    On Error GoTo 0

    Dim ts As String, line As String
    ts = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    line = ts & " | ERR | " & procName & " | #" & CStr(errNo) & " | " & errDesc & " | line " & CStr(erl)

    ' Global debug log (always)
    If Len(mDebugLogPath) = 0 Then
        If Len(gLogPath) > 0 Then
            mDebugLogPath = Left$(gLogPath, InStrRev(gLogPath, Application.PathSeparator)) & "macro-debug.log"
        End If
    End If
    If Len(mDebugLogPath) > 0 Then
        DebugWriteLine mDebugLogPath, line
    End If

    ' Per-run log (only during a report run)
    If mInReportRun And Len(gLogPath) > 0 Then
        DebugWriteLine gLogPath, line
    End If
End Sub

' Append one line to a log file. Opens/closes per call so concurrent runs don't
' collide and a partial write never leaves the file locked.
Private Sub DebugWriteLine(ByVal path As String, ByVal line As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open path For Append As #ff
    Print #ff, line
    Close #ff
    On Error GoTo 0
End Sub
```

### Key design points

- **Global log path is fixed** (`macro-debug.log` next to the per-run log), so errors from every run accumulate in one place.
- **Per-run logging is gated** by `mInReportRun` — set `True` at the start of your build, `False` on both success and error paths. Outside a build, errors only hit the global log.
- `DebugWriteLine` is defensive: `On Error Resume Next` around file I/O so a logging failure can never mask the original error.

______________________________________________________________________

## 5. Wiring into your entry point

In your main build routine (e.g. `RunReport`), set the flag at the start and reset it on every exit path:

```vba
Public Function RunReport(ByVal sourcePath As String, _
                          ByVal outputFolder As String, _
                          Optional ByVal priorPath As String = "") As String
    Dim stage As String

    LogReset
    LogOpen outputFolder
    mInReportRun = True                      ' <-- enable per-run error logging
    LogLine "--- RunReport: source=" & sourcePath & "  prior=" & _
            IIf(Len(priorPath) = 0, "(auto)", priorPath)

    ' ... state capture, FastModeOn, open source, load data ...

    On Error GoTo Fail
    ' ... build manager tabs, dashboard, save ...

    LogLine "=== RUN OK in " & Format$(Timer - gLogT0, "0.0") & "s"
    SetStatus ""
    mInReportRun = False                     ' <-- reset on success
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    RunReport = "OK|" & savePath & "|" & CStr(pmTotal) & "|" & CStr(skippedRows) & "|" & priorPath
    Exit Function

Fail:
    ' Capture the error BEFORE any cleanup statement can touch Err.
    Dim errNo As Long, errDesc As String, errStage As String
    errNo = Err.Number: errDesc = Err.Description: errStage = stage

    LogErr "RunReport"                       ' <-- log with proc name + line no
    LogLine "!!! ERROR " & CStr(errNo) & ": " & errDesc & "   [stage: " & errStage & "]"
    LogLine "    last breadcrumb above is where it stopped."

    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    SetStatus ""
    mInReportRun = False                     ' <-- reset on error too
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    On Error GoTo 0
    RunReport = "ERROR|" & CStr(errNo) & "|" & errDesc & " @ " & errStage
End Function
```

______________________________________________________________________

## 6. Realtime stage tracking & breadcrumbs

Use a module-level `stage` string plus a tiny helper:

```vba
Private stage As String

Private Sub LogStage(ByVal label As String)
    LogLine "[stage] " & label
End Sub
```

Then label every major transition:

```vba
stage = "fastmode": LogStage "fastmode"
stage = "source open": LogStage "source open"
stage = "data read": LogStage "data read"
stage = "classified": LogStage "classified"
stage = "manager tabs": LogStage "manager tabs"
stage = "dashboard": LogStage "dashboard"
stage = "saved": LogStage "saved"
```

And log per-item progress inside loops so a hang is visible line-by-line:

```vba
For i = LBound(managers) To UBound(managers)
    manager = CStr(managers(i))
    SetStatus "Creating manager tab: " & manager
    LogLine "manager tab " & CStr(i) & "/" & CStr(UBound(managers)) & ": " & manager & _
            " (" & CStr(rowsForManager.Count) & " rows)"
    ' ... build tab ...
    LogLine "manager tab done: " & manager
Next i
```

**Why this matters:** every log line is `yyyy-MM-dd hh:mm:ss  +<ms>ms  message`. A hang is one line followed by silence — and the line *before* the silence is the step that hung.

______________________________________________________________________

## 7. Audit counters (data-quality checks)

During classification, count anomalies and emit one summary line at the end:

```vba
Dim auditMissingDate As Long, auditMissingStatus As Long
Dim auditFutureDate As Long, auditOldDate As Long
Dim auditDuplicateName As Long, auditMissingKey As Long

' ... inside the row loop:
If IsEmpty(curStart) Then auditMissingDate = auditMissingDate + 1
If Len(curStatus) = 0 Then auditMissingStatus = auditMissingStatus + 1
If curDate > CDbl(Date) Then auditFutureDate = auditFutureDate + 1
If curDate < oldDateCutoff Then auditOldDate = auditOldDate + 1
If seenNames contains curName Then auditDuplicateName = auditDuplicateName + 1
If Len(repName) = 0 Then auditMissingKey = auditMissingKey + 1

' ... one summary line:
LogLine "classified: managers=" & CStr(mgrNames.Count) & _
        "  skipped(no manager)=" & CStr(skippedRows) & _
        "  missingDate=" & CStr(auditMissingDate) & _
        "  missingStatus=" & CStr(auditMissingStatus) & _
        "  futureDate=" & CStr(auditFutureDate) & _
        "  oldDate=" & CStr(auditOldDate) & _
        "  duplicateName=" & CStr(auditDuplicateName) & _
        "  missingKey=" & CStr(auditMissingKey)
```

All zeros means clean data; anything above zero is a data-quality issue to review.

______________________________________________________________________

## 8. Compile gate enhancements

### Why VBA needs a real compile gate

Excel's compile-on-demand means **running a procedure does not compile its neighbours**. A typo in a procedure nobody calls survives every `Application.Run`. Compile errors surface as modal dialogs, never as exceptions — so a headless harness sees a hang, not a message.

### `scripts\test-compile.py`

The gate:

1. Copies the master `.xlsm` to `Logs\_compile_test.xlsm` (master never touched).
1. Starts a dialog watcher (`watch-compile-error.py`) that records and dismisses every modal dialog to `Logs\watcher.log`.
1. Imports your `.bas` into the copy (tests the source, not the embedded module).
1. Triggers real **Debug > Compile VBAProject** (VBE command bar, Debug menu, control id `578`).
1. Polls the watcher log for compile-error dialog text.
1. On failure, asks the VBE for the exact offending token: `module=<name> line=<n> token=[<text>] lineText=<full line>`.
1. Runs an execution probe to prove the module still runs, not just parses.
1. Writes verdict: `COMPILE OK`, `COMPILE FAILED`, `PROBE FAILED`, or `INCONCLUSIVE`.

### Persistent compile-error log

Add this to the failure branch so every failed gate leaves a durable record:

```powershell
if ($compileError) {
    $errLog = Join-Path $logDir "compile-errors.log"
    $errLine = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | COMPILE FAILED - $compileError"
    if ($where) { $errLine += " | $where" }
    Add-Content -LiteralPath $errLog -Value $errLine

    Log "RESULT: COMPILE FAILED - $compileError"
    if ($where) { Log "         $where" }
    Log "Full error details written to Logs\compile-errors.log"
}
```

### Self-test mode

`test-compile.py -SelfTest` injects a deliberately broken procedure, asserts the gate **fails**, verifies the fixture is really in the module, removes it, and asserts the gate passes — proving the detector actually detects.

______________________________________________________________________

## 9. Example log output

### Per-run log (`Logs\PM_Report_YYYYMMDD_HHMMSS.log`)

```
2026-11-15 19:02:35  +0000ms  manager tab 9/12: Derek Owens (10 rows)
2026-11-15 19:02:35  +0312ms  manager tab done: Derek Owens
2026-11-15 19:02:35  +0000ms  [stage] manager tabs
2026-11-15 19:02:36  +0000ms  Building dashboard...
2026-11-15 19:02:36  +0000ms  [stage] dashboard
2026-11-15 19:02:37 | ERR | AddDashboardLogo | #-2147024809 | The item with the specified name wasn't found. | line 0
2026-11-15 19:02:37  +0289ms  [stage] dashboard moved
2026-11-15 19:02:37  +0008ms  [stage] raw copied
2026-11-15 19:02:37  +0000ms  [stage] tabs colored
2026-11-15 19:02:37  +0000ms  [stage] view set
2026-11-15 19:02:37  +0047ms  [stage] saved
2026-11-15 19:02:37  +0016ms  SUMMARY: saved=...\Reports\PM Requests 2026-11-15 (9).xlsx pmTotal=124 managers=12 newPMs=0 skipped=0
2026-11-15 19:02:37  +0000ms  === RUN OK in 4.9s
```

Notice the `AddDashboardLogo` error: it was caught and logged, but the build continued because the logo is non-fatal. Without `LogErr`, this would have been a silent failure or a modal dialog.

### Global debug log (`Logs\macro-debug.log`)

```
2026-11-15 19:02:32 | ERR | PriorNameOverlap | #0 |  | line 0
2026-11-15 19:02:37 | ERR | AddDashboardLogo | #-2147024809 | The item with the specified name wasn't found. | line 0
```

Accumulates across every run — one place to check for historical errors.

### Compile error log (`Logs\compile-errors.log`)

```
2026-11-15 18:59:15 | COMPILE FAILED - the project does not compile | module=PM_Report_Macro line=123 token=[UndefinedName] lineText=x = UndefinedName
```

______________________________________________________________________

## 10. Copy-paste starter kit

Add these to your module, then apply the `LocalFail` pattern procedure-by-procedure.

```vba
' ================= ERROR LOGGING STATE =================
Private mInReportRun As Boolean
Private mRunLogPath As String
Private mDebugLogPath As String
Private stage As String

' ================= ERROR LOGGING HELPERS =================
Private Sub LogErr(ByVal procName As String)
    Dim errNo As Long, errDesc As String, erl As Long
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    erl = Erl
    On Error GoTo 0

    Dim ts As String, line As String
    ts = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    line = ts & " | ERR | " & procName & " | #" & CStr(errNo) & " | " & errDesc & " | line " & CStr(erl)

    If Len(mDebugLogPath) = 0 Then
        If Len(gLogPath) > 0 Then
            mDebugLogPath = Left$(gLogPath, InStrRev(gLogPath, Application.PathSeparator)) & "macro-debug.log"
        End If
    End If
    If Len(mDebugLogPath) > 0 Then
        DebugWriteLine mDebugLogPath, line
    End If

    If mInReportRun And Len(gLogPath) > 0 Then
        DebugWriteLine gLogPath, line
    End If
End Sub

Private Sub DebugWriteLine(ByVal path As String, ByVal line As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open path For Append As #ff
    Print #ff, line
    Close #ff
    On Error GoTo 0
End Sub

Private Sub LogStage(ByVal label As String)
    LogLine "[stage] " & label
End Sub
```

Then in each procedure:

```vba
Private Sub YourProcedure(ws As Worksheet)
    On Error GoTo LocalFail
    ' ... body ...
    Exit Sub
LocalFail:
    LogErr "YourProcedure"
End Sub
```

______________________________________________________________________

## 11. Cross-project adaptation checklist

- [ ] Add `mInReportRun`, `mRunLogPath`, `mDebugLogPath` to module declarations.
- [ ] Add `LogErr`, `DebugWriteLine`, `LogStage` helpers.
- [ ] Set `mInReportRun = True` at start of your build; `False` on success **and** error exit paths.
- [ ] Convert every Excel-object-touching procedure to `On Error GoTo LocalFail` + `Exit Sub` + `LocalFail: LogErr "<name>"`.
- [ ] Keep `On Error Resume Next` only around intentional "not found" lookups (`SheetExists`, etc.).
- [ ] Add `stage = "name": LogStage "name"` at major transitions.
- [ ] Add `LogLine` per-item progress inside long loops (per-sheet, per-manager, per-row batch).
- [ ] Add audit counters for data-quality issues that matter to your domain.
- [ ] Copy `scripts\test-compile.py` and adapt `$Bas`, `$ModuleName`, `$CompileCommandId`.
- [ ] Add the `compile-errors.log` write block to the gate's failure branch.
- [ ] Run `test-compile.py -SelfTest` to prove the detector detects.
- [ ] Run your build end-to-end and verify no unexpected `ERR` lines appear.

______________________________________________________________________

*Adapted from the PM Requests Report Builder (`PM_Report_Macro.bas`). The same two files exist in the peer project, so this guide runs correctly from either side.*
