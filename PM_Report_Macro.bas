Attribute VB_Name = "PM_Report_Macro"
Option Explicit

' ============================================================
' WEEKLY SALES MANAGER REPORT BUILDER - v1.4 (Excel) - SPLIT INTO MODULES
'
' Canonical module. Holds the public surface, shared declarations,
' the run-log infrastructure, and the core RunReport pipeline.
' Supporting modules: PM_Report_UI, PM_Report_Host, PM_Report_Setup,
' PM_Report_Data, PM_Report_Output.
'
' Public surface: BuildManagerTabs, CreateEmailFromReport, EditCCEmails,
'   EditEmailTemplate, GenerateManagerEmail, GenerateReportEmail, LogPath,
'   RebuildStartSheet, RunReport, SetupEmailButtons, SetupGenerateEmailButton,
'   SetupReportButton
' ============================================================
Public Const RAW_SHEET As String = "Raw Data"
Public Const SETTINGS_SHEET As String = "PM Settings"
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const MANAGER_HEADER As String = "Assigned Sales Manager"
Public Const NAME_HEADER As String = "Name"
Public Const STATUS_HEADER As String = "Treatment Status"
Public Const BUTTON_NAME As String = "btnCreatePMReport"
Public Const BUTTON_TEXT As String = "CREATE PM REQUEST REPORT"
Public Const SETTING_OUTPUT_FOLDER As String = "OutputFolder"
Public Const LOGO_SHAPE As String = "AOTI_Logo"
Public Const START_SHEET As String = "START HERE"
Public Const LOGO_FILE As String = "aoti-logo.png"
Public Const LOGO_DIR As String = "documentation"
Public Const SHEET_COL_WIDTH As Double = 16
Public Const UI_FONT As String = "Segoe UI"
Public Const MANAGER_LAST_COL As Long = 8

' Run-log state (must live in the declarations section, not between procedures).
Public gLogPath As String
Public gLogPrev As Double
Public gLogT0 As Double

' Realtime error-logging state (mirrors the Pending Pickup pattern).
' mInReportRun  - true while RunReport is executing, so LogErr also hits the
'                 per-run log. False means only the global debug log is written.
' mRunLogPath   - path to the current per-run log (mirrors gLogPath).
' mDebugLogPath - path to the global debug log (appended across every run).
Public mInReportRun As Boolean
Public mRunLogPath As String
Public mDebugLogPath As String


' ==================== RUN LOG ====================
' Centralized realtime error-logging infrastructure.
' Every procedure that can fail logs via LogErr from its LocalFail block.
Public Function RunReport(ByVal sourcePath As String, ByVal outputFolder As String, Optional ByVal priorPath As String = "") As String
    Dim wb As Workbook, srcWb As Workbook, raw As Worksheet, outWb As Workbook
    Dim headerRow As Long, managerCol As Long, statusCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant, r As Long, i As Long
    Dim manager As String
    Dim mgrNames As Collection, mgrRows As Collection
    Dim managers As Variant, outWs As Worksheet
    Dim rowsForManager As Collection, arrOut As Variant
    Dim firstOutput As Boolean, reportDate As String, savePath As String
    Dim skippedRows As Long, pmTotal As Long, idx As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim priorNames As Collection, priorMgrNames As Collection, priorMgrCounts As Collection
    Dim mgrDataNames As Collection
    Dim priorPMs As Long
    Dim repNames As Collection
    Dim repCounts() As Long, repN As Long
    Dim newPMs As Long, nameCol As Long
    Dim repName As String, nm As String
    Dim stage As String
    Dim auditMissingDate As Long, auditMissingStatus As Long, auditFutureDate As Long
    Dim auditOldDate As Long, auditDuplicateName As Long, auditMissingKey As Long
    Dim seenNames As Collection, curName As String
    Dim curStart As Variant, curStatus As Variant, curDate As Double, oldDateCutoff As Double

    On Error GoTo Fail

    ' Fail fast on a missing destination: everything below reads the source and
    ' builds 12 tabs, and a bad folder argument would only surface at SaveAs.
    stage = "args"
    If Len(Trim$(outputFolder)) = 0 Then
        Err.Raise vbObjectError + 2204, , "No output folder was given."
    End If

    Set wb = ThisWorkbook

    ' One log per run: without the reset, a second build in the same Excel session
    ' appended to the first run's file and reported that run's elapsed time.
    LogReset
    LogOpen outputFolder
    mInReportRun = True
    LogLine "--- RunReport: source=" & sourcePath & "  prior=" & IIf(Len(priorPath) = 0, "(auto)", priorPath)

    oldCalc = -4135: oldScreen = False: oldEvents = False: oldAlerts = False
    On Error Resume Next
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Fail

    FastModeOn
    stage = "fastmode": LogStage "fastmode"
    Set srcWb = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    stage = "source open": LogStage "source open"
    Set raw = FindSourceWorksheet(srcWb, headerRow)
    If raw Is Nothing Then
        Err.Raise vbObjectError + 2202, , "I could not find a worksheet containing both required headers: " & MANAGER_HEADER & " and " & STATUS_HEADER & "."
    End If

    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, STATUS_HEADER)
    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)
    ' Only the first MANAGER_LAST_COL columns belong on manager tabs; the
    ' treatment-detail columns (PM date, end date, setup date, technician)
    ' are source-only and are not carried into the report.
    If lastCol > MANAGER_LAST_COL Then lastCol = MANAGER_LAST_COL
    LogLine "source sheet '" & raw.Name & "': headerRow=" & CStr(headerRow) & " lastRow=" & CStr(lastRow) & _
             " lastCol=" & CStr(lastCol) & " managerCol=" & CStr(managerCol) & " statusCol=" & CStr(statusCol)

    If lastRow <= headerRow Then
        Err.Raise vbObjectError + 2203, , "The selected current report has no data rows."
    End If

    SetStatus "Loading current report into Raw Data..."

    LoadSourceIntoRawData wb, raw, headerRow, lastRow, lastCol
    stage = "raw load": LogStage "raw load"
    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    Set raw = wb.Worksheets(RAW_SHEET)
    headerRow = FindHeaderRow(raw, MANAGER_HEADER, STATUS_HEADER)
    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, STATUS_HEADER)
    nameCol = FindHeaderColumn(raw, headerRow, NAME_HEADER)
    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)
    ' Only the first MANAGER_LAST_COL columns belong on manager tabs; the
    ' treatment-detail columns (PM date, end date, setup date, technician)
    ' are source-only and are not carried into the report.
    If lastCol > MANAGER_LAST_COL Then lastCol = MANAGER_LAST_COL

    SetStatus "Building manager tabs..."

    data = raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Value2
    stage = "data read": LogStage "data read"
    Set mgrNames = New Collection
    Set mgrRows = New Collection
    Set mgrDataNames = New Collection
    Set repNames = New Collection
    ReDim repCounts(1 To 1)
    repN = 0
    stage = "collections set": LogStage "collections set"
    ' Null-safe classification: a row is a PM request when it has a manager.
    ' Blank manager cells are skipped, not fatal (CStr(Null) would raise
    ' error 94). Collection-based to avoid the Scripting.Dictionary COM
    ' dependency. Treatment Status is required in the export and carried into
    ' the report, but it no longer filters anything.
    Set seenNames = New Collection
    oldDateCutoff = CLng(DateSerial(Year(Date) - 5, Month(Date), Day(Date)))   ' 5 years ago
    For r = 2 To UBound(data, 1)
        manager = Trim$(SafeStr(data(r, managerCol)))

        If Len(manager) > 0 Then
            idx = FindManager(mgrNames, manager)
            If idx = 0 Then
                Set rowsForManager = New Collection
                mgrNames.Add manager
                mgrRows.Add rowsForManager
            Else
                Set rowsForManager = mgrRows(idx)
            End If
            rowsForManager.Add r
            If nameCol > 0 Then
                nm = NormalizeName(data(r, nameCol))
                If Len(nm) > 0 Then mgrDataNames.Add nm
            End If
            ' Rep totals (column A = Assigned Sales Representative). Counts live
            ' in a Long array because Collection items cannot be reassigned.
            repName = Trim$(SafeStr(data(r, 1)))
            If Len(repName) > 0 Then
                idx = FindManager(repNames, repName)
                If idx = 0 Then
                    repN = repN + 1
                    ReDim Preserve repCounts(1 To repN)
                    repNames.Add repName
                    repCounts(repN) = 1
                Else
                    repCounts(idx) = repCounts(idx) + 1
                End If
            End If

            ' ---- audit: missing pickup date (col 8) / future / very old ----
            curStart = data(r, 8)
            If IsEmpty(curStart) Then
                auditMissingDate = auditMissingDate + 1
            ElseIf IsDate(curStart) Then
                curDate = CDbl(CDate(curStart))
                If curDate > CDbl(Date) Then
                    auditFutureDate = auditFutureDate + 1
                ElseIf curDate < oldDateCutoff Then
                    auditOldDate = auditOldDate + 1
                End If
            ElseIf IsNumeric(curStart) Then
                ' Value2 returns dates as Double serials; IsDate misses them.
                curDate = CDbl(curStart)
                If curDate > CDbl(Date) Then
                    auditFutureDate = auditFutureDate + 1
                ElseIf curDate < oldDateCutoff Then
                    auditOldDate = auditOldDate + 1
                End If
            Else
                auditMissingDate = auditMissingDate + 1
            End If

            ' ---- audit: missing status (col 7) ----
            curStatus = Trim$(SafeStr(data(r, 7)))
            If Len(curStatus) = 0 Then auditMissingStatus = auditMissingStatus + 1

            ' ---- audit: duplicate name (col 3) ----
            If nameCol > 0 Then
                curName = NormalizeName(data(r, nameCol))
                If Len(curName) > 0 Then
                    On Error Resume Next
                    seenNames.Add curName, curName
                    If Err.Number <> 0 Then
                        auditDuplicateName = auditDuplicateName + 1
                        Err.Clear
                    End If
                    On Error GoTo Fail
                End If
            End If
        Else
            skippedRows = skippedRows + 1
        End If

        ' ---- audit: missing key (col 1 rep name blank) ----
        If Len(Trim$(SafeStr(data(r, 1)))) = 0 Then auditMissingKey = auditMissingKey + 1
    Next r
    stage = "classified": LogStage "classified"
    LogLine "classified: managers=" & CStr(mgrNames.Count) & _
            "  skipped(no manager)=" & CStr(skippedRows) & "  reps=" & CStr(repN) & _
            "  missingDate=" & CStr(auditMissingDate) & "  missingStatus=" & CStr(auditMissingStatus) & _
            "  futureDate=" & CStr(auditFutureDate) & "  oldDate=" & CStr(auditOldDate) & _
            "  duplicateName=" & CStr(auditDuplicateName) & "  missingKey=" & CStr(auditMissingKey)
    ' Weekly comparison: with no explicit prior, use the previous week's
    ' report found in the output folder (name overlap). An unreadable prior
    ' degrades to empty comparison data instead of failing the build.
    If Len(priorPath) = 0 Then
        SetStatus "Looking for the prior week's report..."
        stage = "detect prior": LogStage "detect prior"
        priorPath = DetectPriorReport(outputFolder, data, nameCol, lastRow)
    End If

    If Len(priorPath) > 0 Then
        SetStatus "Comparing to the prior report..."
        LoadPriorComparison priorPath, priorNames, priorMgrNames, priorMgrCounts
    Else
        Set priorNames = New Collection
        Set priorMgrNames = New Collection
        Set priorMgrCounts = New Collection
    End If
    priorPMs = 0
    For i = 1 To priorMgrCounts.Count: priorPMs = priorPMs + CLng(priorMgrCounts(i)): Next i

    ' New PMs: names assigned to a manager this week that were not on the prior
    ' report's manager tabs. Computed before the log line that reports it - the
    ' old order printed newPMs=0 on every run, which sent the reader to the
    ' dashboard for the one number the log claimed to have.
    If priorNames.Count > 0 Then
        For i = 1 To mgrDataNames.Count
            If Not PriorNameExists(priorNames, CStr(mgrDataNames(i))) Then newPMs = newPMs + 1
        Next i
    End If

    LogLine "prior report: " & IIf(Len(priorPath) = 0, "(none)", priorPath) & _
            "  priorPMs=" & CStr(priorPMs) & "  priorNames=" & CStr(priorNames.Count) & "  newPMs=" & CStr(newPMs)

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    ' Remove the default blank "Sheet1" so the report ships with only Dashboard,
    ' manager tabs, and Raw Data - no orphan tab.
    On Error Resume Next
    outWb.Worksheets("Sheet1").Delete
    On Error GoTo Fail
    firstOutput = True
    stage = "out workbook": LogStage "out workbook"
    If mgrNames.Count > 0 Then
        ReDim managers(1 To mgrNames.Count)
        For i = 1 To mgrNames.Count: managers(i) = mgrNames(i): Next i
        SortManagersByLastName managers
        For i = LBound(managers) To UBound(managers)
            manager = CStr(managers(i))
            SetStatus "Creating manager tab: " & manager
            Set rowsForManager = mgrRows(FindManager(mgrNames, manager))
            LogLine "manager tab " & CStr(i) & "/" & CStr(UBound(managers)) & ": " & manager & _
                    " (" & CStr(rowsForManager.Count) & " rows)"
            Set outWs = GetOrCreateOutputSheet(outWb, SafeSheetName(manager), firstOutput)
            firstOutput = False
            arrOut = BuildOutputArray(data, rowsForManager, lastCol)
            ' Table header on row 9, data from row 10 (KPI cards above).
            outWs.Range(outWs.Cells(9, 1), outWs.Cells(8 + UBound(arrOut, 1), lastCol)).Value = arrOut
            For idx = 1 To rowsForManager.Count
                CopyRowFormats raw, outWs, CLng(rowsForManager(idx)), idx + 9, lastCol
            Next idx
            FormatManagerSheet outWs, manager, priorNames, priorMgrNames, priorMgrCounts, nameCol
            LogLine "manager tab done: " & manager
        Next i
    End If
    stage = "manager tabs": LogStage "manager tabs"
    ' PM totals: sum of rows actually written to manager tabs.
    pmTotal = 0
    For i = 1 To mgrRows.Count: pmTotal = pmTotal + mgrRows(i).Count: Next i

    reportDate = Format(Date, "yyyy-mm-dd")
    savePath = GetUniqueOutputPath(outputFolder, "PM Requests " & reportDate & ".xlsx")

    SetStatus "Building dashboard..."
    stage = "dashboard": LogStage "dashboard"
    BuildDashboard outWb, managers, mgrRows, mgrNames, data, nameCol, _
                  pmTotal, newPMs, _
                  priorNames, priorMgrNames, priorMgrCounts, priorPMs, _
                  repNames, repCounts, repN, _
                  Mid$(savePath, InStrRev(savePath, Application.PathSeparator) + 1), priorPath
    outWb.Worksheets(DASHBOARD_SHEET).Move Before:=outWb.Worksheets(1)
    stage = "dashboard moved": LogStage "dashboard moved"
    ' Copy the loaded Raw Data sheet into the report (last tab).
    raw.Copy After:=outWb.Worksheets(outWb.Worksheets.Count)
    stage = "raw copied": LogStage "raw copied"

    ' Block out unused cells on Raw Data so it shows only the copied data,
    ' nothing else.
    Dim rws As Worksheet
    Set rws = outWb.Worksheets(outWb.Worksheets.Count)
    If StrComp(rws.Name, RAW_SHEET, vbTextCompare) = 0 Then
        Dim rLastRow As Long, rLastCol As Long
        rLastRow = LastUsedRow(rws): rLastCol = LastUsedColumn(rws)
        HideUnusedCells rws, rLastRow, rLastCol
    End If
    ColorReportTabs outWb
    stage = "tabs colored": LogStage "tabs colored"
    ' Open the report the way the user wants to see it: maximized, on the
    ' Dashboard, with A1 selected.
    With outWb.Worksheets(DASHBOARD_SHEET)
        .Activate
        .Range("A1").Select
    End With
    On Error Resume Next
    ActiveWindow.WindowState = xlMaximized
    ActiveWindow.Zoom = 90
    On Error GoTo 0
    stage = "view set": LogStage "view set"
    SetStatus "Saving finished report..."

    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    stage = "saved": LogStage "saved"
    outWb.Close SaveChanges:=False
    Set outWb = Nothing

    RunReport = "OK|" & savePath & "|" & CStr(pmTotal) & "|" & CStr(skippedRows) & "|" & _
                IIf(Len(priorPath) = 0, "", Mid$(priorPath, InStrRev(priorPath, Application.PathSeparator) + 1))

    LogLine "SUMMARY: saved=" & savePath & "  pmTotal=" & CStr(pmTotal) & _
            "  managers=" & CStr(mgrNames.Count) & "  newPMs=" & CStr(newPMs) & _
            "  skipped=" & CStr(skippedRows)
    LogLine "=== RUN OK in " & Format$(Timer - gLogT0, "0.0") & "s"

    SetStatus ""
    mInReportRun = False
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    Exit Function

Fail:
    ' Capture the error before any cleanup statement can touch Err.
    Dim errNo As Long, errDesc As String, errStage As String
    errNo = Err.Number: errDesc = Err.Description: errStage = stage

    LogErr "RunReport"
    LogLine "!!! ERROR " & CStr(errNo) & ": " & errDesc & "   [stage: " & errStage & "]"
    LogLine "    last breadcrumb above is where it stopped."

    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    SetStatus ""
    mInReportRun = False
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    On Error GoTo 0
    RunReport = "ERROR|" & CStr(errNo) & "|" & errDesc & " @ " & errStage
End Function


' ====================== HOST / DIALOG HELPERS ======================
' Owns everything that needs a user: file picker, folder picker, status bar.
'
'   PickSourceFile    - GetOpenFilename; returns False when cancelled.
'   PickOutputFolder  - FileDialog(4); returns "" when cancelled.
'   SetStatus         - status bar text + a log line; also used with "" to clear.
'
' Contracts:
'   * Cancel is signalled by False / "" and the caller decides what to do; these
'     helpers never raise on cancel.
'   * SetStatus logs what it shows, so the status bar text and the log can never
'     disagree.
'

' Source file: Excel open dialog (False = user cancelled).

Public Sub SetStatus(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = text
    On Error GoTo 0
    LogLine text
End Sub


' ====================== RUN LOG ======================
' Owns the run log, which is the only evidence a headless run leaves behind.
' Every run writes <workbook folder>\Logs\PM_Report_yyyyMMdd_hhmmss.log, i.e. the
' project root next to the master workbook. Each line carries the wall time and
' the milliseconds since the previous line, so a hang shows up as one line
' followed by silence - and the line before the silence names the step that hung.
'
'   LogOpen    - opens today's log beside the workbook (never in the report
'                folder, so /logs always knows where to look).
'   LogReset   - forgets the current log so the next LogOpen starts a new file.
'   LogLine    - "yyyy-MM-dd hh:mm:ss  +NNNNms  msg", appended, also Debug.Print.
'   LogStage   - "[stage] label" breadcrumb used by RunReport's error handler.
'   LogPath    - Public; the one line the UI and /logs use to name the file.
'   FastModeOn - screen updating/events/alerts off; paired with RestoreExcelState.
'
' Contracts:
'   * The log line format is parsed by .kilo\command\logs.md; changing it breaks
'     /logs.
'   * gLogPath / gLogPrev / gLogT0 are module-level state and must stay in the
'     declarations section: a Private declaration between two procedures is not
'     visible to the rest of the module.
'   * A log file belongs to exactly one run. LogReset is the run boundary; the
'     file name carries the run's start time, not the Excel session's.
'   * Timer rolls over at midnight; LogLine adds 86400 to keep deltas positive.
'


Public Sub LogReset()
    gLogPath = ""
    gLogPrev = 0
    gLogT0 = 0
End Sub



Public Sub LogOpen(ByVal outputFolder As String)
    Dim fld As String, logDir As String

    If Len(gLogPath) > 0 Then Exit Sub

    On Error Resume Next
    gLogT0 = Timer
    gLogPrev = gLogT0

    ' Logs live beside the master workbook, never in the report output folder:
    ' the /logs command always knows where to look.
    fld = ThisWorkbook.Path
    If Len(fld) = 0 Then fld = Trim$(outputFolder)
    If Len(fld) = 0 Then Exit Sub
    If Right$(fld, 1) = Application.PathSeparator Then fld = Left$(fld, Len(fld) - 1)

    logDir = fld & Application.PathSeparator & "Logs"
    If Not FolderExists(logDir) Then MkDir logDir

    gLogPath = logDir & Application.PathSeparator & "PM_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".log"
    LogLine "=== RUN START " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  (PM Report Macro)"
    LogLine "workbook: " & ThisWorkbook.Name
    LogLine "workbook folder: " & fld
    LogLine "report output folder: " & IIf(Len(Trim$(outputFolder)) = 0, "(not given)", Trim$(outputFolder))
    LogLine "excel: " & CStr(Application.Version) & "  user: " & CStr(Environ$("USERNAME"))
    On Error GoTo 0
End Sub



Public Sub LogLine(ByVal msg As String)
    Dim ff As Integer, nowT As Double, since As Long

    If Len(gLogPath) = 0 Then Exit Sub

    On Error Resume Next
    nowT = Timer
    If nowT < gLogPrev Then nowT = nowT + 86400   ' midnight rollover
    since = CLng((nowT - gLogPrev) * 1000)
    If since < 0 Then since = 0

    ff = FreeFile
    Open gLogPath For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:mm:ss") & "  +" & Format$(since, "0000") & "ms  " & msg
    Close #ff
    gLogPrev = nowT
    On Error GoTo 0

    Debug.Print Format$(Now, "hh:mm:ss") & " +" & since & "ms  " & msg
End Sub


' Fine-grained breadcrumb: the `stage` label from RunReport's error handler.

Public Sub LogStage(ByVal label As String)
    LogLine "[stage] " & label
End Sub


' One line naming the current log file, for the UI and the /logs command.

Public Function LogPath() As String
    LogPath = gLogPath
End Function


' ---- realtime error logging (one helper, used by every LocalFail handler) ----
' Writes one line to BOTH logs:
'   * macro-debug.log  - global, appended across every run (never rotated)
'   * per-run log      - only when mInReportRun = True
' Each line: wall-time | procedure | Err.Number | Err.Description | Erl (line no)

Public Sub LogErr(ByVal procName As String)
    Dim errNo As Long, errDesc As String, erl As Long
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    erl = Erl
    On Error GoTo 0

    Dim ts As String, line As String
    ts = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    line = ts & " | ERR | " & procName & " | #" & CStr(errNo) & " | " & errDesc & " | line " & CStr(erl)

    ' Global debug log (always) - fixed path, appended across every run.
    If Len(mDebugLogPath) = 0 Then
        If Len(gLogPath) > 0 Then
            mDebugLogPath = Left$(gLogPath, InStrRev(gLogPath, Application.PathSeparator)) & "macro-debug.log"
        End If
    End If
    If Len(mDebugLogPath) > 0 Then
        DebugWriteLine mDebugLogPath, line
    End If

    ' Per-run log (only during a report run).
    If mInReportRun And Len(gLogPath) > 0 Then
        DebugWriteLine gLogPath, line
    End If
End Sub

' Append one line to a log file (opens/closes each call so concurrent runs don't
' collide, and so a partial write never leaves the file locked).

Public Sub DebugWriteLine(ByVal path As String, ByVal line As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open path For Append As #ff
    Print #ff, line
    Close #ff
    On Error GoTo 0
End Sub



Public Sub FastModeOn()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    On Error GoTo 0
End Sub


' ====================== WORKBOOK / SHEET SETUP ======================
' Owns the sheets of the master workbook: the START HERE canvas, the Raw Data
' sheet the export is copied into, and the hidden PM Settings sheet.
'
'   GetOrCreateStartSheet - rebuilds the canvas idempotently: Cells.Clear +
'                           UnMerge, then a fixed grid B2:F21.
'   AddStartStep          - numbered step row (navy number gutter in B, text
'                           merged C:F, alternating fill).
'   AddStartBullet        - bullet row (ChrW(8226) + two spaces + text) merged C:F.
'   PositionStartLogo     - keeps the embedded AOTI_Logo picture top-right.
'   FindSourceWorksheet   - the export sheet that holds both required headers.
'   LoadSourceIntoRawData - copies the export into Raw Data (values + formats).
'   GetOrCreateRawSheet   - Raw Data sheet in the master.
'   GetOrCreateSettingsSheet - hidden PM Settings sheet (may be VeryHidden).
'   GetRememberedOutputFolder / SaveOutputFolder - the OutputFolder setting.
'
' Contracts:
'   * The canvas geometry is fixed and rebuilt every run: columns A-F are 2.6 /
'     5 / 28 / 28 / 28 / 2, rows 1-21 have explicit heights, and the CTA band is
'     B14:F14. SetupReportButton positions the button from that range, so moving
'     a row silently moves the button off the band.
'   * Cells.Clear does not remove shapes. Anything that adds a shape must delete
'     the old one by name first (SetupReportButton, PositionStartLogo) or a
'     second run leaves two of them.
'   * The footer is the only version string a user sees: v1.4 + output name.
'   * Raw Data is a copy, not a link: after LoadSourceIntoRawData the source
'     workbook is closed and every bound is recomputed from the copy.
'

' START HERE canvas, rebuilt idempotently on every run (shapes survive
' Cells.Clear). Canvas design: narrow gutters, navy label bars, numbered step
' rows, a tall CTA row, and a footer.

