Attribute VB_Name = "PM_Report_Macro"
Option Explicit

' ============================================================
' WEEKLY SALES MANAGER REPORT BUILDER - v1.4 (Excel)
'
' Host: Excel (Windows/macOS). The workbook is the report
' template: the START HERE button runs the weekly build.
'
' Structure:
'   BuildManagerTabs   - interactive shell (dialogs + message).
'   RunReport          - core pipeline; returns a result string:
'                        "OK|savePath|pmTotal|skipped|priorFileName"
'                        "ERROR|number|description @ stage"
'                        Public so the shell and tests can call it.
'                        priorFileName is "" when nothing was compared.
'
' Public surface (everything else in this module is Private):
'   SetupReportButton, RebuildStartSheet, BuildManagerTabs, RunReport, LogPath.
'   The shipped button calls BuildManagerTabs by name, so renaming it breaks
'   the button with no compile error.
'
' One-button weekly workflow:
'   - Select the current raw-data workbook; no copy/paste.
'   - The prior report is found automatically in the output folder
'     (Cancel/none = no comparison); prior tabs are read read-only.
'   - Automatically finds the source worksheet by required headers.
'   - Validates required headers before processing.
'   - Null-safe: blank manager/status/header cells are skipped, not fatal.
'   - Full source formatting (dates, number formats) carried into output.
'   - Every data row with a manager counts as a PM request.
'   - New PM rows are light green + bold (names absent from prior report).
'   - Dashboard summary (KPIs + manager table) is the first tab.
'   - Raw Data is copied into the report as the last tab.
'   - Remembers the output folder in a hidden PM Settings sheet.
'   - First run asks for an output folder; later runs reuse it.
'   - Prevents accidental filename overwrite (unique filenames).
'   - Repairs/installs the START HERE button automatically.
'   - Keeps the source workbook read-only and closes it without saving.
'
' IMPORTANT BUSINESS RULES:
'   * Every data row assigned to a manager is a PM request on that manager's tab.
'   * Treatment Status is carried into the report but does not filter rows.
'   * PM Needed equals the number of data rows on the manager tab.
'   * Treatment PM Date is not a PM-needed filter.
'
' MAINTENANCE:
'   * The MAP below is a snapshot of section starts and procedure names. If you
'     move code, update the MAP in the same change; scripts\check-toc.ps1 checks
'     it against the file and prints drift (read-only, no Excel).
'   * One section per concern. New helpers go in the section they serve, not in
'     a new section, unless the concern is genuinely new.
' ============================================================
' ============================================================
' MAP OF THIS MODULE - snapshot of section starts and procedures.
' Generated and checked by scripts\check-toc.ps1 (add -Fix to rewrite).
' Reading order for a change: CORE first, then the section you touch.
'
' 10 sections, 79 procedures.
' Public (outside callers): BuildManagerTabs, CreateEmailFromReport, EditCCEmails,
'        EditEmailTemplate, GenerateManagerEmail, GenerateReportEmail, LogPath,
'        RebuildStartSheet, RunReport, SetupEmailButtons, SetupGenerateEmailButton,
'        SetupReportButton
'
'   #  SECTION (line)                          procedures
'
'   [  132] INTERACTIVE ENTRY (button)
'        SetupReportButton, SetupEmailButtons, SetupGenerateEmailButton, GenerateReportEmail,
'        RebuildStartSheet, BuildManagerTabs
'   [  416] CORE (host-neutral)
'        RunReport
'   [  785] HOST / DIALOG HELPERS
'        PickSourceFile, PickOutputFolder, SetStatus
'   [  830] RUN LOG
'        LogReset, LogOpen, LogLine, LogStage, LogPath, LogErr, DebugWriteLine, FastModeOn
'   [  982] WORKBOOK / SHEET SETUP
'        GetOrCreateStartSheet, AddStartStep, AddStartBullet, PositionStartLogo,
'        FindSourceWorksheet, LoadSourceIntoRawData, GetOrCreateRawSheet,
'        GetOrCreateSettingsSheet, GetRememberedOutputFolder, SaveOutputFolder
'   [ 1337] PATH / STATE HELPERS
'        FolderExists, NormalizeFolder, GetUniqueOutputPath, RestoreExcelState
'   [ 1420] DATA / FORMAT HELPERS
'        BuildOutputArray, GetHeaderArray, CopyRowFormats, SortManagersByLastName, LastNameKey,
'        FindHeaderRow, FindHeaderColumn, LastUsedRow, LastUsedColumn, SheetExists,
'        SafeSheetName, SafeStr, FindManager, GetOrCreateOutputSheet, UniqueSheetName
'   [ 1678] PRIOR REPORT COMPARISON
'        DetectPriorReport, PriorNameOverlap, SortPathsNewestFirst, LoadPriorComparison,
'        IsManagerTab, FindHeaderRowForSingle, NormalizeName, PriorNameExists, CountDiff,
'        PriorMgrCount
'   [ 1984] OUTPUT FORMATTING
'        FormatManagerSheet, HideUnusedCells, CountInSheet, AddPolishedCard,
'        AddPolishedTextCard, ChangeColor, ApplyReaderFriendlyFormatting
'   [ 2386] DASHBOARD
'        BuildDashboard, ColorReportTabs, AddManagerChart, WriteTableHeader,
'        GetOrCreateDashboardSheet, ResolveLogoPath, AddDashboardLogo, PositionDashboardLogo,
'        SortByColumnA, CreateEmailFromReport, GetEmailTemplate, GetCCEmails, EditEmailTemplate,
'        EditCCEmails, GenerateManagerEmail
'
' ============================================================
Private Const RAW_SHEET As String = "Raw Data"
Private Const SETTINGS_SHEET As String = "PM Settings"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const MANAGER_HEADER As String = "Assigned Sales Manager"
Private Const NAME_HEADER As String = "Name"
Private Const STATUS_HEADER As String = "Treatment Status"
Private Const BUTTON_NAME As String = "btnCreatePMReport"
Private Const BUTTON_TEXT As String = "CREATE PM REQUEST REPORT"
Private Const SETTING_OUTPUT_FOLDER As String = "OutputFolder"
Private Const LOGO_SHAPE As String = "AOTI_Logo"
Private Const START_SHEET As String = "START HERE"
Private Const LOGO_FILE As String = "aoti-logo.png"
Private Const LOGO_DIR As String = "documentation"
Private Const SHEET_COL_WIDTH As Double = 16
Private Const UI_FONT As String = "Segoe UI"
Private Const MANAGER_LAST_COL As Long = 8

' Run-log state (must live in the declarations section, not between procedures).
Private gLogPath As String
Private gLogPrev As Double
Private gLogT0 As Double

' Realtime error-logging state (mirrors the Pending Pickup pattern).
' mInReportRun  - true while RunReport is executing, so LogErr also hits the
'                 per-run log. False means only the global debug log is written.
' mRunLogPath   - path to the current per-run log (mirrors gLogPath).
' mDebugLogPath - path to the global debug log (appended across every run).
Private mInReportRun As Boolean
Private mRunLogPath As String
Private mDebugLogPath As String


' ==================== INTERACTIVE ENTRY (button) ====================
' Owns the UI shell only: the START HERE button and the dialogs around a build.
'
'   SetupReportButton  - (re)builds the canvas and the CTA button, wires it to
'                        '<workbook name>'!BuildManagerTabs. Silent=True is the
'                        automation path (sync_template.ps1, RebuildStartSheet).
'   RebuildStartSheet  - what sync_template.ps1 calls; no report is run.
'   BuildManagerTabs   - the real button action: source dialog, output folder,
'                        RunReport, final message.
'
' Contracts:
'   * The button shape is deleted and recreated every run, so its style cannot
'     drift. Its name (BUTTON_NAME) is what OnAction replacement relies on.
'   * OnAction is a string built from ThisWorkbook.Name. Renaming the workbook or
'     BuildManagerTabs breaks the button silently - no compile error, no log.
'   * Every cancel path must clear the status bar (SetStatus "") or the user is
'     left looking at "Step 1 of 3..." from the previous attempt.
'

Public Sub SetupReportButton(Optional ByVal Silent As Boolean = False)
    Dim ws As Worksheet, shp As Shape
    Dim action As String
    Dim btnWidth As Double, btnHeight As Double
    Dim btnLeft As Double, btnTop As Double

    On Error GoTo ButtonFail

    action = "'" & ThisWorkbook.Name & "'!BuildManagerTabs"

    Set ws = GetOrCreateStartSheet()

    ' One CTA, recreated fresh so the shape type/style never drifts.
    On Error Resume Next
    ws.Shapes(BUTTON_NAME).Delete
    On Error GoTo ButtonFail

    btnWidth = 360
    btnHeight = 48
    ' Center on the content area (C:F), shrinking if the columns are narrower.
    If ws.Range("C14:F14").Width - 8 < btnWidth Then btnWidth = ws.Range("C14:F14").Width - 8
    btnLeft = ws.Range("C14").Left + (ws.Range("C14:F14").Width - btnWidth) / 2
    btnTop = ws.Range("B14").Top + (ws.Rows(14).Height - btnHeight) / 2

    Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)  ' rounded rectangle
    shp.Name = BUTTON_NAME

    On Error Resume Next
    With shp
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(46, 96, 148)
        .Line.Weight = 1.5
        With .TextFrame
            .Characters.Text = BUTTON_TEXT
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = 14
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
        End With
        .Placement = xlMove
    End With
    shp.OnAction = action
    On Error GoTo ButtonFail

    If Not Silent Then MsgBox "The CREATE PM REQUEST REPORT button is ready.", vbInformation
    Exit Sub

ButtonFail:
    LogErr "SetupReportButton"
    If Not Silent Then MsgBox "I could not install the button." & vbCrLf & Err.Description, vbCritical
End Sub


' Public entry point for sync_template.ps1: rebuild START HERE + button so the
' shipped workbook matches the current module without a full report run.
Public Sub SetupEmailButtons(Optional ByVal Silent As Boolean = False)
    Dim ws As Worksheet, shp As Shape
    Dim action1 As String, action2 As String
    Dim btnWidth As Double, btnHeight As Double
    Dim btnLeft As Double, btnTop As Double

    On Error GoTo ButtonFail

    action1 = "'" & ThisWorkbook.Name & "'!SetupManagerEmail"
    action2 = "'" & ThisWorkbook.Name & "'!SetupDirectorEmail"

    Set ws = GetOrCreateStartSheet()

    ' Remove old buttons if they exist
    On Error Resume Next
    ws.Shapes("BtnManagerTemplate").Delete
    ws.Shapes("BtnDirectorTemplate").Delete
    ws.Shapes("BtnGenerateEmail").Delete
    ws.Shapes("BtnEditEmailTemplate").Delete
    ws.Shapes("BtnEditToEmails").Delete
    ws.Shapes("BtnEditCCEmails").Delete
    On Error GoTo ButtonFail

    btnWidth = 160
    btnHeight = 30
    btnLeft = ws.Range("C23").Left + 4
    btnTop = ws.Range("B24").Top + (ws.Rows(24).Height - btnHeight) / 2

    Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
    shp.Name = "BtnManagerTemplate"
    With shp
        .Fill.ForeColor.RGB = RGB(70, 130, 180)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(46, 96, 148)
        .Line.Weight = 1
        With .TextFrame
            .Characters.Text = "Manager Template"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = 10
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
        End With
        .Placement = xlMove
    End With
    shp.OnAction = action1

    btnLeft = btnLeft + btnWidth + 8
    Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
    shp.Name = "BtnDirectorTemplate"
    With shp
        .Fill.ForeColor.RGB = RGB(70, 130, 180)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(46, 96, 148)
        .Line.Weight = 1
        With .TextFrame
            .Characters.Text = "Director Template"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = 10
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
        End With
        .Placement = xlMove
    End With
    shp.OnAction = action2

    If Not Silent Then MsgBox "Email setup buttons are ready.", vbInformation
    Exit Sub

ButtonFail:
    LogErr "SetupEmailButtons"
    If Not Silent Then MsgBox "I could not install the email buttons." & vbCrLf & Err.Description, vbCritical
End Sub

' Also add a button to generate the email
Public Sub SetupGenerateEmailButton(Optional ByVal Silent As Boolean = False)
    Dim ws As Worksheet, shp As Shape
    Dim action As String
    Dim btnWidth As Double, btnHeight As Double
    Dim btnLeft As Double, btnTop As Double

    On Error GoTo ButtonFail

    action = "'" & ThisWorkbook.Name & "'!GenerateReportEmail"

    Set ws = GetOrCreateStartSheet()

    On Error Resume Next
    ws.Shapes("BtnGenerateEmail").Delete
    On Error GoTo ButtonFail

    btnWidth = 160
    btnHeight = 30
    btnLeft = ws.Range("C23").Left + 4 + 160 + 8 + 160 + 8
    btnTop = ws.Range("B24").Top + (ws.Rows(24).Height - btnHeight) / 2

    Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
    shp.Name = "BtnGenerateEmail"
    With shp
        .Fill.ForeColor.RGB = RGB(34, 139, 34)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(0, 100, 0)
        .Line.Weight = 1
        With .TextFrame
            .Characters.Text = "Generate Email"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = 10
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
        End With
        .Placement = xlMove
    End With
    shp.OnAction = action

    If Not Silent Then MsgBox "Generate Email button is ready.", vbInformation
    Exit Sub

ButtonFail:
    LogErr "SetupGenerateEmailButton"
    If Not Silent Then MsgBox "I could not install the button." & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub SetupDoneButtons(Optional ByVal Silent As Boolean = False)
    Dim ws As Worksheet, shp As Shape
    Dim action As String
    Dim btnWidth As Double, btnHeight As Double
    Dim btnLeft As Double, btnTop As Double

    On Error GoTo ButtonFail

    action = "'" & ThisWorkbook.Name & "'!DoneEditing"

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Manager Email Template" Or ws.Name = "Director Email Template" Then
            ' Unprotect to allow button creation
            On Error Resume Next
            ws.Unprotect
            ws.Shapes("BtnDone").Delete
            On Error GoTo ButtonFail

            btnWidth = 180
            btnHeight = 24
            btnLeft = ws.Range("K2").Left + ws.Range("K2").Width - btnWidth - 4
            btnTop = 4

            Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
            shp.Name = "BtnDone"
            With shp
                .Fill.ForeColor.RGB = RGB(31, 78, 121)
                .Fill.Solid
                .Line.ForeColor.RGB = RGB(31, 78, 121)
                .Line.Weight = 1
                .TextFrame2.TextRange.Text = "Done - Return to START HERE"
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .TextFrame2.TextRange.Font.Name = "Segoe UI"
                .TextFrame2.TextRange.Font.Size = 9
                .TextFrame2.TextRange.Font.Bold = msoTrue
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Placement = xlMove
            End With
            shp.OnAction = action

            ' Re-protect
            ws.Protect
            ws.EnableSelection = xlUnlockedCells
        End If
    Next ws

    If Not Silent Then MsgBox "Done buttons are ready.", vbInformation
    Exit Sub

ButtonFail:
    LogErr "SetupDoneButtons"
    If Not Silent Then MsgBox "I could not install the Done buttons." & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub GenerateReportEmail()
    On Error GoTo LocalFail
    Dim wb As Workbook
    ' Use the most recent report if available, otherwise prompt
    Dim reportPath As String
    reportPath = InputBox("Path to the report .xlsx to email:", "Generate Email", "")
    If Len(reportPath) = 0 Then Exit Sub
    If Dir(reportPath) = "" Then
        MsgBox "File not found: " & reportPath, vbCritical
        Exit Sub
    End If
    Set wb = Workbooks.Open(reportPath)
    CreateEmailFromReport wb
    Exit Sub
LocalFail:
    LogErr "GenerateReportEmail"
End Sub
Public Sub SetupSheetVisibility()
    On Error GoTo LocalFail
    Dim ws As Worksheet
    ' Move START HERE to first position
    ThisWorkbook.Sheets("START HERE").Move Before:=ThisWorkbook.Sheets(1)
    ' Hide all sheets except START HERE
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "START HERE" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    Exit Sub
LocalFail:
    LogErr "SetupSheetVisibility"
End Sub

Public Sub SetupManagerEmail()
    On Error GoTo LocalFail
    ThisWorkbook.Sheets("Manager Email Template").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Manager Email Template").Activate
    ThisWorkbook.Sheets("Manager Email Template").Range("D5").Select
    Exit Sub
LocalFail:
    LogErr "SetupManagerEmail"
End Sub

Public Sub SetupDirectorEmail()
    On Error GoTo LocalFail
    ThisWorkbook.Sheets("Director Email Template").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Director Email Template").Activate
    ThisWorkbook.Sheets("Director Email Template").Range("D5").Select
    Exit Sub
LocalFail:
    LogErr "SetupDirectorEmail"
End Sub

Public Sub DoneEditing()
    On Error GoTo LocalFail
    If ActiveSheet.Name = "Manager Email Template" Or ActiveSheet.Name = "Director Email Template" Then
        ActiveSheet.Visible = xlSheetHidden
    End If
    ThisWorkbook.Sheets("START HERE").Activate
    Exit Sub
LocalFail:
    LogErr "DoneEditing"
End Sub

Public Sub RebuildStartSheet()
    SetupSheetVisibility
    SetupReportButton True
    SetupEmailButtons True
    SetupDoneButtons True
End Sub


Public Sub BuildManagerTabs()
    Dim sourcePath As Variant
    Dim outputFolder As String
    Dim resultParts() As String
    Dim result As String

    On Error GoTo CleanFail

    SetupReportButton True

    SetStatus "Step 1 of 3 - Select the current raw-data report..."

    sourcePath = PickSourceFile()
    If VarType(sourcePath) = vbBoolean And sourcePath = False Then SetStatus "": Exit Sub

    If StrComp(CStr(sourcePath), ThisWorkbook.fullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 2201, , "The selected raw-data file is the PM Requests master workbook. Please select the current report export instead."
    End If

    ' Step 2 of 3 - the prior report is found automatically from the output
    ' folder (previous week's saved report). No dialog: one click is the goal.
    outputFolder = GetRememberedOutputFolder(ThisWorkbook)
    If Len(outputFolder) = 0 Or Not FolderExists(outputFolder) Then
        outputFolder = NormalizeFolder(PickOutputFolder())
        If Len(outputFolder) = 0 Then SetStatus "": Exit Sub
        ' Save the master immediately so the remembered folder survives a close
        ' without a save prompt answer.
        SaveOutputFolder ThisWorkbook, outputFolder
        ThisWorkbook.Save
    End If

    SetStatus "Step 2 of 3 - Building the PM Requests report..."
    ' RunReport owns the log (one file per run); logging here would write the
    ' picker lines into the previous run's file.
    result = RunReport(CStr(sourcePath), outputFolder, "")

    resultParts = Split(result, "|")
    If UBound(resultParts) >= 0 Then
        If resultParts(0) = "OK" Then
            MsgBox "Done!" & vbCrLf & vbCrLf & _
                   "Created: " & vbCrLf & resultParts(1) & vbCrLf & vbCrLf & _
                   "Manager PM requests: " & resultParts(2) & vbCrLf & _
                   IIf(CLng(resultParts(3)) > 0, "Skipped rows (no manager): " & resultParts(3) & vbCrLf, "") & _
                   IIf(UBound(resultParts) >= 4 And Len(resultParts(4)) > 0, "Compared against: " & resultParts(4) & vbCrLf, "No prior report found - comparison skipped." & vbCrLf), vbInformation
        Else
            MsgBox "The PM report could not be created." & vbCrLf & vbCrLf & _
                   "Error " & resultParts(1) & ": " & resultParts(2), vbCritical
        End If
    End If
    Exit Sub

CleanFail:
    LogErr "BuildManagerTabs"
    LogLine "!!! ERROR " & CStr(Err.Number) & ": " & Err.Description & "   [BuildManagerTabs]"
    SetStatus ""
    MsgBox "The PM report could not be created." & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical
End Sub


' ====================== CORE (host-neutral) ======================
' Owns the whole build. Nothing here touches a dialog or a MsgBox: RunReport is
' callable from the button, from build-worker.ps1, or from a test.
'
' Pipeline, in order (each step is a `stage` breadcrumb in the run log):
'   fastmode -> source open -> raw load -> data read -> classified ->
'   [detect prior] -> manager tabs -> dashboard -> dashboard moved ->
'   raw copied -> tabs colored -> view set -> saved.
'
' Contracts:
'   * Result string is the only output: "OK|savePath|pmTotal|skipped|prior"
'     or "ERROR|number|description @ stage". scripts\run-report.ps1 parses it;
'     the regression build must be 124 PMs and 0 skipped.
'   * A row is a PM request when Assigned Sales Manager is non-empty. Blank
'     manager = skipped and counted, never fatal (CStr(Null) raises error 94).
'   * Manager tab geometry is fixed: KPI cards above, table header on row 9,
'     data from row 10. FormatManagerSheet, ApplyReaderFriendlyFormatting and
'     AddManagerChart all assume it.
'   * The source workbook is opened ReadOnly and closed with SaveChanges:=False;
'     the report is built from the Raw Data copy, so every bound is recomputed
'     after the copy.
'   * Fast mode is saved and restored on both the success and the Fail path.
'   * `stage` is captured into errStage before any cleanup runs, so the error
'     line names the step that failed. Do not reset `stage` in Fail.
'

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
Private Function PickSourceFile() As Variant
    PickSourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Reports (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Select the CURRENT raw-data report")
End Function


' Output folder: Excel folder picker ("" = user cancelled).
Private Function PickOutputFolder() As String
    Dim fd As Object

    On Error Resume Next
    Set fd = Application.FileDialog(4)
    If Not fd Is Nothing Then
        fd.Title = "Choose the folder for finished PM Requests reports"
        fd.AllowMultiSelect = False
        If fd.Show = -1 Then PickOutputFolder = fd.SelectedItems(1)
    End If
    On Error GoTo 0
End Function


Private Sub SetStatus(ByVal text As String)
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

Private Sub LogReset()
    gLogPath = ""
    gLogPrev = 0
    gLogT0 = 0
End Sub


Private Sub LogOpen(ByVal outputFolder As String)
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


Private Sub LogLine(ByVal msg As String)
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
Private Sub LogStage(ByVal label As String)
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
Private Sub DebugWriteLine(ByVal path As String, ByVal line As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open path For Append As #ff
    Print #ff, line
    Close #ff
    On Error GoTo 0
End Sub


Private Sub FastModeOn()
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
Private Function GetOrCreateStartSheet() As Worksheet
    Dim ws As Worksheet

    On Error GoTo LocalFail
    Set ws = ThisWorkbook.Worksheets(START_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = START_SHEET
    End If

    ' Rebuild the canvas every run; stale merges would block the new layout.
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.UnMerge
    On Error GoTo 0

    ' Gutters + content columns: narrow number gutter in B, wide content in
    ' C:E, thin right gutter in F.
    ws.Columns("A").ColumnWidth = 2.6
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E").ColumnWidth = 28
    ws.Columns("F").ColumnWidth = 2

    ws.Rows(1).RowHeight = 8
    ws.Rows(2).RowHeight = 28
    ws.Rows(3).RowHeight = 22
    ws.Rows(4).RowHeight = 5
    ws.Rows(5).RowHeight = 10
    ws.Rows(6).RowHeight = 22
    ws.Rows(7).RowHeight = 10
    ws.Rows(8).RowHeight = 24
    ws.Rows("9:12").RowHeight = 26
    ws.Rows(13).RowHeight = 22
    ws.Rows(14).RowHeight = 64
    ws.Rows(15).RowHeight = 12
    ws.Rows(16).RowHeight = 24
    ws.Rows("17:19").RowHeight = 20
    ws.Rows(20).RowHeight = 12
    ws.Rows(21).RowHeight = 18
    ws.Rows(22).RowHeight = 12
    ws.Rows(23).RowHeight = 24
    ws.Rows(24).RowHeight = 36
    ws.Rows(25).RowHeight = 12
    ws.Rows(26).RowHeight = 18

    ' Two-tone rule under the title: navy across the title, gold to the right edge.
    ws.Range("B4:D4").Interior.Color = RGB(31, 78, 121)
    ws.Range("E4:F4").Interior.Color = RGB(246, 206, 62)

    ' Title (navy, no fill) with the logo zone to its right.
    With ws.Range("B2:D3")
        .Merge
        .Value = "PM REQUESTS REPORT BUILDER"
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B6:F6")
        .Merge
        .Value = "Build this week's PM Requests report in one click - no copy-paste, no formatting."
        .Font.Size = 11
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' HOW IT WORKS label bar + numbered steps.
    With ws.Range("B8:F8")
        .Merge
        .Value = "HOW IT WORKS"
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B8").IndentLevel = 1

    AddStartStep ws, 9, "Select this week's raw treatment export.", RGB(242, 247, 252)
    AddStartStep ws, 10, "The prior week's report is found automatically and used for comparison.", RGB(255, 255, 255)
    AddStartStep ws, 11, "Choose where to save. The folder is remembered for next time.", RGB(242, 247, 252)
    AddStartStep ws, 12, "Open the finished PM Requests YYYY-MM-DD.xlsx report.", RGB(255, 255, 255)

    With ws.Range("B13:F13")
        .Merge
        .Value = "Ready to build this week's report?"
        .Font.Italic = False
        .Font.Size = 11
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' CTA band: the button sits on this light blue strip.
    ws.Range("B14:F14").Interior.Color = RGB(221, 235, 247)

    ' Email templates bar + buttons area.
    With ws.Range("B20:F20")
        .Merge
        .Value = "EMAIL TEMPLATES"
        .Interior.Color = RGB(70, 130, 180)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B20").IndentLevel = 1

    ' GOOD TO KNOW bar + bullets.
    With ws.Range("B16:F16")
        .Merge
        .Value = "GOOD TO KNOW"
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B16").IndentLevel = 1

    AddStartBullet ws, 17, "One tab per Assigned Sales Manager, plus the Dashboard and Raw Data."
    AddStartBullet ws, 18, "Nothing is overwritten: if the dated report exists, a numbered copy is saved beside it."
    AddStartBullet ws, 19, "PMs new since the prior week's report are highlighted green on each manager tab."

    With ws.Range("B21:F21")
        .Merge
        .Value = "v1.4  " & ChrW(8226) & "  PM Requests Report Builder  " & ChrW(8226) & "  Output: PM Requests YYYY-MM-DD.xlsx"
        .Font.Size = 9
        .Font.Color = RGB(191, 191, 191)
        .HorizontalAlignment = xlCenter
    End With

    ' Email templates bar + buttons area.
    With ws.Range("B23:F23")
        .Merge
        .Value = "EMAIL TEMPLATES"
        .Interior.Color = RGB(70, 130, 180)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B23").IndentLevel = 1

    ws.Range("B2:F26").Font.Name = UI_FONT

    ' Logo: keep the embedded START HERE picture, top-right of the title row.
    PositionStartLogo ws

    ' Cleaner canvas: hide gridlines on this sheet (best-effort).
    On Error Resume Next
    ws.Tab.Color = RGB(31, 78, 121)   ' navy tab to match the title band
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ' Protect the sheet: lock all cells, allow button clicks
    ws.Cells.Locked = True
    ws.Protect
    ws.EnableSelection = xlUnlockedCells

    Set GetOrCreateStartSheet = ws
    Exit Function
LocalFail:
    LogErr "GetOrCreateStartSheet"
End Function


' One numbered step row: navy number cell + merged text on the given fill.
Private Sub AddStartStep(ws As Worksheet, ByVal rowNum As Long, ByVal text As String, ByVal fillColor As Long)
    On Error GoTo LocalFail
    With ws.Cells(rowNum, "B")
        .Value = CStr(rowNum - 8)
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(rowNum, "C"), ws.Cells(rowNum, "F"))
        .Merge
        .Value = text
        .Interior.Color = fillColor
        .Font.Size = 11
        .Font.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(rowNum, "C").IndentLevel = 1
    Exit Sub
LocalFail:
    LogErr "AddStartStep"
End Sub


' One "good to know" bullet row. The glyph is built with ChrW so this module stays
' ASCII-only and imports identically whatever the codepage of the .bas file is.
Private Sub AddStartBullet(ws As Worksheet, ByVal rowNum As Long, ByVal text As String)
    On Error GoTo LocalFail
    With ws.Range(ws.Cells(rowNum, "B"), ws.Cells(rowNum, "F"))
        .Merge
        .Value = ChrW(8226) & "  " & text
        .Font.Size = 10
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(rowNum, "B").IndentLevel = 1
    Exit Sub
LocalFail:
    LogErr "AddStartBullet"
End Sub


' Places the embedded AOTI_Logo picture at the top-right of the title rows.
Private Sub PositionStartLogo(ws As Worksheet)
    Dim shp As Shape, logoW As Double, logoH As Double
    Dim rowTop As Double, rowHeight As Double

    On Error GoTo LocalFail
    Set shp = ws.Shapes(LOGO_SHAPE)
    If shp Is Nothing Then Exit Sub

    ' Match the sibling builder: 150 pt wide, height keeps the picture's own aspect.
    If shp.Width > 0 Then
        logoH = shp.Height / shp.Width * 150
        logoW = 150
        shp.Width = logoW
        shp.Height = logoH
    Else
        logoW = shp.Width
        logoH = shp.Height
    End If
    If logoH <= 0 Then logoH = 43

    rowTop = ws.Range("B2").Top
    rowHeight = ws.Range("B2:B3").Height

    On Error Resume Next
    shp.Left = ws.Range("E2").Left + 1
    shp.Top = rowTop + (rowHeight - logoH) / 2
    shp.Placement = xlMove
    Exit Sub
LocalFail:
    LogErr "PositionStartLogo"
End Sub


Private Function FindSourceWorksheet(wb As Workbook, ByRef headerRowOut As Long) As Worksheet
    Dim ws As Worksheet, hr As Long

    On Error GoTo LocalFail
    For Each ws In wb.Worksheets
        hr = FindHeaderRow(ws, MANAGER_HEADER, STATUS_HEADER)
        If hr > 0 Then
            headerRowOut = hr
            Set FindSourceWorksheet = ws
            Exit Function
        End If
    Next ws
    Exit Function
LocalFail:
    LogErr "FindSourceWorksheet"
End Function


' Loads the source range into Raw Data AND copies the full range formats,
' so date/number formats survive into the report (Value2 alone loses them).
Private Sub LoadSourceIntoRawData(masterWb As Workbook, src As Worksheet, headerRow As Long, lastRow As Long, lastCol As Long)
    Dim dst As Worksheet
    Dim srcRange As Range
    Dim r As Long

    On Error GoTo LocalFail
    Set dst = GetOrCreateRawSheet(masterWb)
    dst.Cells.Clear
    Set srcRange = src.Range(src.Cells(headerRow, 1), src.Cells(lastRow, lastCol))
    dst.Range("A1").Resize(srcRange.Rows.Count, srcRange.Columns.Count).Value2 = srcRange.Value2
    For r = 1 To srcRange.Rows.Count
        CopyRowFormats src, dst, headerRow + r - 1, r, lastCol
    Next r
    dst.Columns.AutoFit
    Exit Sub
LocalFail:
    LogErr "LoadSourceIntoRawData"
End Sub


Private Function GetOrCreateRawSheet(wb As Workbook) As Worksheet
    On Error GoTo LocalFail
    If SheetExists(RAW_SHEET, wb) Then
        Set GetOrCreateRawSheet = wb.Worksheets(RAW_SHEET)
    Else
        Set GetOrCreateRawSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateRawSheet.Name = RAW_SHEET
    End If
    Exit Function
LocalFail:
    LogErr "GetOrCreateRawSheet"
End Function


Private Function GetOrCreateSettingsSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error GoTo LocalFail
    If SheetExists(SETTINGS_SHEET, wb) Then
        Set ws = wb.Worksheets(SETTINGS_SHEET)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SETTINGS_SHEET
        ws.Range("A1").Value = "Setting"
        ws.Range("D7").Value = "Value"
        ws.Range("A2").Value = SETTING_OUTPUT_FOLDER
    End If
    ws.Visible = xlSheetVisible
    Set GetOrCreateSettingsSheet = ws
    Exit Function
LocalFail:
    LogErr "GetOrCreateSettingsSheet"
End Function


Private Function GetRememberedOutputFolder(wb As Workbook) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    On Error GoTo LocalFail
    If Not ws Is Nothing Then GetRememberedOutputFolder = Trim$(SafeStr(ws.Range("B2").Value))
    Exit Function
LocalFail:
    LogErr "GetRememberedOutputFolder"
End Function


Private Sub SaveOutputFolder(wb As Workbook, folderPath As String)
    Dim ws As Worksheet

    On Error GoTo LocalFail
    Set ws = GetOrCreateSettingsSheet(wb)
    ws.Range("B2").Value = folderPath
    Exit Sub
LocalFail:
    LogErr "SaveOutputFolder"
End Sub


' ====================== PATH / STATE HELPERS ======================
' Owns path hygiene and Excel's saved state. Which sheet a file goes into is
' WORKBOOK / SHEET SETUP; this section only answers "where" and "put it back".
'
'   FolderExists / NormalizeFolder - path hygiene (trailing separator, ~$).
'   GetUniqueOutputPath     - never overwrite: "PM Requests <date> (2).xlsx".
'   RestoreExcelState       - puts calculation/screen/events/alerts back.
'
' Contracts:
'   * Nothing is written outside the project folder. Scratch and logs live under
'     the workbook's own folder (see scripts\ for the PowerShell side of this).
'   * GetUniqueOutputPath is what makes a re-run safe; if it ever returns an
'     existing path, a user report is lost.
'   * RestoreExcelState runs on both the success and the Fail path of RunReport.
'

Private Function FolderExists(folderPath As String) As Boolean
    Dim clean As String

    On Error GoTo LocalFail
    clean = NormalizeFolder(folderPath)
    FolderExists = (Len(Dir(clean, vbDirectory)) > 0)
    Exit Function
LocalFail:
    LogErr "FolderExists"
End Function


' Strip any trailing path separator so joins never double up.
Private Function NormalizeFolder(ByVal folderPath As String) As String
    Dim s As String

    On Error GoTo LocalFail
    s = Trim$(folderPath)
    Do While Len(s) > 1 And Right$(s, 1) = Application.PathSeparator
        s = Left$(s, Len(s) - 1)
    Loop
    NormalizeFolder = s
    Exit Function
LocalFail:
    LogErr "NormalizeFolder"
End Function


Private Function GetUniqueOutputPath(folderPath As String, baseFileName As String) As String
    Dim candidate As String, n As Long, dotPos As Long

    On Error GoTo LocalFail
    folderPath = NormalizeFolder(folderPath)
    candidate = folderPath & Application.PathSeparator & baseFileName
    If Len(Dir(candidate)) = 0 Then
        GetUniqueOutputPath = candidate
        Exit Function
    End If
    dotPos = InStrRev(baseFileName, ".")
    n = 2
    Do
        If dotPos > 0 Then
            candidate = folderPath & Application.PathSeparator & Left$(baseFileName, dotPos - 1) & " (" & CStr(n) & ")" & Mid$(baseFileName, dotPos)
        Else
            candidate = folderPath & Application.PathSeparator & baseFileName & " (" & CStr(n) & ")"
        End If
        n = n + 1
    Loop While Len(Dir(candidate)) > 0
    GetUniqueOutputPath = candidate
    Exit Function
LocalFail:
    LogErr "GetUniqueOutputPath"
End Function


Private Sub RestoreExcelState(oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean)
    On Error GoTo LocalFail
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Exit Sub
LocalFail:
    LogErr "RestoreExcelState"
End Sub


' ====================== DATA / FORMAT HELPERS ======================
' Owns the small primitives everything else is built from: reading a header
' row, finding a column, reading used bounds, copying formats, sorting names.
'
'   BuildOutputArray / GetHeaderArray - row blocks as Variant arrays.
'   CopyRowFormats                    - source formatting onto an output row.
'   SortManagersByLastName / LastNameKey - manager tab order.
'   FindHeaderRow / FindHeaderColumn  - header lookup by text (raises when the
'                                       header is missing).
'   LastUsedRow / LastUsedColumn      - used bounds of a sheet.
'   SheetExists / SafeSheetName / UniqueSheetName - sheet naming (31 chars,
'                                       no : \ / ? * [ ]).
'   SafeStr / FindManager / GetOrCreateOutputSheet.
'
' Contracts:
'   * SafeStr is the null guard. Call it on every cell value that came from a
'     Variant array: CStr(Null) raises error 94 and takes the run down.
'   * Lookups are Collection-based on purpose: no Scripting.Dictionary COM
'     dependency, so the module works on a locked-down desktop.
'   * Collection items cannot be reassigned, so counters live in a Long array
'     (repCounts) and are ReDim Preserve'd as they grow.
'   * FindHeaderColumn is a hard failure by design - a missing required header
'     means the wrong export, and the message names the header.
'

Private Function BuildOutputArray(ByRef data As Variant, ByVal rows As Collection, ByVal lastCol As Long) As Variant
    Dim result() As Variant, i As Long, c As Long, srcRow As Long

    On Error GoTo LocalFail
    ReDim result(1 To rows.Count + 1, 1 To lastCol)
    For c = 1 To lastCol: result(1, c) = data(1, c): Next c
    For i = 1 To rows.Count
        srcRow = CLng(rows(i))
        For c = 1 To lastCol: result(i + 1, c) = data(srcRow, c): Next c
    Next i
    BuildOutputArray = result
    Exit Function
LocalFail:
    LogErr "BuildOutputArray"
End Function


Private Function GetHeaderArray(ByRef data As Variant, ByVal lastCol As Long) As Variant
    Dim a() As Variant, c As Long

    On Error GoTo LocalFail
    ReDim a(1 To 1, 1 To lastCol)
    For c = 1 To lastCol: a(1, c) = data(1, c): Next c
    GetHeaderArray = a
    Exit Function
LocalFail:
    LogErr "GetHeaderArray"
End Function


' Per-cell format copy (number formats, fonts) for one row. Property-based
' (no clipboard) so it cannot fail on a clipboard lock or in invisible
' automation. Column-level NumberFormat is unreliable when cells carry
' individual formats, so formats are copied cell by cell.
Private Sub CopyRowFormats(src As Worksheet, dst As Worksheet, ByVal srcRow As Long, ByVal dstRow As Long, ByVal lastCol As Long)
    Dim c As Long
    Dim sf As Font, df As Font

    On Error GoTo LocalFail
    For c = 1 To lastCol
        dst.Cells(dstRow, c).NumberFormat = src.Cells(srcRow, c).NumberFormat
        Set sf = src.Cells(srcRow, c).Font
        Set df = dst.Cells(dstRow, c).Font
        df.Name = sf.Name
        df.Size = sf.Size
        df.Bold = sf.Bold
        df.Color = sf.Color
    Next c
    Exit Sub
LocalFail:
    LogErr "CopyRowFormats"
End Sub


Private Sub SortManagersByLastName(ByRef a As Variant)
    Dim i As Long, j As Long, tmp As Variant

    On Error GoTo LocalFail
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If LastNameKey(CStr(a(j))) < LastNameKey(CStr(a(i))) Then
                tmp = a(i): a(i) = a(j): a(j) = tmp
            End If
        Next j
    Next i
    Exit Sub
LocalFail:
    LogErr "SortManagersByLastName"
End Sub


Private Function LastNameKey(ByVal fullName As String) As String
    Dim s As String, p As Long, commaPos As Long
    s = Trim$(fullName)
    commaPos = InStr(1, s, ",")
    If commaPos > 0 Then
        ' Source names are commonly stored as "Last, First".
        LastNameKey = LCase$(Trim$(Left$(s, commaPos - 1)) & "|" & Trim$(Mid$(s, commaPos + 1)))
    Else
        p = InStrRev(s, " ")
        If p > 0 Then
            LastNameKey = LCase$(Mid$(s, p + 1) & "|" & Left$(s, p - 1))
        Else
            LastNameKey = LCase$(s)
        End If
    End If
End Function


Private Function FindHeaderRow(ws As Worksheet, header1 As String, header2 As String) As Long
    Dim r As Long, topRow As Long

    On Error GoTo LocalFail
    topRow = LastUsedRow(ws)
    If topRow > 30 Then topRow = 30
    For r = 1 To topRow
        If FindHeaderColumn(ws, r, header1) > 0 And FindHeaderColumn(ws, r, header2) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
    Exit Function
LocalFail:
    LogErr "FindHeaderRow"
End Function


' Null-safe: CStr(Null) raises error 94, so blank header cells used to crash
' header detection. Now they are simply skipped.
Private Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim c As Long, lastCol As Long

    On Error GoTo LocalFail
    If headerRow < 1 Then Exit Function
    lastCol = LastUsedColumn(ws)
    For c = 1 To lastCol
        If StrComp(Trim$(SafeStr(ws.Cells(headerRow, c).Value)), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    Exit Function
LocalFail:
    LogErr "FindHeaderColumn"
End Function


Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range

    On Error GoTo LocalFail
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
    Exit Function
LocalFail:
    LogErr "LastUsedRow"
End Function


Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range

    On Error GoTo LocalFail
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = f.Column
    Exit Function
LocalFail:
    LogErr "LastUsedColumn"
End Function


Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function


Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, x As Variant
    s = Trim$(s)
    bad = Array("/", "\", "?", "*", "[", "]", ":")
    For Each x In bad: s = Replace(s, CStr(x), "-"): Next x
    s = Replace(s, "'", "")
    If Len(s) = 0 Then s = "Unknown Manager"
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function


' Null-safe CStr: returns "" for Null/Empty instead of raising error 94.
Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function


' Case-insensitive manager lookup in mgrNames (1-based index, 0 = not found).
Private Function FindManager(ByVal mgrNames As Collection, ByVal name As String) As Long
    Dim i As Long

    On Error GoTo LocalFail
    For i = 1 To mgrNames.Count
        If StrComp(CStr(mgrNames(i)), name, vbTextCompare) = 0 Then
            FindManager = i
            Exit Function
        End If
    Next i
    Exit Function
LocalFail:
    LogErr "FindManager"
End Function


' Collision-proof tab naming: managers like "A/B" and "A\B" both sanitize to
' "A-B", and a name matching a reserved sheet name would collide with it.
Private Function GetOrCreateOutputSheet(wb As Workbook, sheetName As String, ByVal firstOutput As Boolean) As Worksheet
    Dim ws As Worksheet

    On Error GoTo LocalFail
    If firstOutput And wb.Worksheets.Count = 1 Then
        Set ws = wb.Worksheets(1)
        ws.Name = UniqueSheetName(wb, sheetName)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = UniqueSheetName(wb, sheetName)
    End If
    Set GetOrCreateOutputSheet = ws
    Exit Function
LocalFail:
    LogErr "GetOrCreateOutputSheet"
End Function


Private Function UniqueSheetName(wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, n As Long

    On Error GoTo LocalFail
    candidate = baseName
    n = 1
    Do While SheetExists(candidate, wb)
        n = n + 1
        candidate = Left$(baseName, 31 - Len(" (" & CStr(n) & ")")) & " (" & CStr(n) & ")"
    Loop
    UniqueSheetName = candidate
    Exit Function
LocalFail:
    LogErr "UniqueSheetName"
End Function


' ================== PRIOR REPORT COMPARISON ==================
' Owns "what changed since last week". Reads a prior PM Requests report
' (2.0.18 layout or this builder): every sheet except Dashboard / Raw Data is a
' manager tab; its Name column feeds the prior-name set.
'
'   DetectPriorReport   - weekly mode: with no prior given, the previous week's
'                         report is found automatically in the output folder.
'                         Candidates are the saved "PM Requests *.xlsx" files,
'                         newest first; each is scored by how many of the current
'                         report's names it also contains. The newest high scorer
'                         wins (an older report with the same score never
'                         displaces a newer one), and scoring stops at the six
'                         newest files or at a near-perfect match. "" = no
'                         plausible prior, so the build runs without a
'                         comparison.
'   PriorNameOverlap    - score for one candidate; any error scores 0.
'   SortPathsNewestFirst- candidate order (string compare on the file name).
'   LoadPriorComparison - prior names / manager names / manager counts.
'   IsManagerTab / FindHeaderRowForSingle / NormalizeName / PriorNameExists /
'   CountDiff / PriorMgrCount - the lookup primitives.
'
' Contracts:
'   * "" is a normal outcome, not an error: first run of the season, a moved
'     folder, or a renamed export all land here and the build still succeeds.
'   * An unreadable prior degrades to empty comparison sets. Never raise from
'     this section - a bad prior file must not fail a build.
'   * A plausible prior explains at least max(count/4, 5) of this week's names;
'     at 90% the search stops early.
'

' Weekly mode: with no prior report given, the previous week's report is found
' automatically in the output folder. Candidates are the saved
' "PM Requests *.xlsx" files, newest first; each is scored by how many of the
' current report's names it also contains. The newest high scorer wins (an
' older report with the same score never displaces a newer one), and scoring
' stops at the six newest files or at a near-perfect match. "" = no plausible
' prior, so the build runs without a comparison.
Private Function DetectPriorReport(ByVal outputFolder As String, _
                                   ByVal data As Variant, _
                                   ByVal nameCol As Long, _
                                   ByVal currentLastRow As Long) As String
    Dim paths() As String, n As Long, i As Long
    Dim fileName As String, folder As String
    Dim currentNames As Collection
    Dim bestPath As String, bestScore As Long, score As Long
    Dim threshold As Long, r As Long, nm As String

    On Error GoTo LocalFail

    If Len(outputFolder) = 0 Or nameCol < 1 Then Exit Function
    folder = NormalizeFolder(outputFolder)
    If Not FolderExists(folder) Then Exit Function

    Set currentNames = New Collection
    For r = 2 To currentLastRow
        nm = NormalizeName(data(r, nameCol))
        If Len(nm) > 0 Then currentNames.Add nm
    Next r
    If currentNames.Count = 0 Then Exit Function

    ReDim paths(1 To 1)
    n = 0
    fileName = Dir$(folder & Application.PathSeparator & "PM Requests *.xlsx")
    Do While Len(fileName) > 0
        If InStr(1, fileName, "~$", vbTextCompare) = 0 Then
            n = n + 1
            ReDim Preserve paths(1 To n)
            paths(n) = folder & Application.PathSeparator & fileName
        End If
        fileName = Dir$()
    Loop
    If n = 0 Then Exit Function

    SortPathsNewestFirst paths, n

    ' A prior that explains at least a quarter of the current names (minimum 5)
    ' is treated as plausible; at 90% the search stops early.
    threshold = currentNames.Count / 4
    If threshold < 5 Then threshold = 5

    For i = 1 To n
        If i > 6 Then Exit For
        score = PriorNameOverlap(paths(i), currentNames)
        If score > bestScore Then
            bestScore = score
            bestPath = paths(i)
        End If
        If bestScore >= (currentNames.Count * 9) / 10 Then Exit For
    Next i

    If bestScore < threshold Then Exit Function
    DetectPriorReport = bestPath
    Exit Function
LocalFail:
    LogErr "DetectPriorReport"
End Function


' How many of currentNames also appear in the prior report's name column.
' Any error (unreadable file, no recognizable header) scores 0.
Private Function PriorNameOverlap(ByVal priorPath As String, _
                                  ByVal currentNames As Collection) As Long
    Dim wb As Workbook, ws As Worksheet
    Dim headerRow As Long, nameCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant, priorNames As Collection
    Dim r As Long, missing As Long

    On Error GoTo Fail

    Set wb = Workbooks.Open(Filename:=priorPath, ReadOnly:=True, _
                            UpdateLinks:=False, AddToMru:=False)

    Set priorNames = New Collection

    ' Prefer the builder's Raw Data tab, then any sheet carrying both
    ' required headers. A 2.0.18 prior has neither - it stores one manager
    ' per tab with the Name column (usually C) - so fall back to reading
    ' every manager tab's name column instead of scoring zero.
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(RAW_SHEET)
    On Error GoTo Fail
    If ws Is Nothing Then Set ws = FindSourceWorksheet(wb, headerRow)

    If Not ws Is Nothing Then
        headerRow = FindHeaderRow(ws, MANAGER_HEADER, STATUS_HEADER)
        If headerRow < 1 Then GoTo Fail
        nameCol = FindHeaderColumn(ws, headerRow, NAME_HEADER)
        If nameCol < 1 Then GoTo Fail

        lastRow = LastUsedRow(ws)
        lastCol = LastUsedColumn(ws)
        If lastRow > headerRow Then
            data = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol)).Value2
            For r = 2 To UBound(data, 1)
                priorNames.Add NormalizeName(data(r, nameCol))
            Next r
        End If
    Else
        For Each ws In wb.Worksheets
            If IsManagerTab(ws.Name) Then
                headerRow = FindHeaderRowForSingle(ws, NAME_HEADER)
                If headerRow = 0 Then headerRow = 1
                nameCol = FindHeaderColumn(ws, headerRow, NAME_HEADER)
                If nameCol = 0 Then nameCol = 3
                lastRow = LastUsedRow(ws)
                If lastRow > headerRow Then
                    For r = headerRow + 1 To lastRow
                        priorNames.Add NormalizeName(ws.Cells(r, nameCol).Value)
                    Next r
                End If
            End If
        Next ws
    End If

    missing = CountDiff(currentNames, priorNames)
    PriorNameOverlap = currentNames.Count - missing
    If PriorNameOverlap < 0 Then PriorNameOverlap = 0

Fail:
    LogErr "PriorNameOverlap"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function


' Sorts a 1-based path array by file time, newest first (insertion sort).
Private Sub SortPathsNewestFirst(ByRef paths() As String, ByVal n As Long)
    Dim i As Long, j As Long, key As String
    Dim keyTime As Date, curTime As Date

    On Error GoTo LocalFail
    For i = 2 To n
        key = paths(i)
        keyTime = FileDateTime(paths(i))
        j = i - 1
        Do While j >= 1
            curTime = FileDateTime(paths(j))
            If curTime >= keyTime Then Exit Do
            paths(j + 1) = paths(j)
            j = j - 1
        Loop
        paths(j + 1) = key
    Next i
    Exit Sub
LocalFail:
    LogErr "SortPathsNewestFirst"
End Sub


Private Sub LoadPriorComparison(ByVal priorPath As String, _
                                ByRef priorNames As Collection, _
                                ByRef priorMgrNames As Collection, _
                                ByRef priorMgrCounts As Collection)
    Dim wb As Workbook, ws As Worksheet
    Dim headerRow As Long, nameCol As Long, lastRow As Long, r As Long

    Set priorNames = New Collection
    Set priorMgrNames = New Collection
    Set priorMgrCounts = New Collection

    On Error GoTo Fail
    Set wb = Workbooks.Open(priorPath, ReadOnly:=True, UpdateLinks:=False)

    For Each ws In wb.Worksheets
        If IsManagerTab(ws.Name) Then
            headerRow = FindHeaderRowForSingle(ws, NAME_HEADER)
            If headerRow = 0 Then headerRow = 1
            nameCol = FindHeaderColumn(ws, headerRow, NAME_HEADER)
            If nameCol = 0 Then nameCol = 3
            lastRow = LastUsedRow(ws)
            If lastRow > headerRow Then
                priorMgrNames.Add ws.Name
                priorMgrCounts.Add lastRow - headerRow
                For r = headerRow + 1 To lastRow
                    priorNames.Add NormalizeName(ws.Cells(r, nameCol).Value)
                Next r
            End If
        End If
    Next ws

    wb.Close SaveChanges:=False
    Exit Sub

Fail:
    LogErr "LoadPriorComparison"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ' Comparison degrades to "no prior" instead of failing the build.
End Sub


Private Function IsManagerTab(ByVal sheetName As String) As Boolean
    On Error GoTo LocalFail
    If StrComp(sheetName, DASHBOARD_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, RAW_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, SETTINGS_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, "START HERE", vbTextCompare) = 0 Then Exit Function
    IsManagerTab = True
    Exit Function
LocalFail:
    LogErr "IsManagerTab"
End Function


' Searches the first rows for a lone header name (prior tabs may have
' title rows above the table).
Private Function FindHeaderRowForSingle(ws As Worksheet, ByVal headerName As String) As Long
    Dim r As Long, maxRow As Long

    On Error GoTo LocalFail
    maxRow = LastUsedRow(ws)
    If maxRow > 30 Then maxRow = 30
    For r = 1 To maxRow
        If FindHeaderColumn(ws, r, headerName) > 0 Then
            FindHeaderRowForSingle = r
            Exit Function
        End If
    Next r
    Exit Function
LocalFail:
    LogErr "FindHeaderRowForSingle"
End Function


Private Function NormalizeName(ByVal v As Variant) As String
    NormalizeName = LCase$(Trim$(SafeStr(v)))
End Function


Private Function PriorNameExists(ByVal priorNames As Collection, ByVal nm As String) As Boolean
    Dim i As Long
    For i = 1 To priorNames.Count
        If StrComp(CStr(priorNames(i)), nm, vbTextCompare) = 0 Then
            PriorNameExists = True
            Exit Function
        End If
    Next i
End Function


' Items in a that are not in b (normalized names; a and b hold names
' already passed through NormalizeName).
Private Function CountDiff(ByVal a As Collection, ByVal b As Collection) As Long
    Dim i As Long, n As Long
    For i = 1 To a.Count
        If Not PriorNameExists(b, CStr(a(i))) Then n = n + 1
    Next i
    CountDiff = n
End Function


' Prior PM count for one manager tab; -1 = not found on the prior report.
Private Function PriorMgrCount(ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, ByVal manager As String) As Long
    Dim i As Long
    For i = 1 To priorMgrNames.Count
        If StrComp(CStr(priorMgrNames(i)), manager, vbTextCompare) = 0 Then
            PriorMgrCount = CLng(priorMgrCounts(i))
            Exit Function
        End If
    Next i
    PriorMgrCount = -1
End Function


' ================== OUTPUT FORMATTING ==================
' Owns how one manager tab looks: KPI cards, the table, banding, highlighting.
'
'   FormatManagerSheet          - the whole tab: cards, header, body, filters.
'   CountInSheet                - value count used by the cards.
'   AddPolishedCard / AddPolishedTextCard - KPI card (eyebrow, label, big value);
'                                   ChangeColor derives the text colour from the
'                                   card fill.
'   ApplyReaderFriendlyFormatting - banding, borders, freeze panes, autofilter,
'                                   print setup, number formats.
'
' Contracts:
'   * The cards occupy rows 1-8; the table header is row 9 and data starts at
'     row 10. CORE writes rows with that offset, so both must change together.
'   * New PM rows (names absent from the prior report) are light green + bold.
'     With no prior comparison there are no new rows and nothing is highlighted.
'   * Every formatting call runs under On Error Resume Next: a cosmetic failure
'     must not fail a finished report.
'

' Manager tab: banner row 1, four equal KPI cards on rows 2-4, a full-width
' "PMs by Rep" card on rows 5-7, table header on row 9, data from row 10.
' Rows are banded per employee (column A) rather than per row.
Private Sub FormatManagerSheet(ws As Worksheet, ByVal manager As String, _
                               ByVal priorNames As Collection, _
                               ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, _
                               ByVal nameCol As Long)
    Dim lastCol As Long, lastRow As Long
    Dim priorCount As Long, curCount As Long, newCount As Long, completedCount As Long
    Dim r As Long, isNew As Boolean, bandOn As Boolean
    Dim repNames As Collection
    Dim repText As String, i As Long, j As Long, repName As String, prevRep As String, nm As String
    Dim rNames() As String, rCounts() As Long, tmpC As Long

    On Error GoTo LocalFail

    lastCol = LastUsedColumn(ws): lastRow = LastUsedRow(ws)
    If lastRow < 1 Then Exit Sub
    ' Manager tabs carry only MANAGER_LAST_COL (8) data columns; cap used
    ' range so nothing beyond column 8 can reappear on the tab.
    If lastCol > MANAGER_LAST_COL Then lastCol = MANAGER_LAST_COL

    ' Uniform columns: the four KPI cards below are exactly three columns wide
    ' each, so they come out the same size without disturbing the data.
    ws.Columns("A").ColumnWidth = SHEET_COL_WIDTH
    ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).ColumnWidth = SHEET_COL_WIDTH
    ws.Columns(lastCol).ColumnWidth = SHEET_COL_WIDTH

    ' Table: header on row 9, data rows 10..lastRow.
    ApplyReaderFriendlyFormatting ws, 9, lastRow, lastCol
    If lastRow >= 10 Then SortByColumnA ws, 9, lastRow, lastCol

    ' KPI values.
    priorCount = PriorMgrCount(priorMgrNames, priorMgrCounts, SafeSheetName(manager))
    If priorCount < 0 Then priorCount = 0
    curCount = lastRow - 9
    newCount = 0
    If priorNames.Count > 0 And nameCol > 0 Then
        For r = 10 To lastRow
            nm = NormalizeName(ws.Cells(r, nameCol).Value)
            If Len(nm) > 0 And Not PriorNameExists(priorNames, nm) Then newCount = newCount + 1
        Next r
    End If
    completedCount = priorCount - (curCount - newCount)
    If completedCount < 0 Then completedCount = 0

    ' Banner: "Manager | PM SUMMARY".
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .Value = manager & "  |  PM Summary"
        .Font.Name = UI_FONT
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 34

    ' Four equal KPI cards across the full width (A:H = 8 data columns).
    AddPolishedCard ws, "A2:B4", "PMs This Week", curCount, RGB(221, 235, 247), False
    AddPolishedCard ws, "C2:D4", "Prior Report", priorCount, RGB(242, 242, 242), False, "LAST WEEK"
    AddPolishedCard ws, "E2:F4", "Completed", completedCount, RGB(226, 239, 218), False, "SINCE LAST WEEK"
    AddPolishedCard ws, "G2:H4", "New PMs", newCount, RGB(226, 239, 218), False, "ADDED THIS WEEK"

    ' PMs by Rep card (counts read from the written column A).
    Set repNames = New Collection
    If lastRow >= 10 Then
        For r = 10 To lastRow
            repName = Trim$(SafeStr(ws.Cells(r, 1).Value))
            If Len(repName) > 0 Then
                i = FindManager(repNames, repName)   ' reuse case-insensitive lookup
                If i = 0 Then repNames.Add repName
            End If
        Next r
    End If
    If repNames.Count > 0 Then
        ReDim rNames(1 To repNames.Count)
        ReDim rCounts(1 To repNames.Count)
        For i = 1 To repNames.Count
            rNames(i) = CStr(repNames(i))
            rCounts(i) = CountInSheet(ws, rNames(i), lastRow)
        Next i
        ' Busiest rep first, ties alphabetical: the card reads as a ranking
        ' instead of "whoever appeared first in the export".
        For i = 1 To repNames.Count - 1
            For j = i + 1 To repNames.Count
                If (rCounts(j) > rCounts(i)) Or _
                   ((rCounts(j) = rCounts(i)) And (StrComp(rNames(j), rNames(i), vbTextCompare) < 0)) Then
                    repName = rNames(i): rNames(i) = rNames(j): rNames(j) = repName
                    tmpC = rCounts(i): rCounts(i) = rCounts(j): rCounts(j) = tmpC
                End If
            Next j
        Next i
        repText = ""
        For i = 1 To repNames.Count
            If Len(repText) > 0 Then repText = repText & "   |   "
            repText = repText & rNames(i) & ": " & CStr(rCounts(i))
        Next i
    End If
    AddPolishedTextCard ws, "A5:H7", "Sales Rep Summary", repText, RGB(247, 249, 251)

    ws.Rows(8).RowHeight = 10

    ' Per-employee banding: all rows of one employee share a colour, the next
    ' employee gets the other colour. New PM rows keep their green highlight.
    If lastRow >= 10 Then
        bandOn = True
        prevRep = ""
        For r = 10 To lastRow
            repName = Trim$(SafeStr(ws.Cells(r, 1).Value))
            If StrComp(repName, prevRep, vbTextCompare) <> 0 Then
                bandOn = Not bandOn
                prevRep = repName
            End If
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
                If bandOn Then
                    .Interior.Color = RGB(255, 255, 255)
                Else
                    .Interior.Color = RGB(235, 242, 248)
                End If
            End With

            isNew = False
            If priorNames.Count > 0 And nameCol > 0 Then
                isNew = Not PriorNameExists(priorNames, NormalizeName(ws.Cells(r, nameCol).Value))
            End If
            If isNew Then
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
                    .Interior.Color = RGB(226, 239, 218)
                    .Font.Bold = True
                End With
            End If
        Next r
    End If

    ' Print setup: one page wide, header row repeated.
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$9:$9"
        .CenterHorizontally = True
    End With
    On Error GoTo 0

    ' Clear and hide any columns beyond the 8 data columns so treatment-detail
    ' columns removed from source never appear on manager tabs - not even as
    ' empty formatted cells in UsedRange (which screenshot export reads).
    If MANAGER_LAST_COL < ws.Columns.Count Then
        ws.Columns(MANAGER_LAST_COL + 1).Resize(, ws.Columns.Count - MANAGER_LAST_COL).Clear
        ws.Columns(MANAGER_LAST_COL + 1).Resize(, ws.Columns.Count - MANAGER_LAST_COL).Hidden = True
    End If

    ' Block out every row below the last data row (and re-confirm columns
    ' beyond the data columns are hidden) so the tab shows only the report
    ' - nothing else is visible, on screen or in a screenshot.
    HideUnusedCells ws, lastRow, lastCol
    ' Add Generate Email button at top-right of the tab
    On Error Resume Next
    ws.Shapes("BtnGenEmail").Delete
    On Error GoTo LocalFail
    Dim btnShp As Shape
    Set btnShp = ws.Shapes.AddShape(5, 900, ws.Range("A1").Top + 2, 200, 24)
    btnShp.Name = "BtnGenEmail"
    btnShp.Width = 200
    btnShp.Height = 24
    With btnShp
        .Fill.ForeColor.RGB = RGB(34, 139, 34)
        .Fill.Solid
        .TextFrame.Characters.Text = "Create Email - " & manager
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "'" & ThisWorkbook.Name & "'!GenerateManagerEmail"
    End With

    Exit Sub

LocalFail:
    LogErr "FormatManagerSheet"
End Sub


' Hide every row below the last data row and every column to the right of the
' last data column so a tab shows only the report; nothing else is visible.
Private Sub HideUnusedCells(ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    On Error Resume Next
    If lastRow > 0 And lastRow < ws.Rows.Count Then
        ws.Rows(CStr(lastRow + 1) & ":" & CStr(ws.Rows.Count)).Hidden = True
    End If
    If lastCol > 0 And lastCol < ws.Columns.Count Then
        ws.Columns(CStr(lastCol + 1) & ":" & CStr(ws.Columns.Count)).Hidden = True
    End If
    On Error GoTo 0
End Sub


' Case-insensitive count of a value in a sheet column (rep totals).
Private Function CountInSheet(ws As Worksheet, ByVal value As String, ByVal lastRow As Long) As Long
    Dim r As Long, n As Long

    On Error GoTo LocalFail
    For r = 10 To lastRow
        If StrComp(Trim$(SafeStr(ws.Cells(r, 1).Value)), value, vbTextCompare) = 0 Then n = n + 1
    Next r
    CountInSheet = n
    Exit Function
LocalFail:
    LogErr "CountInSheet"
End Function


' 2.0.18 KPI card: eyebrow + label header strip, big centered value.
Private Sub AddPolishedCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal value As Long, ByVal fillColor As Long, ByVal colorChange As Boolean, Optional ByVal eyebrow As String = "")
    Dim box As Range, hdr As Range, valBox As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    Dim accent As Long

    On Error GoTo LocalFail
    Set box = ws.Range(address)
    firstRow = box.Row: firstCol = box.Column
    lastRow = box.Row + box.Rows.Count - 1: lastCol = box.Column + box.Columns.Count - 1
    box.UnMerge
    Set hdr = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(firstRow, lastCol))
    Set valBox = ws.Range(ws.Cells(firstRow + 1, firstCol), ws.Cells(lastRow, lastCol))
    hdr.Merge: valBox.Merge

    If colorChange Then accent = ChangeColor(value) Else accent = RGB(31, 78, 121)

    hdr.Value = UCase$(label)
    If Len(eyebrow) > 0 Then hdr.Value = UCase$(eyebrow) & " " & ChrW(183) & " " & UCase$(label)
    valBox.Value = CStr(value)

    With box
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(210, 218, 226)
    End With
    With hdr
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = False
        .Font.Name = UI_FONT
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = accent
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    With valBox
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = False
        .Font.Name = UI_FONT
        .Font.Bold = True: .Font.Size = 24: .Font.Color = RGB(31, 45, 61)
        .Interior.Color = fillColor
    End With
    If colorChange Then valBox.Font.Color = ChangeColor(value)

    ws.Rows(firstRow).RowHeight = 20
    ws.Rows(firstRow + 1).RowHeight = 24
    If lastRow > firstRow + 1 Then ws.Rows((firstRow + 1) & ":" & lastRow).RowHeight = 24
    Exit Sub
LocalFail:
    LogErr "AddPolishedCard"
End Sub


' 2.0.18 text card: eyebrow + label header strip, wrapped text value.
Private Sub AddPolishedTextCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal valueText As String, ByVal fillColor As Long, Optional ByVal eyebrow As String = "")
    Dim box As Range, hdr As Range, valBox As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long

    On Error GoTo LocalFail
    Set box = ws.Range(address)
    firstRow = box.Row: firstCol = box.Column
    lastRow = box.Row + box.Rows.Count - 1: lastCol = box.Column + box.Columns.Count - 1
    box.UnMerge
    Set hdr = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(firstRow, lastCol))
    Set valBox = ws.Range(ws.Cells(firstRow + 1, firstCol), ws.Cells(lastRow, lastCol))
    hdr.Merge: valBox.Merge

    hdr.Value = UCase$(label)
    If Len(eyebrow) > 0 Then hdr.Value = UCase$(eyebrow) & vbLf & UCase$(label)
    valBox.Value = valueText

    With box
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(210, 218, 226)
    End With
    With hdr
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Name = UI_FONT
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(91, 110, 126)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    With valBox
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Name = UI_FONT
        .Font.Bold = False: .Font.Size = 9: .Font.Color = RGB(31, 45, 61)
        .Interior.Color = fillColor
    End With

    ws.Rows(firstRow).RowHeight = 30
    ws.Rows(firstRow + 1).RowHeight = 24
    If lastRow > firstRow + 1 Then ws.Rows((firstRow + 1) & ":" & lastRow).RowHeight = 24
    Exit Sub
LocalFail:
    LogErr "AddPolishedTextCard"
End Sub


' Green above target, red below (2.0.18 semantics: >0 green, <0 red).
Private Function ChangeColor(ByVal v As Long) As Long
    If v > 0 Then
        ChangeColor = RGB(0, 128, 0)
    ElseIf v < 0 Then
        ChangeColor = RGB(192, 0, 0)
    Else
        ChangeColor = RGB(31, 78, 121)
    End If
End Function


' Reader-friendly formatting for a data table. headerRow = the table's header
' row (manager tabs use row 9 under the KPI cards).
' Borders, header style, filter and frozen header. Row banding is left to
' FormatManagerSheet, which bands by employee instead of by row.
Private Sub ApplyReaderFriendlyFormatting(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim dataRange As Range

    On Error GoTo LocalFail
    Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
    dataRange.Borders.LineStyle = xlContinuous
    dataRange.Borders.Weight = xlHairline
    dataRange.VerticalAlignment = xlCenter
    dataRange.WrapText = True
    dataRange.Font.Name = UI_FONT

    ' Strong, easy-to-scan header.
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        If lastRow > headerRow Then
            .Font.Name = UI_FONT
            .Font.Size = ws.Cells(headerRow + 1, 1).Font.Size
        End If
    End With
    ws.Rows(headerRow).RowHeight = 34

    ' Data rows keep the automatic height so wrapped cells stay readable.
    If lastRow > headerRow Then
        On Error Resume Next
        ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, lastCol)).EntireRow.AutoFit
        On Error GoTo 0
    End If

    ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol)).AutoFilter

    ' Keep the header visible while scrolling (best-effort: invisible
    ' automation has no usable window for freeze panes).
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(headerRow + 1, 1)).Select
        ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
    Exit Sub
LocalFail:
    LogErr "ApplyReaderFriendlyFormatting"
End Sub


' ==================== DASHBOARD ====================
' Owns the report's summary sheet: a fixed A:L grid, rebuilt every run.
'
' Weekly summary on a fixed A:L grid: logo + title band, a meta strip,
' six KPI cards (two rows of three), a manager table with data bars and a
' TOTAL row, a two-block sales-rep grid with prior-report counts, a footnote,
' and a bar chart of PMs per manager below it.
' Prior columns show 0 when no prior report was compared.
'
'   BuildDashboard          - the whole sheet (title, meta, cards, tables,
'                             total row, chart). CORE moves it to index 1.
'   ColorReportTabs         - navy summary, light blue managers, grey raw data.
'   AddManagerChart         - clustered column of PMs per manager; deletes any
'                             previous chart first (the sheet is reused).
'   WriteTableHeader        - one header row in the builder's navy/white style.
'   GetOrCreateDashboardSheet - get or create, then Cells.Clear.
'   ResolveLogoPath / AddDashboardLogo / PositionDashboardLogo - logo art.
'   SortByColumnA           - name order inside the dashboard tables.
'
' Contracts:
'   * Row plan: cardRow = 4, hdrRow = cardRow + 9 = 13. Everything below the
'     cards is derived from hdrRow, so a layout change is one number.
'   * Six cards, three per row, each four columns wide: A:D, E:H, I:L. The
'     column widths are chosen so those three groups are equal (46 characters
'     each), which is what makes the cards the same size.
'   * The footnote sits below whichever table is taller, and the chart starts two
'     rows under it, so neither can land on top of the other. The chart width is
'     the table width (columns A:L), not a fixed number: a fixed width overhangs
'     column L and widens the printed page.
'   * One typeface (UI_FONT) for the whole sheet; without it the dashboard
'     inherits Excel's default font and the workbook shows two typefaces.
'   * The TOTAL row is the sum of the manager rows, so the table has to agree
'     with the PMs card.
'   * The meta strip names the report and the prior it was compared to; both
'     come from CORE (savePath and priorPath), so a rename there shows up here.
'   * Logo lookup order: documentation\aoti-logo.png, then the workbook root,
'     then LOGO_FILE as a bare relative name. All three are tried with Dir(),
'     and a missing logo is not fatal - the sheet just has no picture.
'   * The sheet is reused, not recreated: Cells.Clear plus chart deletion must
'     run before drawing, or a second run stacks shapes.
'
Private Sub BuildDashboard(wb As Workbook, ByVal managers As Variant, _
                           ByVal mgrRows As Collection, ByVal mgrNames As Collection, _
                           ByVal data As Variant, ByVal nameCol As Long, _
                           ByVal pmTotal As Long, ByVal newPMs As Long, _
                           ByVal priorNames As Collection, _
                           ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, _
                           ByVal priorPMs As Long, _
                           ByVal repNames As Collection, repCounts() As Long, ByVal repN As Long, _
                           ByVal currentReportName As String, ByVal priorPath As String)
    Dim ws As Worksheet
    Dim r As Long, i As Long, idx As Long
    Dim rowsForManager As Collection
    Dim priorCount As Long, mgrNew As Long, mgrCompleted As Long, mgrOpen As Long
    Dim completedTotal As Long, retainedPMs As Long
    Dim repEnd As Long, repRows As Long, repCount As Long, repGroup As Long
    Dim repRowOffset As Long, repBaseCol As Long
    Dim sortNames() As String, sortCounts() As Long, sortPrior() As Long
    Dim tS As String, tC As Long, tP As Long, j As Long, nm As String
    Dim mgrEnd As Long
    Dim hdrRow As Long, cardRow As Long
    Dim totalRow As Long, sumCol As Long
    Dim footRow As Long

    On Error GoTo LocalFail

    ' Row plan, in one place: two card rows, then the tables. Everything below
    ' the cards is derived from hdrRow so a layout change is one number.
    cardRow = 4
    hdrRow = cardRow + 9

    Set ws = GetOrCreateDashboardSheet(wb)

    ' Title band: white logo zone (A:B) + navy title (C:L).
    With ws.Range("A1:B1")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Range("C1:L1")
        .Merge
        .Value = "PM Request Summary"
        .Font.Bold = True
        .Font.Size = 20
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 34
    ws.Tab.Color = RGB(31, 78, 121)

    ' Meta strip: which file this is, which file it is compared to, when built.
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = "Report: " & currentReportName
    ws.Range("E2:H2").Merge
    If Len(priorPath) > 0 Then
        ws.Range("E2").Value = "Prior Report: " & Mid$(priorPath, InStrRev(priorPath, Application.PathSeparator) + 1)
    Else
        ws.Range("E2").Value = "Prior Report: (none)"
    End If
    ws.Range("I2:L2").Merge
    ws.Range("I2").Value = "Built: " & Format(Now, "mm/dd/yyyy h:mm AM/PM")
    With ws.Range("A2:L2")
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = RGB(235, 242, 248)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(2).RowHeight = 26

    ' KPI cards: two rows of three. Each card spans four columns, and the column
    ' widths below are tuned so A:D, E:H and I:L are the same width - so all six
    ' cards are identical without disturbing the two tables underneath.
    AddPolishedCard ws, "A4:D6", "PMs", pmTotal, RGB(221, 235, 247), False, "THIS WEEK"
    AddPolishedCard ws, "E4:H6", "PMs", priorPMs, RGB(242, 242, 242), False, "LAST WEEK"
    If priorPMs > 0 Then
        completedTotal = priorPMs - (pmTotal - newPMs)
        If completedTotal < 0 Then completedTotal = 0
    Else
        completedTotal = 0
    End If
    AddPolishedCard ws, "I4:L6", "Closed", completedTotal, RGB(226, 239, 218), False, "SINCE LAST WEEK"

    ' Without a prior report there is nothing carried over.
    If priorPMs > 0 Then
        retainedPMs = pmTotal - newPMs
        If retainedPMs < 0 Then retainedPMs = 0
    Else
        retainedPMs = 0
    End If
    AddPolishedCard ws, "A8:D10", "New", newPMs, RGB(226, 239, 218), False, "THIS WEEK"
    AddPolishedCard ws, "E8:H10", "Reps", repN, RGB(221, 235, 247), False, "ACTIVE"
    AddPolishedCard ws, "I8:L10", "Open", retainedPMs, _
                      IIf(retainedPMs = 0, RGB(226, 239, 218), RGB(244, 204, 204)), False, "CARRIED OVER"
    If retainedPMs = 0 Then ws.Range("I9:L10").Font.Color = RGB(0, 97, 0) Else ws.Range("I9:L10").Font.Color = RGB(156, 0, 6)

    ws.Rows(cardRow - 1).RowHeight = 8
    ws.Rows(cardRow + 3).RowHeight = 8
    ws.Rows(cardRow + 7).RowHeight = 8

    ' Section labels above the two tables.
    With ws.Range(ws.Cells(hdrRow - 1, 1), ws.Cells(hdrRow - 1, 6))
        .Merge
        .Value = "Sales Manager Summary"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range(ws.Cells(hdrRow - 1, 7), ws.Cells(hdrRow - 1, 12))
        .Merge
        .Value = "Sales Rep Summary"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With

    ' Table header row.
    r = hdrRow
    WriteTableHeader ws, r, 1, Array("Sales Manager", "PMs", "Prior", "Completed", "New", "Still Open")

    ' Rep grid: two blocks of three columns (G:I, J:L) with the prior count.
    If repN > 0 Then
        ReDim sortNames(1 To repN)
        ReDim sortCounts(1 To repN)
        ReDim sortPrior(1 To repN)
        For i = 1 To repN
            sortNames(i) = CStr(repNames(i))
            sortCounts(i) = repCounts(i)
            sortPrior(i) = 0
        Next i

        ' Prior counts: how many of each rep's current PMs were on the prior report.
        If priorNames.Count > 0 And nameCol > 0 Then
            If IsArray(managers) Then
                For i = LBound(managers) To UBound(managers)
                    Set rowsForManager = mgrRows(FindManager(mgrNames, CStr(managers(i))))
                    For idx = 1 To rowsForManager.Count
                        j = FindManager(repNames, Trim$(SafeStr(data(rowsForManager(idx), 1))))
                        If j > 0 Then
                            If PriorNameExists(priorNames, NormalizeName(data(rowsForManager(idx), nameCol))) Then sortPrior(j) = sortPrior(j) + 1
                        End If
                    Next idx
                Next i
            End If
        End If

        For i = 1 To repN - 1
            For j = i + 1 To repN
                If StrComp(sortNames(j), sortNames(i), vbTextCompare) < 0 Then
                    tS = sortNames(i): sortNames(i) = sortNames(j): sortNames(j) = tS
                    tC = sortCounts(i): sortCounts(i) = sortCounts(j): sortCounts(j) = tC
                    tP = sortPrior(i): sortPrior(i) = sortPrior(j): sortPrior(j) = tP
                End If
            Next j
        Next i

        repCount = repN
        repRows = (repCount + 1) \ 2
        If repRows < 1 Then repRows = 1
        For repGroup = 0 To 1
            repBaseCol = 7 + (repGroup * 3)
            WriteTableHeader ws, r, repBaseCol, Array("Sales Rep", "PMs", "Prior")
        Next repGroup
        For i = 1 To repCount
            repGroup = (i - 1) \ repRows
            repRowOffset = (i - 1) Mod repRows
            repBaseCol = 7 + (repGroup * 3)
            ws.Cells(r + 1 + repRowOffset, repBaseCol).Value = sortNames(i)
            ws.Cells(r + 1 + repRowOffset, repBaseCol + 1).Value = sortCounts(i)
            ws.Cells(r + 1 + repRowOffset, repBaseCol + 2).Value = sortPrior(i)
        Next i
        repEnd = r + repRows
        For repGroup = 0 To 1
            repBaseCol = 7 + (repGroup * 3)
            With ws.Range(ws.Cells(r, repBaseCol), ws.Cells(repEnd, repBaseCol + 2))
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .VerticalAlignment = xlCenter
            End With
            ws.Range(ws.Cells(r + 1, repBaseCol + 1), ws.Cells(repEnd, repBaseCol + 2)).HorizontalAlignment = xlCenter
        Next repGroup
    Else
        WriteTableHeader ws, r, 7, Array("Sales Rep", "PMs", "Prior")
        repEnd = r
    End If

    ' Manager rows.
    If IsArray(managers) Then
        For i = LBound(managers) To UBound(managers)
            r = r + 1
            idx = FindManager(mgrNames, CStr(managers(i)))
            Set rowsForManager = mgrRows(idx)
            priorCount = PriorMgrCount(priorMgrNames, priorMgrCounts, SafeSheetName(CStr(managers(i))))
            If priorCount < 0 Then priorCount = 0

            mgrNew = 0
            If priorNames.Count > 0 And nameCol > 0 Then
                For idx = 1 To rowsForManager.Count
                    nm = NormalizeName(data(rowsForManager(idx), nameCol))
                    If Len(nm) > 0 And Not PriorNameExists(priorNames, nm) Then mgrNew = mgrNew + 1
                Next idx
            End If
            mgrCompleted = priorCount - (rowsForManager.Count - mgrNew)
            If mgrCompleted < 0 Then mgrCompleted = 0
            mgrOpen = rowsForManager.Count - mgrNew
            If mgrOpen < 0 Then mgrOpen = 0

            ws.Cells(r, 1).Value = CStr(managers(i))
            ws.Cells(r, 2).Value = rowsForManager.Count
            ws.Cells(r, 3).Value = priorCount
            ws.Cells(r, 4).Value = mgrCompleted
            ws.Cells(r, 5).Value = mgrNew
            ws.Cells(r, 6).Value = mgrOpen
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", _
                               SubAddress:="'" & SafeSheetName(CStr(managers(i))) & "'!A1", _
                               TextToDisplay:=CStr(managers(i))
            ws.Cells(r, 4).Font.Color = RGB(0, 97, 0)
            If mgrOpen > 0 Then ws.Cells(r, 6).Font.Color = RGB(156, 0, 6)
        Next i
    End If
    mgrEnd = r

    ' Table styling.
    ws.Rows(mgrEnd).RowHeight = 26
    If mgrEnd > hdrRow Then
        For i = hdrRow + 1 To mgrEnd
            If i Mod 2 = 1 Then ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).Interior.Color = RGB(247, 249, 251)
        Next i
        ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(mgrEnd, 6)).VerticalAlignment = xlCenter
        With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(mgrEnd, 6))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        For i = hdrRow + 1 To mgrEnd
            ws.Range(ws.Cells(i, 2), ws.Cells(i, 6)).HorizontalAlignment = xlCenter
        Next i
    End If

    ' Data bars on the manager PM column (visual weight, no extra width).
    If mgrEnd > hdrRow + 1 Then
        On Error Resume Next
        With ws.Range(ws.Cells(hdrRow + 1, 2), ws.Cells(mgrEnd, 2))
            .FormatConditions.Delete
            .FormatConditions.AddDatabar
            With .FormatConditions(1)
                .BarFill.Type = xlDataBarFillSolid
                .BarFill.Color = RGB(91, 155, 213)
                .MinPoint.Modify newtype:=xlConditionValueLowestValue
                .MaxPoint.Modify newtype:=xlConditionValueHighestValue
            End With
        End With
        On Error GoTo 0
    End If

    ' Total row: the table has to add up to the PMs card without a calculator.
    If mgrEnd > hdrRow Then
        totalRow = mgrEnd + 1
        ws.Cells(totalRow, 1).Value = "TOTAL"
        For sumCol = 2 To 6
            ws.Cells(totalRow, sumCol).Value = Application.WorksheetFunction.Sum( _
                ws.Range(ws.Cells(hdrRow + 1, sumCol), ws.Cells(mgrEnd, sumCol)))
        Next sumCol
        With ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 6))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(totalRow, 2), ws.Cells(totalRow, 6)).HorizontalAlignment = xlCenter
    End If

    ' Chart: PMs per manager, below the tables (keeps the A:L width).
    AddManagerChart ws, hdrRow, mgrEnd, repEnd

    ' Footnote: every comparison column, in one line. It goes below BOTH tables,
    ' not below the rep grid: with fewer reps than managers the rep grid ends
    ' first and repEnd + 2 lands inside the manager table, where the merge would
    ' wipe a manager row.
    footRow = repEnd
    If mgrEnd > footRow Then footRow = mgrEnd
    footRow = footRow + 2
    With ws.Range(ws.Cells(footRow, 1), ws.Cells(footRow, 12))
        .Merge
        .Value = "Prior = PMs on this manager's tab last week. Completed = prior PMs no longer open. " & _
                 "New = PMs not on last week's report. Still Open = carried over from last week (target 0)."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlLeft
    End With

    ' Column widths (fixed A:L footprint). A:D, E:H and I:L are equal (46
    ' characters each) so the six KPI cards come out the same size.
    ws.Columns("A").ColumnWidth = 21
    ws.Columns("B").ColumnWidth = 8: ws.Columns("C").ColumnWidth = 8: ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 9: ws.Columns("F").ColumnWidth = 11
    ws.Columns("G").ColumnWidth = 18: ws.Columns("H").ColumnWidth = 8: ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 18: ws.Columns("K").ColumnWidth = 8: ws.Columns("L").ColumnWidth = 8
    ws.Range(ws.Cells(cardRow, 1), ws.Cells(cardRow + 6, 12)).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(cardRow, 1), ws.Cells(cardRow + 6, 12)).WrapText = True

    ' One typeface for the whole sheet. Without this the dashboard inherits
    ' Excel's default font while START HERE uses Segoe UI, so a reader sees two
    ' different typefaces in the same workbook.
    ws.Cells.Font.Name = UI_FONT

    ' Print setup: one page wide, landscape.
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    On Error GoTo 0

    AddDashboardLogo ws

    ' Dashboard canvas: no gridlines, no row/column headings, and no horizontal
    ' scrollbar - that hands the freed strip to the sheet tabs so more of them
    ' are visible at once.
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ws.Range("A1").Select

    ' Block out everything beyond the used data so the dashboard shows only
    ' the report, nothing else.
    Dim dLastRow As Long, dLastCol As Long
    dLastRow = LastUsedRow(ws): dLastCol = LastUsedColumn(ws)
    HideUnusedCells ws, dLastRow, dLastCol

    Exit Sub
LocalFail:
    LogErr "BuildDashboard"
End Sub


' Tab colours: navy for the summary, light blue for managers (dark text stays
' legible), grey for the raw data.
Private Sub ColorReportTabs(wb As Workbook)
    Dim ws As Worksheet

    On Error GoTo LocalFail
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, DASHBOARD_SHEET, vbTextCompare) = 0 Then
            ws.Tab.Color = RGB(31, 78, 121)
        ElseIf StrComp(ws.Name, RAW_SHEET, vbTextCompare) = 0 Then
            ws.Tab.Color = RGB(191, 191, 191)
        Else
            ws.Tab.Color = RGB(157, 195, 230)
        End If
    Next ws
    Exit Sub
LocalFail:
    LogErr "ColorReportTabs"
End Sub


' Clustered column chart of PMs per manager, placed below the tables.
Private Sub AddManagerChart(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal repEnd As Long)
    Dim co As ChartObject, ch As Chart
    Dim topRow As Long, leftPos As Double, topPos As Double, tableW As Double

    On Error GoTo LocalFail

    ' Remove any chart from a previous run (the sheet is reused).
    On Error Resume Next
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    On Error GoTo LocalFail

    If lastRow <= headerRow Then Exit Sub

    ' The footnote sits on the row after the tables, so the chart starts below it.
    topRow = IIf(repEnd > lastRow, repEnd, lastRow) + 4
    leftPos = ws.Range("A1").Left
    topPos = ws.Range(ws.Cells(topRow, 1), ws.Cells(topRow + 1, 1)).Top
    ' Chart width follows the table, not a fixed number: a fixed width overhangs
    ' column L and widens the printed page.
    tableW = ws.Range(ws.Cells(1, 1), ws.Cells(1, 12)).Width

    Set co = ws.ChartObjects.Add(leftPos, topPos, tableW, 280)
    Set ch = co.Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, 2))
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.ChartTitle.Text = "PMs by Sales Manager"
    On Error Resume Next
    With ch.ChartTitle.Font
        .Name = "Segoe UI"
        .Size = 12
        .Bold = True
        .Color = RGB(31, 78, 121)
    End With
    With ch.PlotArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
    End With
    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
    End With
    With ch.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(91, 155, 213)
    End With
    ch.HasLegend = False
    With ch.Axes(xlValue)
        .HasMajorGridlines = True
        .MinimumScale = 0
        .TickLabels.Font.Size = 9
    End With
    With ch.Axes(xlCategory)
        .TickLabels.Font.Size = 9
    End With
    On Error GoTo LocalFail
    Exit Sub
LocalFail:
    LogErr "AddManagerChart"
End Sub


Private Sub WriteTableHeader(ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal vals As Variant)
    Dim i As Long

    On Error GoTo LocalFail
    For i = LBound(vals) To UBound(vals): ws.Cells(rowNum, colNum + i).Value = vals(i): Next i
    With ws.Range(ws.Cells(rowNum, colNum), ws.Cells(rowNum, colNum + UBound(vals) - LBound(vals)))
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(31, 78, 121): .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(rowNum).RowHeight = 28
    Exit Sub
LocalFail:
    LogErr "WriteTableHeader"
End Sub


Private Function GetOrCreateDashboardSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(DASHBOARD_SHEET)
    On Error GoTo LocalFail
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = DASHBOARD_SHEET
    End If
    ws.Cells.Clear
    Set GetOrCreateDashboardSheet = ws
    Exit Function
LocalFail:
    LogErr "GetOrCreateDashboardSheet"
End Function


' Logo art lives with the docs: documentation\aoti-logo.png next to the workbook.
' Falls back to the workbook root so older folders keep working.
Private Function ResolveLogoPath() As String
    Dim base As String, p As String

    On Error GoTo LocalFail
    base = ThisWorkbook.Path
    If Len(base) = 0 Then ResolveLogoPath = "": Exit Function

    p = base & Application.PathSeparator & LOGO_DIR & Application.PathSeparator & LOGO_FILE
    If Dir(p) <> "" Then ResolveLogoPath = p: Exit Function

    p = base & Application.PathSeparator & LOGO_FILE
    If Dir(p) <> "" Then ResolveLogoPath = p: Exit Function

    ResolveLogoPath = base & Application.PathSeparator & LOGO_DIR & Application.PathSeparator & LOGO_FILE
    Exit Function
LocalFail:
    LogErr "ResolveLogoPath"
End Function


' Adds the company logo to the dashboard's white logo zone (A1:B1).
' Prefers documentation\aoti-logo.png beside the workbook (works headless, no
' clipboard); falls back to the workbook root, then to copying the embedded
' AOTI_Logo shape from START HERE when Excel is visible. Any existing logo on the
' sheet is removed first, so re-runs stay idempotent. Silent no-op when no source
' is available.
Private Sub AddDashboardLogo(ws As Worksheet)
    Dim src As Shape, dst As Shape
    Dim logoH As Double, logoW As Double, logoPath As String
    Dim oldSU As Boolean

    On Error GoTo LocalFail

    logoH = 30
    oldSU = Application.ScreenUpdating
    Application.ScreenUpdating = True

    ' Remove any logo already on this dashboard.
    ws.Shapes(LOGO_SHAPE).Delete

    ' 1) Load from documentation\aoti-logo.png (automation-safe).
    logoPath = ResolveLogoPath()
    Set dst = ws.Shapes.AddPicture(Filename:=logoPath, LinkToFile:=False, _
                                   SaveWithDocument:=True, Left:=0, Top:=0, Width:=100, Height:=logoH)
    If Not dst Is Nothing Then
        dst.Name = LOGO_SHAPE
        logoW = dst.Width
        PositionDashboardLogo ws, dst, logoH, logoW
        Application.ScreenUpdating = oldSU
        Exit Sub
    End If

    ' 2) Clipboard copy of the START HERE logo (visible Excel only).
    If Not Application.Visible Then
        Application.ScreenUpdating = oldSU
        Exit Sub
    End If

    Set src = ThisWorkbook.Worksheets(START_SHEET).Shapes(LOGO_SHAPE)
    If src Is Nothing Then
        Application.ScreenUpdating = oldSU
        Exit Sub
    End If

    src.Copy
    ws.Paste
    Set dst = ws.Shapes(ws.Shapes.Count)
    If dst Is Nothing Then
        Application.ScreenUpdating = oldSU
        Exit Sub
    End If

    dst.Name = LOGO_SHAPE
    If src.Height > 0 Then logoW = src.Width * (logoH / src.Height) Else logoW = 105
    dst.Height = logoH
    dst.Width = logoW
    PositionDashboardLogo ws, dst, logoH, logoW
    Application.ScreenUpdating = oldSU
    Exit Sub
LocalFail:
    Application.ScreenUpdating = oldSU
    LogErr "AddDashboardLogo"
End Sub


' Top-left placement inside the white A1:B1 logo zone, vertically centered.
Private Sub PositionDashboardLogo(ws As Worksheet, shp As Shape, ByVal logoH As Double, ByVal logoW As Double)
    On Error GoTo LocalFail
    shp.Left = ws.Range("A1").Left + 6
    shp.Top = ws.Range("A1").Top + (ws.Rows(1).Height - logoH) / 2
    shp.Placement = xlMove
    Exit Sub
LocalFail:
    LogErr "PositionDashboardLogo"
End Sub


Private Sub SortByColumnA(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    On Error GoTo LocalFail
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    Exit Sub
LocalFail:
    LogErr "SortByColumnA"
End Sub
' ===========================================================================
' EMAIL - builds an Outlook email from the current report data
' ===========================================================================
' Contracts:
'   * CreateEmailFromReport(outWb) - creates a new Outlook mail item, copies
'     the dashboard data range with source formatting, pastes into the body,
'     sets subject to "Annual Equipment Preventative Maintenance MM/DD/YYYY",
'     and displays (does not send) so the user can add recipients.
'   * Uses late binding so no reference to Outlook is required.
' ===========================================================================

Public Sub CreateEmailFromReport(outWb As Workbook)
    On Error GoTo LocalFail

    Dim olApp As Object, olMail As Object
    Dim ws As Worksheet
    Dim rng As Range
    Dim wdDoc As Object, wdSel As Object
    Dim subjectDate As String
    Dim templateText As String
    Dim ccEmails As String
    Dim toEmails As String
    Dim i As Long

    subjectDate = Format$(Date, "mm/dd/yyyy")
    templateText = GetEmailTemplate()
    ccEmails = GetCCEmails()
    toEmails = GetToEmails()

    ' Find the dashboard sheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = outWb.Sheets("Dashboard")
    On Error GoTo LocalFail
    If ws Is Nothing Then
        For i = 1 To outWb.Sheets.Count
            If outWb.Sheets(i).Name <> "Raw Data" And outWb.Sheets(i).Name <> "START HERE" Then
                Set ws = outWb.Sheets(i)
                Exit For
            End If
        Next i
    End If
    If ws Is Nothing Then
        LogLine "CreateEmailFromReport: no suitable sheet found"
        Exit Sub
    End If

    Set rng = ws.UsedRange
    If rng Is Nothing Then
        LogLine "CreateEmailFromReport: empty sheet"
        Exit Sub
    End If

    rng.Copy

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .Subject = "Annual Equipment Preventative Maintenance " & subjectDate
        If Len(toEmails) > 0 Then .To = toEmails
        If Len(ccEmails) > 0 Then .CC = ccEmails
        .Display

        Set wdDoc = .GetInspector.WordEditor
        Set wdSel = wdDoc.Application.Selection

        ' Write template text
        If Len(templateText) > 0 Then
            wdSel.TypeText templateText
            wdSel.TypeParagraph
            wdSel.TypeParagraph
        End If

        ' Paste Excel data with source formatting
        wdSel.Paste
    End With

    LogLine "CreateEmailFromReport: email created and displayed"
    On Error Resume Next
    ThisWorkbook.Sheets("Manager Email Template").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Director Email Template").Visible = xlSheetHidden
    On Error GoTo LocalFail
    Exit Sub

LocalFail:
    LogErr "CreateEmailFromReport"
End Sub
' ===========================================================================
' EMAIL TEMPLATE - stored inside the workbook (hidden sheet)
' ===========================================================================
' Contracts:
'   * EmailTemplate sheet holds the template body text in A1 (merged).
'   * Settings sheet holds CC emails in B2 (comma-separated).
'   * EditEmailTemplate() - prompts user to edit the template text.
'   * EditCCEmails() - prompts user to edit the CC email list.
'   * CreateEmailFromReport() reads template + CC from the workbook.
' ===========================================================================

Private Function GetEmailTemplate() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Director Email Template")
    On Error GoTo LocalFail
    If ws Is Nothing Then
        ' Create the sheet
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Director Email Template"
        ws.Visible = xlSheetVisible
        ws.Range("A1").Value = "Hi" & vbCrLf & vbCrLf & _
            "Our records indicate that the following patient's equipment has been on an " & _
            "active treatment plan over 10 months. Please submit an Annual Equipment " & _
            "Preventative Maintenance Form in GoCanvas and perform a PM on this equipment " & _
            "within one week of this email." & vbCrLf & vbCrLf & _
            "* if the equipment was purchased no PM required let me know to remove from my file." & vbCrLf & _
            "* If an exchange of equipment has taken place, please let us know so we can update the PM due date." & vbCrLf & _
            "* If it is in pending pickup status once it is picked up it will drop off my list." & vbCrLf & _
            "* If you think you have completed this for any of the patients listed below let me know and I will double check." & vbCrLf
        ws.Range("A1").WrapText = True
    End If
    GetEmailTemplate = CStr(ws.Range("A1").Value)
    Exit Function
LocalFail:
    LogErr "GetEmailTemplate"
    GetEmailTemplate = ""
End Function

Private Function GetCCEmails() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Manager Email Template")
    On Error GoTo LocalFail
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Manager Email Template"
        ws.Visible = xlSheetVisible
        ws.Range("A1").Value = "CC Emails (comma-separated):"
        ws.Range("D7").Value = ""
    End If
    GetCCEmails = CStr(ws.Range("D7").Value)
    Exit Function
LocalFail:
    LogErr "GetCCEmails"
    GetCCEmails = ""
End Function

Public Sub EditCCEmails()
    On Error GoTo LocalFail
    ThisWorkbook.Sheets("Manager Email Template").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Manager Email Template").Activate
    ThisWorkbook.Sheets("Manager Email Template").Range("B1").Select
    Exit Sub
LocalFail:
    LogErr "EditCCEmails"
End Sub

Public Sub EditToEmails()
    On Error GoTo LocalFail
    ThisWorkbook.Sheets("Manager Email Template").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Manager Email Template").Activate
    ThisWorkbook.Sheets("Manager Email Template").Range("B3").Select
    Exit Sub
LocalFail:
    LogErr "EditToEmails"
End Sub

Private Function GetToEmails() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Manager Email Template")
    GetToEmails = CStr(ws.Range("D5").Value)
    Exit Function
LocalFail:
    LogErr "GetToEmails"
    GetToEmails = ""
End Function

Private Function GetEmailSubject() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Manager Email Template")
    Dim subj As String
    subj = CStr(ws.Range("D9").Value)
    If Len(subj) = 0 Then subj = "PM Requests - {{DATE}}"
    GetEmailSubject = subj
    Exit Function
LocalFail:
    LogErr "GetEmailSubject"
    GetEmailSubject = "PM Requests - {{DATE}}"
End Function

Private Function GetEmailBody() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Manager Email Template")
    GetEmailBody = CStr(ws.Range("D11").Value)
    Exit Function
LocalFail:
    LogErr "GetEmailBody"
    GetEmailBody = ""
End Function


Public Sub GenerateManagerEmail()
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No sheet selected.", vbExclamation
        Exit Sub
    End If

    Dim olApp As Object, olMail As Object
    Dim templateHtml As String
    Dim ccEmails As String
    Dim imagePath As String
    Dim subjectDate As String
    Dim htmlBody As String
    Dim dataHtml As String
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim cellVal As String

    subjectDate = Format(Date, "mm/dd/yyyy")

    ' Read from Manager Email Template
    Dim wsMgrTemplate As Worksheet
    Set wsMgrTemplate = ThisWorkbook.Sheets("Manager Email Template")
    Dim emailSubject As String
    Dim emailBody As String
    emailSubject = CStr(wsMgrTemplate.Range("D9").Value)
    emailBody = CStr(wsMgrTemplate.Range("D11").Value)
    emailSubject = Replace(emailSubject, "{{DATE}}", subjectDate)
    ccEmails = CStr(wsMgrTemplate.Range("D7").Value)
    Dim toEmails As String
    toEmails = CStr(wsMgrTemplate.Range("D5").Value)
    imagePath = ""  ' Use EmailSettings image path if needed

    templateHtml = emailBody
    ' Convert plain text body to HTML, replacing {{REPORT}} with data table
    templateHtml = Replace(templateHtml, vbCrLf, "<br>")
    templateHtml = Replace(templateHtml, "{{REPORT}}", "{{REPORT_PLACEHOLDER}}")

    ' Build HTML table from the manager tab data
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count
    dataHtml = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:10pt'>"
    Dim i As Long, j As Long
    For i = 1 To lastRow
        dataHtml = dataHtml & "<tr>"
        For j = 1 To lastCol
            cellVal = CStr(ws.Cells(i, j).Value)
            If i = 1 Then
                dataHtml = dataHtml & "<th style='background:#1F4E79;color:white;padding:4px'>" & cellVal & "</th>"
            Else
                dataHtml = dataHtml & "<td style='padding:4px'>" & cellVal & "</td>"
            End If
        Next j
        dataHtml = dataHtml & "</tr>"
    Next i
    dataHtml = dataHtml & "</table>"

    ' Replace placeholder with data table
    htmlBody = "<html><body style=""font-family:Calibri,sans-serif;font-size:11pt"">" & htmlBody & "</body></html>"
    htmlBody = Replace(htmlBody, "{{REPORT_PLACEHOLDER}}", dataHtml)
    htmlBody = Replace(htmlBody, "{{REPORT}}", dataHtml)
    htmlBody = Replace(htmlBody, "{{REPORT_PLACEHOLDER}}", dataHtml)
    htmlBody = Replace(htmlBody, "{{REPORT}}", dataHtml)

    ' Create email
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .Subject = emailSubject
        If Len(ccEmails) > 0 Then .CC = ccEmails
        '' Add image reference at bottom if image exists
        If Len(imagePath) > 0 Then
            If Dir(imagePath) <> "" Then
                htmlBody = htmlBody & "<br><br><img src=""cid:emailImage"" style=""max-width:600px"">"
            End If
        End If

        .HTMLBody = htmlBody

        ' Embed image if path exists
        If Len(imagePath) > 0 Then
            If Dir(imagePath) <> "" Then
                Dim att As Object
                Set att = .Attachments.Add(imagePath)
                att.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "emailImage"
            End If
        End If

        .Display
    End With

    LogLine "GenerateManagerEmail: email created for " & ws.Name
    On Error Resume Next
    ThisWorkbook.Sheets("Manager Email Template").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Director Email Template").Visible = xlSheetHidden
    On Error GoTo LocalFail
    Exit Sub

LocalFail:
    LogErr "GenerateManagerEmail"
End Sub
