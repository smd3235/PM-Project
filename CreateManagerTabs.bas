Attribute VB_Name = "CreateManagerTabs"
Option Explicit

' ============================================================
' WEEKLY SALES MANAGER REPORT BUILDER - v1.2 (Excel)
'
' Host: Excel (Windows/macOS). The workbook is the report
' template: the START HERE button runs the weekly build.
'
' Structure:
'   BuildManagerTabs   - interactive shell (dialogs + message).
'   RunReport          - core pipeline; returns a result string:
'                        "OK|savePath|managerPms|pending|skipped"
'                        "ERROR|number|description"
'                        Public so the shell and tests can call it.
'
' One-button weekly workflow:
'   - Select the current raw-data workbook; no copy/paste.
'   - Optionally select the prior report to compare against
'     (Cancel = no comparison); prior tabs are read read-only.
'   - Automatically finds the source worksheet by required headers.
'   - Validates required headers before processing.
'   - Null-safe: blank manager/status/header cells are skipped, not fatal.
'   - Full source formatting (dates, number formats) carried into output.
'   - Every manager data row counts as a PM request.
'   - Pending Pickup rows are excluded from manager tabs and PM counts.
'   - New PM rows are light green + bold (names absent from prior report).
'   - Dashboard summary (KPIs + manager table) is the first tab.
'   - Pending Pickup tab carries a prior/now/new/cleared comparison card.
'   - Raw Data is copied into the report as the last tab.
'   - Remembers the output folder in a hidden PM Settings sheet.
'   - First run asks for an output folder; later runs reuse it.
'   - Prevents accidental filename overwrite (unique filenames).
'   - Repairs/installs the START HERE button automatically.
'   - Keeps the source workbook read-only and closes it without saving.
'
' IMPORTANT BUSINESS RULES:
'   * Every non-Pending Pickup data row assigned to a manager is a PM request.
'   * Pending Pickup rows go only to the Pending Pickup tab.
'   * PM Needed equals the number of data rows on the manager tab.
'   * Treatment PM Date is not a PM-needed filter.
' ============================================================

Private Const RAW_SHEET As String = "Raw Data"
Private Const SETTINGS_SHEET As String = "PM Settings"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const MANAGER_HEADER As String = "Assigned Sales Manager"
Private Const NAME_HEADER As String = "Name"
Private Const STATUS_HEADER As String = "Treatment Status"
Private Const PENDING_TEXT As String = "Pending Pickup"
Private Const BUTTON_NAME As String = "btnCreatePMReport"
Private Const BUTTON_TEXT As String = "CREATE PM REQUEST REPORT"
Private Const SETTING_OUTPUT_FOLDER As String = "OutputFolder"
Private Const LOGO_SHAPE As String = "AOTI_Logo"
Private Const START_SHEET As String = "START HERE"
Private Const LOGO_FILE As String = "aoti-logo.png"


' ==================== INTERACTIVE ENTRY (button) ====================

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

    btnWidth = 340
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
    If Not Silent Then MsgBox "I could not install the button." & vbCrLf & Err.Description, vbCritical
End Sub


' Public entry point for sync_template.ps1: rebuild START HERE + button so the
' shipped workbook matches the current module without a full report run.
Public Sub RebuildStartSheet()
    SetupReportButton True
End Sub


Public Sub BuildManagerTabs()
    Dim sourcePath As Variant
    Dim outputFolder As String
    Dim priorPath As String
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
        if Len(outputFolder) = 0 Then SetStatus "": Exit Sub
        ' Save the master immediately so the remembered folder survives a close
        ' without a save prompt answer.
        SaveOutputFolder ThisWorkbook, outputFolder
        ThisWorkbook.Save
    End If

    SetStatus "Step 2 of 3 - Building the PM Requests report..."
    result = RunReport(CStr(sourcePath), outputFolder, priorPath)

    resultParts = Split(result, "|")
    If UBound(resultParts) >= 0 Then
        If resultParts(0) = "OK" Then
            MsgBox "Done!" & vbCrLf & vbCrLf & _
                   "Created: " & vbCrLf & resultParts(1) & vbCrLf & vbCrLf & _
                   "Manager PM requests: " & resultParts(2) & vbCrLf & _
                   "Pending Pickup: " & resultParts(3) & vbCrLf & _
                   IIf(CLng(resultParts(4)) > 0, "Skipped rows (no manager, not pending): " & resultParts(4) & vbCrLf, "") & _
                   IIf(UBound(resultParts) >= 5 And Len(resultParts(5)) > 0, "Compared against: " & resultParts(5) & vbCrLf, "No prior report found - comparison skipped." & vbCrLf), vbInformation
        Else
            MsgBox "The PM report could not be created." & vbCrLf & vbCrLf & _
                   "Error " & resultParts(1) & ": " & resultParts(2), vbCritical
        End If
    End If
    Exit Sub

CleanFail:
    SetStatus ""
    MsgBox "The PM report could not be created." & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical
End Sub


' ====================== CORE (host-neutral) ======================

Public Function RunReport(ByVal sourcePath As String, ByVal outputFolder As String, Optional ByVal priorPath As String = "") As String
    Dim wb As Workbook, srcWb As Workbook, raw As Worksheet, outWb As Workbook
    Dim headerRow As Long, managerCol As Long, statusCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant, r As Long, i As Long
    Dim manager As String, status As String
    Dim mgrNames As Collection, mgrRows As Collection, pendingRows As Collection
    Dim managers As Variant, outWs As Worksheet, pendingWs As Worksheet
    Dim rowsForManager As Collection, arrOut As Variant, arrPending As Variant
    Dim firstOutput As Boolean, reportDate As String, savePath As String
    Dim skippedRows As Long, pmTotal As Long, idx As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim priorNames As Collection, priorMgrNames As Collection, priorMgrCounts As Collection
    Dim priorPendingNames As Collection, priorPending As Long, priorPMs As Long
    Dim curPendingNames As Collection, mgrDataNames As Collection
    Dim pendMgrNames As Collection, repNames As Collection
    Dim repCounts() As Long, repN As Long
    Dim newPMs As Long, nameCol As Long
    Dim repName As String
    Dim stage As String

    On Error GoTo Fail

    Set wb = ThisWorkbook

    oldCalc = -4135: oldScreen = False: oldEvents = False: oldAlerts = False
    On Error Resume Next
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Fail

    FastModeOn
    stage = "fastmode"

    Set srcWb = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    stage = "source open"

    Set raw = FindSourceWorksheet(srcWb, headerRow)
    If raw Is Nothing Then
        Err.Raise vbObjectError + 2202, , "I could not find a worksheet containing both required headers: " & MANAGER_HEADER & " and " & STATUS_HEADER & "."
    End If

    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, STATUS_HEADER)
    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)

    If lastRow <= headerRow Then
        Err.Raise vbObjectError + 2203, , "The selected current report has no data rows."
    End If

    SetStatus "Loading current report into Raw Data..."

    LoadSourceIntoRawData wb, raw, headerRow, lastRow, lastCol
    stage = "raw load"

    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    Set raw = wb.Worksheets(RAW_SHEET)
    headerRow = FindHeaderRow(raw, MANAGER_HEADER, STATUS_HEADER)
    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, STATUS_HEADER)
    nameCol = FindHeaderColumn(raw, headerRow, NAME_HEADER)
    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)

    SetStatus "Building manager and Pending Pickup tabs..."

    data = raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Value2
    stage = "data read"

    Set mgrNames = New Collection
    Set mgrRows = New Collection
    Set pendingRows = New Collection
    Set curPendingNames = New Collection
    Set mgrDataNames = New Collection
    Set pendMgrNames = New Collection
    Set repNames = New Collection
    ReDim repCounts(1 To 1)
    repN = 0
    stage = "collections set"

    ' Null-safe classification: blank manager/status cells are skipped,
    ' not fatal (CStr(Null) would raise error 94). Collection-based to
    ' avoid the Scripting.Dictionary COM dependency.
    For r = 2 To UBound(data, 1)
        manager = Trim$(SafeStr(data(r, managerCol)))
        status = SafeStr(data(r, statusCol))

        If InStr(1, status, PENDING_TEXT, vbTextCompare) > 0 Then
            pendingRows.Add r
            If nameCol > 0 Then curPendingNames.Add NormalizeName(data(r, nameCol))
            pendMgrNames.Add Trim$(SafeStr(data(r, managerCol)))
        ElseIf Len(manager) > 0 Then
            idx = FindManager(mgrNames, manager)
            If idx = 0 Then
                Set rowsForManager = New Collection
                mgrNames.Add manager
                mgrRows.Add rowsForManager
            Else
                Set rowsForManager = mgrRows(idx)
            End If
            rowsForManager.Add r
            If nameCol > 0 Then mgrDataNames.Add NormalizeName(data(r, nameCol))
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
        Else
            skippedRows = skippedRows + 1
        End If
    Next r
    stage = "classified"

    ' Weekly comparison: with no explicit prior, use the previous week's
    ' report found in the output folder (name overlap). An unreadable prior
    ' degrades to empty comparison data instead of failing the build.
    If Len(priorPath) = 0 Then
        SetStatus "Looking for the prior week's report..."
        stage = "detect prior"
        priorPath = DetectPriorReport(outputFolder, data, nameCol, lastRow)
    End If

    If Len(priorPath) > 0 Then
        SetStatus "Comparing to the prior report..."
        LoadPriorComparison priorPath, priorNames, priorMgrNames, priorMgrCounts, priorPendingNames, priorPending
    Else
        Set priorNames = New Collection
        Set priorMgrNames = New Collection
        Set priorMgrCounts = New Collection
        Set priorPendingNames = New Collection
        priorPending = 0
    End If
    priorPMs = 0
    For i = 1 To priorMgrCounts.Count: priorPMs = priorPMs + CLng(priorMgrCounts(i)): Next i

    ' New PMs: assigned (non-pending) names that were not on the prior
    ' report's manager tabs.
    If priorNames.Count > 0 Then
        For i = 1 To mgrDataNames.Count
            If Not PriorNameExists(priorNames, CStr(mgrDataNames(i))) Then newPMs = newPMs + 1
        Next i
    End If

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    firstOutput = True
    stage = "out workbook"

    If mgrNames.Count > 0 Then
        ReDim managers(1 To mgrNames.Count)
        For i = 1 To mgrNames.Count: managers(i) = mgrNames(i): Next i
        SortManagersByLastName managers
        For i = LBound(managers) To UBound(managers)
            manager = CStr(managers(i))
            SetStatus "Creating manager tab: " & manager
            Set outWs = GetOrCreateOutputSheet(outWb, SafeSheetName(manager), firstOutput)
            firstOutput = False
            Set rowsForManager = mgrRows(FindManager(mgrNames, manager))
            arrOut = BuildOutputArray(data, rowsForManager, lastCol)
            ' Table header on row 9, data from row 10 (KPI cards above).
            outWs.Range(outWs.Cells(9, 1), outWs.Cells(8 + UBound(arrOut, 1), lastCol)).Value = arrOut
            For idx = 1 To rowsForManager.Count
                CopyRowFormats raw, outWs, CLng(rowsForManager(idx)), idx + 9, lastCol
            Next idx
            FormatManagerSheet outWs, manager, priorNames, priorMgrNames, priorMgrCounts, nameCol
        Next i
    End If
    stage = "manager tabs"

    SetStatus "Creating Pending Pickup tab..."
    Set pendingWs = GetOrCreateOutputSheet(outWb, "Pending Pickup", firstOutput)

    If pendingRows.Count > 0 Then
        arrPending = BuildOutputArray(data, pendingRows, lastCol)
        pendingWs.Range(pendingWs.Cells(1, 1), pendingWs.Cells(UBound(arrPending, 1), lastCol)).Value = arrPending
        For idx = 1 To pendingRows.Count
            CopyRowFormats raw, pendingWs, CLng(pendingRows(idx)), idx + 1, lastCol
        Next idx
    Else
        pendingWs.Range(pendingWs.Cells(1, 1), pendingWs.Cells(1, lastCol)).Value = GetHeaderArray(data, lastCol)
    End If

    FormatPendingPickupSheet pendingWs, priorPendingNames, curPendingNames, priorPending
    pendingWs.Move After:=outWb.Worksheets(outWb.Worksheets.Count)
    stage = "pending done"

    ' PM totals: sum of rows actually written to manager tabs.
    pmTotal = 0
    For i = 1 To mgrRows.Count: pmTotal = pmTotal + mgrRows(i).Count: Next i

    If Len(outputFolder) = 0 Then
        Err.Raise vbObjectError + 2204, , "No output folder was given."
    End If
    reportDate = Format(Date, "yyyy-mm-dd")
    savePath = GetUniqueOutputPath(outputFolder, "PM Requests " & reportDate & ".xlsx")

    SetStatus "Building dashboard..."
    stage = "dashboard"
    BuildDashboard outWb, managers, mgrRows, mgrNames, data, nameCol, _
                  pmTotal, newPMs, pendingRows.Count, curPendingNames, _
                  priorNames, priorMgrNames, priorMgrCounts, priorPMs, priorPending, priorPendingNames, _
                  pendMgrNames, repNames, repCounts, repN, _
                  Mid$(savePath, InStrRev(savePath, Application.PathSeparator) + 1), priorPath
    outWb.Worksheets(DASHBOARD_SHEET).Move Before:=outWb.Worksheets(1)
    stage = "dashboard moved"

    ' Copy the loaded Raw Data sheet into the report (last tab), like the
    ' 2.0.18 build. Sheet copy leaves the copy active, so re-activate the
    ' Dashboard before saving: the report then always opens on it.
    raw.Copy After:=outWb.Worksheets(outWb.Worksheets.Count)
    stage = "raw copied"
    outWb.Worksheets(DASHBOARD_SHEET).Activate
    outWb.Worksheets(DASHBOARD_SHEET).Range("A1").Select

    SetStatus "Saving finished report..."

    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    stage = "saved"
    outWb.Close SaveChanges:=False
    Set outWb = Nothing

    RunReport = "OK|" & savePath & "|" & CStr(pmTotal) & "|" & CStr(pendingRows.Count) & "|" & CStr(skippedRows) & "|" & _
                IIf(Len(priorPath) = 0, "", Mid$(priorPath, InStrRev(priorPath, Application.PathSeparator) + 1))

    SetStatus ""
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    Exit Function

Fail:
    ' Capture the error before any cleanup statement can touch Err.
    Dim errNo As Long, errDesc As String, errStage As String
    errNo = Err.Number: errDesc = Err.Description: errStage = stage

    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    SetStatus ""
    RestoreExcelState oldCalc, oldScreen, oldEvents, oldAlerts
    On Error GoTo 0
    RunReport = "ERROR|" & CStr(errNo) & "|" & errDesc & " @ " & errStage
End Function


' ====================== HOST / DIALOG HELPERS ======================

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
        if fd.Show = -1 Then PickOutputFolder = fd.SelectedItems(1)
    End If
    On Error GoTo 0
End Function


Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = text
    On Error GoTo 0
End Sub


Private Sub FastModeOn()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo 0
End Sub


' ====================== WORKBOOK / SHEET SETUP ======================

' START HERE canvas, rebuilt idempotently on every run (shapes survive
' Cells.Clear). Layout mirrors the Pending Pickup builder: narrow gutters,
' navy label bars, numbered step rows, a tall CTA row, and a footer.
Private Function GetOrCreateStartSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(START_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = START_SHEET
    End If

    ' Rebuild the canvas every run; stale merges would block the new layout.
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.UnMerge
    On Error GoTo 0

    ' Gutters + content columns.
    ws.Columns("A").ColumnWidth = 2.6
    ws.Columns("B").ColumnWidth = 5.6
    ws.Columns("C").ColumnWidth = 28.6
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 2.6

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
        .Font.Italic = True
        .Font.Size = 11
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

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

    With ws.Range("B17:F17")
        .Merge
        .Value = "-  One tab per Assigned Sales Manager, plus Pending Pickup and Raw Data."
    End With
    With ws.Range("B18:F18")
        .Merge
        .Value = "-  Rows with Treatment Status 'Pending Pickup' appear only on Pending Pickup."
    End With
    With ws.Range("B19:F19")
        .Merge
        .Value = "-  PMs new since the prior report are highlighted green on each manager tab."
    End With
    With ws.Range("B17:F19")
        .Font.Size = 10
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("B21:F21")
        .Merge
        .Value = "v1.3  |  PM Requests Report Builder  |  Output: PM Requests YYYY-MM-DD.xlsx"
        .Font.Size = 9
        .Font.Color = RGB(191, 191, 191)
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("B2:F21").Font.Name = "Segoe UI"

    ' Logo: keep the embedded START HERE picture, top-right of the title row.
    PositionStartLogo ws

    ' Cleaner canvas: hide gridlines on this sheet (best-effort).
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    Set GetOrCreateStartSheet = ws
End Function


' One numbered step row: navy number cell + merged text on the given fill.
Private Sub AddStartStep(ws As Worksheet, ByVal rowNum As Long, ByVal text As String, ByVal fillColor As Long)
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
End Sub


' Places the embedded AOTI_Logo picture at the top-right of the title rows.
Private Sub PositionStartLogo(ws As Worksheet)
    Dim shp As Shape, logoW As Double, logoH As Double
    Dim rowTop As Double, rowHeight As Double

    On Error Resume Next
    Set shp = ws.Shapes(LOGO_SHAPE)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    logoW = shp.Width
    logoH = shp.Height
    If logoH <= 0 Then logoH = 30

    rowTop = ws.Range("B2").Top
    rowHeight = ws.Range("B2:B3").Height

    On Error Resume Next
    shp.Left = ws.Range("E2").Left + 1
    shp.Top = rowTop + (rowHeight - logoH) / 2
    shp.Placement = xlMove
    On Error GoTo 0
End Sub


Private Function FindSourceWorksheet(wb As Workbook, ByRef headerRowOut As Long) As Worksheet
    Dim ws As Worksheet, hr As Long
    For Each ws In wb.Worksheets
        hr = FindHeaderRow(ws, MANAGER_HEADER, STATUS_HEADER)
        If hr > 0 Then
            headerRowOut = hr
            Set FindSourceWorksheet = ws
            Exit Function
        End If
    Next ws
End Function


' Loads the source range into Raw Data AND copies the full range formats,
' so date/number formats survive into the report (Value2 alone loses them).
Private Sub LoadSourceIntoRawData(masterWb As Workbook, src As Worksheet, headerRow As Long, lastRow As Long, lastCol As Long)
    Dim dst As Worksheet
    Dim srcRange As Range
    Dim r As Long
    Set dst = GetOrCreateRawSheet(masterWb)
    dst.Cells.Clear
    Set srcRange = src.Range(src.Cells(headerRow, 1), src.Cells(lastRow, lastCol))
    dst.Range("A1").Resize(srcRange.Rows.Count, srcRange.Columns.Count).Value2 = srcRange.Value2
    For r = 1 To srcRange.Rows.Count
        CopyRowFormats src, dst, headerRow + r - 1, r, lastCol
    Next r
    dst.Columns.AutoFit
End Sub


Private Function GetOrCreateRawSheet(wb As Workbook) As Worksheet
    If SheetExists(RAW_SHEET, wb) Then
        Set GetOrCreateRawSheet = wb.Worksheets(RAW_SHEET)
    Else
        Set GetOrCreateRawSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateRawSheet.Name = RAW_SHEET
    End If
End Function


Private Function GetOrCreateSettingsSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If SheetExists(SETTINGS_SHEET, wb) Then
        Set ws = wb.Worksheets(SETTINGS_SHEET)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SETTINGS_SHEET
        ws.Range("A1").Value = "Setting"
        ws.Range("B1").Value = "Value"
        ws.Range("A2").Value = SETTING_OUTPUT_FOLDER
    End If
    ws.Visible = xlSheetVeryHidden
    Set GetOrCreateSettingsSheet = ws
End Function


Private Function GetRememberedOutputFolder(wb As Workbook) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    If Not ws Is Nothing Then GetRememberedOutputFolder = Trim$(SafeStr(ws.Range("B2").Value))
    On Error GoTo 0
End Function


Private Sub SaveOutputFolder(wb As Workbook, folderPath As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSettingsSheet(wb)
    ws.Range("B2").Value = folderPath
End Sub


' ====================== PATH / STATE HELPERS ======================

Private Function FolderExists(folderPath As String) As Boolean
    Dim clean As String
    clean = NormalizeFolder(folderPath)
    On Error Resume Next
    FolderExists = (Len(Dir(clean, vbDirectory)) > 0)
    On Error GoTo 0
End Function


' Strip any trailing path separator so joins never double up.
Private Function NormalizeFolder(ByVal folderPath As String) As String
    Dim s As String
    s = Trim$(folderPath)
    Do While Len(s) > 1 And Right$(s, 1) = Application.PathSeparator
        s = Left$(s, Len(s) - 1)
    Loop
    NormalizeFolder = s
End Function


Private Function GetUniqueOutputPath(folderPath As String, baseFileName As String) As String
    Dim candidate As String, n As Long, dotPos As Long
    folderPath = NormalizeFolder(folderPath)
    candidate = folderPath & Application.PathSeparator & baseFileName
    If Len(Dir(candidate)) = 0 Then
        GetUniqueOutputPath = candidate
        Exit Function
    End If
    dotPos = InStrRev(baseFileName, ".")
    n = 2
    Do
        candidate = folderPath & Application.PathSeparator & Left$(baseFileName, dotPos - 1) & " (" & CStr(n) & ")" & Mid$(baseFileName, dotPos)
        n = n + 1
    Loop While Len(Dir(candidate)) > 0
    GetUniqueOutputPath = candidate
End Function


Private Sub RestoreExcelState(oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean)
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub


' ====================== DATA / FORMAT HELPERS ======================

Private Function BuildOutputArray(ByRef data As Variant, ByVal rows As Collection, ByVal lastCol As Long) As Variant
    Dim result() As Variant, i As Long, c As Long, srcRow As Long
    ReDim result(1 To rows.Count + 1, 1 To lastCol)
    For c = 1 To lastCol: result(1, c) = data(1, c): Next c
    For i = 1 To rows.Count
        srcRow = CLng(rows(i))
        For c = 1 To lastCol: result(i + 1, c) = data(srcRow, c): Next c
    Next i
    BuildOutputArray = result
End Function


Private Function GetHeaderArray(ByRef data As Variant, ByVal lastCol As Long) As Variant
    Dim a() As Variant, c As Long
    ReDim a(1 To 1, 1 To lastCol)
    For c = 1 To lastCol: a(1, c) = data(1, c): Next c
    GetHeaderArray = a
End Function


' Per-cell format copy (number formats, fonts) for one row. Property-based
' (no clipboard) so it cannot fail on a clipboard lock or in invisible
' automation. Column-level NumberFormat is unreliable when cells carry
' individual formats, so formats are copied cell by cell.
Private Sub CopyRowFormats(src As Worksheet, dst As Worksheet, ByVal srcRow As Long, ByVal dstRow As Long, ByVal lastCol As Long)
    Dim c As Long
    Dim sf As Font, df As Font
    For c = 1 To lastCol
        dst.Cells(dstRow, c).NumberFormat = src.Cells(srcRow, c).NumberFormat
        Set sf = src.Cells(srcRow, c).Font
        Set df = dst.Cells(dstRow, c).Font
        df.Name = sf.Name
        df.Size = sf.Size
        df.Bold = sf.Bold
        df.Color = sf.Color
    Next c
End Sub


Private Sub SortManagersByLastName(ByRef a As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If LastNameKey(CStr(a(j))) < LastNameKey(CStr(a(i))) Then
                tmp = a(i): a(i) = a(j): a(j) = tmp
            End If
        Next j
    Next i
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
    topRow = LastUsedRow(ws)
    If topRow > 30 Then topRow = 30
    For r = 1 To topRow
        If FindHeaderColumn(ws, r, header1) > 0 And FindHeaderColumn(ws, r, header2) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function


' Null-safe: CStr(Null) raises error 94, so blank header cells used to crash
' header detection. Now they are simply skipped.
Private Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim c As Long, lastCol As Long
    If headerRow < 1 Then Exit Function
    lastCol = LastUsedColumn(ws)
    For c = 1 To lastCol
        If StrComp(Trim$(SafeStr(ws.Cells(headerRow, c).Value)), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function


Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
End Function


Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = f.Column
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
    For i = 1 To mgrNames.Count
        If StrComp(CStr(mgrNames(i)), name, vbTextCompare) = 0 Then
            FindManager = i
            Exit Function
        End If
    Next i
End Function


' Collision-proof tab naming: managers like "A/B" and "A\B" both sanitize to
' "A-B"; a manager literally named "Pending Pickup" would clash with that tab.
Private Function GetOrCreateOutputSheet(wb As Workbook, sheetName As String, ByVal firstOutput As Boolean) As Worksheet
    Dim ws As Worksheet
    If firstOutput And wb.Worksheets.Count = 1 Then
        Set ws = wb.Worksheets(1)
        ws.Name = UniqueSheetName(wb, sheetName)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = UniqueSheetName(wb, sheetName)
    End If
    Set GetOrCreateOutputSheet = ws
End Function


Private Function UniqueSheetName(wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, n As Long
    candidate = baseName
    n = 1
    Do While SheetExists(candidate, wb)
        n = n + 1
        candidate = Left$(baseName, 31 - Len(" (" & CStr(n) & ")")) & " (" & CStr(n) & ")"
    Loop
    UniqueSheetName = candidate
End Function


' ================== PRIOR REPORT COMPARISON ==================
' Reads a prior PM Requests report (2.0.18 layout or this builder):
' every sheet except Dashboard / Pending Pickup / Raw Data / START HERE
' is a manager tab; its Name column feeds the prior-name set.

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
    Dim threshold As Long, r As Long

    If Len(outputFolder) = 0 Or nameCol < 1 Then Exit Function
    folder = NormalizeFolder(outputFolder)
    If Not FolderExists(folder) Then Exit Function

    Set currentNames = New Collection
    For r = 2 To currentLastRow
        currentNames.Add NormalizeName(data(r, nameCol))
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

    ' Prefer the builder's Raw Data tab; fall back to any sheet with the
    ' required headers (covers 2.0.18-built reports).
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(RAW_SHEET)
    On Error GoTo Fail
    If ws Is Nothing Then Set ws = FindSourceWorksheet(wb, headerRow)

    If ws Is Nothing Then GoTo Fail
    headerRow = FindHeaderRow(ws, MANAGER_HEADER, STATUS_HEADER)
    If headerRow < 1 Then GoTo Fail
    nameCol = FindHeaderColumn(ws, headerRow, NAME_HEADER)
    If nameCol < 1 Then GoTo Fail

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)
    If lastRow <= headerRow Then GoTo Fail

    data = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol)).Value2
    Set priorNames = New Collection
    For r = 2 To UBound(data, 1)
        priorNames.Add NormalizeName(data(r, nameCol))
    Next r

    missing = CountDiff(currentNames, priorNames)
    PriorNameOverlap = currentNames.Count - missing
    If PriorNameOverlap < 0 Then PriorNameOverlap = 0

Fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function


' Sorts a 1-based path array by file time, newest first (insertion sort).
Private Sub SortPathsNewestFirst(ByRef paths() As String, ByVal n As Long)
    Dim i As Long, j As Long, key As String
    Dim keyTime As Date, curTime As Date

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
End Sub


Private Sub LoadPriorComparison(ByVal priorPath As String, _
                                ByRef priorNames As Collection, _
                                ByRef priorMgrNames As Collection, _
                                ByRef priorMgrCounts As Collection, _
                                ByRef priorPendingNames As Collection, _
                                ByRef priorPending As Long)
    Dim wb As Workbook, ws As Worksheet
    Dim headerRow As Long, nameCol As Long, lastRow As Long, r As Long

    Set priorNames = New Collection
    Set priorMgrNames = New Collection
    Set priorMgrCounts = New Collection
    Set priorPendingNames = New Collection
    priorPending = 0

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
        ElseIf StrComp(ws.Name, PENDING_TEXT, vbTextCompare) = 0 Then
            headerRow = FindHeaderRowForSingle(ws, NAME_HEADER)
            If headerRow = 0 Then headerRow = 1
            nameCol = FindHeaderColumn(ws, headerRow, NAME_HEADER)
            If nameCol = 0 Then nameCol = 3
            lastRow = LastUsedRow(ws)
            If lastRow > headerRow Then priorPending = lastRow - headerRow
            For r = headerRow + 1 To lastRow
                priorPendingNames.Add NormalizeName(ws.Cells(r, nameCol).Value)
            Next r
        End If
    Next ws

    wb.Close SaveChanges:=False
    Exit Sub

Fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ' Comparison degrades to "no prior" instead of failing the build.
End Sub


Private Function IsManagerTab(ByVal sheetName As String) As Boolean
    If StrComp(sheetName, DASHBOARD_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, PENDING_TEXT, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, RAW_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, SETTINGS_SHEET, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, "START HERE", vbTextCompare) = 0 Then Exit Function
    IsManagerTab = True
End Function


' Searches the first rows for a lone header name (prior tabs may have
' title rows above the table).
Private Function FindHeaderRowForSingle(ws As Worksheet, ByVal headerName As String) As Long
    Dim r As Long, maxRow As Long
    maxRow = LastUsedRow(ws)
    If maxRow > 30 Then maxRow = 30
    For r = 1 To maxRow
        If FindHeaderColumn(ws, r, headerName) > 0 Then
            FindHeaderRowForSingle = r
            Exit Function
        End If
    Next r
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

' Manager tab: 2.0.18-style KPI summary cards (rows 1-8), table header on
' row 9, data from row 10. Cards: PMs Requested / Prior PMs / Completed /
' New PMs / PMs by Rep (per-manager rep breakdown read from column A).
Private Sub FormatManagerSheet(ws As Worksheet, ByVal manager As String, _
                               ByVal priorNames As Collection, _
                               ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, _
                               ByVal nameCol As Long)
    Dim lastCol As Long, lastRow As Long
    Dim priorCount As Long, curCount As Long, newCount As Long, completedCount As Long
    Dim r As Long, isNew As Boolean
    Dim repNames As Collection
    Dim repText As String, i As Long, repName As String

    lastCol = LastUsedColumn(ws): lastRow = LastUsedRow(ws)
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    ' Table: header on row 9, data rows 10..lastRow.
    ApplyReaderFriendlyFormatting ws, 9, lastRow, lastCol
    If lastRow >= 10 Then SortByColumnA ws, 9, lastRow, lastCol

    ' New PM rows: light green + bold, overriding the banding (matches
    ' the 2.0.18 build used for work comparisons).
    If priorNames.Count > 0 And nameCol > 0 Then
        For r = 10 To lastRow
            isNew = Not PriorNameExists(priorNames, NormalizeName(ws.Cells(r, nameCol).Value))
            If isNew Then
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
                    .Interior.Color = RGB(226, 239, 218)
                    .Font.Bold = True
                End With
            End If
        Next r
    End If

    ' KPI values.
    priorCount = PriorMgrCount(priorMgrNames, priorMgrCounts, SafeSheetName(manager))
    If priorCount < 0 Then priorCount = 0
    curCount = lastRow - 9
    newCount = 0
    If priorNames.Count > 0 And nameCol > 0 Then
        For r = 10 To lastRow
            If Not PriorNameExists(priorNames, NormalizeName(ws.Cells(r, nameCol).Value)) Then newCount = newCount + 1
        Next r
    End If
    completedCount = priorCount - (curCount - newCount)
    If completedCount < 0 Then completedCount = 0

    ' Banner: "Manager | PM SUMMARY".
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .Value = manager & "  |  PM SUMMARY"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30

    ' KPI cards (2.0.18 arrangement).
    AddPolishedCard ws, "A2:B4", "PMs Requested", curCount, RGB(221, 235, 247), False, "CURRENT"
    AddPolishedCard ws, "C2:D4", "Prior PMs", priorCount, RGB(242, 242, 242), False, "PRIOR"
    AddPolishedCard ws, "E2:F4", "Completed", completedCount, RGB(226, 239, 218), False, "SINCE PRIOR"
    AddPolishedCard ws, "A5:B7", "New PMs", newCount, RGB(226, 239, 218), False, "ADDED"

    ' PMs by Rep card (counts read from the written column A). Rep names stay
    ' in first-appearance order: SortManagersByLastName only accepts arrays,
    ' not a Collection.
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
        repText = ""
        For i = 1 To repNames.Count
            If Len(repText) > 0 Then repText = repText & "   |   "
            repText = repText & CStr(repNames(i)) & ": " & CStr(CountInSheet(ws, CStr(repNames(i)), lastRow))
        Next i
    End If
    AddPolishedTextCard ws, "C5:F7", "PMs BY REP", repText, RGB(247, 249, 251), "REP BREAKDOWN"

    ws.Rows(8).RowHeight = 10
    ws.Columns.AutoFit
    If lastCol >= 1 Then ws.Columns(1).ColumnWidth = 14
End Sub


' Case-insensitive count of a value in a sheet column (rep totals).
Private Function CountInSheet(ws As Worksheet, ByVal value As String, ByVal lastRow As Long) As Long
    Dim r As Long, n As Long
    For r = 10 To lastRow
        If StrComp(Trim$(SafeStr(ws.Cells(r, 1).Value)), value, vbTextCompare) = 0 Then n = n + 1
    Next r
    CountInSheet = n
End Function


' 2.0.18 KPI card: eyebrow + label header strip, big centered value.
Private Sub AddPolishedCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal value As Long, ByVal fillColor As Long, ByVal colorChange As Boolean, Optional ByVal eyebrow As String = "")
    Dim box As Range, hdr As Range, valBox As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    Dim accent As Long

    Set box = ws.Range(address)
    firstRow = box.Row: firstCol = box.Column
    lastRow = box.Row + box.Rows.Count - 1: lastCol = box.Column + box.Columns.Count - 1
    box.UnMerge
    Set hdr = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(firstRow, lastCol))
    Set valBox = ws.Range(ws.Cells(firstRow + 1, firstCol), ws.Cells(lastRow, lastCol))
    hdr.Merge: valBox.Merge

    If colorChange Then accent = ChangeColor(value) Else accent = RGB(31, 78, 121)

    hdr.Value = UCase$(label)
    If Len(eyebrow) > 0 Then hdr.Value = UCase$(eyebrow) & vbLf & UCase$(label)
    valBox.Value = CStr(value)

    With box
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(210, 218, 226)
    End With
    With hdr
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = accent
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    With valBox
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Bold = True: .Font.Size = 24: .Font.Color = RGB(31, 45, 61)
        .Interior.Color = fillColor
    End With
    If colorChange Then valBox.Font.Color = ChangeColor(value)

    ws.Rows(firstRow).RowHeight = 30
    ws.Rows(firstRow + 1).RowHeight = 24
    If lastRow > firstRow + 1 Then ws.Rows((firstRow + 1) & ":" & lastRow).RowHeight = 24
End Sub


' 2.0.18 text card: eyebrow + label header strip, wrapped text value.
Private Sub AddPolishedTextCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal valueText As String, ByVal fillColor As Long, Optional ByVal eyebrow As String = "")
    Dim box As Range, hdr As Range, valBox As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long

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
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(91, 110, 126)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    With valBox
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Bold = False: .Font.Size = 9: .Font.Color = RGB(31, 45, 61)
        .Interior.Color = fillColor
    End With

    ws.Rows(firstRow).RowHeight = 30
    ws.Rows(firstRow + 1).RowHeight = 24
    If lastRow > firstRow + 1 Then ws.Rows((firstRow + 1) & ":" & lastRow).RowHeight = 24
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
' row (manager tabs use row 9 under the KPI cards; Pending Pickup uses 1).
Private Sub ApplyReaderFriendlyFormatting(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim r As Long
    Dim dataRange As Range

    ' Hairline borders across the data area.
    Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
    dataRange.Borders.LineStyle = xlContinuous
    dataRange.Borders.Weight = xlHairline
    dataRange.VerticalAlignment = xlCenter

    ' Strong, easy-to-scan header.
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        If lastRow > headerRow Then
            .Font.Name = ws.Cells(headerRow + 1, 1).Font.Name
            .Font.Size = ws.Cells(headerRow + 1, 1).Font.Size
        End If
    End With
    ws.Rows(headerRow).RowHeight = 34

    ' Alternating light banding for the data rows.
    If lastRow > headerRow Then
        For r = headerRow + 1 To lastRow
            If r Mod 2 = 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(242, 242, 242)
            Else
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 255)
            End If
            ws.Rows(r).RowHeight = 22
        Next r
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
End Sub


Private Sub FormatPendingPickupSheet(ws As Worksheet, ByVal priorPendingNames As Collection, ByVal curPendingNames As Collection, ByVal priorPending As Long)
    Dim lastCol As Long, lastRow As Long

    lastCol = LastUsedColumn(ws): lastRow = LastUsedRow(ws)
    If lastCol < 1 Then Exit Sub

    ' Format the header even in a week with no pending rows.
    ApplyReaderFriendlyFormatting ws, 1, lastRow, lastCol
    AddPendingComparisonCard ws, curPendingNames, priorPendingNames, priorPending, lastCol
    If lastRow <= 1 Then Exit Sub

    SortPendingPickup ws, lastRow, lastCol

    ' Every Pending Pickup row is red text.
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Font.Color = RGB(192, 0, 0)
End Sub


' Compact Pending Pickup comparison card to the right of the data
' (mirrors the KPI card treatment of the 2.0.18 build).
Private Sub AddPendingComparisonCard(ws As Worksheet, ByVal curPendingNames As Collection, ByVal priorPendingNames As Collection, ByVal priorPending As Long, ByVal lastCol As Long)
    Dim box As Range
    Dim newPending As Long, clearedPending As Long
    Dim text As String

    If priorPendingNames.Count = 0 Then Exit Sub

    newPending = CountDiff(curPendingNames, priorPendingNames)
    clearedPending = CountDiff(priorPendingNames, curPendingNames)

    text = "PENDING PICKUP" & vbLf & _
           "Prior: " & CStr(priorPending) & vbLf & _
           "Now: " & CStr(curPendingNames.Count) & vbLf & _
           "New: " & CStr(newPending) & vbLf & _
           "Cleared: " & CStr(clearedPending)

    Set box = ws.Range(ws.Cells(1, lastCol + 2), ws.Cells(5, lastCol + 3))
    box.Merge
    box.Value = text
    With box
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    ws.Columns(lastCol + 2).ColumnWidth = 14
    ws.Columns(lastCol + 3).ColumnWidth = 14
End Sub


' ==================== DASHBOARD ====================

' 2.0.18-style dashboard: logo zone + title band, meta strip (current/prior/
' created), six KPI cards, manager table with per-manager pending counts, and
' the three-block rep/PM grid. Prior columns show 0 when no prior was compared.
Private Sub BuildDashboard(wb As Workbook, ByVal managers As Variant, _
                           ByVal mgrRows As Collection, ByVal mgrNames As Collection, _
                           ByVal data As Variant, ByVal nameCol As Long, _
                           ByVal pmTotal As Long, ByVal newPMs As Long, _
                           ByVal curPendingCount As Long, ByVal curPendingNames As Collection, _
                           ByVal priorNames As Collection, _
                           ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, _
                           ByVal priorPMs As Long, ByVal priorPending As Long, _
                           ByVal priorPendingNames As Collection, _
                           ByVal pendMgrNames As Collection, _
                           ByVal repNames As Collection, repCounts() As Long, ByVal repN As Long, _
                           ByVal currentReportName As String, ByVal priorPath As String)
    Dim ws As Worksheet
    Dim r As Long, i As Long, idx As Long, mgrPending As Long
    Dim rowsForManager As Collection
    Dim priorCount As Long, mgrNew As Long, mgrCompleted As Long
    Dim completedTotal As Long, retainedPMs As Long
    Dim repEnd As Long, repRows As Long, repCount As Long, repGroup As Long
    Dim repRowOffset As Long, repBaseCol As Long
    Dim sortNames() As String, sortCounts() As Long, tS As String, tC As Long, j As Long

    Set ws = GetOrCreateDashboardSheet(wb)

    ' Title band: white logo zone (A:B) + navy title (C:L).
    With ws.Range("A1:B1")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Range("C1:L1")
        .Merge
        .Value = "PM REQUEST DASHBOARD"
        .Font.Bold = True
        .Font.Size = 20
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 34
    ws.Tab.Color = RGB(31, 78, 121)

    ' Meta strip.
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = "CURRENT: " & currentReportName
    ws.Range("E2:H2").Merge
    If Len(priorPath) > 0 Then
        ws.Range("E2").Value = "PRIOR: " & Mid$(priorPath, InStrRev(priorPath, Application.PathSeparator) + 1)
    Else
        ws.Range("E2").Value = "PRIOR: None selected"
    End If
    ws.Range("I2:L2").Merge
    ws.Range("I2").Value = "CREATED: " & Format(Now, "mm/dd/yyyy h:mm AM/PM")
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
    ws.Rows(2).RowHeight = 28

    ' KPI cards.
    AddPolishedCard ws, "A4:B6", "PMs Requested", pmTotal, RGB(221, 235, 247), False, "CURRENT"
    AddPolishedCard ws, "C4:D6", "Prior PMs", priorPMs, RGB(242, 242, 242), False, "PRIOR"
    If priorPMs > 0 Then
        completedTotal = priorPMs - (pmTotal - newPMs)
        If completedTotal < 0 Then completedTotal = 0
    Else
        completedTotal = 0
    End If
    AddPolishedCard ws, "E4:F6", "Completed", completedTotal, RGB(226, 239, 218), False, "SINCE PRIOR"
    AddPolishedCard ws, "G4:H6", "New PMs", newPMs, RGB(226, 239, 218), False, "ADDED"
    AddPolishedCard ws, "I4:J6", "Pending Pickups", curPendingCount, RGB(255, 242, 204), False, "SEPARATE QUEUE"
    ' Without a prior report there is nothing completed and nothing carried over.
    If priorPMs > 0 Then
        retainedPMs = pmTotal - newPMs
        If retainedPMs < 0 Then retainedPMs = 0
    Else
        retainedPMs = 0
    End If
    AddPolishedCard ws, "K4:L6", "Prior PMs Still Open", retainedPMs, _
                      IIf(retainedPMs = 0, RGB(226, 239, 218), RGB(244, 204, 204)), False, "TARGET = 0"
    If retainedPMs = 0 Then ws.Range("K5:L6").Font.Color = RGB(0, 97, 0) Else ws.Range("K5:L6").Font.Color = RGB(156, 0, 6)

    ws.Rows(7).RowHeight = 8

    ' Manager table.
    r = 8
    WriteTableHeader ws, r, 1, Array("Manager", "PMs Requested", "Prior PMs", "Completed", "New PMs", "Pending Pickups")

    ' Rep grid (three side-by-side blocks), names + counts sorted in parallel.
    If repN > 0 Then
        ReDim sortNames(1 To repN)
        ReDim sortCounts(1 To repN)
        For i = 1 To repN
            sortNames(i) = CStr(repNames(i))
            sortCounts(i) = repCounts(i)
        Next i
        For i = 1 To repN - 1
            For j = i + 1 To repN
                If StrComp(sortNames(j), sortNames(i), vbTextCompare) < 0 Then
                    tS = sortNames(i): sortNames(i) = sortNames(j): sortNames(j) = tS
                    tC = sortCounts(i): sortCounts(i) = sortCounts(j): sortCounts(j) = tC
                End If
            Next j
        Next i
        repCount = repN
        repRows = (repCount + 2) \ 3
        For repGroup = 0 To 2
            repBaseCol = 7 + (repGroup * 2)
            WriteTableHeader ws, r, repBaseCol, Array("Sales Rep", "PM Count")
        Next repGroup
        For i = 1 To repCount
            repGroup = (i - 1) \ repRows
            repRowOffset = (i - 1) Mod repRows
            repBaseCol = 7 + (repGroup * 2)
            ws.Cells(r + 1 + repRowOffset, repBaseCol).Value = sortNames(i)
            ws.Cells(r + 1 + repRowOffset, repBaseCol + 1).Value = sortCounts(i)
        Next i
        repEnd = r + repRows
        For repGroup = 0 To 2
            repBaseCol = 7 + (repGroup * 2)
            ws.Range(ws.Cells(r, repBaseCol), ws.Cells(repEnd, repBaseCol + 1)).Borders.LineStyle = xlContinuous
            ws.Range(ws.Cells(r, repBaseCol), ws.Cells(repEnd, repBaseCol + 1)).Borders.Weight = xlThin
        Next repGroup
    Else
        WriteTableHeader ws, r, 7, Array("Sales Rep", "PM Count")
    End If

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
                    If Not PriorNameExists(priorNames, NormalizeName(data(rowsForManager(idx), nameCol))) Then mgrNew = mgrNew + 1
                Next idx
            End If
            mgrCompleted = priorCount - (rowsForManager.Count - mgrNew)
            If mgrCompleted < 0 Then mgrCompleted = 0

            ' Pending pickups assigned to this manager (from the pending rows).
            mgrPending = 0
            For idx = 1 To pendMgrNames.Count
                If StrComp(Trim$(CStr(pendMgrNames(idx))), CStr(managers(i)), vbTextCompare) = 0 Then mgrPending = mgrPending + 1
            Next idx

            ws.Cells(r, 1).Value = CStr(managers(i))
            ws.Cells(r, 2).Value = rowsForManager.Count
            ws.Cells(r, 3).Value = priorCount
            ws.Cells(r, 4).Value = mgrCompleted
            ws.Cells(r, 5).Value = mgrNew
            ws.Cells(r, 6).Value = mgrPending
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", _
                               SubAddress:="'" & SafeSheetName(CStr(managers(i))) & "'!A1", _
                               TextToDisplay:=CStr(managers(i))
            ws.Cells(r, 4).Font.Color = RGB(0, 97, 0)
        Next i
    End If

    ' Table styling.
    ws.Rows(r).RowHeight = 28
    With ws.Range("A8:F8")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    If r > 9 Then
        ws.Range("E9:E" & r).Interior.Color = RGB(226, 239, 218)
        ws.Range("F9:F" & r).Interior.Color = RGB(255, 242, 204)
        For i = 9 To r
            If i Mod 2 = 1 Then ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).Interior.Color = RGB(247, 249, 251)
        Next i
        ws.Range(ws.Cells(9, 1), ws.Cells(r, 6)).VerticalAlignment = xlCenter
        ws.Range(ws.Cells(9, 1), ws.Cells(r, 6)).Borders.Weight = xlThin
    End If

    ' Column widths (2.0.18 footprint A:L).
    ws.Columns("A").ColumnWidth = 23
    ws.Columns("B").ColumnWidth = 12: ws.Columns("C").ColumnWidth = 11: ws.Columns("D").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 10: ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 17: ws.Columns("H").ColumnWidth = 9
    ws.Columns("I").ColumnWidth = 17: ws.Columns("J").ColumnWidth = 9
    ws.Columns("K").ColumnWidth = 17: ws.Columns("L").ColumnWidth = 9
    ws.Range("A4:L6").VerticalAlignment = xlCenter
    ws.Range("A4:L6").WrapText = True

    AddDashboardLogo ws
    ws.Activate
    ws.Range("A1").Select
    On Error Resume Next
    ActiveWindow.Zoom = 90
    On Error GoTo 0
End Sub


Private Sub WriteTableHeader(ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal vals As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals): ws.Cells(rowNum, colNum + i).Value = vals(i): Next i
    With ws.Range(ws.Cells(rowNum, colNum), ws.Cells(rowNum, colNum + UBound(vals) - LBound(vals)))
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(31, 78, 121): .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(rowNum).RowHeight = 28
End Sub


Private Function GetOrCreateDashboardSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = DASHBOARD_SHEET
    End If
    ws.Cells.Clear
    Set GetOrCreateDashboardSheet = ws
End Function


' Adds the company logo to the dashboard's white logo zone (A1:B1).
' Prefers aoti-logo.png beside the workbook (works headless, no clipboard);
' falls back to copying the embedded AOTI_Logo shape from START HERE when
' Excel is visible. Any existing logo on the sheet is removed first, so
' re-runs stay idempotent. Silent no-op when no source is available.
Private Sub AddDashboardLogo(ws As Worksheet)
    Dim src As Shape, dst As Shape
    Dim logoH As Double, logoW As Double, logoPath As String
    Dim oldSU As Boolean

    logoH = 30
    oldSU = Application.ScreenUpdating
    Application.ScreenUpdating = True

    ' Remove any logo already on this dashboard.
    On Error Resume Next
    ws.Shapes(LOGO_SHAPE).Delete
    On Error GoTo 0

    ' 1) Load from aoti-logo.png next to the workbook (automation-safe).
    logoPath = ThisWorkbook.Path & Application.PathSeparator & LOGO_FILE
    On Error Resume Next
    Set dst = ws.Shapes.AddPicture(Filename:=logoPath, LinkToFile:=False, _
                                   SaveWithDocument:=True, Left:=0, Top:=0, Width:=100, Height:=logoH)
    On Error GoTo 0
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

    On Error Resume Next
    Set src = ThisWorkbook.Worksheets(START_SHEET).Shapes(LOGO_SHAPE)
    If src Is Nothing Then
        Application.ScreenUpdating = oldSU
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    src.Copy
    ws.Paste
    On Error Resume Next
    Set dst = ws.Shapes(ws.Shapes.Count)
    On Error GoTo 0
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
End Sub


' Top-left placement inside the white A1:B1 logo zone, vertically centered.
Private Sub PositionDashboardLogo(ws As Worksheet, shp As Shape, ByVal logoH As Double, ByVal logoW As Double)
    shp.Left = ws.Range("A1").Left + 6
    shp.Top = ws.Range("A1").Top + (ws.Rows(1).Height - logoH) / 2
    shp.Placement = xlMove
End Sub


Private Sub SortPendingPickup(ws As Worksheet, lastRow As Long, lastCol As Long)
    If lastRow <= 2 Then Exit Sub
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub


Private Sub SortByColumnA(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub