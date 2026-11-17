Attribute VB_Name = "PM_Report_UI"
Option Explicit

' ==================== INTERACTIVE ENTRY (button) ====================
' UI shell: START HERE button and the dialogs around a build.
' Entry points: SetupReportButton, SetupEmailButtons,
'   SetupGenerateEmailButton, RebuildStartSheet, BuildManagerTabs
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


' Public entry point for sync_template.py: rebuild START HERE + button so the
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
' callable from the button, from build-worker.py, or from a test.
'
' Pipeline, in order (each step is a `stage` breadcrumb in the run log):
'   fastmode -> source open -> raw load -> data read -> classified ->
'   [detect prior] -> manager tabs -> dashboard -> dashboard moved ->
'   raw copied -> tabs colored -> view set -> saved.
'
' Contracts:
'   * Result string is the only output: "OK|savePath|pmTotal|skipped|prior"
'     or "ERROR|number|description @ stage". scripts\run-report.py parses it;
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


