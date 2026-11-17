Attribute VB_Name = "PM_Report_Setup"
Option Explicit

' ==================== WORKBOOK / SHEET SETUP + PATH / STATE ====================
' Sheet creation/retrieval, settings, output folder, Excel state.
' Entry points: GetOrCreateStartSheet, GetOrCreateRawSheet,
'   GetOrCreateSettingsSheet, GetUniqueOutputPath, RestoreExcelState
Public Function GetOrCreateStartSheet() As Worksheet
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

    ' Protect the sheet: lock all cells, allow only button clicks
    ws.Cells.Locked = True
    ws.Protect
    ws.EnableSelection = xlNoSelection

    Set GetOrCreateStartSheet = ws
    Exit Function
LocalFail:
    LogErr "GetOrCreateStartSheet"
End Function


' One numbered step row: navy number cell + merged text on the given fill.

Public Sub AddStartStep(ws As Worksheet, ByVal rowNum As Long, ByVal text As String, ByVal fillColor As Long)
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

Public Sub AddStartBullet(ws As Worksheet, ByVal rowNum As Long, ByVal text As String)
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

Public Sub PositionStartLogo(ws As Worksheet)
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



Public Function FindSourceWorksheet(wb As Workbook, ByRef headerRowOut As Long) As Worksheet
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

Public Sub LoadSourceIntoRawData(masterWb As Workbook, src As Worksheet, headerRow As Long, lastRow As Long, lastCol As Long)
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



Public Function GetOrCreateRawSheet(wb As Workbook) As Worksheet
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



Public Function GetOrCreateSettingsSheet(wb As Workbook) As Worksheet
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



Public Function GetRememberedOutputFolder(wb As Workbook) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    On Error GoTo LocalFail
    If Not ws Is Nothing Then GetRememberedOutputFolder = Trim$(SafeStr(ws.Range("B2").Value))
    Exit Function
LocalFail:
    LogErr "GetRememberedOutputFolder"
End Function



Public Sub SaveOutputFolder(wb As Workbook, folderPath As String)
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


Public Function FolderExists(folderPath As String) As Boolean
    Dim clean As String

    On Error GoTo LocalFail
    clean = NormalizeFolder(folderPath)
    FolderExists = (Len(Dir(clean, vbDirectory)) > 0)
    Exit Function
LocalFail:
    LogErr "FolderExists"
End Function


' Strip any trailing path separator so joins never double up.

Public Function NormalizeFolder(ByVal folderPath As String) As String
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



Public Function GetUniqueOutputPath(folderPath As String, baseFileName As String) As String
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



Public Sub RestoreExcelState(oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean)
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


