Attribute VB_Name = "PM_Report_Output"
Option Explicit

' ==================== OUTPUT FORMATTING + DASHBOARD + EMAIL ====================
' Manager-tab formatting, dashboard, charts, email composition.
' Entry points: FormatManagerSheet, BuildDashboard, AddManagerChart,
'   CreateEmailFromReport, GenerateManagerEmail
Public Sub FormatManagerSheet(ws As Worksheet, ByVal manager As String, _
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
    Set btnShp = ws.Shapes.AddShape(5, ws.Range("K1").Left + 200, ws.Range("A3").Top + 2, 200, 24)
    btnShp.Name = "BtnGenEmail"
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

Public Sub HideUnusedCells(ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
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

Public Function CountInSheet(ws As Worksheet, ByVal value As String, ByVal lastRow As Long) As Long
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

Public Sub AddPolishedCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal value As Long, ByVal fillColor As Long, ByVal colorChange As Boolean, Optional ByVal eyebrow As String = "")
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

Public Sub AddPolishedTextCard(ws As Worksheet, ByVal address As String, ByVal label As String, ByVal valueText As String, ByVal fillColor As Long, Optional ByVal eyebrow As String = "")
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

Public Function ChangeColor(ByVal v As Long) As Long
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

Public Sub ApplyReaderFriendlyFormatting(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
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

Public Sub BuildDashboard(wb As Workbook, ByVal managers As Variant, _
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

Public Sub ColorReportTabs(wb As Workbook)
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

Public Sub AddManagerChart(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal repEnd As Long)
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



Public Sub WriteTableHeader(ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal vals As Variant)
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



Public Function GetOrCreateDashboardSheet(wb As Workbook) As Worksheet
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

Public Function ResolveLogoPath() As String
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

Public Sub AddDashboardLogo(ws As Worksheet)
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

Public Sub PositionDashboardLogo(ws As Worksheet, shp As Shape, ByVal logoH As Double, ByVal logoW As Double)
    On Error GoTo LocalFail
    shp.Left = ws.Range("A1").Left + 6
    shp.Top = ws.Range("A1").Top + (ws.Rows(1).Height - logoH) / 2
    shp.Placement = xlMove
    Exit Sub
LocalFail:
    LogErr "PositionDashboardLogo"
End Sub



Public Sub SortByColumnA(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
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


Public Function GetEmailTemplate() As String
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


Public Function GetCCEmails() As String
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


Public Function GetToEmails() As String
    On Error GoTo LocalFail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Manager Email Template")
    GetToEmails = CStr(ws.Range("D5").Value)
    Exit Function
LocalFail:
    LogErr "GetToEmails"
    GetToEmails = ""
End Function


Public Function GetEmailSubject() As String
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


Public Function GetEmailBody() As String
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

