Attribute VB_Name = "PM_Report_Data"
Option Explicit

' ==================== DATA / FORMAT + PRIOR COMPARISON ====================
' Row classification, manager sorting, header resolution, prior-report diff.
' Entry points: BuildOutputArray, FindHeaderRow, LastUsedRow, SheetExists,
'   FindManager, GetOrCreateOutputSheet, DetectPriorReport, LoadPriorComparison
Public Function BuildOutputArray(ByRef data As Variant, ByVal rows As Collection, ByVal lastCol As Long) As Variant
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



Public Function GetHeaderArray(ByRef data As Variant, ByVal lastCol As Long) As Variant
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

Public Sub CopyRowFormats(src As Worksheet, dst As Worksheet, ByVal srcRow As Long, ByVal dstRow As Long, ByVal lastCol As Long)
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



Public Sub SortManagersByLastName(ByRef a As Variant)
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



Public Function LastNameKey(ByVal fullName As String) As String
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



Public Function FindHeaderRow(ws As Worksheet, header1 As String, header2 As String) As Long
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

Public Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerName As String) As Long
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



Public Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range

    On Error GoTo LocalFail
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
    Exit Function
LocalFail:
    LogErr "LastUsedRow"
End Function



Public Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range

    On Error GoTo LocalFail
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = f.Column
    Exit Function
LocalFail:
    LogErr "LastUsedColumn"
End Function



Public Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function



Public Function SafeSheetName(ByVal s As String) As String
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

Public Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function


' Case-insensitive manager lookup in mgrNames (1-based index, 0 = not found).

Public Function FindManager(ByVal mgrNames As Collection, ByVal name As String) As Long
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

Public Function GetOrCreateOutputSheet(wb As Workbook, sheetName As String, ByVal firstOutput As Boolean) As Worksheet
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



Public Function UniqueSheetName(wb As Workbook, ByVal baseName As String) As String
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

Public Function DetectPriorReport(ByVal outputFolder As String, _
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

Public Function PriorNameOverlap(ByVal priorPath As String, _
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

Public Sub SortPathsNewestFirst(ByRef paths() As String, ByVal n As Long)
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



Public Sub LoadPriorComparison(ByVal priorPath As String, _
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



Public Function IsManagerTab(ByVal sheetName As String) As Boolean
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

Public Function FindHeaderRowForSingle(ws As Worksheet, ByVal headerName As String) As Long
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



Public Function NormalizeName(ByVal v As Variant) As String
    NormalizeName = LCase$(Trim$(SafeStr(v)))
End Function



Public Function PriorNameExists(ByVal priorNames As Collection, ByVal nm As String) As Boolean
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

Public Function CountDiff(ByVal a As Collection, ByVal b As Collection) As Long
    Dim i As Long, n As Long
    For i = 1 To a.Count
        If Not PriorNameExists(b, CStr(a(i))) Then n = n + 1
    Next i
    CountDiff = n
End Function


' Prior PM count for one manager tab; -1 = not found on the prior report.

Public Function PriorMgrCount(ByVal priorMgrNames As Collection, ByVal priorMgrCounts As Collection, ByVal manager As String) As Long
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

