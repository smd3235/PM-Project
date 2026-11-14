# dump-sheet.ps1
# Read-only structural dump of one sheet: merged areas, column widths, row heights,
# per-cell font/fill/alignment, and shapes. Used to verify START HERE (or any layout
# sheet) after a change to the VBA.
#
#   .\scripts\dump-sheet.ps1
#   .\scripts\dump-sheet.ps1 -Path .\Reports\PM Requests 2026-11-14.xlsx -Sheet Dashboard
#
# Opens the workbook read-only, prints, and closes without saving. Nothing is written.
# Kills any EXCEL process left behind by COM.

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Sheet = "START HERE"
)

$ErrorActionPreference = "Stop"
# Excel resolves relative paths against its own current directory, not PowerShell's.
$Path = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3
$wb = $null
try {
    $wb = $xl.Workbooks.Open($Path, $false, $true)
    "SHEETS: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ", ")

    $ws = $wb.Worksheets.Item($Sheet)
    $ur = $ws.UsedRange
    $rowStart = $ur.Row
    $rowEnd = $ur.Row + $ur.Rows.Count - 1
    $colStart = $ur.Column
    $colEnd = $ur.Column + $ur.Columns.Count - 1
    "USED: $($ur.Address(0,0))  rows=$($ur.Rows.Count) cols=$($ur.Columns.Count)"
    "TABCOLOR: $($ws.Tab.Color)  GRIDLINES: $($ws.DisplayGridlines)  ZOOM: $($ws.Zoom)"

    $mergeAreas = @{}
    for ($r = $rowStart; $r -le $rowEnd; $r++) {
        for ($c = $colStart; $c -le $colEnd; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            if ($cell.MergeCells) {
                $addr = $cell.MergeArea.Address(0, 0)
                if (-not $mergeAreas.ContainsKey($addr)) {
                    $mergeAreas[$addr] = $cell.MergeArea.Cells.Item(1).Text
                }
            }
        }
    }
    foreach ($k in ($mergeAreas.Keys | Sort-Object)) { "MERGE $k = [$($mergeAreas[$k])]" }

    "COLWIDTH: " + (($colStart..$colEnd | ForEach-Object { "$($ws.Columns($_).ColumnWidth)" }) -join " | ")

    "ROWS:"
    for ($r = $rowStart; $r -le $rowEnd; $r++) {
        "  row $r height=$($ws.Rows($r).RowHeight)"
    }

    "CELLS:"
    for ($r = $rowStart; $r -le $rowEnd; $r++) {
        for ($c = $colStart; $c -le $colEnd; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            $txt = $cell.Text
            $fill = $cell.Interior.Color
            if ($txt -eq $null -or $txt -eq "") {
                if ($fill -ne 16777215) { "  $($cell.Address(0,0)) EMPTY fill=$fill" }
                continue
            }
            $f = $cell.Font
            "  $($cell.Address(0,0)) [$txt] font=$($f.Name),sz=$($f.Size),b=$($f.Bold),i=$($f.Italic),color=$($f.Color) fill=$fill halign=$($cell.HorizontalAlignment) valign=$($cell.VerticalAlignment) wrap=$($cell.WrapText) indent=$($cell.IndentLevel)"
        }
    }

    "SHAPES: $($ws.Shapes.Count)"
    foreach ($sh in $ws.Shapes) {
        $txt = ""
        try { if ($sh.TextFrame.HasText) { $txt = $sh.TextFrame.Characters.Text } } catch { }
        "  [$($sh.Name)] type=$($sh.Type) L=$([math]::Round($sh.Left,1)) T=$([math]::Round($sh.Top,1)) W=$([math]::Round($sh.Width,1)) H=$([math]::Round($sh.Height,1)) text=[$txt]"
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}