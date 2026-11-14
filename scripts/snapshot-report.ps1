# snapshot-report.ps1
# Exports one or more sheets of a workbook as PNG images, for documentation and
# design comparison. Each sheet is copied as a picture, pasted onto a throwaway
# chart, and exported from there: this captures fills, fonts, merges, and shapes
# exactly as Excel renders them, which a cell-by-cell dump cannot.
#
#   .\scripts\snapshot-report.ps1 -Path "PM_Requests_Master_Template.xlsm" -Sheets "START HERE"
#   .\scripts\snapshot-report.ps1 -Path "Reports\PM Requests 2026-11-13.xlsx" -Sheets Dashboard,"Derek Owens"
#
# Defaults to the newest workbook in Reports\. Images land in
# documentation\screenshots\ unless -OutFolder says otherwise.
# Exit codes: 0 = all requested sheets exported, 2 = no workbook, 3 = nothing exported.

[CmdletBinding()]
param(
    [string]   $Path,
    [string[]]$Sheets,
    [string]   $OutFolder,
    [string]   $Prefix,
    [switch]   $List
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $candidates = @(Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "Reports") -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending)
    if ($candidates.Count -eq 0) {
        Write-Host "No workbook given and Reports\ holds no .xlsx. Pass -Path." -ForegroundColor Red
        exit 2
    }
    $Path = $candidates[0].FullName
    Write-Host "Using newest report: $Path" -ForegroundColor DarkGray
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "Workbook not found: $Path" -ForegroundColor Red
    exit 2
}
$Path = (Resolve-Path -LiteralPath $Path).Path

if (-not $OutFolder) { $OutFolder = Join-Path $ProjectRoot "documentation\screenshots" }
if (-not (Test-Path -LiteralPath $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder | Out-Null }
if (-not $Prefix) { $Prefix = "report" }

function Save-SheetImage {
    param($Workbook, $Worksheet, [string]$PngPath)

    $used = $Worksheet.UsedRange
    $range = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($used.Row + $used.Rows.Count, $used.Column + $used.Columns.Count))
    $range.CopyPicture(1, 2) | Out-Null

    $chart = $Workbook.Charts.Add()
    try {
        $chart.Paste() | Out-Null
        Start-Sleep -Milliseconds 250
        if (Test-Path -LiteralPath $PngPath) { Remove-Item -LiteralPath $PngPath -Force }
        $chart.Export($PngPath)
    }
    finally {
        $chart.Delete()
    }
}

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
$wb = $null
$exported = 0
try {
    $wb = $xl.Workbooks.Open($Path, $false, $true)

    if ($List -or -not $Sheets) {
        Write-Host "Sheets in $(Split-Path -Leaf $Path):"
        foreach ($s in $wb.Worksheets) {
            $u = $s.UsedRange
            Write-Host ("  {0,-28} used {1}" -f $s.Name, $u.Address(0, 0))
        }
    }

    $targets = @()
    if ($Sheets) {
        foreach ($name in $Sheets) {
            $found = $null
            foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { $found = $s; break } }
            if ($found) { $targets += $found }
            else { Write-Host "Sheet not found, skipped: $name" -ForegroundColor Yellow }
        }
    }

    foreach ($ws in $targets) {
        $safe = ($ws.Name -replace "[^\w\-]+", "-").Trim("-").ToLower()
        $png = Join-Path $OutFolder ("{0}-{1}.png" -f $Prefix, $safe)
        Save-SheetImage -Workbook $wb -Worksheet $ws -PngPath $png
        $kb = [math]::Round((Get-Item -LiteralPath $png).Length / 1KB, 1)
        Write-Host ("  exported {0}  ({1} KB)" -f $png, $kb) -ForegroundColor Green
        $exported++
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

if ($exported -eq 0 -and -not $List -and $Sheets) {
    Write-Host "Nothing exported." -ForegroundColor Red
    exit 3
}
exit 0