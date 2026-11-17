# build-and-package.ps1
#
# Full build pipeline: compile gate -> sync template -> run report ->
# regenerate macro-free template -> rebuild ZIP -> commit & push to git.
#
# Usage:
#   .\scripts\build-and-package.ps1
#   .\scripts\build-and-package.ps1 -SkipPush   # commit only, don't push
#   .\scripts\build-and-package.ps1 -SkipRun     # skip the report run (compile+sync only)

param(
    [switch]$SkipPush,
    [switch]$SkipRun,
    [string]$Source = (Join-Path (Split-Path -Parent $PSScriptRoot) "AOTIOpsTreatmentPMDatesResults_FAKE.xlsx")
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$scripts = $PSScriptRoot
$logDir = Join-Path $root "Logs"
$xlsm = Join-Path $root "PM_Requests_Master_Template.xlsm"
$macroFree = Join-Path $root "PM_Requests_Master_Template_(No Macros).xlsx"
$zipPath = Join-Path $root "Project-Local-Build.zip"

Write-Host "=== STEP 1: Compile gate ==="
& (Join-Path $scripts "test-compile.ps1")
if ($LASTEXITCODE -ne 0) { Write-Host "BUILD FAILED: compile gate"; exit 1 }

Write-Host "=== STEP 2: Sync template ==="
& (Join-Path $scripts "sync_template.ps1")
if ($LASTEXITCODE -ne 0) { Write-Host "BUILD FAILED: sync template"; exit 1 }

if (-not $SkipRun) {
    Write-Host "=== STEP 3: Run report (full build verification) ==="
    & (Join-Path $scripts "run-report.ps1") -Source $Source
    if ($LASTEXITCODE -ne 0) { Write-Host "BUILD FAILED: report run"; exit 1 }
}

Write-Host "=== STEP 4: Regenerate macro-free template ==="
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsm)
if (Test-Path -LiteralPath $macroFree) { Remove-Item -LiteralPath $macroFree -Force }
# xlOpenXMLWorkbook = 51
$wb.SaveAs($macroFree, 51)
$wb.Close()
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "  Saved: $macroFree"

Write-Host "=== STEP 5: Rebuild Project-Local-Build.zip ==="
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
$basFiles = Get-ChildItem -Path $root -Filter "*.bas" | Select-Object -ExpandProperty FullName
$clsFiles = Get-ChildItem -Path $root -Filter "*.cls" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
$frmFiles = Get-ChildItem -Path $root -Filter "*.frm" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
$buildFiles = @($macroFree) + $basFiles + $clsFiles + $frmFiles + (Join-Path $root "BUILD-INSTRUCTIONS.txt")
Compress-Archive -Path $buildFiles -DestinationPath $zipPath -Force
Write-Host "  ZIP: $zipPath"

Write-Host "=== STEP 6: Commit & push ==="
Push-Location $root
git add "PM_Requests_Master_Template_(No Macros).xlsx" "Project-Local-Build.zip" "BUILD-INSTRUCTIONS.txt" "PM_Report_Macro.bas" "PM_Requests_Master_Template.xlsm"
$status = git status --porcelain
if ($status) {
    git commit -m "chore: rebuild build package after successful build"
    if (-not $SkipPush) {
        git push
    }
} else {
    Write-Host "  Nothing to commit."
}
Pop-Location

Write-Host ""
Write-Host "=== BUILD & PACKAGE COMPLETE ==="
