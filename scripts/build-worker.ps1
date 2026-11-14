# build_worker.ps1
# Headless worker for run_report.ps1: runs one report build in its own Excel
# instance and prints RESULT=... / EXCEPTION=... on stdout.
#
#   powershell -File build_worker.ps1 -Source <xlsx> -OutputFolder <dir>

param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$OutputFolder,
    [string]$Master = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Requests_Master_Template.xlsm")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Master)) { Write-Output "EXCEPTION=master not found: $Master"; exit 1 }
if (-not (Test-Path -LiteralPath $Source)) { Write-Output "EXCEPTION=source not found: $Source"; exit 1 }

$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$app.EnableEvents = $false
$app.AutomationSecurity = 1

$wb = $null
try {
    $wb = $app.Workbooks.Open($Master)
    $result = $app.Run("'$([IO.Path]::GetFileName($Master))'!RunReport", $Source, $OutputFolder, "")
    Write-Output "RESULT=$result"
} catch {
    Write-Output "EXCEPTION=$($_.Exception.Message)"
} finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch { } }
    try { $app.Quit() } catch { }
}