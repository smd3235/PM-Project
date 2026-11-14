# sync_template.ps1
# Imports PM_Report_Macro.bas into the master workbook's embedded VBA project
# and rebuilds the START HERE tab (layout + CTA button).
#
#   .\scripts\sync_template.ps1
#   .\scripts\sync_template.ps1 -Template ..\Other.xlsm -Bas ..\Other.bas
#
# Idempotent: when the embedded module already matches the .bas file nothing is
# written, so no backup pile-up. A backup is taken only when the module really
# changes, and only the newest two backups are kept.
# Run this after every change to the .bas file.

param(
    [string]$Template    = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Requests_Master_Template.xlsm"),
    [string]$Bas         = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Report_Macro.bas"),
    [string]$ModuleName  = "PM_Report_Macro",
    [string]$LegacyModuleName = "CreateManagerTabs",   # pre-rename name; removed if still present
    [switch]$SkipUiRebuild
)

$ProjectRoot = Split-Path -Parent $PSScriptRoot

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Template)) { throw "Template not found: $Template" }
if (-not (Test-Path -LiteralPath $Bas))      { throw "Bas not found: $Bas" }

function Get-NormalizedBody([string]$Text) {
    # One canonical form: no Attribute lines, LF endings, no trailing blanks.
    $t = $Text -replace "\r\n", "`n"
    $t = $t -replace "(?m)^Attribute\s+VB_Name[^\n]*\n", ""
    return $t.TrimEnd()
}

# --- .bas body --------------------------------------------------------------
$basBody = Get-NormalizedBody (Get-Content -LiteralPath $Bas -Raw)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1      # msoAutomationSecurityLow: lets RebuildStartSheet run headless
$wb = $null
try {
    $wb = $excel.Workbooks.Open($Template)

    try {
        $vbp = $wb.VBProject
    } catch {
        throw "Cannot read the VBA project (needs 'Trust access to the VBA project object model' in Trust Center): $($_.Exception.Message)"
    }

    # --- drop the pre-rename module so no stale copy survives ----------------
    if ($LegacyModuleName -and ($LegacyModuleName -ne $ModuleName)) {
        foreach ($c in @($vbp.VBComponents)) {
            if ($c.Name -eq $LegacyModuleName) {
                $vbp.VBComponents.Remove($c)
                Write-Host "Removed legacy module '$LegacyModuleName'."
            }
        }
    }

    # --- embedded module body ----------------------------------------------
    $component = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $ModuleName) { $component = $c; break } }

    if ($component -ne $null) {
        $cm = $component.CodeModule
        $current = ""
        if ($cm.CountOfLines -gt 0) { $current = $cm.Lines(1, $cm.CountOfLines) }
    }

    if ($component -ne $null -and (Get-NormalizedBody $current) -eq $basBody) {
        Write-Host "Module '$ModuleName' is already up to date - nothing to do."
        $wb.Close($false)
        $wb = $null
        return
    }

    # --- backup (only when the module actually changes) ---------------------
    $backupDir = Join-Path $ProjectRoot "backups"
    if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
    $stamp  = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backup = Join-Path $backupDir ("{0}_{1}.xlsm" -f [IO.Path]::GetFileNameWithoutExtension($Template), $stamp)
    Copy-Item -LiteralPath $Template -Destination $backup
    Write-Host "Backup: $backup"

    # Keep only the newest two backups (names sort chronologically).
    Get-ChildItem -LiteralPath $backupDir -Filter "*.xlsm" |
        Sort-Object Name -Descending |
        Select-Object -Skip 2 |
        Remove-Item -Force

    # --- replace the module so there is never a duplicate ------------------
    if ($component -ne $null) { $vbp.VBComponents.Remove($component) }
    $null = $vbp.VBComponents.Import($Bas)
    Write-Host "Imported '$ModuleName' from $Bas"

    # --- rebuild START HERE + button (also the compile gate) ---------------
    if (-not $SkipUiRebuild) {
        $excel.Run(("'{0}'!RebuildStartSheet" -f $wb.Name))
        Write-Host "Rebuilt START HERE."
    }

    # Save to a scratch file inside the project and swap it in: Save() on this
    # workbook can block on an invisible modal in headless automation. Nothing is
    # written outside the project folder.
    $scratchDir = Join-Path $ProjectRoot "Logs"
    if (-not (Test-Path -LiteralPath $scratchDir)) { New-Item -ItemType Directory -Path $scratchDir | Out-Null }
    $tmp = Join-Path $scratchDir ("_sync_{0}.xlsm" -f $stamp)
    $wb.SaveAs($tmp, 52)            # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    $wb = $null
    [IO.File]::Copy($tmp, $Template, $true)
    Remove-Item -LiteralPath $tmp -Force
    Write-Host "Saved $Template"
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch { } }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}