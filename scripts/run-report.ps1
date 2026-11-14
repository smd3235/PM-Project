# run_report.ps1
# Builds one report with the real master template and a watchdog.
#
#   .\run_report.ps1
#   .\run_report.ps1 -Source .\Some_Export.xlsx -StallSeconds 15
#
# Watchdog rule: the macro writes one line per step to Logs\PM_Report_*.log.
# If no new line appears for -StallSeconds (default 10), the run is considered
# stuck: Excel is killed, the last log lines are printed, and the script exits 2
# so the offending step is visible instead of a silent hang.

param(
    [string]$Source        = (Join-Path (Split-Path -Parent $PSScriptRoot) "AOTIOpsTreatmentPMDatesResults_FAKE.xlsx"),
    [string]$OutputFolder  = (Join-Path (Split-Path -Parent $PSScriptRoot) "Reports"),
    [int]$StallSeconds     = 10,
    [int]$MaxSeconds       = 900
)

$ErrorActionPreference = "Stop"

$logDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Logs"
New-Item -ItemType Directory -Force -Path $logDir, $OutputFolder | Out-Null

if (-not (Test-Path -LiteralPath $Source))       { Write-Host "Source not found: $Source"; exit 1 }
if (-not (Test-Path -LiteralPath $OutputFolder)) { Write-Host "Output folder not found: $OutputFolder"; exit 1 }

function Kill-Excel {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

function Get-NewLog([string[]]$Before) {
    $cand = @(Get-ChildItem $logDir -Filter "PM_Report_*.log" -ErrorAction SilentlyContinue |
              Where-Object { $Before -notcontains $_.Name })
    if ($cand.Count -eq 0) { return $null }
    return ($cand | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}

function Show-Tail([string]$Path, [int]$Lines) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { Write-Host "(no log file)"; return }
    Write-Host "--- tail of $([IO.Path]::GetFileName($Path)) ---"
    Get-Content -LiteralPath $Path -Tail $Lines | ForEach-Object { Write-Host $_ }
    $gaps = Get-Content -LiteralPath $Path | Select-String -Pattern '\+(\d+)ms' -AllMatches |
            ForEach-Object { foreach ($m in $_.Matches) { [int]$m.Groups[1].Value } }
    if ($gaps) {
        $max = ($gaps | Measure-Object -Maximum).Maximum
        Write-Host "--- slowest step: $max ms (threshold $($StallSeconds * 1000) ms) ---"
    }
}

Write-Host "Killing stale Excel/Python..."
Kill-Excel
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$before    = @(Get-ChildItem $logDir -Filter "PM_Report_*.log" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
$workerOut = Join-Path $logDir "_worker_out.txt"
Remove-Item -LiteralPath $workerOut -ErrorAction SilentlyContinue

Write-Host "Starting build (source: $([IO.Path]::GetFileName($Source)))..."
$worker = Join-Path $PSScriptRoot "build-worker.ps1"
$workerArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ('"{0}"' -f $worker),
                 "-Master", ('"{0}"' -f (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Requests_Master_Template.xlsm")),
                 "-Source", ('"{0}"' -f $Source), "-OutputFolder", ('"{0}"' -f $OutputFolder))
$proc = Start-Process powershell -PassThru -WindowStyle Hidden `
    -ArgumentList $workerArgs -RedirectStandardOutput $workerOut

$startedAt = Get-Date
$deadline  = $startedAt.AddSeconds($MaxSeconds)
$lastWrite = $startedAt
$newLog    = $null
$stall     = $false
$timeout   = $false

try {
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 300
        if ($proc.HasExited) { break }

        $newLog = Get-NewLog $before

        if ($null -eq $newLog) {
            if (((Get-Date) - $startedAt).TotalSeconds -gt $StallSeconds) {
                $stall = $true
                Write-Host ""
                Write-Host "STALLED: no log file was created within $StallSeconds s."
                Write-Host "         Excel never reached the first logged step - look for a modal dialog"
                Write-Host "         (macro security, compile error, file picker) or a blocked COM call."
                break
            }
            continue
        }

        if ($newLog.LastWriteTime -gt $lastWrite) {
            $lastWrite = $newLog.LastWriteTime
            continue
        }

        if (((Get-Date) - $lastWrite).TotalSeconds -gt $StallSeconds) {
            $stall = $true
            Write-Host ""
            Write-Host "STALLED: no new log line for $StallSeconds s in $([IO.Path]::GetFileName($newLog.FullName))."
            Write-Host "         The last line names the step that hung:"
            Get-Content -LiteralPath $newLog.FullName -Tail 3 | ForEach-Object { Write-Host "         $_" }
            break
        }
    }

    if (-not $stall -and -not $proc.HasExited) {
        $timeout = $true
        Write-Host ""
        Write-Host "TIMEOUT: build exceeded $MaxSeconds s."
    }
}
finally {
    if ($stall -or $timeout) {
        Write-Host ""
        Write-Host "Killing Excel..."
        try { if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } } catch { }
        Kill-Excel
    }
}

if ($stall -or $timeout) {
    Write-Host ""
    Show-Tail $newLog.FullName 40
    Write-Host ""
    Write-Host "No report was produced. Fix the step above, then re-run."
    exit 2
}

# Normal completion.
$out = ""
if (Test-Path -LiteralPath $workerOut) { $out = (Get-Content -LiteralPath $workerOut -Raw) }
Write-Host ""
Write-Host ($out.Trim())
if ($out -match 'RESULT=OK') { $ok = $true } else { $ok = $false }

$newLog = Get-NewLog $before
if ($null -ne $newLog) { Show-Tail $newLog.FullName 25 }

if ($ok) { Write-Host "BUILD OK"; exit 0 }
Write-Host "BUILD FAILED"; exit 3