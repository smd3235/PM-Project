# test-compile.ps1
# Compile gate for PM_Report_Macro.bas: compiles the WHOLE project, not just the
# procedures a probe happens to call.
#
#   .\scripts\test-compile.ps1
#   .\scripts\test-compile.ps1 -SelfTest      # proves the detector actually detects
#
# How it works:
#  1. copies the master to Logs\_compile_test.xlsm (the master is never touched),
#  2. starts watch-compile-error.ps1, which records and dismisses every modal
#     dialog (Logs\watcher.log),
#  3. imports PM_Report_Macro.bas into the copy, so the gate tests the SOURCE and
#     not whatever happens to be embedded in the master,
#  4. triggers "Debug > Compile VBAProject" through the VBE command bar (Debug
#     menu, control id 578). That compiles every procedure in every module, which
#     plain Application.Run does NOT do: VBA compiles one procedure at a time, so
#     a typo in a procedure nobody calls only surfaces here.
#  5. treats the result as a PASS only when all three hold: the compile command
#     was available and executed, no error dialog appeared, and the probe marker
#     was written. "No dialog" alone is not a pass - a silent no-op would look
#     identical.
#  6. on failure prints the dialog text plus, when the VBE is still alive, the
#     exact offending token, module, and line.
#
# -SelfTest injects a deliberately broken procedure, asserts the gate FAILS, then
# removes it and asserts the gate passes. Use it after editing this script, and
# whenever a gate result looks too good: a fixture built but never saved is the
# classic way to get a false pass.
#
# Logs: Logs\compile-test.log (this script), Logs\watcher.log (dialog text).

param(
    [string]$Master           = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Requests_Master_Template.xlsm"),
    [string]$Bas              = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Report_Macro.bas"),
    [string]$ModuleName       = "PM_Report_Macro",
    [string]$DebugMenuCaption = "Debug",
    [int]$CompileCommandId    = 578,
    [int]$ProbeTimeoutSec     = 30,
    [int]$CompileWaitSec      = 1,
    [int]$CompilePollMs       = 100,
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$logDir      = Join-Path $ProjectRoot "Logs"
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$testCopy = Join-Path $logDir "_compile_test.xlsm"
$marker   = Join-Path $logDir "_compile_probe_ran.txt"
$watcher  = Join-Path $PSScriptRoot "watch-compile-error.ps1"
$watchLog = Join-Path $logDir "watcher.log"
$testLog  = Join-Path $logDir "compile-test.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -LiteralPath $testLog -Value $line
}

function Stop-Excel {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

function Get-CompileControl($vbe) {
    # The Debug menu is Item(6) of the VBE menu bar, but match on the caption so a
    # different VBE layout still works. Index access is deliberate: foreach over
    # these COM collections yields nothing in PowerShell 5.1.
    $menuBar = $vbe.CommandBars.Item(1)
    for ($i = 1; $i -le $menuBar.Controls.Count; $i++) {
        $menu = $menuBar.Controls.Item($i)
        if ($menu.Caption -notlike "*$DebugMenuCaption*") { continue }
        for ($j = 1; $j -le $menu.Controls.Count; $j++) {
            $item = $menu.Controls.Item($j)
            if ($item.Id -eq $CompileCommandId) { return $item }
            if ($item.Caption -like "*Compile*")    { return $item }
        }
    }
    return $null
}

function Get-WatcherTail([int]$fromLine) {
    if (-not (Test-Path -LiteralPath $watchLog)) { return @() }
    $all = @(Get-Content -LiteralPath $watchLog)
    if ($all.Count -le $fromLine) { return @() }
    return @($all[$fromLine..($all.Count - 1)])
}

function Get-CompileError([string[]]$fresh) {
    foreach ($l in $fresh) {
        if ($l -match 'Compile error|Variable not defined|Sub or Function not defined|Expected array|Expected:|Invalid outside') { return $true }
    }
    return $false
}

# Runs one whole-project compile. Returns @{ Error; Text; Waited }
function Invoke-Compile {
    $ctrl = $null
    try { $ctrl = Get-CompileControl $excel.VBE }
    catch { return @{ Error = "could not read the VBE command bars: $($_.Exception.Message)"; Text = @() } }

    if ($null -eq $ctrl) { return @{ Error = "Debug > Compile control not found in the VBE"; Text = @() } }

    try {
        if (-not $ctrl.Enabled) { return @{ Error = "the compile command is disabled (no active VBA project?)"; Text = @() } }
    } catch { }

    try { $ctrl.Execute() }
    catch { return @{ Error = "compile command raised: $($_.Exception.Message)"; Text = @() } }

    # Execute() returns in ~2 ms even though the compile runs after it, so poll.
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $waited = 0
    while ($sw.Elapsed.TotalSeconds -lt $CompileWaitSec) {
        Start-Sleep -Milliseconds $CompilePollMs
        $waited = [int]$sw.Elapsed.TotalSeconds
        $fresh = Get-WatcherTail $script:mark
        if ((Get-CompileError $fresh)) {
            return @{ Error = "the project does not compile"; Text = $fresh; Waited = $waited }
        }
    }
    return @{ Error = $null; Text = @(); Waited = $waited }
}

# 0. Clean slate
Stop-Excel
if (Test-Path -LiteralPath $testCopy) { Remove-Item -LiteralPath $testCopy -Force }
if (Test-Path -LiteralPath $marker)   { Remove-Item -LiteralPath $marker -Force }
Copy-Item -LiteralPath $Master -Destination $testCopy -Force
Log "=== Compile test started (master: $([IO.Path]::GetFileName($Master))) ==="

# 1. Dialog watcher in its own process
$watchProc = Start-Process powershell -WindowStyle Hidden -PassThru `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$watcher`""
Log "Dialog watcher started (pid $($watchProc.Id)) -> $watchLog"
Start-Sleep -Seconds 3

$script:mark = 0
$script:mark = @(Get-Content -LiteralPath $watchLog -ErrorAction SilentlyContinue).Count

# 2. Our own Excel instance
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel   = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$wb = $excel.Workbooks.Open($testCopy)

# 2b. Test the SOURCE, not whatever happens to be embedded in the master.
try {
    if (-not (Test-Path -LiteralPath $Bas)) { throw "Bas not found: $Bas" }
    foreach ($c in @($wb.VBProject.VBComponents)) {
        if ($c.Name -eq $ModuleName -or $c.Name -eq "CreateManagerTabs") { $wb.VBProject.VBComponents.Remove($c) }
    }
    $null = $wb.VBProject.VBComponents.Import($Bas)
    Log "Imported $ModuleName from $([IO.Path]::GetFileName($Bas)) (gate tests the source)."
} catch {
    Log "Could not import the .bas: $($_.Exception.Message)"
}

$after   = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$newPids = @($after | Where-Object { $before -notcontains $_ })
$excelPid = if ($newPids.Count -gt 0) { $newPids[0] } else { $null }
Log "Excel instance: $(if ($excelPid) { "pid $excelPid" } else { 'attached to an existing instance' })"

# 3. Watchdog: anything slower than ProbeTimeoutSec is a stall, not a compile
$watchdog = $null
if ($excelPid) {
    $watchdog = Start-Job -ScriptBlock {
        param($tp, $secs)
        Start-Sleep -Seconds $secs
        if (Get-Process -Id $tp -ErrorAction SilentlyContinue) {
            Stop-Process -Id $tp -Force -ErrorAction SilentlyContinue
        }
    } -ArgumentList $excelPid, $ProbeTimeoutSec
}

$selfTestNote = ""
$selfTestBroken = $false

# 4. Optional self-test: prove the detector detects.
if ($SelfTest) {
    Log "SELFTEST: injecting a deliberately broken procedure..."
    $mod = $wb.VBProject.VBComponents($ModuleName).CodeModule
    $mod.AddFromString("Private Sub SelfTestBroken()" + "`r`n" + "    NoSuchNameHere = 1" + "`r`n" + "End Sub")

    # Verify the fixture really contains the defect. A fixture that was never
    # saved is the classic false pass.
    if ($mod.Lines(1, $mod.CountOfLines) -notmatch "SelfTestBroken") {
        $selfTestBroken = $true
        $selfTestNote = "SELFTEST FAILED: the broken fixture is not in the module, so a pass would mean nothing."
    } else {
        $script:mark = @(Get-Content -LiteralPath $watchLog -ErrorAction SilentlyContinue).Count
        $r = Invoke-Compile
        if (-not $r.Error) {
            $selfTestBroken = $true
            $selfTestNote = "SELFTEST FAILED: the injected compile error was not detected."
        } else {
            $selfTestNote = "SELFTEST: injected error detected as expected."
            Log $selfTestNote

            # Advance the mark past the lines the broken compile just produced,
            # otherwise the clean compile below re-reads them and looks like a
            # second failure.
            $script:mark = @(Get-Content -LiteralPath $watchLog -ErrorAction SilentlyContinue).Count

            $start = $mod.ProcStartLine("SelfTestBroken", 0)
            $count = $mod.ProcCountLines("SelfTestBroken", 0)
            $mod.DeleteLines($start, $count)
            Log "SELFTEST: removed the broken procedure; compiling the clean project."
        }
    }
}

# 5. The real compile
$compileError = $null
$compileText  = @()
$res = Invoke-Compile
if ($res.Error) {
    $compileError = $res.Error
    $compileText  = $res.Text
    Log "compile result: FAILED - $($res.Error) (waited $($res.Waited)s)"
} else {
    Log "compile result: clean (waited $($res.Waited)s)"
}

# 6. Where exactly? Ask the VBE which token it selected.
$where = ""
if ($compileError -and $excelPid) {
    try {
        $pane = $excel.VBE.ActiveCodePane
        if ($pane) {
            $ln = 0; $sc = 0; $el = 0; $ec = 0
            $pane.GetSelection([ref]$ln, [ref]$sc, [ref]$el, [ref]$ec)
            $cm = $pane.CodeModule
            $lineText = $cm.Lines($ln, 1)
            $token = ""
            if ($ec -gt $sc) { $token = $lineText.Substring([Math]::Max(0, $sc - 1), $ec - $sc) }
            $where = "module=$($cm.Name) line=$ln token=[$token] lineText=$lineText"
        }
    } catch {
        $where = "(VBE cursor unavailable: $($_.Exception.Message))"
    }
}

# 7. Execution probe: proves the module still runs, not just parses
$probeCode = @"
Public Sub Probe()
    Dim ws As Worksheet, c1 As Collection, c2 As Collection, c3 As Collection
    Set c1 = New Collection: Set c2 = New Collection: Set c3 = New Collection
    c1.Add "Ann Smith"
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Cells(9, 1).Value = "Name"
    ws.Cells(9, 2).Value = "PMs"
    ws.Cells(10, 1).Value = "Zed Alpha"
    ws.Cells(10, 2).Value = 1
    ws.Cells(11, 1).Value = "Ann Smith"
    ws.Cells(11, 2).Value = 2
    FormatManagerSheet ws, "Probe", c1, c2, c3, 1
    ApplyReaderFriendlyFormatting ws, 9, 11, 2
    SortByColumnA ws, 9, 11, 2
    AddPolishedCard ws, "A1:C3", "Probe", 1, RGB(1, 2, 3), False, "TEST"
    AddPolishedTextCard ws, "A5:L7", "Probe", "text", RGB(1, 2, 3), "TEST"
    ColorReportTabs ThisWorkbook
    AddManagerChart ws, 9, 11, 13
    AddDashboardLogo ws
    Open "$marker" For Output As #1
    Print #1, "ran"
    Close #1
End Sub
"@

$runError = $null
if (-not $compileError) {
    try {
        $mod = $wb.VBProject.VBComponents($ModuleName).CodeModule
        $mod.AddFromString($probeCode)
        Log "Running probe (checks the formatting call graph executes)..."
        $excel.Run("Probe")
    } catch {
        $runError = $_.Exception.Message
        Log "Probe raised: $runError"
    }
} else {
    Log "Skipping the probe: the project does not compile."
}

# 8. Verdict - a pass needs all three: command ran, no dialog, marker written
Start-Sleep -Milliseconds 400
$markerOk = (Test-Path -LiteralPath $marker)
$ok = ((-not $compileError) -and (-not $runError) -and $markerOk -and (-not $selfTestBroken))

if ($compileError) {
    Log "RESULT: COMPILE FAILED - $compileError"
    if ($where) { Log "         $where" }
    if ($compileText.Count -gt 0) {
        Log "Captured dialog text (Logs\watcher.log):"
        foreach ($l in $compileText) { if ("$l".Trim()) { Log "  $l" } }
    }
} elseif ($runError) {
    Log "RESULT: COMPILE OK, PROBE FAILED - $runError"
    $ok = $false
} elseif (-not $markerOk) {
    Log "RESULT: INCONCLUSIVE - no dialog, but the probe marker was never written (Excel killed, or the probe never ran)."
    $ok = $false
} elseif ($ok) {
    Log "RESULT: COMPILE OK - whole project compiled and the probe executed."
} else {
    Log "RESULT: FAILED - see the self-test note below."
}

if ($selfTestNote) { Log $selfTestNote }

# 9. Cleanup
if ($watchdog) { Stop-Job $watchdog -ErrorAction SilentlyContinue; Remove-Job $watchdog -Force -ErrorAction SilentlyContinue }
try { if ($wb) { $wb.Close($false) } } catch { Log "Close failed: $($_.Exception.Message)" }
try { $excel.Quit() } catch { }
Stop-Excel
try { if ($watchProc) { Stop-Process -Id $watchProc.Id -Force -ErrorAction SilentlyContinue } } catch { }
Remove-Item -LiteralPath $testCopy -Force -ErrorAction SilentlyContinue
Log "=== Compile test finished: $(if ($ok) { 'OK' } else { 'FAILED' }) ==="

if ($ok) { exit 0 } else { exit 1 }