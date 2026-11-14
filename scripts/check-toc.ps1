# check-toc.ps1
# Verifies the MAP block at the top of PM_Report_Macro.bas against the file.
#
#   .\scripts\check-toc.ps1            # read-only check, prints drift
#   .\scripts\check-toc.ps1 -Fix       # rewrites the MAP block in place
#   .\scripts\check-toc.ps1 -Bas ..\Other.bas
#
# The MAP is a snapshot: section start lines and the procedure names in each
# section. It is documentation, so it is checked like documentation. -Fix
# rewrites it and re-checks, because a rewrite changes every line number after
# the block; it stops as soon as the block is stable.
#
# Exit code 0 = MAP matches, 1 = drift (and -Fix could not settle).

param(
    [string]$Bas = (Join-Path (Split-Path -Parent $PSScriptRoot) "PM_Report_Macro.bas"),
    [switch]$Fix
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Bas)) { throw "Bas not found: $Bas" }

$bannerRe = "^' =+ (?<title>.+?) =+\s*$"
$procRe   = "^(?:Public |Private )?(?:Sub|Function) (?<name>[A-Za-z_][A-Za-z0-9_]*)"
$mapStartRe = "^' MAP OF THIS MODULE"
$mapEndRe   = "^' =+\s*$"
$entryRe  = "^'\s*\[\s*(?<line>\d+)\]\s*(?<title>.+?)\s*$"
$itemRe   = "^'\s{2,}(?<items>[A-Za-z_].*?)\s*$"

function Read-Map {
    param([string[]]$Lines)

    $map = $null
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match $mapStartRe) {
            $start = $i
            break
        }
    }
    if ($null -eq $start) { return $null }

    $entries = New-Object System.Collections.ArrayList
    $cur = $null
    for ($i = $start; $i -lt $Lines.Count; $i++) {
        $l = $Lines[$i]
        if ($i -gt $start -and $l -match $mapEndRe) { break }
        if ($l -match $entryRe) {
            $cur = [pscustomobject]@{ Line = [int]$Matches['line']; Title = $Matches['title'].Trim(); Items = New-Object System.Collections.ArrayList }
            [void]$entries.Add($cur)
            continue
        }
        if ($null -ne $cur -and $l -match $itemRe) {
            foreach ($n in ($Matches['items'] -split '\s*,\s*')) {
                if ($n) { [void]$cur.Items.Add($n.Trim()) }
            }
        }
    }
    return ,$entries
}

function Read-Actual {
    param([string[]]$Lines)

    $sections = New-Object System.Collections.ArrayList
    $cur = $null
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $l = $Lines[$i]
        if ($l -match $bannerRe) {
            $cur = [pscustomobject]@{ Line = $i + 1; Title = $Matches['title'].Trim(); Items = New-Object System.Collections.ArrayList }
            [void]$sections.Add($cur)
            continue
        }
        if ($null -ne $cur -and $l -match $procRe) { [void]$cur.Items.Add($Matches['name']) }
    }
    return $sections
}

function Format-Map {
    param($Sections, $PublicProcs)

    $out = New-Object System.Collections.ArrayList
    [void]$out.Add("' ============================================================")
    [void]$out.Add("' MAP OF THIS MODULE - snapshot of section starts and procedures.")
    [void]$out.Add("' Generated and checked by scripts\check-toc.ps1 (add -Fix to rewrite).")
    [void]$out.Add("' Reading order for a change: CORE first, then the section you touch.")
    [void]$out.Add("'")

    $totalProcs = 0
    foreach ($s in $Sections) { $totalProcs += $s.Items.Count }
    [void]$out.Add(("' {0} sections, {1} procedures." -f $Sections.Count, $totalProcs))

    if ($PublicProcs -and @($PublicProcs).Count -gt 0) {
        $line = "' Public (outside callers):"
        foreach ($p in @($PublicProcs)) {
            $piece = $p
            if ($p -ne @($PublicProcs)[-1]) { $piece = $piece + "," }
            if (($line.Length + 1 + $piece.Length) -gt 96) { [void]$out.Add($line); $line = "'        " }
            if ($line.Trim().Length -le 1) { $line = $line + $piece } else { $line = $line + " " + $piece }
        }
        [void]$out.Add($line)
    }

    [void]$out.Add("'")
    [void]$out.Add("'   #  SECTION (line)                          procedures")
    [void]$out.Add("'")

    foreach ($s in $Sections) {
        $head = "[{0,5}] {1}" -f $s.Line, $s.Title
        [void]$out.Add("'   " + $head)
        $line = "'"
        for ($k = 0; $k -lt $s.Items.Count; $k++) {
            $piece = $s.Items[$k]
            if ($k -lt $s.Items.Count - 1) { $piece = $piece + "," }
            if (($line.Length + 1 + $piece.Length) -gt 96 -and $line.Trim().Length -gt 1) {
                [void]$out.Add($line)
                $line = "'"
            }
            if ($line.Trim().Length -le 1) { $line = "'        " + $piece } else { $line = $line + " " + $piece }
        }
        if ($line.Trim().Length -gt 1) { [void]$out.Add($line) }
    }

    [void]$out.Add("'")
    [void]$out.Add("' ============================================================")
    return $out
}

$lines = [System.IO.File]::ReadAllLines($Bas)
$map = Read-Map $lines
$actual = Read-Actual $lines

$drift = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt [Math]::Max($map.Count, $actual.Count); $i++) {
    $m = if ($i -lt $map.Count) { $map[$i] } else { $null }
    $a = if ($i -lt $actual.Count) { $actual[$i] } else { $null }

    if ($null -eq $a) { [void]$drift.Add("MAP lists '$($m.Title)' but the file has no such section"); continue }
    if ($null -eq $m) { [void]$drift.Add("file has section '$($a.Title)' at line $($a.Line) that the MAP does not list"); continue }

    if ($m.Line -ne $a.Line) { [void]$drift.Add("$($a.Title): MAP says line $($m.Line), file has $($a.Line)") }
    if ($m.Title -ne $a.Title) { [void]$drift.Add("line $($a.Line): MAP title '$($m.Title)' vs file '$($a.Title)'") }

    $missing = @($a.Items | Where-Object { $m.Items -notcontains $_ })
    $extra   = @($m.Items | Where-Object { $a.Items -notcontains $_ })
    foreach ($x in $missing) { [void]$drift.Add("$($a.Title): '$x' is in the file but not in the MAP") }
    foreach ($x in $extra)   { [void]$drift.Add("$($a.Title): '$x' is in the MAP but not in the file") }
}

if ($null -eq $map -and -not $Fix) {
    Write-Host "No MAP block found in $Bas (expected a line starting with '$ MAP OF THIS MODULE')."
    Write-Host "Run .\scripts\check-toc.ps1 -Fix to create it."
    exit 1
}

if ($drift.Count -eq 0 -and $null -ne $map) {
    $procs = 0
    foreach ($s in $actual) { $procs += $s.Items.Count }
    Write-Host "TOC OK: $($actual.Count) sections, $procs procedures, line numbers match."
    exit 0
}

if ($null -ne $map) {
    Write-Host "TOC DRIFT ($($drift.Count)):"
    foreach ($d in $drift) { Write-Host "  - $d" }
}

if (-not $Fix) {
    Write-Host ""
    Write-Host "Run .\scripts\check-toc.ps1 -Fix to rewrite the MAP block."
    exit 1
}

# Rewrite. The block length changes the numbering of everything after it, so
# regenerate from the rewritten text until the numbers stop moving.
$blockRe = "(?ms)^' =+\r?\n' MAP OF THIS MODULE.*?^' =+\s*(?=\r?\n|$)"

for ($attempt = 1; $attempt -le 4; $attempt++) {
    $lines = [System.IO.File]::ReadAllLines($Bas)

    $text = [System.IO.File]::ReadAllText($Bas)
    $lines = $text -split "\r\n|\r|\n"
    $sections = Read-Actual $lines
    $publicProcs = @($lines | ForEach-Object {
        if ($_ -match "^Public (?:Sub|Function) ([A-Za-z_][A-Za-z0-9_]*)") { $Matches[1] }
    } | Sort-Object)
    $block = (Format-Map $sections $publicProcs) -join "`r`n"

    if ($text -match $blockRe) {
        $newText = [regex]::Replace($text, $blockRe, [System.Text.RegularExpressions.MatchEvaluator] { param($m) $block }, 1)
    } else {
        # No block yet: insert after the leading comment header, before the
        # first real code line (line 1 is the Attribute line, not a comment).
        $idx = $lines.Count
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $t = $lines[$i].Trim()
            if ($t -eq "") { continue }
            if ($t -match "^'") { continue }
            if ($t -match "^Attribute VB_") { continue }
            if ($t -match "^Option (Explicit|Base|Compare|Private Module)") { continue }
            $idx = $i
            break
        }
        $head = if ($idx -gt 0) { $lines[0..($idx - 1)] -join "`r`n" } else { "" }
        $tail = if ($idx -lt $lines.Count) { $lines[$idx..($lines.Count - 1)] -join "`r`n" } else { "" }
        $newText = ($head.TrimEnd("`r", "`n") + "`r`n") + $block + "`r`n" + $tail
    }

    $newText = $newText.TrimEnd("`r", "`n") + "`r`n"
    [System.IO.File]::WriteAllText($Bas, $newText, (New-Object System.Text.UTF8Encoding($false)))

    $check = [System.IO.File]::ReadAllLines($Bas)
    $m2 = Read-Map $check
    $a2 = Read-Actual $check
    $mapLines = @($check | Where-Object { $_ -match $mapStartRe }).Count
    $stable = ($null -ne $m2 -and $mapLines -eq 1)
    if ($stable) {
        for ($i = 0; $i -lt [Math]::Max($m2.Count, $a2.Count); $i++) {
            $mm = if ($i -lt $m2.Count) { $m2[$i] } else { $null }
            $aa = if ($i -lt $a2.Count) { $a2[$i] } else { $null }
            if ($null -eq $mm -or $null -eq $aa -or $mm.Line -ne $aa.Line) { $stable = $false; break }
        }
    }

    Write-Host "attempt $attempt : MAP rewritten ($((($block -split "`r`n").Count)) lines, $mapLines MAP block(s) in file)"
    if ($stable) {
        $procs = 0
        foreach ($s in $a2) { $procs += $s.Items.Count }
        Write-Host "TOC OK after fix: $($a2.Count) sections, $procs procedures."
        exit 0
    }
}

Write-Host "TOC still drifting after 4 rewrites - inspect the block by hand."
exit 1