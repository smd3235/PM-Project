---
description: Show the PM report run log and diagnose a stuck run
argument-hint: "[log-name-fragment] [tail-lines]"
---

Show the PM report run log for this project and diagnose stalls.

Logs are written by the macro to `Logs\` in the project root, one file per run,
named `PM_Report_yyyyMMdd_hhmmss.log`. Every line is:

```
yyyy-MM-dd hh:mm:ss  +<ms since previous line>ms  message
```

So a hang looks like one line followed by silence, and the line *before* the
silence names the step that hung. The full date is on every line, so an old line
can never be mistaken for a fresh one.

Log files, oldest first so the newest run is the **last** block printed (bottom):

!`powershell -NoProfile -Command "Get-ChildItem 'Logs\PM_Report_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | ForEach-Object { '=== ' + $_.Name + '  (' + $_.LastWriteTime + ')  ' + [math]::Round($_.Length/1KB,1) + ' KB'; Get-Content $_.FullName -Tail 40; '' }"

Slowest steps across the five newest logs (top 10 gaps, newest run last):

!`powershell -NoProfile -Command "Get-ChildItem 'Logs\PM_Report_*.log' | Sort-Object LastWriteTime | Select-Object -Last 5 | ForEach-Object { $n=$_.Name; Get-Content $_.FullName | Select-String -Pattern '\+(\d+)ms' -AllMatches | ForEach-Object { foreach ($m in $_.Matches) { [pscustomobject]@{ log=$n; ms=[int]$m.Groups[1].Value; line=$_.Line } } } } | Sort-Object ms -Descending | Select-Object -First 10 | Format-Table -AutoSize"

If `$ARGUMENTS` is given:
- First word: a filename fragment; show only matching logs (e.g. `/logs 20261114`).
- Second word: tail length per log (default 40).

Then report, concisely:
1. Which run is the newest, its last step, and the size of the final gap.
2. If any gap exceeds 10000 ms, name the step that hung and propose the fix. Stage
   labels map to `RunReport`: `fastmode`, `source open`, `raw load`, `data read`,
   `collections set`, `classified`, `detect prior`, `out workbook`, `manager tabs`,
   `pending skipped`, `dashboard`, `dashboard moved`, `raw copied`, `tabs colored`,
   `view set`, `saved`.
3. Any `!!! ERROR` line with its stage.

Do not rebuild anything unless asked; this command only reads logs.