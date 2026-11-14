# answer-dialogs.ps1
# Answers the PM report macro's modal dialogs so a run can finish unattended.
# Used with the real user path (BuildManagerTabs), which prompts for:
#   - the CURRENT raw-data report  (GetOpenFilename, title "Select the CURRENT ...")
#   - the output folder             (FileDialog, title "Choose the folder ...")
# and shows a final "Done!" / error MsgBox.
#
#   .\scripts\answer-dialogs.ps1 -SourcePath C:\path\to+export.xlsx
#
# Log: Logs\answer-dialogs.log, one timestamped line per answered dialog.
# Any other dialog is dismissed with Enter and logged; nothing is silently closed.
# Stops after 10 minutes.

param(
    [string]$SourcePath = (Join-Path (Split-Path -Parent $PSScriptRoot) "AOTIOpsTreatmentPMDatesResults_FAKE.xlsx"),
    [string]$LogPath    = (Join-Path (Split-Path -Parent $PSScriptRoot) "Logs\answer-dialogs.log")
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class Dlg {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, string lp);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    static string T(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
    static string C(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static List<IntPtr> ExcelDialogs() {
        var list = new List<IntPtr>();
        var pids = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("EXCEL")) pids.Add(p.Id);
        if (pids.Count == 0) return list;
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains((int)pid)) return true;
            list.Add(h);
            return true;
        }, IntPtr.Zero);
        return list;
    }
    public static string Text(IntPtr h) { return T(h); }
    public static string Class(IntPtr h) { return C(h); }
    public static string ChildText(IntPtr parent) {
        var sb = new StringBuilder();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) {
            var t = T(h);
            if (t.Length > 0) { if (sb.Length > 0) sb.Append(" | "); sb.Append(t); }
            return true;
        }, IntPtr.Zero);
        return sb.ToString();
    }
    public static void Focus(IntPtr h) { SetForegroundWindow(h); }
    public static IntPtr FindChild(IntPtr parent, string cls) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) {
            if (C(h) == cls) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@

$logDir = Split-Path -Parent $LogPath
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

function Log { param($m) $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $m"; Write-Output $line; Add-Content -LiteralPath $LogPath -Value $line }

Log "answerer started (pid $PID)"
$answered = @{}
$deadline = (Get-Date).AddMinutes(10)
try {
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 400
        $live = @{}
        foreach ($h in [Dlg]::ExcelDialogs()) { $live["$h"] = $h }
        foreach ($k in @($answered.Keys)) { if (-not $live.ContainsKey($k)) { [void]$answered.Remove($k) } }

        foreach ($k in $live.Keys) {
            if ($answered.ContainsKey($k)) { continue }
            $h     = $live[$k]
            $title = [Dlg]::Text($h)
            $cls   = [Dlg]::Class($h)

            if ($cls -eq "bosa_sdm_MsoFiles" -or $title -match 'Select the CURRENT') {
                Start-Sleep -Milliseconds 500
                [Dlg]::Focus($h) | Out-Null
                Start-Sleep -Milliseconds 200
                $edit = [Dlg]::FindChild($h, "Edit")
                if ($edit -ne [IntPtr]::Zero) {
                    [Dlg]::SendMessage($edit, 0x000C, [IntPtr]::Zero, $SourcePath) | Out-Null   # WM_SETTEXT
                    [Dlg]::SendMessage($h, 0x0111, [IntPtr]1, [IntPtr]::Zero) | Out-Null         # WM_COMMAND IDOK
                    Log "answered CURRENT picker -> $SourcePath"
                } else {
                    [System.Windows.Forms.SendKeys]::SendWait("^a")
                    [System.Windows.Forms.SendKeys]::SendWait($SourcePath)
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Log "answered CURRENT picker via SendKeys -> $SourcePath"
                }
                $answered[$k] = 1
            }
            elseif ($cls -eq "bosa_sdm_MsoFiles" -or $title -match 'Choose the folder') {
                Start-Sleep -Milliseconds 400
                [Dlg]::Focus($h) | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $answered[$k] = 1
                Log "accepted output folder dialog (title=$title)"
            }
            else {
                # MsgBox / prompt / anything else: log the full text, then dismiss.
                Start-Sleep -Milliseconds 300
                $text = [Dlg]::ChildText($h)
                Log "dialog: class=$cls title=$title text=$text"
                [Dlg]::Focus($h) | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $answered[$k] = 1
                Log "dismissed dialog (title=$title)"
            }
        }
    }
} finally { Log "answerer stopped" }