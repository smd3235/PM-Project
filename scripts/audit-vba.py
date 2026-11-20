# audit-vba.py
# Deep inspection audit of PM_Report_*.bas VBA modules.
# Checks the VBA-specific rules from AGENTS.md section 4 plus runtime-risk
# patterns that the compile gate cannot catch (compile-on-demand means a typo
# in an uncalled procedure survives every Application.Run).
#
# Severity levels:
#   CRITICAL - likely runtime bugs (type mismatch, null object, no error handling on I/O)
#   HIGH     - error-handling gaps (no LogErr in Fail handler, missing Exit before Fail)
#   MEDIUM   - logging / traceability (missing stage breadcrumbs)
#   INFO     - style / maintainability (long procedures, unused vars)
#
# Usage: python scripts/audit-vba.py [--fix]   (--fix is accepted for compat, read-only here)
import argparse
import os
import re
import sys

parser = argparse.ArgumentParser(description="Deep VBA code inspection audit")
parser.add_argument("--fix", action="store_true", help="accepted for compatibility (read-only)")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

bas_files = sorted(
    f for f in os.listdir(project_root)
    if f.startswith("PM_Report_") and f.endswith(".bas")
)
if not bas_files:
    print("No PM_Report_*.bas module files found in project root.", file=sys.stderr)
    sys.exit(1)

# ---------------------------------------------------------------------------
# Regexes
# ---------------------------------------------------------------------------
proc_re = re.compile(
    r"^\s*(Public|Private)\s+(Sub|Function|Property\s+(?:Get|Let|Set)?)\s+(\w+)",
    re.MULTILINE,
)
error_go_to_re = re.compile(r"On\s+Error\s+GoTo\s+\w+", re.IGNORECASE)
error_resume_next_re = re.compile(r"On\s+Error\s+Resume\s+Next", re.IGNORECASE)
error_go_to_0_re = re.compile(r"On\s+Error\s+GoTo\s+0", re.IGNORECASE)
exit_re = re.compile(r"^\s*Exit\s+(Sub|Function|Property)", re.IGNORECASE | re.MULTILINE)
fail_label_re = re.compile(r"^\s*(?:Fail|LocalFail|CleanUp|ExitProc):\s*", re.IGNORECASE | re.MULTILINE)
logerr_re = re.compile(r"LogErr\s*[\( ]", re.IGNORECASE)
logstage_re = re.compile(r"LogStage\s*[\( ]", re.IGNORECASE)
setstatus_re = re.compile(r"SetStatus\s*[\( ]", re.IGNORECASE)
cdbl_clng_re = re.compile(r"\b(?:CDbl|CLng|CInt|CSng)\s*\(", re.IGNORECASE)
isnumeric_re = re.compile(r"\bIsNumeric\s*\(", re.IGNORECASE)
ubound_re = re.compile(r"\bUBound\s*\(", re.IGNORECASE)
nothing_check_re = re.compile(r"\bIs\s+Nothing\b|\bNot\s+Is\s+Nothing\b|\bIsNot\s+\w+\b", re.IGNORECASE)
module_decl_re = re.compile(
    r"^\s*(Public|Private)\s+(?:Const\s+)?\w+\s+As\s", re.MULTILINE
)
err_raise_re = re.compile(r"Err\.Raise\s*[\( ]", re.IGNORECASE)
set_obj_re = re.compile(r"^\s*Set\s+(\w+)\s*=\s*", re.IGNORECASE | re.MULTILINE)
stage_assign_re = re.compile(r"^\s*stage\s*=\s*", re.IGNORECASE | re.MULTILINE)
dim_re = re.compile(r"^\s*Dim\s+(.+)", re.IGNORECASE | re.MULTILINE)

# Cell value access patterns that should be IsNumeric-guarded before conversion
cell_value_re = re.compile(r"\b(?:Cells|Range|data)\s*\([^)]*\)\.?(?:Value|Value2)?\b", re.IGNORECASE)

problems = []
checked = 0

for bas in bas_files:
    path = os.path.join(project_root, bas)
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    content = "".join(lines)
    checked += 1

    # ---- Find all procedures with their line spans ----
    procs = []
    for m in proc_re.finditer(content):
        start_line = content[: m.start()].count("\n") + 1  # 1-indexed
        procs.append((m.group(1), m.group(2), m.group(3), start_line))

    for idx, (vis, kind, name, start) in enumerate(procs):
        end = procs[idx + 1][3] - 1 if idx + 1 < len(procs) else len(lines)
        proc_lines = lines[start - 1:end]
        proc_text = "".join(proc_lines)
        # Strip comment lines (starting with ') for code-pattern matching
        code_only_lines = [l for l in proc_lines if not l.lstrip().startswith("'")]
        code_only_text = "".join(code_only_lines)

        has_err_handler = bool(error_go_to_re.search(code_only_text))
        has_resume_next = bool(error_resume_next_re.search(code_only_text))
        has_go_to_0 = bool(error_go_to_0_re.search(code_only_text))
        has_exit = bool(exit_re.search(code_only_text))
        has_logerr = bool(logerr_re.search(code_only_text))
        has_fail_label = bool(fail_label_re.search(code_only_text))
        has_logstage = bool(logstage_re.search(code_only_text))
        has_setstatus = bool(setstatus_re.search(code_only_text))
        has_err_raise = bool(err_raise_re.search(code_only_text))
        has_stage_assign = bool(stage_assign_re.search(code_only_text))
        proc_len = end - start + 1

        # CRITICAL: procedure raising Err.Raise but no error handler
        if has_err_raise and not has_err_handler:
            problems.append(("CRITICAL", f"{bas}:{start} {name}: Err.Raise used but no On Error GoTo handler"))

        # CRITICAL: CDbl/CLng/CInt/CSng on cell values without IsNumeric guard
        for m in cdbl_clng_re.finditer(proc_text):
            conv_line_no = start + proc_text[: m.start()].count("\n")
            conv_line = lines[conv_line_no - 1]
            # Only flag if the conversion argument looks like a cell value access
            match_text = m.group(0)
            if re.search(cell_value_re, match_text) or re.search(cell_value_re, conv_line):
                same_line = bool(isnumeric_re.search(conv_line))
                prev_line = conv_line_no > 1 and bool(isnumeric_re.search(lines[conv_line_no - 2]))
                if not same_line and not prev_line:
                    func_name = match_text.split("(")[0].split()[-1]
                    problems.append(("CRITICAL", f"{bas}:{conv_line_no} {func_name}() on cell value without IsNumeric guard -- type mismatch risk"))

        # CRITICAL: no error handling on procedures that do file I/O / workbook access
        io_patterns = [r"Workbooks\.Open", r"\.SaveAs", r"\.Export", r"FileSystemObject", r"Open\s+", r"\.Print\s*#"]
        does_io = any(re.search(p, proc_text, re.IGNORECASE) for p in io_patterns)
        if does_io and vis.lower() == "public" and not has_err_handler and not has_resume_next:
            problems.append(("CRITICAL", f"{bas}:{start} {name}: does file I/O / workbook access with no error handling"))

        # HIGH: procedure with Fail label but no Exit Sub/Function before it
        if has_fail_label and not has_exit:
            problems.append(("HIGH", f"{bas}:{start} {name}: Fail label present but no Exit Sub/Function -- falls through to error handler on success"))

        # HIGH: Fail/error label present but no LogErr call
        if has_fail_label and not has_logerr:
            problems.append(("HIGH", f"{bas}:{start} {name}: Fail/error label present but no LogErr call -- error swallowed without logging"))

        # HIGH: On Error Resume Next without matching On Error GoTo 0 (unguarded resume)
        if has_resume_next and not has_go_to_0 and not has_err_handler:
            problems.append(("HIGH", f"{bas}:{start} {name}: On Error Resume Next with no On Error GoTo 0 to restore error handling"))

        # MEDIUM: public procedure with stage assignment but no LogStage/SetStatus
        if vis.lower() == "public" and kind.lower() in ("sub", "function"):
            if has_stage_assign and not has_logstage and not has_setstatus:
                problems.append(("MEDIUM", f"{bas}:{start} {name}: assigns 'stage' but no LogStage/SetStatus -- breadcrumb never written"))

        # MEDIUM: public long-running procedure with no LogStage/SetStatus
        if vis.lower() == "public" and kind.lower() in ("sub", "function") and proc_len > 40:
            if not has_logstage and not has_setstatus:
                problems.append(("MEDIUM", f"{bas}:{start} {name}: public {kind} ({proc_len} lines) with no LogStage/SetStatus -- no audit trace"))

        # INFO: procedure longer than 80 lines (candidate for refactoring)
        if proc_len > 80:
            problems.append(("INFO", f"{bas}:{start} {name}: {proc_len} lines -- candidate for refactoring into smaller procedures"))

    # ---- Declaration placement: Public/Private between procedures ----
    proc_start_lines = sorted(p[3] for p in procs)
    first_proc = proc_start_lines[0] if proc_start_lines else len(lines) + 1

    for i, line in enumerate(lines, 1):
        if not module_decl_re.match(line):
            continue
        if i < first_proc:
            continue  # in the declarations section at the top of the module
        containing_proc = None
        for idx, pstart in enumerate(proc_start_lines):
            pnext = (
                proc_start_lines[idx + 1]
                if idx + 1 < len(proc_start_lines)
                else len(lines) + 1
            )
            if pstart <= i < pnext:
                containing_proc = pstart
                break
        if containing_proc is None:
            problems.append(("HIGH", f"{bas}:{i} module-level declaration '{line.strip()}' appears between procedures -- belongs in the declarations section at top"))

    # ---- Bounds checking before array indexing data(r, col) ----
    array_index_re = re.compile(r"\bdata\s*\(\s*r\s*,\s*\w+\s*\)", re.IGNORECASE)
    last_ubound_line = 0
    for m in ubound_re.finditer(content):
        last_ubound_line = max(last_ubound_line, content[: m.start()].count("\n") + 1)
    last_lur_line = 0
    lur_re = re.compile(r"\bLastUsedRow\s*\(", re.IGNORECASE)
    for m in lur_re.finditer(content):
        last_lur_line = max(last_lur_line, content[: m.start()].count("\n") + 1)
    for m in array_index_re.finditer(content):
        line_no = content[: m.start()].count("\n") + 1
        if line_no <= last_ubound_line + 10 or line_no <= last_lur_line + 10:
            continue
        problems.append(("CRITICAL", f"{bas}:{line_no} array index 'data(r, ...)' without recent UBound/LastUsedRow bounds check -- index out of range risk"))

print(f"Audit: {checked} module(s) scanned")
if problems:
    by_sev: dict[str, list] = {"CRITICAL": [], "HIGH": [], "MEDIUM": [], "INFO": []}
    for sev, p in problems:
        by_sev[sev].append(p)

    total = sum(len(v) for v in by_sev.values())
    print(f"\nPROBLEMS FOUND: {total} "
          f"({len(by_sev['CRITICAL'])} critical, {len(by_sev['HIGH'])} high, "
          f"{len(by_sev['MEDIUM'])} medium, {len(by_sev['INFO'])} info)")

    for sev in ("CRITICAL", "HIGH", "MEDIUM", "INFO"):
        if by_sev[sev]:
            print(f"\n--- {sev} ---")
            for p in sorted(by_sev[sev]):
                print(f"  - {p}")
    sys.exit(1)
print("OK: no deep-inspection problems found in any module.")
