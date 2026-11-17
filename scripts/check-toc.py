# check-toc.py
# Verifies the split VBA modules: each PM_Report_*.bas has Option Explicit,
# a unique Attribute VB_Name, and lists all procedures across modules.
# Usage: python scripts/check-toc.py [--fix]   (--fix is accepted for compat, read-only here)
import argparse
import os
import re
import sys

parser = argparse.ArgumentParser(description="Check VBA module split (PM_Report_*.bas)")
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

proc_re = re.compile(r"^\s*(Public|Private)\s+(Sub|Function|Property)\s+(\w+)", re.MULTILINE)
name_re = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.MULTILINE)

total = 0
seen_names = set()
problems = []
for bas in bas_files:
    path = os.path.join(project_root, bas)
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    names = name_re.findall(content)
    vbname = names[0] if names else None
    if vbname and vbname != os.path.splitext(bas)[0]:
        problems.append(f"{bas}: VB_Name={vbname!r} does not match filename {bas!r}")
    if not vbname:
        problems.append(f"{bas}: missing Attribute VB_Name")
    if "Option Explicit" not in content:
        problems.append(f"{bas}: missing Option Explicit")
    procs = proc_re.findall(content)
    total += len(procs)
    for _, _, pname in procs:
        if pname in seen_names:
            problems.append(f"{bas}: duplicate procedure {pname} (already in another module)")
        seen_names.add(pname)
    print(f"{bas}: {len(procs)} procedures (VB_Name={vbname})")
    for _, _, pname in procs:
        print(f"    {pname}")

print(f"\nTOC check: {len(bas_files)} modules, {total} procedures")
if problems:
    print("PROBLEMS:")
    for p in problems:
        print("  - " + p)
    sys.exit(1)
print("OK: all modules have Option Explicit, matching VB_Name, and unique procedure names.")
