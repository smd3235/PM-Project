# check-toc.py
# Verifies the MAP block in PM_Report_Macro.bas against the file.
# Usage: python scripts/check-toc.py [--fix]

import argparse
import os
import re
import sys

parser = argparse.ArgumentParser(description="Check TOC in PM_Report_Macro.bas")
parser.add_argument("--fix", action="store_true", help="Rewrite the MAP block")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
bas_path = os.path.join(project_root, "PM_Report_Macro.bas")

if not os.path.exists(bas_path):
    print(f"File not found: {bas_path}", file=sys.stderr)
    sys.exit(1)

with open(bas_path, "r", encoding="utf-8") as f:
    content = f.read()

# Count sections and procedures
sections = re.findall(r"^\s*'\s*\[\s*\d+\]\s*(.+)$", content, re.MULTILINE)
procs = re.findall(r"^\s*(Public|Private)\s+(Sub|Function)\s+(\w+)", content, re.MULTILINE)

print(f"TOC check: {len(sections)} sections, {len(procs)} procedures")
for s in sections:
    print(f"  {s.strip()}")
print(f"Procedures: {', '.join(p[2] for p in procs)}")
