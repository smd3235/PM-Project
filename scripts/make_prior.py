r"""Build a prior-week report by removing rows from a finished report, so the next
build has to flag new PMs (green rows) and a non-zero Completed count.

    python scripts\make_prior.py
    python scripts\make_prior.py --source "Reports\PM Requests 2026-11-14.xlsx" --drop 4

Manager tabs keep their header on row 9 and data from row 10, so that is where the
rows are removed from.
"""
import argparse
import datetime
import glob
import os
import re
import shutil

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPORTS = os.path.join(ROOT, "Reports")
SKIP = ("Dashboard", "Raw Data")
DATA_START = 10


def newest_report():
    files = glob.glob(os.path.join(REPORTS, "PM Requests *.xlsx"))
    if not files:
        raise SystemExit("no 'PM Requests *.xlsx' in Reports\\ - build a report first")
    return max(files, key=os.path.getmtime)


def prior_name(source):
    """Last week's name for a report, so the next build finds it automatically."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(source))
    if not m:
        return "PRIOR_PM_Requests.xlsx"
    d = datetime.date.fromisoformat(m.group(1)) - datetime.timedelta(days=7)
    return "PM Requests %s.xlsx" % d.isoformat()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=None, help="report to copy (default: newest in Reports\\)")
    ap.add_argument("--out", default=None, help="destination (default: last week's PM Requests name)")
    ap.add_argument("--drop", type=int, default=3, help="rows removed per manager tab")
    args = ap.parse_args()

    src = args.source or newest_report()
    if not os.path.exists(src):
        raise SystemExit("source not found: " + src)
    dst = args.out or os.path.join(REPORTS, prior_name(src))

    if os.path.exists(dst):
        os.remove(dst)
    shutil.copyfile(src, dst)

    wb = openpyxl.load_workbook(dst)
    removed = 0
    tabs = 0
    for name in wb.sheetnames:
        if name in SKIP:
            continue
        ws = wb[name]
        if ws.max_row < DATA_START:
            continue
        n = min(args.drop, ws.max_row - DATA_START + 1)
        ws.delete_rows(DATA_START, n)
        removed += n
        tabs += 1
    wb.save(dst)
    wb.close()

    print("prior built:", dst)
    print("tabs:", tabs, " removed rows:", removed)
    print("expect next run: Completed > 0, New PMs > 0, green rows > 0")


if __name__ == "__main__":
    main()