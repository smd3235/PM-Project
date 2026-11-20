#!/usr/bin/env python3
"""db_query.py
ODBC database query tool for extracting data into Excel reports.
Uses pyodbc to connect to ODBC data sources and export query results.
Usage: python scripts/db_query.py --dsn <DSN_NAME> --query <SQL_QUERY> [--output <OUTPUT_FILE>]
"""

import argparse
import os
import subprocess
import sys
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "db_query.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    parser = argparse.ArgumentParser(description="ODBC database query tool")
    parser.add_argument("--dsn", required=True, help="ODBC DSN name to connect to")
    parser.add_argument("--query", required=True, help="SQL query to execute")
    parser.add_argument("--output", default=None, help="Output CSV file path (default: Logs/query_results.csv)")
    parser.add_argument("--uid", default=None, help="Database username")
    parser.add_argument("--pwd", default=None, help="Database password")
    args = parser.parse_args()

    log("=== ODBC Database Query started ===")

    try:
        import pyodbc
    except ImportError:
        log("ERROR: pyodbc is not installed. Run: pip install pyodbc")
        sys.exit(1)

    # Build connection string
    conn_str = f"DSN={args.dsn}"
    if args.uid:
        conn_str += f";UID={args.uid}"
    if args.pwd:
        conn_str += f";PWD={args.pwd}"

    log(f"Connecting to DSN: {args.dsn}")

    try:
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        log(f"ERROR: Failed to connect to ODBC data source: {e}")
        sys.exit(1)

    log(f"Connected to {args.dsn} successfully")

    # Execute query
    try:
        cursor = conn.cursor()
        cursor.execute(args.query)

        # Get column names
        columns = [column[0] for column in cursor.description]

        # Fetch all rows
        rows = cursor.fetchall()

        log(f"Query returned {len(rows)} row(s)")

        # Determine output file
        output_file = args.output or os.path.join(log_dir, "query_results.csv")

        # Write results to CSV
        import csv
        with open(output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(columns)
            writer.writerows(rows)

        log(f"Query results written to {output_file} ({len(rows)} rows, {len(columns)} columns)")

        cursor.close()
        conn.close()

    except Exception as e:
        log(f"ERROR: Query execution failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()