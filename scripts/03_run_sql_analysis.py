"""
03_run_sql_analysis.py
-----------------------
Runs every query in sql/analysis_queries.sql against the SQLite DB and
saves each result set as a CSV, plus prints a readable summary.
"""

import sqlite3
import pandas as pd
import re
import os

DB_PATH = "sales_project/sales.db"
SQL_PATH = "sales_project/sql/analysis_queries.sql"
OUT_DIR = "sales_project/sql/results"

os.makedirs(OUT_DIR, exist_ok=True)

with open(SQL_PATH) as f:
    sql_text = f.read()

# Split into individual queries using the "-- Q<n>." markers as delimiters
blocks = re.split(r"-- Q(\d+)\.", sql_text)[1:]  # drop preamble
queries = {}
for i in range(0, len(blocks), 2):
    qnum = blocks[i]
    body = blocks[i + 1]
    # the real SQL starts at the first SELECT or WITH keyword; everything
    # before that is comment/header text
    match = re.search(r"^\s*(SELECT|WITH)\b", body, re.IGNORECASE | re.MULTILINE)
    stmt = body[match.start():].strip() if match else ""
    queries[f"Q{qnum}"] = stmt

conn = sqlite3.connect(DB_PATH)

for qname, stmt in queries.items():
    df = pd.read_sql_query(stmt, conn)
    out_path = os.path.join(OUT_DIR, f"{qname}.csv")
    df.to_csv(out_path, index=False)
    print(f"\n===== {qname} =====")
    print(df.head(10).to_string(index=False))

conn.close()
