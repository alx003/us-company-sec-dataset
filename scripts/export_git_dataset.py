#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".python_deps"))

import duckdb


TABLES = [
    "companies",
    "filings",
    "raw_facts",
    "metric_map",
    "curated_metrics",
    "workbook_validation",
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Export DuckDB research tables to git-friendly dataset files.")
    parser.add_argument("--database", default="outputs/us_company_research.duckdb")
    parser.add_argument("--output-dir", default="dataset")
    args = parser.parse_args()

    db_path = (ROOT / args.database).resolve()
    output_dir = (ROOT / args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    con = duckdb.connect(str(db_path), read_only=True)
    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "database": str(db_path),
        "tables": {},
    }
    for table in TABLES:
        count = con.execute(f"SELECT count(*) FROM {table}").fetchone()[0]
        csv_path = output_dir / f"{table}.csv"
        con.execute(f"COPY (SELECT * FROM {table}) TO '{csv_path}' (HEADER, DELIMITER ',')")
        manifest["tables"][table] = {
            "rows": count,
            "file": f"{table}.csv",
            "columns": [row[1] for row in con.execute(f"PRAGMA table_info('{table}')").fetchall()],
        }
    (output_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    con.close()


if __name__ == "__main__":
    main()
