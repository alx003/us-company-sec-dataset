#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / ".python_deps"))

from sec_research.workbook import extract_company_seeds


DEFAULT_WORKBOOKS = [
    Path("/Users/allisonxu/Desktop/BrianProject/1.  WB - MASTER - 2026.05.23.xlsx"),
    Path("/Users/allisonxu/Desktop/BrianProject/L&C - AI Project Dashboard.xlsx"),
]


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract workbook company seeds into data/seeds/manual_companies.csv.")
    parser.add_argument("--output", default="data/seeds/manual_companies.csv")
    args = parser.parse_args()

    output = ROOT / args.output
    existing_manual = ROOT / "data/seeds/manual_companies.csv"
    seeds = extract_company_seeds([path for path in DEFAULT_WORKBOOKS if path.exists()], existing_manual)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["ticker", "cik", "name", "country", "sector", "industry", "source_note"],
        )
        writer.writeheader()
        for seed in seeds:
            writer.writerow({
                "ticker": seed.ticker,
                "cik": seed.cik or "",
                "name": seed.name,
                "country": seed.country or "",
                "sector": seed.sector or "",
                "industry": seed.industry or "",
                "source_note": f"{seed.source_workbook}:{seed.source_sheet}:row {seed.source_row}",
            })
    print(f"Wrote {len(seeds)} company seeds to {output}")


if __name__ == "__main__":
    main()
