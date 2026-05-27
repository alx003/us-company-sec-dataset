from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

import duckdb

from .metrics import METRIC_MAP
from .workbook import CompanySeed


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS collection_runs (
    run_id VARCHAR PRIMARY KEY,
    started_at TIMESTAMP,
    finished_at TIMESTAMP,
    status VARCHAR,
    note VARCHAR
);

CREATE TABLE IF NOT EXISTS companies (
    ticker VARCHAR PRIMARY KEY,
    cik VARCHAR,
    name VARCHAR,
    country VARCHAR,
    exchange VARCHAR,
    sector VARCHAR,
    industry VARCHAR,
    source_workbook VARCHAR,
    source_sheet VARCHAR,
    source_row INTEGER,
    source_note VARCHAR,
    resolution_status VARCHAR
);

CREATE TABLE IF NOT EXISTS filings (
    cik VARCHAR,
    ticker VARCHAR,
    company_name VARCHAR,
    accession_number VARCHAR,
    form VARCHAR,
    report_date DATE,
    filing_date DATE,
    fiscal_year INTEGER,
    fiscal_period VARCHAR,
    primary_document VARCHAR,
    source_url VARCHAR,
    PRIMARY KEY (cik, accession_number)
);

CREATE TABLE IF NOT EXISTS raw_facts (
    cik VARCHAR,
    ticker VARCHAR,
    company_name VARCHAR,
    taxonomy VARCHAR,
    concept VARCHAR,
    label VARCHAR,
    description VARCHAR,
    unit VARCHAR,
    value DOUBLE,
    accession_number VARCHAR,
    form VARCHAR,
    filed_date DATE,
    fiscal_year INTEGER,
    fiscal_period VARCHAR,
    period_start DATE,
    period_end DATE,
    frame VARCHAR,
    source_url VARCHAR,
    PRIMARY KEY (cik, taxonomy, concept, unit, accession_number, fiscal_year, fiscal_period, period_end)
);

CREATE TABLE IF NOT EXISTS metric_map (
    metric VARCHAR,
    taxonomy VARCHAR,
    concept VARCHAR,
    preferred_unit VARCHAR,
    priority INTEGER,
    PRIMARY KEY (metric, taxonomy, concept, preferred_unit, priority)
);

CREATE TABLE IF NOT EXISTS curated_metrics (
    cik VARCHAR,
    ticker VARCHAR,
    company_name VARCHAR,
    metric VARCHAR,
    value DOUBLE,
    unit VARCHAR,
    form VARCHAR,
    fiscal_year INTEGER,
    fiscal_period VARCHAR,
    period_end DATE,
    filed_date DATE,
    accession_number VARCHAR,
    taxonomy VARCHAR,
    concept VARCHAR,
    source_url VARCHAR,
    PRIMARY KEY (cik, metric, fiscal_year, fiscal_period, period_end, accession_number)
);

CREATE TABLE IF NOT EXISTS workbook_validation (
    workbook VARCHAR,
    sheet VARCHAR,
    row INTEGER,
    column_index INTEGER,
    value VARCHAR,
    classification VARCHAR,
    v1_status VARCHAR
);
"""


class ResearchDatabase:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.con = duckdb.connect(str(path))
        self.con.execute(SCHEMA_SQL)
        self.load_metric_map()

    def close(self) -> None:
        self.con.close()

    def start_run(self, run_id: str, note: str) -> None:
        self.con.execute(
            "INSERT OR REPLACE INTO collection_runs VALUES (?, ?, NULL, ?, ?)",
            [run_id, datetime.now(timezone.utc), "running", note],
        )

    def clear_derived_data(self) -> None:
        for table in ["workbook_validation", "curated_metrics", "raw_facts", "filings", "companies"]:
            self.con.execute(f"DELETE FROM {table}")

    def finish_run(self, run_id: str, status: str, note: str = "") -> None:
        self.con.execute(
            "UPDATE collection_runs SET finished_at = ?, status = ?, note = ? WHERE run_id = ?",
            [datetime.now(timezone.utc), status, note, run_id],
        )

    def load_metric_map(self) -> None:
        self.con.execute("DELETE FROM metric_map")
        self.con.executemany("INSERT INTO metric_map VALUES (?, ?, ?, ?, ?)", METRIC_MAP)

    def upsert_companies(self, seeds: Iterable[CompanySeed], ticker_lookup: dict[str, dict]) -> None:
        rows = []
        for seed in seeds:
            lookup = ticker_lookup.get(seed.ticker.upper(), {})
            cik = seed.cik or _cik(lookup.get("cik"))
            rows.append([
                seed.ticker.upper(),
                cik,
                seed.name,
                seed.country,
                lookup.get("exchange"),
                seed.sector,
                seed.industry,
                seed.source_workbook,
                seed.source_sheet,
                seed.source_row,
                seed.source_note,
                "resolved" if cik else "unresolved",
            ])
        self.con.executemany(
            """
            INSERT OR REPLACE INTO companies
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )

    def resolved_companies(self, sample_size: int | None = None) -> list[dict]:
        sql = "SELECT * FROM companies WHERE cik IS NOT NULL ORDER BY ticker"
        if sample_size:
            sql += f" LIMIT {int(sample_size)}"
        return [dict(zip([c[0] for c in self.con.description], row)) for row in self.con.execute(sql).fetchall()]

    def insert_filings(self, rows: list[list]) -> None:
        if rows:
            self.con.executemany("INSERT OR REPLACE INTO filings VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", rows)

    def insert_raw_facts(self, rows: list[list]) -> None:
        if rows:
            self.con.executemany(
                "INSERT OR REPLACE INTO raw_facts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                rows,
            )

    def refresh_curated_metrics(self) -> None:
        self.con.execute("DELETE FROM curated_metrics")
        self.con.execute(
            """
            INSERT OR REPLACE INTO curated_metrics
            SELECT
                cik, ticker, company_name, metric, value, unit, form, fiscal_year, fiscal_period,
                period_end, filed_date, accession_number, taxonomy, concept, source_url
            FROM (
                SELECT
                    rf.*,
                    mm.metric,
                    row_number() OVER (
                        PARTITION BY rf.cik, mm.metric, rf.fiscal_year, rf.fiscal_period, rf.period_end
                        ORDER BY mm.priority, rf.filed_date DESC, rf.accession_number DESC
                    ) AS rn
                FROM raw_facts rf
                JOIN metric_map mm
                    ON rf.taxonomy = mm.taxonomy
                    AND rf.concept = mm.concept
                WHERE rf.form IN ('10-K', '10-K/A', '10-Q', '10-Q/A')
                    AND rf.value IS NOT NULL
                    AND rf.fiscal_year IS NOT NULL
            )
            WHERE rn = 1
            """
        )
        self.con.execute(
            """
            INSERT OR REPLACE INTO curated_metrics
            SELECT
                ocf.cik, ocf.ticker, ocf.company_name, 'free_cash_flow',
                ocf.value - capex.value, ocf.unit, ocf.form, ocf.fiscal_year, ocf.fiscal_period,
                ocf.period_end, ocf.filed_date, ocf.accession_number,
                'derived', 'operating_cash_flow_minus_capex', ocf.source_url
            FROM curated_metrics ocf
            JOIN curated_metrics capex
                ON ocf.cik = capex.cik
                AND ocf.fiscal_year = capex.fiscal_year
                AND ocf.fiscal_period = capex.fiscal_period
                AND ocf.period_end = capex.period_end
            WHERE ocf.metric = 'operating_cash_flow'
                AND capex.metric = 'capex'
                AND ocf.unit = capex.unit
            """
        )

    def write_validation(self, rows: list[dict]) -> None:
        self.con.execute("DELETE FROM workbook_validation")
        if rows:
            self.con.executemany(
                "INSERT INTO workbook_validation VALUES (?, ?, ?, ?, ?, ?, ?)",
                [[r[k] for k in ["workbook", "sheet", "row", "column_index", "value", "classification", "v1_status"]] for r in rows],
            )

    def export_parquet(self, output_dir: Path) -> None:
        output_dir.mkdir(parents=True, exist_ok=True)
        for table in ["companies", "filings", "raw_facts", "metric_map", "curated_metrics", "collection_runs"]:
            self.con.execute(f"COPY {table} TO '{output_dir / (table + '.parquet')}' (FORMAT PARQUET)")


def _cik(value: object) -> str | None:
    if value is None or value == "":
        return None
    return str(int(value)).zfill(10)
