#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import gzip
import json
import os
import time
import urllib.request
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

FORMS = {"10-K", "10-K/A", "10-Q", "10-Q/A"}
MIN_YEAR = 2018
DEFAULT_OUTPUT_DIR = Path.home() / "OneDrive" / "ImportantFiles" / "Brian" / "sec-workbooks"
USER_AGENT = os.environ.get("SEC_USER_AGENT", "SEC workbook automation contact@example.com")

METRIC_MAP = {
    "revenue": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues"],
    "gross_profit": ["GrossProfit"],
    "operating_income": ["OperatingIncomeLoss"],
    "net_income": ["NetIncomeLoss"],
    "eps": ["EarningsPerShareDiluted", "EarningsPerShareBasic"],
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "debt": ["DebtCurrent", "LongTermDebtCurrent", "LongTermDebtNoncurrent"],
    "operating_cash_flow": ["NetCashProvidedByUsedInOperatingActivities"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "shares_outstanding": ["EntityCommonStockSharesOutstanding"],
    "rd": ["ResearchAndDevelopmentExpense"],
    "sga": ["SellingGeneralAndAdministrativeExpense"],
    "inventory": ["InventoryNet"],
    "assets": ["Assets"],
    "liabilities": ["Liabilities"],
    "equity": ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
}


def get_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"},
    )
    with urllib.request.urlopen(request, timeout=90) as response:
        payload = response.read()
        if response.headers.get("Content-Encoding") == "gzip" or payload.startswith(b"\x1f\x8b"):
            payload = gzip.decompress(payload)
    time.sleep(0.15)
    return json.loads(payload.decode("utf-8"))


def sec_url(cik: str, accession: str | None = None) -> str:
    if not accession:
        return f"https://www.sec.gov/edgar/browse/?CIK={int(cik)}"
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession.replace('-', '')}/"


def load_ticker_lookup() -> dict[str, dict[str, Any]]:
    data = get_json("https://www.sec.gov/files/company_tickers_exchange.json")
    fields = data["fields"]
    lookup = {}
    for row in data["data"]:
        item = dict(zip(fields, row))
        ticker = str(item["ticker"]).upper()
        lookup[ticker] = item
        lookup[ticker.replace("-", ".")] = item
        lookup[ticker.replace(".", "-")] = item
    return lookup


def seed_tickers(path: Path | None) -> list[str]:
    if not path or not path.exists():
        return []
    tickers = []
    with path.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            ticker = (row.get("ticker") or "").strip().upper()
            if ticker:
                tickers.append(ticker)
    return tickers


def at(data: dict[str, list[Any]], key: str, index: int) -> Any:
    values = data.get(key) or []
    return values[index] if index < len(values) else ""


def to_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def filing_rows(submissions: dict[str, Any], cik: str) -> list[dict[str, Any]]:
    recent = submissions.get("filings", {}).get("recent", {})
    rows = []
    for i, form in enumerate(recent.get("form", [])):
        fiscal_year = to_int(at(recent, "fy", i))
        if form not in FORMS or (fiscal_year and fiscal_year < MIN_YEAR):
            continue
        accession = at(recent, "accessionNumber", i)
        rows.append({
            "accession_number": accession,
            "form": form,
            "fiscal_year": fiscal_year,
            "fiscal_period": at(recent, "fp", i),
            "report_date": at(recent, "reportDate", i),
            "filing_date": at(recent, "filingDate", i),
            "primary_document": at(recent, "primaryDocument", i),
            "source_url": sec_url(cik, accession),
        })
    return rows


def raw_fact_rows(company: dict[str, Any], facts: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for taxonomy, concepts in facts.get("facts", {}).items():
        for concept, payload in concepts.items():
            for unit, values in payload.get("units", {}).items():
                for item in values:
                    fiscal_year = to_int(item.get("fy"))
                    if item.get("form") not in FORMS or (fiscal_year and fiscal_year < MIN_YEAR):
                        continue
                    accession = item.get("accn")
                    rows.append({
                        "ticker": company["ticker"],
                        "company_name": company["name"],
                        "taxonomy": taxonomy,
                        "concept": concept,
                        "label": payload.get("label", ""),
                        "unit": unit,
                        "value": item.get("val"),
                        "form": item.get("form"),
                        "fiscal_year": fiscal_year,
                        "fiscal_period": item.get("fp"),
                        "period_start": item.get("start", ""),
                        "period_end": item.get("end", ""),
                        "filed_date": item.get("filed", ""),
                        "accession_number": accession,
                        "source_url": sec_url(company["cik"], accession),
                    })
    return rows


def curated_rows(raw_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    concept_to_metric = {concept: metric for metric, concepts in METRIC_MAP.items() for concept in concepts}
    rows = []
    seen = set()
    for row in raw_rows:
        metric = concept_to_metric.get(row["concept"])
        if not metric:
            continue
        key = (row["ticker"], metric, row["fiscal_year"], row["fiscal_period"], row["period_end"], row["accession_number"])
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "ticker": row["ticker"],
            "company_name": row["company_name"],
            "metric": metric,
            "value": row["value"],
            "unit": row["unit"],
            "form": row["form"],
            "fiscal_year": row["fiscal_year"],
            "fiscal_period": row["fiscal_period"],
            "period_end": row["period_end"],
            "filed_date": row["filed_date"],
            "xbrl_concept": row["concept"],
            "source_url": row["source_url"],
        })
    return rows


def add_sheet(wb: Workbook, name: str, rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(name)
    headers = list(rows[0].keys()) if rows else ["message"]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, column in enumerate(ws.columns, start=1):
        values = [str(cell.value) for cell in column[:200] if cell.value is not None]
        width = min(max([len(value) for value in values] + [10]) + 2, 55)
        ws.column_dimensions[get_column_letter(idx)].width = width


def write_workbook(output_path: Path, company: dict[str, Any], filings: list[dict[str, Any]], raw_rows: list[dict[str, Any]], clean_rows: list[dict[str, Any]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, "Company Info", [
        {"field": "ticker", "value": company["ticker"]},
        {"field": "company_name", "value": company["name"]},
        {"field": "cik", "value": company["cik"]},
        {"field": "sec_company_page", "value": sec_url(company["cik"])},
        {"field": "filing_rows", "value": len(filings)},
        {"field": "curated_metric_rows", "value": len(clean_rows)},
        {"field": "raw_fact_rows", "value": len(raw_rows)},
    ])
    add_sheet(wb, "Filings", filings)
    add_sheet(wb, "Curated Metrics", clean_rows)
    add_sheet(wb, "Raw XBRL Facts", raw_rows[:50000])
    add_sheet(wb, "Research Notes", [
        {"date": "", "topic": "", "notes": "", "source": "", "follow_up": ""},
        {"date": "", "topic": "", "notes": "", "source": "", "follow_up": ""},
        {"date": "", "topic": "", "notes": "", "source": "", "follow_up": ""},
    ])
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Excel company workbooks from SEC/XBRL data.")
    parser.add_argument("--ticker", action="append", help="Ticker to build. Can be used multiple times.")
    parser.add_argument("--seed", default="data/seeds/manual_companies.csv", help="CSV with a ticker column.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Folder where Excel files are saved.")
    parser.add_argument("--limit", type=int, default=None, help="Optional max number of tickers to build.")
    args = parser.parse_args()

    tickers = [t.upper() for t in (args.ticker or [])]
    if not tickers:
        tickers = seed_tickers(Path(args.seed))
    if args.limit:
        tickers = tickers[:args.limit]
    if not tickers:
        raise SystemExit("No tickers found. Use --ticker FDX or provide data/seeds/manual_companies.csv.")

    lookup = load_ticker_lookup()
    output_dir = Path(args.output_dir).expanduser()
    for ticker in tickers:
        match = lookup.get(ticker)
        if not match:
            print(f"SKIP {ticker}: not found in SEC ticker file")
            continue
        company = {"ticker": ticker, "name": match["name"], "cik": str(match["cik"]).zfill(10)}
        submissions = get_json(f"https://data.sec.gov/submissions/CIK{company['cik']}.json")
        facts = get_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{company['cik']}.json")
        filings = filing_rows(submissions, company["cik"])
        raw_rows = raw_fact_rows(company, facts)
        clean_rows = curated_rows(raw_rows)
        output_path = output_dir / f"{ticker}_sec_xbrl_workbook.xlsx"
        write_workbook(output_path, company, filings, raw_rows, clean_rows)
        print(f"WROTE {output_path}")


if __name__ == "__main__":
    main()
