# U.S. Company SEC Dataset

This repo collects free U.S. public-company data from the SEC.

It pulls company filing data from 10-K and 10-Q reports, organizes it into CSV files, and keeps source links so the numbers can be checked later.

## What Data This Gets

The dataset is built from SEC EDGAR / XBRL data. It focuses on U.S. public companies.

It gets:

- Company list: ticker, CIK, company name, sector, industry, and whether the company was matched to the SEC.
- Filing history: 10-K and 10-Q filing dates, report dates, fiscal year/quarter, accession numbers, and SEC source links.
- Raw XBRL facts: detailed reported SEC facts exactly as structured in filings.
- Clean research metrics: selected useful metrics pulled from the raw facts.

The clean metrics includs: 

- revenue
- gross profit
- operating income
- net income
- EPS
- cash
- debt
- operating cash flow
- capex
- free cash flow
- shares outstanding
- R&D
- SG&A
- inventory
- assets
- liabilities
- equity

This does **not** yet collect RSS feeds, company news, emails, Europe, ASEAN, or LLM summaries. Those should come after the source pipes are designed.

## Which Files To Use

- `dataset/companies.csv`: list of companies and whether they matched to the SEC.
- `dataset/curated_metrics.csv`: the most useful research-ready financial metrics.
- `dataset/filings.csv`: filing history and SEC links.

for more detail:

- `dataset/raw_facts.csv`: all raw SEC/XBRL facts collected from filings.
- `dataset/metric_map.csv`: shows which XBRL tags feed each clean metric.
- `dataset/manifest.json`: row counts and generation info.

## How To Use The Data

For normal research, open `dataset/curated_metrics.csv`.

Typical workflow:

1. Filter by `ticker`, such as `AAPL`.
2. Filter by `metric`, such as `revenue` or `net_income`.
3. Compare `fiscal_year` and `fiscal_period`.
4. Use `source_url` if you need to check the original SEC filing.

For audit/source checking, use `dataset/raw_facts.csv` and `dataset/filings.csv`.

## Auto Updates

`/.github/workflows/update-sec-dataset.yml`

It is set to run weekly:

- every Monday at 14:00 UTC

GitHub Actions for manual refreshes when you want to update immediately.

Important: before relying on the automatic weekly run, update the SEC user-agent email in:

`config/repo_pipeline.yml`

Replace:

```yaml
user_agent: "alx003 SEC research dataset contact@example.com"
```

with a real contact email. The SEC expects automated requests to identify who is making them.

## How To Refresh Locally

If you want to run it from this computer:

```bash
cd /Users/allisonxu/Desktop/Project

PYTHONPATH=/Users/allisonxu/Desktop/Project/.python_deps \
/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
scripts/run_pipeline.py --config config/repo_pipeline.yml

PYTHONPATH=/Users/allisonxu/Desktop/Project/.python_deps \
/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
scripts/export_git_dataset.py --database outputs/us_company_research.duckdb --output-dir dataset
```

Then commit and push:

```bash
git add dataset outputs/workbook_field_inventory.csv outputs/workbook_validation.csv
git commit -m "Refresh SEC dataset"
git push
```

## Current Limits

- U.S. companies only.
- SEC 10-K and 10-Q data only.
- Some workbook companies are non-U.S. or do not map cleanly to SEC tickers.
- Raw SEC facts can be noisy; use `curated_metrics.csv` first.
- RSS feeds and company news flow are not implemented yet.
