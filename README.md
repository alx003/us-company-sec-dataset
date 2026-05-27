# U.S.-First Company Data Collection System

This project builds a local U.S. company research database from the existing BrianProject workbooks and free SEC EDGAR/XBRL data.

## What It Produces

- `outputs/us_company_research.duckdb`: durable research database.
- `outputs/parquet/`: Parquet exports of core tables.
- `outputs/us_company_research_export.xlsx`: analyst-facing Excel views.
- `outputs/workbook_field_inventory.csv`: workbook/sheet/header inventory.
- `outputs/workbook_validation.csv`: hand-collected field classification.

## Quick Start

The dependencies for this workspace were installed into `.python_deps`. To run with the bundled Python:

```bash
PYTHONPATH=/Users/allisonxu/Desktop/Project/.python_deps \
/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
scripts/run_pipeline.py --config config/pipeline.yml --sample-size 5
```

For a full run, omit `--sample-size 5`.

## GitHub Dataset Workflow

This repo includes a weekly GitHub Actions workflow at `.github/workflows/update-sec-dataset.yml`.

The workflow:

1. Installs Python dependencies.
2. Runs the SEC collection pipeline using `config/repo_pipeline.yml`.
3. Exports git-friendly CSV files into `dataset/`.
4. Commits refreshed dataset files back to the repo.

Before enabling it in GitHub:

- Create a dedicated repo under `https://github.com/alx003`, for example `us-company-sec-dataset`.
- Push this project folder to that repo.
- Update `sec.user_agent` in `config/repo_pipeline.yml` with a real contact email.
- Keep large generated binary files out of git; the workflow commits CSV dataset files and the manifest, not the DuckDB or Excel output.

Useful local commands:

```bash
# Refresh the company seed list from the local BrianProject workbooks
PYTHONPATH=/Users/allisonxu/Desktop/Project/src:/Users/allisonxu/Desktop/Project/.python_deps \
/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
scripts/extract_workbook_companies.py

# Export current DuckDB tables to dataset/*.csv
PYTHONPATH=/Users/allisonxu/Desktop/Project/.python_deps \
/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
scripts/export_git_dataset.py --database outputs/us_company_research.duckdb --output-dir dataset
```

## Important Notes

- V1 is U.S.-only and focuses on 10-K / 10-Q filing history plus standardized XBRL facts.
- The pipeline stores all SEC facts it can ingest for selected companies, then creates a curated metric layer for research use.
- If the workbooks do not contain enough U.S. company identifiers, add rows to `data/seeds/manual_companies.csv`.
- SEC requests require a descriptive User-Agent. Update `sec.user_agent` in `config/pipeline.yml` before production use.

## Core Tables

- `companies`: workbook/manual seed companies resolved to SEC CIKs.
- `filings`: 10-K / 10-Q filing metadata.
- `raw_facts`: normalized long-form SEC XBRL facts.
- `metric_map`: preferred XBRL tags for curated metrics.
- `curated_metrics`: source-traced research metrics.
- `collection_runs`: run metadata and status.
