# Company SEC/XBRL Workbook Automation

This repo creates one Excel workbook per company from SEC XBRL data. The current workbook is for FedEx, ticker `FDX`.

## Workbook Structure

Each company gets one Excel workbook.

The FedEx workbook has these tabs:

- `Collection`: automated SEC/XBRL data.
- `Research`: manual analyst notes.
- `Company Info`: company-level fields that are the same for every row, including ticker, company name, CIK, SEC company page, and row counts.
- `Blank`: empty placeholder tab.

The `Collection` tab is designed to be readable first, auditable second. The first columns show:

- `Data Group`
- `Statement / Detail Area`
- `Metric Group`
- `Metric Name`
- `XBRL Label`
- `Value`
- `Unit`

## What The Automation Pulls

The script [scripts/build_fdx_workbook.mjs](scripts/build_fdx_workbook.mjs) automatically pulls FedEx data from the SEC.

It uses three SEC sources:

1. SEC company ticker file
   - URL: `https://www.sec.gov/files/company_tickers_exchange.json`
   - Purpose: finds FedEx's CIK from ticker `FDX`.
2. SEC submissions API
   - URL pattern: `https://data.sec.gov/submissions/CIK##########.json`
   - Purpose: gets recent 10-K and 10-Q filing metadata, accession numbers, fiscal periods, filing dates, and primary filing document names.
3. SEC companyfacts XBRL API
   - URL pattern: `https://data.sec.gov/api/xbrl/companyfacts/CIK##########.json`
   - Purpose: gets standardized company XBRL line items for income statement, balance sheet, cash flow statement, and other tagged facts.

The script also opens recent inline XBRL filing documents listed in the submissions API and searches for dimensional XBRL facts. Those rows become `Segment / Geography` rows when the XBRL context includes dimensions such as segment, geography, product, service, business, domestic, international, Express, Freight, or similar members.

## How The Auto-Scrape Works

Run the script:

```bash
cd /Users/allisonxu/Desktop/Project

/Users/allisonxu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  scripts/build_fdx_workbook.mjs
```

The script then performs this automated process:

1. Looks up `FDX` in the SEC ticker file.
2. Converts FedEx's CIK to the SEC 10-digit CIK format.
3. Downloads FedEx submission metadata from the SEC submissions API.
4. Downloads FedEx companyfacts XBRL from the SEC companyfacts API.
5. Keeps only 10-K, 10-K/A, 10-Q, and 10-Q/A facts from fiscal year 2018 onward.
6. Converts raw XBRL concepts into readable workbook columns:
   - `Statement / Detail Area`
   - `Metric Group`
   - `Metric Name`
   - `Value`
   - `Unit`
   - source filing fields
7. Downloads recent filing HTML documents from SEC archive links.
8. Parses inline XBRL contexts from those filing documents.
9. Finds dimensional facts that look like segment, geography, product, service, or business disclosures.
10. Adds all rows directly into the Excel `Collection` tab.
11. Creates `Research`, `Company Info`, and `Blank` tabs.
