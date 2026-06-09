# Company SEC/XBRL Workbook Automation

This repo creates one Excel workbook per company from SEC XBRL data. The current workbook is for FedEx, ticker `FDX`.

## Quick Start

1. Install Node.js.

   Download it from [nodejs.org](https://nodejs.org/) if it is not already installed.

2. Clone this repo.

   ```bash
   git clone https://github.com/alx003/us-company-sec-dataset.git
   cd us-company-sec-dataset
   ```

3. Install the spreadsheet dependency.

   ```bash
   npm install @oai/artifact-tool
   ```

4. Run the FedEx workbook builder.

   ```bash
   node scripts/build_fdx_workbook.mjs
   ```

5. Open the generated workbook.

   ```text
   outputs/FDX_fedex_sec_collection.xlsx
   ```

No manual copy/paste into Excel is required. To refresh the workbook later, run the same command again.

## Choose Your Own Output Folder

By default, the workbook is saved to:

```text
outputs/FDX_fedex_sec_collection.xlsx
```

To save it somewhere else on your own computer, set `FDX_OUTPUT_DIR` before running the script.

macOS or Linux:

```bash
FDX_OUTPUT_DIR="$HOME/Desktop/sec-workbooks" node scripts/build_fdx_workbook.mjs
```

Windows PowerShell:

```powershell
$env:FDX_OUTPUT_DIR="$env:USERPROFILE\Desktop\sec-workbooks"
node scripts/build_fdx_workbook.mjs
```

The script will create the output folder if it does not already exist.

## Workbook Structure

Each company's workbook has these tabs:

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

This makes it easier to filter for items such as depreciation, assets, capex, revenue, profit, debt, liabilities, and equity without scanning raw XBRL concept names first.

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

When you run:

```bash
node scripts/build_fdx_workbook.mjs
```

the script performs this automated process:

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
12. Saves the workbook to the selected output folder.

## Segment And Geography Examples

Segment and geography data is only available when FedEx tags that detail in inline XBRL.

Example segment dimension:

```text
StatementBusinessSegmentsAxis=FederalExpressSegmentMember
```

This means the value is a FedEx disclosed number tagged to the Federal Express segment.

Example product/service dimension:

```text
ProductOrServiceAxis=BoeingMd11FAircraftMember
```

This means FedEx tagged one of its own filing facts with a product/service member related to Boeing MD-11F aircraft.

The workbook keeps these dimension fields in:

- `Dimensions`
- `Segment / Geography Member`

## Troubleshooting

If `node` is not recognized, install Node.js and open a new terminal window.

If `@oai/artifact-tool` is missing, run:

```bash
npm install @oai/artifact-tool
```

If the SEC blocks or slows requests, wait a few minutes and run the script again. The script includes a short pause between SEC requests, but SEC availability can still vary.
