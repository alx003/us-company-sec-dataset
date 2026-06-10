# Company SEC/XBRL Excel Automation

This project uses Python to pull free SEC/XBRL data and create Excel workbooks.

## What This Gets

For each ticker, the automation creates one Excel workbook with:

- `Company Info`: ticker, company name, CIK, SEC company page, row counts.
- `Financial Statements`: finance-friendly period table for revenue, profit, cash flow, capex, assets, debt, liabilities, and equity.
- `Segment Summary`: grouped segment and sector-bucket view when FedEx tags dimensional XBRL facts.
- `Filings`: recent 10-K and 10-Q filing metadata.
- `Curated Metrics`: cleaner research fields such as revenue, net income, cash, debt, assets, liabilities, equity, capex, and cash flow.
- `Segment Facts`: raw inline XBRL dimensional facts used to build the segment summary.
- `Raw XBRL Facts`: detailed SEC/XBRL facts from company filings.
- `Research Notes`: blank sheet for manual analyst notes.

## Where Excel Files Save

By default, Allison's local files are saved to:

```text
/Users/allisonxu/OneDrive/ImportantFiles/Brian/sec-workbooks
```

Other people can choose their own save folder with `--output-dir`.

## First Time Setup

Open Terminal.

Go to the project folder:

```bash
cd "$HOME/OneDrive/ImportantFiles/Brian"
```

Install the Excel-writing package:

```bash
python3 -m pip install openpyxl
```

Set your SEC contact email:

```bash
export SEC_USER_AGENT="email@example.com"
```

Use a real email. The SEC expects automated scripts to identify who is making requests.

## Create One Company Workbook

Example for FedEx:

```bash
python3 scripts/build_company_workbooks.py --ticker FDX
```

The file will be saved to:

```text
~/OneDrive/ImportantFiles/Brian/sec-workbooks/FDX_sec_xbrl_workbook.xlsx
```

Open that file in Excel.

## Create Several Company Workbooks

```bash
python3 scripts/build_company_workbooks.py --ticker FDX --ticker UPS --ticker AAPL
```

## Save To A Different Folder

Mac OneDrive example:

```bash
python3 scripts/build_company_workbooks.py --ticker FDX --output-dir "$HOME/OneDrive/ImportantFiles/Brian/sec-workbooks"
```

Windows PowerShell OneDrive example:

```powershell
python scripts/build_company_workbooks.py --ticker FDX --output-dir "$env:USERPROFILE\OneDrive\ImportantFiles\Brian\sec-workbooks"
```

## Project Storage Rule

Do not create project work files on the Desktop. Store generated workbooks, scripts, and notes for this project inside `OneDrive/ImportantFiles/Brian`.

## Weekly GitHub Auto-Refresh

The repo includes a GitHub Action that can refresh the workbook without anyone running the script locally.

What it does:

1. Runs once per week on GitHub.
2. Installs Python and `openpyxl`.
3. Runs `scripts/build_company_workbooks.py --ticker FDX --output-dir workbooks`.
4. Commits the refreshed workbook back to the repo if the SEC data changed.

Manual run:

1. Open the GitHub repo.
2. Click **Actions**.
3. Click **Refresh SEC Workbooks**.
4. Click **Run workflow**.
5. Wait for the run to finish.
6. Open or download the newest workbook from `workbooks/FDX_sec_xbrl_workbook.xlsx`.

If the workflow does not commit, check repository settings:

1. Go to **Settings**.
2. Go to **Actions** > **General**.
3. Under **Workflow permissions**, choose **Read and write permissions**.
4. Save.

## Refresh A Workbook Locally

Run the same command again.

Example:

```bash
python3 scripts/build_company_workbooks.py --ticker FDX
```

It overwrites the old workbook with fresh SEC/XBRL data.

## What Not To Do

Do not run:

```bash
npm install @oai/artifact-tool
node scripts/build_fdx_workbook.mjs
```

That old workflow used a Codex-only package and does not work from normal Terminal.

## Troubleshooting

If `python3` is not found, install Python from:

```text
https://www.python.org/downloads/
```

If the SEC request fails, wait a few minutes and run the command again.

If the Excel file does not appear, run with an explicit output folder:

```bash
python3 scripts/build_company_workbooks.py --ticker FDX --output-dir "$HOME/OneDrive/ImportantFiles/Brian/sec-workbooks"
```
