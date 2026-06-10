# Company SEC/XBRL Excel Automation

This project uses **Python** to pull free SEC/XBRL data and create Excel workbooks.

Do **not** run `npm install @oai/artifact-tool`. That package is not public and will fail with a 404 error. Node.js is not needed.

## What This Gets

For each ticker, the automation creates one Excel workbook with:

- `Company Info`: ticker, company name, CIK, SEC company page, row counts.
- `Filings`: recent 10-K and 10-Q filing metadata.
- `Curated Metrics`: cleaner research fields such as revenue, net income, cash, debt, assets, liabilities, equity, capex, and cash flow.
- `Raw XBRL Facts`: detailed SEC/XBRL facts from company filings.
- `Research Notes`: blank sheet for manual analyst notes.

## Where Excel Files Save

By default, files are saved to:

```text
~/OneDrive/ImportantFiles/Brian/sec-workbooks
```

On Allison's Mac, that is usually:

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
export SEC_USER_AGENT="Allison Xu your-email@example.com"
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

Use another folder only if it is still inside your own approved project/work area. For this project, Allison's preferred folder is `OneDrive/ImportantFiles/Brian`.

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

## Refresh A Workbook

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
