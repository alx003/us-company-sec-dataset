import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const ROOT = path.resolve(new URL("..", import.meta.url).pathname);
const OUTPUT_DIR = process.env.FDX_OUTPUT_DIR || "/Users/allisonxu/OneDrive/文件/Brian";
const OUTPUT_FILE = path.join(OUTPUT_DIR, "FDX_fedex_sec_collection.xlsx");
const USER_AGENT = "alx003 SEC research workbook contact@example.com";
const TICKER = "FDX";
const MIN_YEAR = 2018;
const FORMS = new Set(["10-K", "10-K/A", "10-Q", "10-Q/A"]);

const headers = {
  "User-Agent": USER_AGENT,
  "Accept-Encoding": "gzip, deflate",
};

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchJson(url) {
  const response = await fetch(url, { headers });
  if (!response.ok) {
    throw new Error(`SEC request failed ${response.status} ${response.statusText}: ${url}`);
  }
  await sleep(150);
  return response.json();
}

async function fetchText(url) {
  const response = await fetch(url, { headers });
  if (!response.ok) {
    throw new Error(`SEC request failed ${response.status} ${response.statusText}: ${url}`);
  }
  await sleep(150);
  return response.text();
}

function cleanText(value) {
  return String(value ?? "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function attrs(tag) {
  const result = {};
  for (const match of tag.matchAll(/([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g)) {
    result[match[1].toLowerCase()] = match[3] ?? match[4] ?? "";
  }
  return result;
}

function localName(qname) {
  return String(qname || "").split(":").pop();
}

function parseNumber(text, scale = "0", sign = "") {
  const cleaned = cleanText(text).replace(/,/g, "");
  if (!cleaned || cleaned === "-" || cleaned.toLowerCase() === "nil") return null;
  const number = Number(cleaned.replace(/[()]/g, ""));
  if (!Number.isFinite(number)) return null;
  const multiplier = 10 ** Number(scale || 0);
  const negative = sign === "-" || /^\(.*\)$/.test(cleaned);
  return (negative ? -number : number) * multiplier;
}

function statementArea(concept, label, unit, item = {}) {
  const text = `${concept} ${label}`.toLowerCase();
  if (/accumulateddepreciation|accumulated depreciation/.test(text)) {
    return "Balance Sheet";
  }
  if (item.frame || item.start) {
    if (/cash|proceeds|payments|acquir|purchases?|capitalexpenditure|depreciation|amortization/.test(text)) {
      return "Cash Flow Statement";
    }
    if (/revenue|sales|income|loss|earnings|expense|cost|profit|margin|tax|interest|eps|sharebased|benefit/.test(text)) {
      return "Income Statement";
    }
  }
  if (/asset|liabilit|equity|stockholder|cashandcashequivalents|receivable|inventory|debt|lease|goodwill|propertyplant|payable|retainedearnings|treasury/.test(text)) {
    return "Balance Sheet";
  }
  if (/shares|stock|commonstock/.test(text) && unit === "shares") {
    return "Balance Sheet / Equity";
  }
  return "Other SEC XBRL Line Item";
}

function readableMetricName(concept, label = "") {
  if (label) return label;
  return String(concept || "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .replace(/\s+/g, " ")
    .trim();
}

function metricGroup(concept, label = "", statement = "") {
  const text = `${concept} ${label}`.toLowerCase();
  if (/depreciation|amortization|depletion/.test(text)) return "Depreciation & Amortization";
  if (/capitalexpenditure|payments?toacquireproperty|capex|expenditure/.test(text)) return "Capex";
  if (/revenue|sales/.test(text)) return "Revenue";
  if (/impairment|writeoff|write-off/.test(text)) return "Expense / Cost";
  if (/grossprofit|operatingincomeloss|netincomeloss|profit|income|loss|earnings/.test(text)) return "Profit / Income";
  if (/asset|propertyplant|inventory|receivable|cashandcash/.test(text)) return "Assets";
  if (/liabilit|payable/.test(text)) return "Liabilities";
  if (/debt|borrowings|notespayable|leaseobligation/.test(text)) return "Debt";
  if (/equity|stockholder|retainedearnings|treasury/.test(text)) return "Equity";
  if (/cash|operatingactivities|investingactivities|financingactivities|proceeds|payments/.test(text)) return "Cash Flow";
  if (/expense|cost/.test(text)) return "Expense / Cost";
  if (/eps|pershare|share/.test(text)) return "Per Share / Shares";
  if (/tax/.test(text)) return "Tax";
  return statement || "Other";
}

function periodType(item) {
  if (item.start && item.end) return "Duration";
  if (item.end) return "Instant";
  return "";
}

function secCompanyUrl(cik, accession) {
  if (!accession) return `https://www.sec.gov/edgar/browse/?CIK=${Number(cik)}`;
  return `https://www.sec.gov/Archives/edgar/data/${Number(cik)}/${String(accession).replaceAll("-", "")}/`;
}

function accessionBaseUrl(cik, accession) {
  return `https://www.sec.gov/Archives/edgar/data/${Number(cik)}/${String(accession).replaceAll("-", "")}/`;
}

function filingRows(submissions, cik) {
  const recent = submissions.filings?.recent || {};
  const rows = [];
  for (let i = 0; i < (recent.form || []).length; i += 1) {
    const form = recent.form[i];
    const fiscalYear = Number(recent.fy?.[i] || "");
    if (!FORMS.has(form) || (fiscalYear && fiscalYear < MIN_YEAR)) continue;
    rows.push({
      accession: recent.accessionNumber?.[i] || "",
      form,
      fiscalYear,
      fiscalPeriod: recent.fp?.[i] || "",
      reportDate: recent.reportDate?.[i] || "",
      filingDate: recent.filingDate?.[i] || "",
      primaryDocument: recent.primaryDocument?.[i] || "",
      sourceUrl: secCompanyUrl(cik, recent.accessionNumber?.[i] || ""),
    });
  }
  return rows;
}

function factRows(company, facts, filingByAccession) {
  const rows = [];
  for (const [taxonomy, concepts] of Object.entries(facts.facts || {})) {
    for (const [concept, payload] of Object.entries(concepts || {})) {
      for (const [unit, values] of Object.entries(payload.units || {})) {
        for (const item of values || []) {
          if (!FORMS.has(item.form) || (item.fy && Number(item.fy) < MIN_YEAR)) continue;
          const accession = item.accn || "";
          const filing = filingByAccession.get(accession) || {};
          const statement = statementArea(concept, payload.label, unit, item);
          rows.push([
            "Line Item",
            statement,
            metricGroup(concept, payload.label, statement),
            readableMetricName(concept, payload.label),
            payload.label || "",
            item.val ?? "",
            unit,
            item.form || "",
            Number(item.fy || filing.fiscalYear || "") || "",
            item.fp || filing.fiscalPeriod || "",
            periodType(item),
            item.start || "",
            item.end || "",
            item.filed || filing.filingDate || "",
            item.frame || "",
            "",
            "",
            taxonomy,
            concept,
            "",
            accession,
            secCompanyUrl(company.cik, accession),
          ]);
        }
      }
    }
  }
  return rows;
}

function parseContexts(html) {
  const contexts = new Map();
  const contextRe = /<(?:xbrli:)?context\b[^>]*\bid\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/(?:xbrli:)?context>/gi;
  for (const match of html.matchAll(contextRe)) {
    const id = match[1];
    const body = match[2];
    const dimensions = [];
    const dimRe = /<(?:xbrldi:)?explicitmember\b([^>]*)>([\s\S]*?)<\/(?:xbrldi:)?explicitmember>/gi;
    for (const dim of body.matchAll(dimRe)) {
      const attr = attrs(dim[1]);
      dimensions.push(`${localName(attr.dimension)}=${localName(cleanText(dim[2]))}`);
    }
    const start = body.match(/<(?:xbrli:)?startdate>([^<]+)<\/(?:xbrli:)?startdate>/i)?.[1] || "";
    const end = body.match(/<(?:xbrli:)?enddate>([^<]+)<\/(?:xbrli:)?enddate>/i)?.[1]
      || body.match(/<(?:xbrli:)?instant>([^<]+)<\/(?:xbrli:)?instant>/i)?.[1]
      || "";
    contexts.set(id, { dimensions, start, end });
  }
  return contexts;
}

function isSegmentOrGeography(dimensions) {
  const text = dimensions.join(" ").toLowerCase();
  return /segment|geograph|region|product|service|business|operating|country|domestic|international|express|freight|office|ground/.test(text);
}

function isSegmentMetric(concept) {
  return /revenue|sales|income|loss|profit|earnings|capex|capital|expenditure|depreciation|amortization|asset|propertyplant|expense/i.test(concept);
}

function parseInlineSegmentFacts(html, filing, company) {
  const contexts = parseContexts(html);
  const rows = [];
  const factRe = /<ix:nonfraction\b([^>]*)>([\s\S]*?)<\/ix:nonfraction>/gi;
  for (const match of html.matchAll(factRe)) {
    const attr = attrs(match[1]);
    const context = contexts.get(attr.contextref);
    if (!context || !context.dimensions.length) continue;
    const concept = localName(attr.name);
    if (!isSegmentOrGeography(context.dimensions) || !isSegmentMetric(concept)) continue;
    const value = parseNumber(match[2], attr.scale, attr.sign);
    if (value === null) continue;
    const dimensions = context.dimensions.join("; ");
    const statement = /geograph|region|country|domestic|international/i.test(dimensions) ? "Geography XBRL Detail" : "Segment XBRL Detail";
    rows.push([
      "Segment / Geography",
      statement,
      metricGroup(concept, "", statement),
      readableMetricName(concept),
      "",
      value,
      attr.unitref || "",
      filing.form,
      filing.fiscalYear || "",
      filing.fiscalPeriod || "",
      context.start ? "Duration" : "Instant",
      context.start,
      context.end,
      filing.filingDate || "",
      "",
      dimensions,
      dimensions,
      String(attr.name || "").split(":")[0] || "",
      concept,
      filing.primaryDocument,
      filing.accession,
      `${accessionBaseUrl(company.cik, filing.accession)}${filing.primaryDocument}`,
    ]);
  }
  return rows;
}

async function segmentRows(company, filings) {
  const rows = [];
  for (const filing of filings) {
    if (!filing.primaryDocument || !filing.accession) continue;
    const url = `${accessionBaseUrl(company.cik, filing.accession)}${filing.primaryDocument}`;
    try {
      const html = await fetchText(url);
      rows.push(...parseInlineSegmentFacts(html, filing, company));
    } catch (error) {
      rows.push([
        "Segment / Geography",
        "Fetch Warning",
        "Fetch Warning",
        "Filing Download Error",
        error.message,
        "",
        "",
        filing.form,
        filing.fiscalYear || "",
        filing.fiscalPeriod || "",
        "",
        "",
        "",
        filing.filingDate || "",
        "",
        "",
        "",
        "FilingDownloadError",
        "",
        filing.primaryDocument,
        filing.accession,
        url,
      ]);
    }
  }
  return rows;
}

function colLetter(index) {
  let n = index;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

function addValues(sheet, startRow, startCol, rows) {
  if (!rows.length) return;
  const endRow = startRow + rows.length - 1;
  const endCol = startCol + rows[0].length - 1;
  sheet.getRange(`${colLetter(startCol)}${startRow}:${colLetter(endCol)}${endRow}`).values = rows;
}

function styleSheet(sheet, rowCount, colCount) {
  const lastCol = colLetter(colCount);
  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: "#17365D",
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange(`A1:${lastCol}${Math.max(rowCount, 2)}`).format.font = { name: "Aptos", size: 10 };
  sheet.getRange(`A1:${lastCol}${Math.max(rowCount, 2)}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#D9E2F3",
  };
  sheet.getRange(`F2:F${Math.max(rowCount, 2)}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
  sheet.getRange(`I2:I${Math.max(rowCount, 2)}`).format.numberFormat = "0";
  for (let i = 1; i <= colCount; i += 1) {
    const width = [16, 22, 24, 42, 38, 18, 12, 10, 10, 10, 12, 12, 12, 12, 14, 46, 46, 13, 34, 22, 20, 52][i - 1] || 15;
    sheet.getRange(`${colLetter(i)}:${colLetter(i)}`).format.columnWidth = width;
  }
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(4);
}

async function buildWorkbook(collectionRows, company) {
  const workbook = Workbook.create();
  const collection = workbook.worksheets.getOrAdd("Collection", { renameFirstIfOnlyNewSpreadsheet: true });
  const research = workbook.worksheets.add("Research");
  const companyInfo = workbook.worksheets.add("Company Info");
  workbook.worksheets.add("Blank");

  const collectionHeaders = [
    "Data Group",
    "Statement / Detail Area",
    "Metric Group",
    "Metric Name",
    "XBRL Label",
    "Value",
    "Unit",
    "Form",
    "Fiscal Year",
    "Fiscal Period",
    "Period Type",
    "Period Start",
    "Period End",
    "Filed Date",
    "SEC Frame",
    "Dimensions",
    "Segment / Geography Member",
    "Taxonomy",
    "XBRL Concept",
    "Primary Document",
    "Accession Number",
    "Source URL",
  ];
  addValues(collection, 1, 1, [collectionHeaders, ...collectionRows]);
  styleSheet(collection, collectionRows.length + 1, collectionHeaders.length);

  const researchRows = [
    ["Date", "Topic", "Analyst Notes", "Source / Link", "Follow-up"],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
  ];
  addValues(research, 1, 1, researchRows);
  research.getRange("A1:E1").format = {
    fill: "#17365D",
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  research.getRange("A1:E20").format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
  research.getRange("A:A").format.columnWidth = 14;
  research.getRange("B:B").format.columnWidth = 24;
  research.getRange("C:C").format.columnWidth = 56;
  research.getRange("D:D").format.columnWidth = 36;
  research.getRange("E:E").format.columnWidth = 28;
  research.getRange("A2:A20").format.numberFormat = "yyyy-mm-dd";
  research.getRange("C2:E20").format.wrapText = true;
  research.freezePanes.freezeRows(1);

  const companyInfoRows = [
    ["Field", "Value"],
    ["Ticker", company.ticker],
    ["Company", company.name],
    ["CIK", company.cik],
    ["SEC Company Page", secCompanyUrl(company.cik)],
    ["Generated From", "SEC companyfacts XBRL and recent inline XBRL filings"],
    ["Collection Rows", collectionRows.length],
    ["Line Item Rows", collectionRows.filter((row) => row[0] === "Line Item").length],
    ["Segment / Geography Rows", collectionRows.filter((row) => row[0] === "Segment / Geography").length],
  ];
  addValues(companyInfo, 1, 1, companyInfoRows);
  companyInfo.getRange("A1:B1").format = {
    fill: "#17365D",
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  companyInfo.getRange("A1:B9").format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
  companyInfo.getRange("A:A").format.columnWidth = 24;
  companyInfo.getRange("B:B").format.columnWidth = 72;

  return workbook;
}

async function main() {
  const tickerData = await fetchJson("https://www.sec.gov/files/company_tickers_exchange.json");
  const tickerFields = tickerData.fields;
  const tickerRows = tickerData.data.map((row) => Object.fromEntries(tickerFields.map((field, i) => [field, row[i]])));
  const match = tickerRows.find((row) => row.ticker === TICKER);
  if (!match) throw new Error(`Could not find ${TICKER} in SEC ticker file.`);

  const company = {
    ticker: TICKER,
    name: match.name,
    cik: String(match.cik).padStart(10, "0"),
  };

  const submissions = await fetchJson(`https://data.sec.gov/submissions/CIK${company.cik}.json`);
  const facts = await fetchJson(`https://data.sec.gov/api/xbrl/companyfacts/CIK${company.cik}.json`);
  const filings = filingRows(submissions, company.cik);
  const filingByAccession = new Map(filings.map((filing) => [filing.accession, filing]));
  const rows = factRows(company, facts, filingByAccession);
  rows.push(...await segmentRows(company, filings));
  rows.sort((a, b) => {
    const left = `${a[0]} ${a[2]} ${a[8]} ${a[9]} ${a[18]} ${a[15]}`;
    const right = `${b[0]} ${b[2]} ${b[8]} ${b[9]} ${b[18]} ${b[15]}`;
    return left.localeCompare(right);
  });

  const workbook = await buildWorkbook(rows, company);

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
  });
  if (errors.ndjson && errors.ndjson.includes("#")) {
    throw new Error(`Formula error scan found issues:\n${errors.ndjson}`);
  }
  await workbook.render({ sheetName: "Collection", range: "A1:V25", scale: 1 });
  await workbook.render({ sheetName: "Research", range: "A1:E12", scale: 1 });
  await workbook.render({ sheetName: "Company Info", range: "A1:B9", scale: 1 });
  await workbook.render({ sheetName: "Blank", range: "A1:E10", scale: 1 });

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(OUTPUT_FILE);
  console.log(JSON.stringify({
    output: OUTPUT_FILE,
    collectionRows: rows.length,
    lineItemRows: rows.filter((row) => row[0] === "Line Item").length,
    segmentGeographyRows: rows.filter((row) => row[0] === "Segment / Geography").length,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
