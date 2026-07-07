import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

try {
const root = "C:/Users/LENOVO/Documents/Codex/2026-07-01/https-eldcm9-oss-github-io-brd";
const outputPath = path.join(root, "outputs", "Investment_Backend_Template.xlsx");
const scaffoldTemplatePath = path.join(root, "outputs", "laravel-excel-backend-scaffold", "storage", "app", "templates", "Investment_Backend_Template.xlsx");

const workbook = Workbook.create();
const sheets = {
  MASTER_PROJECT: {
    headers: [
      "project_code",
      "brd_number",
      "project_name",
      "division",
      "department",
      "site",
      "category",
      "priority",
      "vendor",
      "pic",
      "budget_requested",
      "budget_approved",
      "actual_cost",
      "planned_start_date",
      "planned_finish_date",
      "actual_start_date",
      "actual_finish_date",
      "progress",
      "status",
      "remarks",
    ],
    required: ["project_code", "project_name", "division", "category", "pic", "budget_approved", "planned_start_date", "planned_finish_date", "status"],
    rows: [
      ["PRJ-ERP-001", "BRD-26002", "ERP Finance Extension", "Finance", "Accounting", "Bandung", "IT System", "High", "SAP Partner Indonesia", "Dina Putri", 9500000000, 9200000000, 2100000000, new Date("2026-03-01"), new Date("2026-12-15"), new Date("2026-03-05"), null, 22, "Running", "Finance module extension for investment dashboard."],
      ["PRJ-WH-001", "BRD-26001", "Warehouse Automation Line", "Supply Chain", "Warehouse", "Jakarta", "Automation", "High", "Siemens", "Ari Wibowo", 20000000000, 18500000000, 13600000000, new Date("2026-01-10"), new Date("2026-09-30"), new Date("2026-01-12"), null, 68, "Running", "Automation line rollout."],
    ],
  },
  PROJECT_ACTIVITY: {
    headers: ["activity_id", "project_code", "activity_date", "activity_type", "activity_title", "description", "pic", "progress", "status", "notes"],
    required: ["project_code", "activity_date", "activity_type", "activity_title", "status"],
    rows: [
      ["ACT-ERP-001", "PRJ-ERP-001", new Date("2026-06-03"), "Milestone", "Procurement Review", "Committee notes closed.", "Dina Putri", 22, "Completed", "Validated in weekly review."],
      ["ACT-WH-001", "PRJ-WH-001", new Date("2026-06-18"), "Review", "PLC Integration Retest", "Retest conveyor PLC integration.", "Ari Wibowo", 68, "Running", "Safety interlock check included."],
    ],
  },
  PROJECT_DOCUMENT: {
    headers: ["document_id", "project_code", "document_type", "document_name", "version", "upload_date", "file_name", "remarks"],
    required: ["project_code", "document_type", "document_name"],
    rows: [
      ["DOC-ERP-001", "PRJ-ERP-001", "Cost Benefit Analysis", "ERP Finance CBA", "v1.0", new Date("2026-06-05"), "erp_finance_cba.pdf", "Metadata only; actual file upload is separate."],
      ["DOC-WH-001", "PRJ-WH-001", "Financial Summary", "Warehouse Automation Financial Summary", "v1.1", new Date("2026-06-20"), "warehouse_financial_summary.xlsx", "Metadata only."],
    ],
  },
  PROJECT_REVIEW: {
    headers: ["review_id", "project_code", "executive_summary", "timeline_performance", "budget_performance", "activity_summary", "benefit_realization", "lessons_learned", "recommendation", "closing_summary"],
    required: ["project_code"],
    rows: [
      ["REV-ERP-001", "PRJ-ERP-001", "ERP extension remains on track.", "Timeline is aligned with planned finish.", "Budget usage is below approved value.", "Procurement review completed.", "Faster finance reporting cycle.", "Keep vendor milestone gates visible.", "Continue execution monitoring.", ""],
      ["REV-WH-001", "PRJ-WH-001", "Warehouse automation requires integration focus.", "PLC retest is in progress.", "Budget utilization is within tolerance.", "Retest and installation activities active.", "Throughput improvement expected.", "Retest windows should be planned earlier.", "Monitor vendor dependency weekly.", ""],
    ],
  },
};

const widths = {
  MASTER_PROJECT: [16, 14, 28, 18, 18, 16, 18, 12, 24, 18, 18, 18, 18, 18, 18, 18, 18, 10, 14, 36],
  PROJECT_ACTIVITY: [16, 16, 16, 18, 28, 36, 18, 10, 14, 34],
  PROJECT_DOCUMENT: [16, 16, 22, 30, 12, 16, 30, 36],
  PROJECT_REVIEW: [16, 16, 34, 34, 34, 34, 34, 34, 34, 34],
};

function columnLetter(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

for (const [name, config] of Object.entries(sheets)) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const colCount = config.headers.length;
  const lastCol = columnLetter(colCount - 1);
  const notes = config.headers.map((header) => config.required.includes(header) ? "Required" : "Optional");

  sheet.getRange(`A1:${lastCol}1`).values = [[`${name} - Investment Backend Import Sheet`, ...Array(colCount - 1).fill("")]];
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: "#0F172A",
    font: { bold: true, color: "#FFFFFF", size: 14 },
  };

  sheet.getRangeByIndexes(1, 0, 1, colCount).values = [notes];
  sheet.getRangeByIndexes(1, 0, 1, colCount).format = {
    fill: "#E0F2FE",
    font: { bold: true, color: "#075985" },
  };

  sheet.getRangeByIndexes(2, 0, 1, colCount).values = [config.headers];
  sheet.getRangeByIndexes(2, 0, 1, colCount).format = {
    fill: "#2563EB",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
  sheet.getRangeByIndexes(3, 0, config.rows.length, colCount).values = config.rows;

  sheet.freezePanes.freezeRows(3);
  sheet.getRangeByIndexes(0, 0, 3 + config.rows.length, colCount).format.borders = { preset: "inside", style: "thin", color: "#CBD5E1" };

  widths[name].forEach((width, index) => {
    sheet.getRange(`${columnLetter(index)}:${columnLetter(index)}`).format.columnWidth = width;
  });

  const dateColumns = config.headers
    .map((header, index) => header.includes("date") ? index : -1)
    .filter((index) => index >= 0);
  for (const index of dateColumns) {
    sheet.getRangeByIndexes(3, index, 100, 1).format.numberFormat = "yyyy-mm-dd";
  }

  for (const moneyColumn of ["budget_requested", "budget_approved", "actual_cost"]) {
    const index = config.headers.indexOf(moneyColumn);
    if (index >= 0) {
      sheet.getRangeByIndexes(3, index, 100, 1).format.numberFormat = '"Rp" #,##0';
    }
  }

  const progressIndex = config.headers.indexOf("progress");
  if (progressIndex >= 0) {
    sheet.getRangeByIndexes(3, progressIndex, 100, 1).format.numberFormat = "0";
  }

  sheet.getRange(`A2:${lastCol}${3 + config.rows.length}`).format.wrapText = true;
  sheet.getRange(`A1:${lastCol}${3 + config.rows.length}`).format.autofitRows();
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(path.dirname(scaffoldTemplatePath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
await fs.copyFile(outputPath, scaffoldTemplatePath);

} catch (error) {
  console.error(error?.stack || error);
  process.exit(1);
}
