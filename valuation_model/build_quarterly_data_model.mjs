import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve("..");
const projectDir = path.join(root, "Eternal-Valuation");
const outputDir = path.join(projectDir, "outputs", "v2.0_quarterly_data");
const masterDir = path.join(projectDir, "data", "master_dataset");
const previewDir = path.join(outputDir, "previews");

const periods = ["Q4FY25", "Q1FY26", "Q2FY26", "Q3FY26", "Q4FY26"];

const sourceIds = [
  ["S5", "Eternal_Q1FY26.pdf", "Quarterly results / shareholder letter", "Q1FY26", "Official quarterly filing", "Used as cross-check for Q1 FY26 values."],
  ["S6", "Eternal_Q2FY26.pdf", "Quarterly results / shareholder letter", "Q2FY26", "Official quarterly filing", "Used as cross-check for Q2 FY26 values."],
  ["S7", "Eternal_Q3FY26.pdf", "Quarterly results / shareholder letter", "Q3FY26", "Official quarterly filing", "Used as cross-check for Q3 FY26 values."],
  ["S8", "Eternal_Q4FY26.pdf", "Quarterly results / shareholder letter", "Q4FY26", "Official quarterly filing", "Primary source for the rolling five-quarter table used in this workbook."],
];

const src = {
  q4A: "S8 | Eternal_Q4FY26.pdf | PDF p12 / filing p11 | Annexure A - Quarterly disclosures",
  q4FoodQc: "S8 | Eternal_Q4FY26.pdf | PDF p13 / filing p12 | Food delivery and quick commerce metrics",
  q4GoHpCash: "S8 | Eternal_Q4FY26.pdf | PDF p14 / filing p13 | Going-out, Hyperpure, cash balance",
  q4Recon: "S8 | Eternal_Q4FY26.pdf | PDF p15 / filing p14 | Annexure B - Adjusted Revenue and Adjusted EBITDA reconciliation",
};

const rows = [
  ["Consolidated", "Food delivery NOV", "INR crore", 8210, 8967, 9423, 9846, 9757, src.q4A, "Verified visually", "Required", "Official quarterly filing", "B2C NOV table"],
  ["Consolidated", "Quick commerce NOV", "INR crore", 7362, 9203, 11679, 13300, 14386, src.q4A, "Verified visually", "Required", "Official quarterly filing", "B2C NOV table"],
  ["Consolidated", "Going-out NOV", "INR crore", 1868, 2013, 2063, 2587, 2736, src.q4A, "Verified visually", "Required", "Official quarterly filing", "B2C NOV table"],
  ["Consolidated", "B2C NOV", "INR crore", 17440, 20183, 23164, 25732, 26880, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Total per filing"],
  ["Consolidated", "B2C NOV YoY growth", "%", 0.534, 0.55, 0.568, 0.548, 0.541, src.q4A, "Verified visually", "Optional", "Official quarterly filing", "Reported YoY growth"],

  ["Consolidated", "Food delivery adjusted revenue", "INR crore", 2409, 2657, 2863, 3053, 3125, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted revenue"],
  ["Consolidated", "Quick commerce adjusted revenue", "INR crore", 1709, 2400, 9891, 12256, 13232, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted revenue"],
  ["Consolidated", "Going-out adjusted revenue", "INR crore", 229, 207, 189, 300, 277, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted revenue"],
  ["Consolidated", "B2B supplies (Hyperpure) adjusted revenue", "INR crore", 1840, 2295, 1023, 1070, 978, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted revenue"],
  ["Consolidated", "Others adjusted revenue", "INR crore", 1, 4, 2, 13, 68, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted revenue"],
  ["Consolidated", "Adjusted revenue", "INR crore", 6188, 7563, 13968, 16692, 17680, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Reported consolidated total"],
  ["Consolidated", "Adjusted revenue YoY growth", "%", 0.598, 0.673, 1.724, 1.905, 1.857, src.q4A, "Verified visually", "Optional", "Official quarterly filing", "Reported YoY growth"],

  ["Consolidated", "Food delivery adjusted EBITDA", "INR crore", 428, 451, 503, 531, 532, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted EBITDA"],
  ["Consolidated", "Quick commerce adjusted EBITDA", "INR crore", -178, -162, -156, 4, 37, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted EBITDA"],
  ["Consolidated", "Going-out adjusted EBITDA", "INR crore", -47, -54, -63, -121, -81, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted EBITDA"],
  ["Consolidated", "B2B supplies (Hyperpure) adjusted EBITDA", "INR crore", -22, -18, -5, 1, 5, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted EBITDA"],
  ["Consolidated", "Others adjusted EBITDA", "INR crore", -16, -45, -55, -51, -64, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Segment adjusted EBITDA"],
  ["Consolidated", "Adjusted EBITDA", "INR crore", 165, 172, 224, 364, 429, src.q4A, "Verified visually", "Required", "Official quarterly filing", "Reported consolidated total"],

  ["Food delivery", "NOV", "INR crore", 8210, 8967, 9423, 9846, 9757, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "NOV YoY growth", "%", 0.142, 0.131, 0.138, 0.166, 0.188, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Adjusted revenue", "INR crore", 2409, 2657, 2863, 3053, 3125, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Adjusted revenue YoY growth", "%", 0.175, 0.178, 0.224, 0.265, 0.297, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Contribution", "INR crore", 842, 885, 976, 1023, 998, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Contribution margin", "% of NOV", 0.103, 0.099, 0.104, 0.104, 0.102, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Adjusted EBITDA", "INR crore", 428, 451, 503, 531, 532, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Adjusted EBITDA margin", "% of NOV", 0.052, 0.05, 0.053, 0.054, 0.055, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Food delivery", "Average monthly transacting customers", "million", 20.9, 22.9, 24.1, 24.9, 25.4, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Operating metrics"],
  ["Food delivery", "Average monthly active restaurant partners", "'000", 314, 313, 327, 336, 344, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Operating metrics"],
  ["Food delivery", "Average monthly active delivery partners", "'000", 444, 509, 555, 567, 576, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Operating metrics"],

  ["Quick commerce", "NOV", "INR crore", 7362, 9203, 11679, 13300, 14386, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "NOV YoY growth", "%", 1.207, 1.266, 1.37, 1.209, 0.954, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Revenue", "INR crore", 1709, 2400, 9891, 12256, 13232, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Revenue YoY growth", "%", 1.222, 1.548, 7.556, 7.761, 6.743, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Gross profit", "INR crore", 1709, 2168, 3132, 3539, 3867, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Gross profit margin", "% of NOV", 0.232, 0.236, 0.268, 0.266, 0.269, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Contribution", "INR crore", 289, 360, 542, 736, 782, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Contribution margin", "% of NOV", 0.039, 0.039, 0.046, 0.055, 0.054, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Adjusted EBITDA", "INR crore", -178, -162, -156, 4, 37, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Adjusted EBITDA margin", "% of NOV", -0.024, -0.018, -0.013, 0, 0.003, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Quick commerce", "Orders", "million", 141.7, 176.7, 222.7, 243.3, 273.9, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Operating metrics"],
  ["Quick commerce", "Net average order value", "INR", 520, 521, 524, 547, 525, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Operating metrics"],
  ["Quick commerce", "Average monthly transacting customers", "million", 13.7, 16.9, 20.8, 23.6, 27.2, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Operating metrics"],
  ["Quick commerce", "Average monthly active delivery partners", "'000", 185, 243, 339, 369, 409, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Operating metrics"],
  ["Quick commerce", "Average NOV per day per store", "INR '000", 736, 734, 771, 750, 768, src.q4FoodQc, "Verified visually", "Optional", "Official quarterly filing", "Operating metrics"],
  ["Quick commerce", "Stores at period end", "#", 1301, 1544, 1816, 2027, 2243, src.q4FoodQc, "Verified visually", "Required", "Official quarterly filing", "Operating metrics"],

  ["Going-out", "NOV", "INR crore", 1868, 2013, 2063, 2587, 2736, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Going-out", "NOV YoY growth", "%", 1.222, 0.953, 0.32, 0.199, 0.465, src.q4GoHpCash, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Going-out", "Revenue", "INR crore", 229, 207, 189, 300, 277, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Going-out", "Revenue YoY growth", "%", 1.462, 1.179, 0.227, 0.158, 0.21, src.q4GoHpCash, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["Going-out", "Adjusted EBITDA", "INR crore", -47, -54, -63, -121, -81, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["Going-out", "Adjusted EBITDA margin", "% of NOV", -0.025, -0.027, -0.031, -0.047, -0.03, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],

  ["B2B supplies (Hyperpure)", "Revenue", "INR crore", 1840, 2295, 1023, 1070, 978, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["B2B supplies (Hyperpure)", "Revenue YoY growth", "%", 0.935, 0.894, -0.305, -0.36, -0.468, src.q4GoHpCash, "Verified visually", "Optional", "Official quarterly filing", "Segment table"],
  ["B2B supplies (Hyperpure)", "Adjusted EBITDA", "INR crore", -22, -18, -5, 1, 5, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],
  ["B2B supplies (Hyperpure)", "Adjusted EBITDA margin", "% of Revenue", -0.012, -0.008, -0.005, 0.001, 0.005, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Segment table"],

  ["Cash bridge", "Adjusted EBITDA", "INR crore", 165, 172, 224, 364, 429, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Treasury income received", "INR crore", 195, 235, 171, 228, 319, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Capital expenditure incurred", "INR crore", -317, -370, -417, -465, -494, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Other items", "INR crore", 148, 107, -39, 69, -157, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Cash burn / surplus", "INR crore", 191, 144, -61, 196, 97, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Increase / decrease in net working capital", "INR crore", -602, -111, -482, -690, 55, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Change in cash", "INR crore", -411, 33, -543, -494, 152, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Opening cash balance", "INR crore", 19235, 18824, 18857, 18314, 17820, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],
  ["Cash bridge", "Closing cash balance", "INR crore", 18824, 18857, 18314, 17820, 17972, src.q4GoHpCash, "Verified visually", "Required", "Official quarterly filing", "Cash balance table"],

  ["Reconciliation", "Revenue from operations", "INR crore", 5833, 7167, 13590, 16315, 17292, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted Revenue reconciliation"],
  ["Reconciliation", "Actual customer delivery charges paid in food delivery", "INR crore", 246, 273, 225, 189, 188, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted Revenue reconciliation"],
  ["Reconciliation", "Gross platform fee and other charges paid on food delivery orders", "INR crore", 186, 213, 254, 304, 315, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted Revenue reconciliation"],
  ["Reconciliation", "Platform fee and other charges already included in revenue", "INR crore", 77, 89, 101, 116, 115, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted Revenue reconciliation"],
  ["Reconciliation", "Adjusted revenue", "INR crore", 6188, 7563, 13968, 16692, 17680, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted Revenue reconciliation"],
  ["Reconciliation", "Adjusted EBITDA", "INR crore", 165, 172, 224, 364, 429, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Other income", "INR crore", 368, 354, 352, 348, 342, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Rental paid pertaining to Ind AS 116 leases", "INR crore", 124, 153, 189, 231, 265, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Depreciation and amortization expense", "INR crore", 287, 314, 376, 439, 468, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Finance cost", "INR crore", 56, 67, 86, 107, 132, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "ESOP expense", "INR crore", 217, 210, 174, 227, 208, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Exceptional items", "INR crore", 0, 0, 0, 0, 0, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Tax expense", "INR crore", 58, 63, 64, 68, 54, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
  ["Reconciliation", "Profit / (loss) for the period", "INR crore", 39, 25, 65, 102, 174, src.q4Recon, "Verified visually", "Required", "Official quarterly filing", "Adjusted EBITDA reconciliation"],
];

const workbook = Workbook.create();
workbook.comments.setSelf({ displayName: "User" });

function setTitle(sheet, title, subtitle) {
  sheet.getRange("A1:J1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#111827",
    font: { color: "#FFFFFF", bold: true, size: 16 },
    horizontalAlignment: "left",
  };
  sheet.getRange("A2:J2").merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    fill: "#E8EEF8",
    font: { color: "#111827", italic: true },
    wrapText: true,
  };
}

function styleHeader(range) {
  range.format = {
    fill: "#1F4E78",
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function styleDataRange(range) {
  range.format.borders = {
    insideHorizontal: { style: "thin", color: "#D9E2F3" },
    bottom: { style: "thin", color: "#B4C6E7" },
  };
}

function formatPeriodBlock(sheet, startRow, endRow, startCol = 4, endCol = 8) {
  const rng = sheet.getRangeByIndexes(startRow - 1, startCol - 1, endRow - startRow + 1, endCol - startCol + 1);
  rng.format.numberFormat = '#,##0;[Red](#,##0);-';
}

function applyPercentFormats(sheet, firstRow, lastRow, unitCol = 3, startCol = 4, endCol = 8) {
  for (let r = firstRow; r <= lastRow; r++) {
    const unit = sheet.getRange(`C${r}`).values?.[0]?.[0];
    if (String(unit).includes("%")) {
      sheet.getRangeByIndexes(r - 1, startCol - 1, 1, endCol - startCol + 1).format.numberFormat = "0.0%;[Red](0.0%);-";
    }
  }
}

function autosize(sheet, widths) {
  widths.forEach(([col, width]) => {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  });
}

const cover = workbook.worksheets.add("00_Cover");
cover.showGridLines = false;
setTitle(cover, "Eternal Limited - Quarterly Data Pack", "Quarterly-only extraction for FY26 source-backed valuation work. Annual reports are intentionally excluded from this file.");
cover.getRange("A4:B11").values = [
  ["Company", "Eternal Limited / Zomato"],
  ["Scope", "Quarterly results only"],
  ["Periods", "Q4FY25 comparator plus Q1FY26-Q4FY26"],
  ["Primary Source", "Eternal_Q4FY26.pdf rolling quarterly disclosure tables"],
  ["Cross-check Sources", "Eternal_Q1FY26.pdf, Eternal_Q2FY26.pdf, Eternal_Q3FY26.pdf"],
  ["Currency / units", "INR crore unless otherwise noted"],
  ["Prepared for", "Phase 2 quarterly data extraction"],
  ["Status", "Quarterly data populated and checks added"],
];
cover.getRange("A4:A11").format = { fill: "#D9EAF7", font: { bold: true } };
cover.getRange("A4:B11").format.borders = { preset: "outside", style: "thin", color: "#B4C6E7" };
cover.getRange("A13:J13").values = [["Next use", "This file should feed 03_Raw_Data only after the annual historical dataset is complete and the final set of annual/quarterly source conventions is locked.", "", "", "", "", "", "", "", ""]];
cover.getRange("A13:J13").merge(true);
cover.getRange("A13").format = { fill: "#FFF2CC", font: { bold: true }, wrapText: true };
autosize(cover, [["A", 26], ["B", 105]]);

const master = workbook.worksheets.add("01_Quarterly_Master");
master.showGridLines = false;
setTitle(master, "Quarterly Master Extraction Table", "One row per metric with period values, source reference, confidence, priority, and extraction notes.");
const masterHeaders = ["Category", "Metric", "Units", ...periods, "Source Reference", "Verification Status", "Priority", "Confidence", "Notes"];
master.getRangeByIndexes(3, 0, 1, masterHeaders.length).values = [masterHeaders];
styleHeader(master.getRangeByIndexes(3, 0, 1, masterHeaders.length));
master.getRangeByIndexes(4, 0, rows.length, masterHeaders.length).values = rows;
styleDataRange(master.getRangeByIndexes(4, 0, rows.length, masterHeaders.length));
formatPeriodBlock(master, 5, rows.length + 4);
applyPercentFormats(master, 5, rows.length + 4);
master.getRange(`A4:M${rows.length + 4}`).format.wrapText = true;
master.freezePanes.freezeRows(4);
master.tables.add(`A4:M${rows.length + 4}`, true, "QuarterlyMasterTable");
autosize(master, [["A", 26], ["B", 48], ["C", 18], ["D", 14], ["E", 14], ["F", 14], ["G", 14], ["H", 14], ["I", 62], ["J", 20], ["K", 13], ["L", 24], ["M", 32]]);

const kpi = workbook.worksheets.add("02_Segment_KPIs");
kpi.showGridLines = false;
setTitle(kpi, "Segment KPI View", "Selected operating and profitability metrics by segment, copied from the master extraction rows.");
const kpiRows = rows.filter(r => ["Food delivery", "Quick commerce", "Going-out", "B2B supplies (Hyperpure)"].includes(r[0]));
kpi.getRangeByIndexes(3, 0, 1, masterHeaders.length).values = [masterHeaders];
styleHeader(kpi.getRangeByIndexes(3, 0, 1, masterHeaders.length));
kpi.getRangeByIndexes(4, 0, kpiRows.length, masterHeaders.length).values = kpiRows;
styleDataRange(kpi.getRangeByIndexes(4, 0, kpiRows.length, masterHeaders.length));
formatPeriodBlock(kpi, 5, kpiRows.length + 4);
applyPercentFormats(kpi, 5, kpiRows.length + 4);
kpi.getRange(`A4:M${kpiRows.length + 4}`).format.wrapText = true;
kpi.freezePanes.freezeRows(4);
kpi.tables.add(`A4:M${kpiRows.length + 4}`, true, "SegmentKPITable");
autosize(kpi, [["A", 26], ["B", 48], ["C", 18], ["D", 14], ["E", 14], ["F", 14], ["G", 14], ["H", 14], ["I", 62], ["J", 20], ["K", 13], ["L", 24], ["M", 32]]);

const checks = workbook.worksheets.add("03_Checks");
checks.showGridLines = false;
setTitle(checks, "Quarterly Data Checks", "Formula-driven checks for source totals, adjusted revenue reconciliation, PAT bridge, and cash movement.");
const checkHeaders = ["Check", "Q4FY25", "Q1FY26", "Q2FY26", "Q3FY26", "Q4FY26", "Tolerance", "Status", "Notes"];
checks.getRange("A4:I4").values = [checkHeaders];
styleHeader(checks.getRange("A4:I4"));
const checkLabels = [
  ["B2C NOV total equals segment NOV", "Reported B2C NOV less sum of food, quick commerce, and going-out NOV"],
  ["Adjusted revenue equals segment sum", "Reported adjusted revenue less segment adjusted revenue total"],
  ["Adjusted EBITDA equals segment sum", "Reported adjusted EBITDA less segment adjusted EBITDA total"],
  ["Adjusted revenue reconciliation", "Revenue from ops + delivery charges + gross platform fees - included platform fees - adjusted revenue"],
  ["PAT bridge from adjusted EBITDA", "Adjusted EBITDA + other income + rent paid - D&A - finance cost - ESOP - exceptional items - tax - PAT"],
  ["Cash bridge subtotal", "Adjusted EBITDA + treasury income + capex + other items - cash burn/surplus"],
  ["Cash change bridge", "Cash burn/surplus + NWC movement - change in cash"],
  ["Closing cash roll-forward", "Opening cash + change in cash - closing cash"],
];
checks.getRangeByIndexes(4, 0, checkLabels.length, 1).values = checkLabels.map(r => [r[0]]);
checks.getRangeByIndexes(4, 8, checkLabels.length, 1).values = checkLabels.map(r => [r[1]]);
const formulas = [
  ["=('01_Quarterly_Master'!D8-SUM('01_Quarterly_Master'!D5:D7))", "=('01_Quarterly_Master'!E8-SUM('01_Quarterly_Master'!E5:E7))", "=('01_Quarterly_Master'!F8-SUM('01_Quarterly_Master'!F5:F7))", "=('01_Quarterly_Master'!G8-SUM('01_Quarterly_Master'!G5:G7))", "=('01_Quarterly_Master'!H8-SUM('01_Quarterly_Master'!H5:H7))"],
  ["=('01_Quarterly_Master'!D15-SUM('01_Quarterly_Master'!D10:D14))", "=('01_Quarterly_Master'!E15-SUM('01_Quarterly_Master'!E10:E14))", "=('01_Quarterly_Master'!F15-SUM('01_Quarterly_Master'!F10:F14))", "=('01_Quarterly_Master'!G15-SUM('01_Quarterly_Master'!G10:G14))", "=('01_Quarterly_Master'!H15-SUM('01_Quarterly_Master'!H10:H14))"],
  ["=('01_Quarterly_Master'!D22-SUM('01_Quarterly_Master'!D17:D21))", "=('01_Quarterly_Master'!E22-SUM('01_Quarterly_Master'!E17:E21))", "=('01_Quarterly_Master'!F22-SUM('01_Quarterly_Master'!F17:F21))", "=('01_Quarterly_Master'!G22-SUM('01_Quarterly_Master'!G17:G21))", "=('01_Quarterly_Master'!H22-SUM('01_Quarterly_Master'!H17:H21))"],
  ["=('01_Quarterly_Master'!D69+'01_Quarterly_Master'!D70+'01_Quarterly_Master'!D71-'01_Quarterly_Master'!D72-'01_Quarterly_Master'!D73)", "=('01_Quarterly_Master'!E69+'01_Quarterly_Master'!E70+'01_Quarterly_Master'!E71-'01_Quarterly_Master'!E72-'01_Quarterly_Master'!E73)", "=('01_Quarterly_Master'!F69+'01_Quarterly_Master'!F70+'01_Quarterly_Master'!F71-'01_Quarterly_Master'!F72-'01_Quarterly_Master'!F73)", "=('01_Quarterly_Master'!G69+'01_Quarterly_Master'!G70+'01_Quarterly_Master'!G71-'01_Quarterly_Master'!G72-'01_Quarterly_Master'!G73)", "=('01_Quarterly_Master'!H69+'01_Quarterly_Master'!H70+'01_Quarterly_Master'!H71-'01_Quarterly_Master'!H72-'01_Quarterly_Master'!H73)"],
  ["=('01_Quarterly_Master'!D74+'01_Quarterly_Master'!D75+'01_Quarterly_Master'!D76-'01_Quarterly_Master'!D77-'01_Quarterly_Master'!D78-'01_Quarterly_Master'!D79-'01_Quarterly_Master'!D80-'01_Quarterly_Master'!D81-'01_Quarterly_Master'!D82)", "=('01_Quarterly_Master'!E74+'01_Quarterly_Master'!E75+'01_Quarterly_Master'!E76-'01_Quarterly_Master'!E77-'01_Quarterly_Master'!E78-'01_Quarterly_Master'!E79-'01_Quarterly_Master'!E80-'01_Quarterly_Master'!E81-'01_Quarterly_Master'!E82)", "=('01_Quarterly_Master'!F74+'01_Quarterly_Master'!F75+'01_Quarterly_Master'!F76-'01_Quarterly_Master'!F77-'01_Quarterly_Master'!F78-'01_Quarterly_Master'!F79-'01_Quarterly_Master'!F80-'01_Quarterly_Master'!F81-'01_Quarterly_Master'!F82)", "=('01_Quarterly_Master'!G74+'01_Quarterly_Master'!G75+'01_Quarterly_Master'!G76-'01_Quarterly_Master'!G77-'01_Quarterly_Master'!G78-'01_Quarterly_Master'!G79-'01_Quarterly_Master'!G80-'01_Quarterly_Master'!G81-'01_Quarterly_Master'!G82)", "=('01_Quarterly_Master'!H74+'01_Quarterly_Master'!H75+'01_Quarterly_Master'!H76-'01_Quarterly_Master'!H77-'01_Quarterly_Master'!H78-'01_Quarterly_Master'!H79-'01_Quarterly_Master'!H80-'01_Quarterly_Master'!H81-'01_Quarterly_Master'!H82)"],
  ["=('01_Quarterly_Master'!D60+'01_Quarterly_Master'!D61+'01_Quarterly_Master'!D62+'01_Quarterly_Master'!D63-'01_Quarterly_Master'!D64)", "=('01_Quarterly_Master'!E60+'01_Quarterly_Master'!E61+'01_Quarterly_Master'!E62+'01_Quarterly_Master'!E63-'01_Quarterly_Master'!E64)", "=('01_Quarterly_Master'!F60+'01_Quarterly_Master'!F61+'01_Quarterly_Master'!F62+'01_Quarterly_Master'!F63-'01_Quarterly_Master'!F64)", "=('01_Quarterly_Master'!G60+'01_Quarterly_Master'!G61+'01_Quarterly_Master'!G62+'01_Quarterly_Master'!G63-'01_Quarterly_Master'!G64)", "=('01_Quarterly_Master'!H60+'01_Quarterly_Master'!H61+'01_Quarterly_Master'!H62+'01_Quarterly_Master'!H63-'01_Quarterly_Master'!H64)"],
  ["=('01_Quarterly_Master'!D64+'01_Quarterly_Master'!D65-'01_Quarterly_Master'!D66)", "=('01_Quarterly_Master'!E64+'01_Quarterly_Master'!E65-'01_Quarterly_Master'!E66)", "=('01_Quarterly_Master'!F64+'01_Quarterly_Master'!F65-'01_Quarterly_Master'!F66)", "=('01_Quarterly_Master'!G64+'01_Quarterly_Master'!G65-'01_Quarterly_Master'!G66)", "=('01_Quarterly_Master'!H64+'01_Quarterly_Master'!H65-'01_Quarterly_Master'!H66)"],
  ["=('01_Quarterly_Master'!D67+'01_Quarterly_Master'!D66-'01_Quarterly_Master'!D68)", "=('01_Quarterly_Master'!E67+'01_Quarterly_Master'!E66-'01_Quarterly_Master'!E68)", "=('01_Quarterly_Master'!F67+'01_Quarterly_Master'!F66-'01_Quarterly_Master'!F68)", "=('01_Quarterly_Master'!G67+'01_Quarterly_Master'!G66-'01_Quarterly_Master'!G68)", "=('01_Quarterly_Master'!H67+'01_Quarterly_Master'!H66-'01_Quarterly_Master'!H68)"],
];
checks.getRangeByIndexes(4, 1, formulas.length, 5).formulas = formulas;
checks.getRange("G5:G12").values = Array.from({ length: checkLabels.length }, () => [1]);
checks.getRange("H5").formulas = [["=IF(AND(ABS(B5)<=G5,ABS(C5)<=G5,ABS(D5)<=G5,ABS(E5)<=G5,ABS(F5)<=G5),\"OK\",\"Review\")"]];
checks.getRange("H5:H12").fillDown();
checks.getRange("B5:G12").format.numberFormat = '#,##0;[Red](#,##0);-';
checks.getRange("A4:I12").format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
checks.getRange("H5:H12").format = { font: { bold: true }, horizontalAlignment: "center" };
autosize(checks, [["A", 42], ["B", 14], ["C", 14], ["D", 14], ["E", 14], ["F", 14], ["G", 13], ["H", 14], ["I", 90]]);

const sources = workbook.worksheets.add("04_Sources_Audit");
sources.showGridLines = false;
setTitle(sources, "Sources Audit", "Source IDs and extraction notes for the quarterly-only data pack.");
const sourceHeaders = ["Source ID", "File", "Document Type", "Period", "Confidence", "Usage Notes"];
sources.getRange("A4:F4").values = [sourceHeaders];
styleHeader(sources.getRange("A4:F4"));
sources.getRangeByIndexes(4, 0, sourceIds.length, sourceHeaders.length).values = sourceIds;
sources.getRange("A10:F10").values = [["Extraction note", "Q4 FY26 rolling tables were rendered and visually checked before entry. Q1-Q3 PDFs are retained as cross-check sources for their own quarterly releases.", "", "", "", ""]];
sources.getRange("A10:F10").merge(true);
sources.getRange("A10").format = { fill: "#FFF2CC", wrapText: true, font: { bold: true } };
sources.getRange("A4:F8").format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
sources.getRange("A4:F8").format.wrapText = true;
autosize(sources, [["A", 14], ["B", 28], ["C", 36], ["D", 14], ["E", 28], ["F", 70]]);

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(masterDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const csvHeader = masterHeaders.join(",");
const csvRows = rows.map(row => row.map(v => {
  const s = String(v ?? "");
  return /[",\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
}).join(","));
await fs.writeFile(path.join(masterDir, "quarterly_master_data_fy26.csv"), [csvHeader, ...csvRows].join("\n"), "utf8");
await fs.writeFile(path.join(masterDir, "quarterly_source_ids.csv"), [sourceHeaders.join(","), ...sourceIds.map(r => r.map(v => `"${String(v).replaceAll('"', '""')}"`).join(","))].join("\n"), "utf8");

for (const sheetName of ["00_Cover", "01_Quarterly_Master", "02_Segment_KPIs", "03_Checks", "04_Sources_Audit"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(previewDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const inspect = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(inspect.ndjson);

const checkInspect = await workbook.inspect({
  kind: "table",
  range: "03_Checks!A4:I12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 9,
});
console.log(checkInspect.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
const outPath = path.join(outputDir, "Eternal_Quarterly_Data_FY26.xlsx");
await xlsx.save(outPath);
await xlsx.save(path.join(masterDir, "Eternal_Quarterly_Data_FY26.xlsx"));
console.log(outPath);
