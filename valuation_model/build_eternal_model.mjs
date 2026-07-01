import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
const projectDir = path.resolve("../Eternal-Valuation");
const outputDir = path.join(projectDir, "outputs", "v1.1_professional_structure");
const excelDir = path.join(projectDir, "excel");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(excelDir, { recursive: true });

const sheets = {};
const sheetNames = {
  "Cover": "01_Cover",
  "Dashboard": "02_Dashboard",
  "Raw Data": "03_Raw_Data",
  "Historical FS": "04_Historical_FS",
  "Assumptions": "05_Assumptions",
  "Operating Model": "06_Operating_Model",
  "Statements": "07_Statements",
  "DCF": "08_DCF",
  "Trading Comps": "09_Trading_Comps",
  "Precedents": "10_Precedents",
  "Football Field": "11_Football_Field",
  "Thesis": "12_Thesis",
  "Checks": "13_Checks",
  "Sources Audit": "14_Sources_Audit",
};
for (const [key, displayName] of Object.entries(sheetNames)) {
  sheets[key] = workbook.worksheets.add(displayName);
  sheets[key].showGridLines = false;
}

const navy = "#0B1F3A";
const steel = "#E7EDF6";
const paleBlue = "#EAF3FF";
const yellow = "#FFF2CC";
const green = "#E2F0D9";
const red = "#FCE4D6";
const grey = "#F4F6F8";
const black = "#000000";
const inputBlue = "#0000FF";
const linkGreen = "#008000";

const numFmt = "#,##0;[Red](#,##0);-";
const pctFmt = "0.0%;[Red](0.0%);-";
const multFmt = "0.0x;[Red](0.0x);-";
const priceFmt = '"₹"0.00;[Red]("₹"0.00);-';
const crFmt = '"₹"#,##0;[Red]("₹"#,##0);-';

function title(sheet, text, range = "A1:H1") {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "middle",
  };
}

function section(sheet, text, range) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "left",
  };
}

function header(range) {
  range.format = {
    fill: steel,
    font: { bold: true, color: black },
    borders: { preset: "inside", style: "thin", color: "#D9E2F3" },
  };
}

function inputs(range) {
  range.format = {
    fill: yellow,
    font: { color: inputBlue },
  };
}

function formulas(range) {
  range.format = {
    font: { color: black },
  };
}

function linked(range) {
  range.format = {
    font: { color: linkGreen },
  };
}

function total(range) {
  range.format = {
    fill: grey,
    font: { bold: true },
    borders: { preset: "doubleBottom", style: "thin", color: "#808080" },
  };
}

function setWidths(sheet, widths) {
  widths.forEach((w, i) => {
    sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = w;
  });
}

function freeze(sheet, rows = 1, cols = 1) {
  sheet.freezePanes.freezeRows(rows);
  if (cols) sheet.freezePanes.freezeColumns(cols);
}

async function addComment(cell, text) {
  try {
    const thread = workbook.comments.addThread({ cell }, text);
    thread.resolve();
  } catch {
    // Comments are helpful audit metadata but not required for workbook calculation.
  }
}

try {
  workbook.comments.setSelf({ displayName: "User" });
} catch {}

// Cover
{
  const s = sheets.Cover;
  title(s, "Eternal Limited Equity Valuation Model", "A1:I1");
  s.getRange("A3:B12").values = [
    ["Company", "Eternal Limited (formerly Zomato)"],
    ["Project", "Full equity valuation model: 3-statement shell + DCF + comps + precedents"],
    ["Primary use", "JPMC / investment banking interview project"],
    ["Currency / units", "INR crores unless stated otherwise"],
    ["Historical period", "FY22A-FY26A input-ready"],
    ["Forecast period", "FY27E-FY31E"],
    ["Valuation date", "2026-07-01"],
    ["Current phase", "Phase 1.1: professional structure upgrade"],
    ["Chosen case", "Eternal was selected over a bank because the target artifact requires DCF, EV multiples, and a football field."],
    ["Next milestone", "Populate Raw_Data with FY22-FY26 actuals from annual reports/results and lock source comments."],
  ];
  s.getRange("A3:A12").format = { font: { bold: true }, fill: steel };
  s.getRange("B3:B12").format = { wrapText: true };
  section(s, "Key Outputs", "D3:I3");
  s.getRange("D4:I10").values = [
    ["Metric", "Value", "Source", "", "Metric", "Value"],
    ["Model status", null, "Checks", "", "DCF target price", null],
    ["DCF value / share", null, "DCF", "", "Comps value / share", null],
    ["Precedents value / share", null, "Precedents", "", "Football field low", null],
    ["Football field high", null, "Football Field", "", "Recommendation", null],
    ["WACC", null, "Assumptions", "", "Terminal growth", null],
    ["Exit multiple", null, "Assumptions", "", "Scenario", "Base case"],
  ];
  s.getRange("E5").formulas = [["='13_Checks'!F14"]];
  s.getRange("E6").formulas = [["='08_DCF'!D29"]];
  s.getRange("I6").formulas = [["='09_Trading_Comps'!H18"]];
  s.getRange("E7").formulas = [["=AVERAGE('10_Precedents'!G14:G15)"]];
  s.getRange("I7").formulas = [["='11_Football_Field'!B8"]];
  s.getRange("E8").formulas = [["='11_Football_Field'!D8"]];
  s.getRange("I8").formulas = [["='12_Thesis'!B10"]];
  s.getRange("E9").formulas = [["='05_Assumptions'!B31"]];
  s.getRange("I9").formulas = [["='05_Assumptions'!B35"]];
  s.getRange("E10").formulas = [["='05_Assumptions'!B36"]];
  header(s.getRange("D4:I4"));
  s.getRange("E6:I9").format.numberFormat = priceFmt;
  s.getRange("E9:I9").format.numberFormat = pctFmt;
  s.getRange("E10").format.numberFormat = multFmt;
  s.getRange("A15:I18").values = [
    ["Color convention", "Blue text / yellow fill = editable assumption; black = formula; green = linked worksheet reference", "", "", "", "", "", "", ""],
    ["Implementation note", "This workbook is intentionally source-first. It computes cleanly now, and it is ready for filing-backed values to be entered without changing formulas.", "", "", "", "", "", "", ""],
    ["Interview angle", "Eternal gives a defensible discussion around food delivery, quick commerce, margin expansion, competition, and whether a sum-of-the-parts frame is warranted.", "", "", "", "", "", "", ""],
    ["Known limitation", "Historical statements are blank input ranges in Phase 1. Do not use outputs as an investment recommendation until actuals and market data are tied.", "", "", "", "", "", "", ""],
  ];
  s.getRange("A15:A18").format = { font: { bold: true }, fill: steel };
  s.getRange("B15:I18").format = { wrapText: true };
  setWidths(s, [20, 42, 4, 22, 18, 18, 4, 22, 18]);
  freeze(s, 1, 0);
}

// Dashboard
{
  const s = sheets.Dashboard;
  title(s, "Eternal Valuation Dashboard", "A1:J1");
  s.getRange("A3:J3").values = [["Metric", "Value", "Unit", "Source", "", "Metric", "Value", "Unit", "Source", "Status"]];
  header(s.getRange("A3:J3"));
  s.getRange("A4:J12").values = [
    ["Current share price", null, "INR / share", "05_Assumptions", "", "Target price midpoint", null, "INR / share", "11_Football_Field", null],
    ["Target price low", null, "INR / share", "11_Football_Field", "", "Target price high", null, "INR / share", "11_Football_Field", null],
    ["Upside / downside", null, "%", "Dashboard", "", "Recommendation", null, "Buy / Hold / Sell", "12_Thesis", null],
    ["Enterprise value", null, "INR cr", "08_DCF", "", "Equity value", null, "INR cr", "08_DCF", null],
    ["WACC", null, "%", "05_Assumptions", "", "Terminal growth", null, "%", "05_Assumptions", null],
    ["Exit EBITDA multiple", null, "x", "05_Assumptions", "", "FY27E-FY31E Revenue CAGR", null, "%", "06_Operating_Model", null],
    ["FY31E EBITDA margin", null, "%", "06_Operating_Model", "", "Model status", null, "Status", "13_Checks", null],
    ["Historical data status", null, "Status", "03_Raw_Data", "", "Comps data status", null, "Status", "09_Trading_Comps", null],
    ["Precedents data status", null, "Status", "10_Precedents", "", "Version", "v1.1", "Model version", "01_Cover", "Professional structure"],
  ];
  s.getRange("B4:B12").formulas = [
    ["='05_Assumptions'!B34"],
    ["='11_Football_Field'!B8"],
    ["=IFERROR(G4/B4-1,0)"],
    ["='08_DCF'!D23"],
    ["='05_Assumptions'!B31"],
    ["='05_Assumptions'!B36"],
    ["='06_Operating_Model'!G6"],
    ["=IF(COUNT('03_Raw_Data'!B4:F15)>=12,\"Entered\",\"Needs input\")"],
    ["=IF(OR('10_Precedents'!H9>0,'10_Precedents'!I9>0),\"Entered\",\"Optional / needs input\")"],
  ];
  s.getRange("G4:G12").formulas = [
    ["='11_Football_Field'!C8"],
    ["='11_Football_Field'!D8"],
    ["='12_Thesis'!B10"],
    ["='08_DCF'!D25"],
    ["='05_Assumptions'!B35"],
    ["=IFERROR(('06_Operating_Model'!G4/'06_Operating_Model'!C4)^(1/4)-1,0)"],
    ["='13_Checks'!F14"],
    ["=IF('09_Trading_Comps'!H11>0,\"Entered\",\"Needs input\")"],
    ["=\"v1.1\""],
  ];
  s.getRange("J4:J12").formulas = [
    ["=IF(B4>0,\"Ready\",\"Needs market price\")"],
    ["=IF(G5>0,\"Ready\",\"Needs valuation data\")"],
    ["=IF(B4>0,\"Ready\",\"Needs market price\")"],
    ["=IF(B7>0,\"Ready\",\"Needs model data\")"],
    ["=IF(B8>G8,\"Ready\",\"Review\")"],
    ["=IF(G9<>0,\"Ready\",\"Needs historical revenue\")"],
    ["=IF(G10=\"OK\",\"Ready\",G10)"],
    ["=IF(B11=\"Entered\",\"Ready\",\"Needs filings\")"],
    ["=\"Ready\""],
  ];
  s.getRange("B4:B5").format.numberFormat = priceFmt;
  s.getRange("G4:G5").format.numberFormat = priceFmt;
  s.getRange("B6").format.numberFormat = pctFmt;
  s.getRange("B7:G7").format.numberFormat = crFmt;
  s.getRange("B8:G9").format.numberFormat = pctFmt;
  s.getRange("B9").format.numberFormat = multFmt;
  s.getRange("A15:J15").values = [["Phase 2 Data Entry Checklist", "Status", "Source", "Model destination", "Notes", "", "Output Readiness", "Status", "Owner", "Next step"]];
  header(s.getRange("A15:J15"));
  s.getRange("A16:J23").values = [
    ["Annual reports downloaded", "Not started", "Eternal IR / exchanges", "data/annual_reports", "FY22-FY25 annual reports", "", "Historical FS", "Waiting", "User/Codex", "Enter raw values"],
    ["Latest FY26 result downloaded", "Not started", "Eternal IR / exchanges", "data/quarterly_results", "Use latest audited/unaudited source", "", "DCF", "Waiting", "User/Codex", "Populate historicals"],
    ["Revenue and EBITDA actuals entered", "Not started", "03_Raw_Data", "04_Historical_FS", "Use INR cr", "", "Dashboard", "Waiting", "User/Codex", "Fill core actuals"],
    ["Cash, debt, capex, NWC entered", "Not started", "03_Raw_Data", "04_Historical_FS", "Needed for FCF bridge", "", "Comps", "Waiting", "User/Codex", "Collect market data"],
    ["WACC inputs sourced", "Not started", "RBI / Damodaran / market data", "05_Assumptions", "Risk-free rate, beta, ERP", "", "Football field", "Waiting", "User/Codex", "Finish valuation range"],
    ["Peer set populated", "Not started", "Filings / market data", "09_Trading_Comps", "Use same data date", "", "Report", "Waiting", "User/Codex", "Draft report"],
    ["Precedent set assessed", "Not started", "Deal releases / filings", "10_Precedents", "Use only if credible", "", "Pitch deck", "Waiting", "User/Codex", "Create slides"],
    ["Investment decision written", "Not started", "Model output", "12_Thesis", "Buy/Hold/Sell plus risks/catalysts", "", "GitHub package", "Waiting", "User/Codex", "README/screenshots"],
  ];
  s.getRange("A16:J23").format.wrapText = false;
  s.getRange("A16:J23").format.rowHeight = 24;
  s.getRange("B16:B23").format = { fill: yellow, font: { color: inputBlue } };
  setWidths(s, [28, 15, 25, 24, 30, 28, 15, 18, 22, 30]);
  freeze(s, 3, 0);
}

// Sources Audit
{
  const s = sheets["Sources Audit"];
  title(s, "Sources & Audit Trail", "A1:I1");
  s.getRange("A3:I3").values = [["Source ID", "Item", "Period / as-of", "Value", "Units", "Source name", "Plain URL", "Model use", "Status / notes"]];
  header(s.getRange("A3:I3"));
  s.getRange("A4:I16").values = [
    ["S1", "Investor relations landing page", "Current", null, null, "Eternal Investor Relations", "https://www.eternal.com/investor-relations", "Primary source hub for results, annual reports, governance", "Use official downloads when populating actuals."],
    ["S2", "FY26 / latest quarterly shareholder letter", "FY26 / latest available", null, "INR cr", "Eternal results page", "https://www.eternal.com/investor-relations/results/", "Latest actuals and segment commentary", "Pull official PDF before locking FY26 actuals."],
    ["S3", "Annual reports", "FY22-FY25", null, "INR cr", "Eternal annual reports page", "https://www.eternal.com/investor-relations/annual-reports/", "Historical 3-statement source", "Download annual reports and tie each line item."],
    ["S4", "NSE company filings", "Current", null, null, "NSE filings", "https://www.nseindia.com/companies-listing/corporate-filings-announcements", "Exchange-filed results and announcements", "Use if official page link rendering is inconvenient."],
    ["S5", "BSE company filings", "Current", null, null, "BSE filings", "https://www.bseindia.com/corporates/ann.html", "Exchange-filed results and announcements", "Cross-check source documents."],
    ["S6", "Risk-free rate", "Valuation date", null, "%", "Government bond yield source", "https://www.rbi.org.in/", "CAPM / WACC", "Enter current 10-year India G-sec yield."],
    ["S7", "Equity risk premium", "Valuation date", null, "%", "Damodaran ERP data", "https://pages.stern.nyu.edu/~adamodar/", "CAPM / WACC", "Use India mature/emerging ERP assumption and cite exact table."],
    ["S8", "Beta", "Valuation date", null, "x", "Market data provider", "https://finance.yahoo.com/quote/ETERNAL.NS", "CAPM / WACC", "Record raw beta and whether levered/unlevered."],
    ["S9", "Share count / market price", "Valuation date", null, "shares / INR", "Exchange or market data provider", "https://www.nseindia.com/get-quotes/equity?symbol=ETERNAL", "Equity value bridge and target price", "Lock as-of date."],
    ["S10", "Peer multiples", "Latest trading", null, "x", "Company filings / market data", "Multiple source URLs", "Trading comps", "Prefer primary filings plus consistent market-data date."],
    ["S11", "Transaction multiples", "Deal announcement date", null, "x", "Deal press releases / filings", "Multiple source URLs", "Precedents", "Use only if deal values and EBITDA/revenue are auditable."],
    ["S12", "Rename / company identity", "2025 onward", null, null, "Company announcement", "https://www.eternal.com/", "Naming and overview", "Use Eternal legal name; note former Zomato brand."],
    ["S13", "Business segments", "Latest annual / quarterly", null, null, "Eternal filings", "https://www.eternal.com/investor-relations/results/", "Food delivery, quick commerce, Hyperpure, going-out segment drivers", "Use segment disclosures where available."],
  ];
  s.getRange("A3:I16").format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
  s.getRange("G4:G16").format = { wrapText: true, font: { color: "#0563C1" } };
  s.getRange("I4:I16").format = { wrapText: true };
  setWidths(s, [11, 28, 18, 12, 12, 26, 54, 34, 42]);
  freeze(s, 3, 1);
}

// Assumptions
{
  const s = sheets.Assumptions;
  title(s, "Assumptions", "A1:H1");
  section(s, "Operating Forecast Drivers", "A3:H3");
  s.getRange("A4:H4").values = [["Driver", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E", "Source / basis", "Notes"]];
  header(s.getRange("A4:H4"));
  s.getRange("A5:H14").values = [
    ["Revenue growth", 0.22, 0.20, 0.18, 0.16, 0.14, "User input", "Initial placeholder ramp; replace with segment build from food delivery + Blinkit + Hyperpure."],
    ["EBITDA margin", 0.04, 0.06, 0.08, 0.10, 0.12, "User input", "Margin expansion should be justified by contribution margin and scale economics."],
    ["D&A / revenue", 0.025, 0.024, 0.023, 0.022, 0.021, "User input", "Tie to PPE/intangibles after full statement build."],
    ["Capex / revenue", 0.035, 0.034, 0.033, 0.032, 0.031, "User input", "Include technology, dark-store, and logistics infrastructure where reported."],
    ["NWC / revenue", 0.02, 0.02, 0.02, 0.02, 0.02, "User input", "Use a working capital schedule once historical actuals are entered."],
    ["Cash tax rate", 0.25, 0.25, 0.25, 0.25, 0.25, "User input", "Use normalized effective tax rate, not early-profitability noise."],
    ["Share count growth", 0.01, 0.01, 0.01, 0.01, 0.01, "User input", "Tie to diluted shares and ESOP dilution."],
    ["Revenue sanity low", 0.05, 0.05, 0.05, 0.05, 0.05, "Check threshold", "Growth below this warns but does not fail the model."],
    ["Revenue sanity high", 0.40, 0.40, 0.40, 0.40, 0.40, "Check threshold", "Growth above this requires explicit explanation."],
    ["Target recommendation", "Hold", null, null, null, null, "User input", "Buy / Hold / Sell after valuation is source-backed."],
  ];
  inputs(s.getRange("B5:F14"));
  inputs(s.getRange("B14"));
  s.getRange("B5:F13").format.numberFormat = pctFmt;
  s.getRange("A17:B18").values = [
    ["Valuation date", new Date("2026-07-01")],
    ["Mid-year convention", "Yes"],
  ];
  inputs(s.getRange("B17:B18"));
  s.getRange("B17").format.numberFormat = "yyyy-mm-dd";
  section(s, "WACC / Terminal Value", "A20:D20");
  s.getRange("A21:B36").values = [
    ["Risk-free rate", 0.069],
    ["Levered beta", 1.2],
    ["Equity risk premium", 0.055],
    ["Size / company-specific premium", 0.0],
    ["Pre-tax cost of debt", 0.085],
    ["Tax rate", 0.25],
    ["Debt / total capital", 0.05],
    ["Equity / total capital", null],
    ["Cost of equity", null],
    ["After-tax cost of debt", null],
    ["WACC", null],
    ["Net debt / (cash)", 0],
    ["Diluted shares outstanding (mm)", 8900],
    ["Current share price", 0],
    ["Terminal growth", 0.04],
    ["Exit EBITDA multiple", 25.0],
  ];
  s.getRange("B28").formulas = [["=1-B27"]];
  s.getRange("B29").formulas = [["=B21+B22*B23+B24"]];
  s.getRange("B30").formulas = [["=B25*(1-B26)"]];
  s.getRange("B31").formulas = [["=B29*B28+B30*B27"]];
  inputs(s.getRange("B21:B27"));
  inputs(s.getRange("B32:B36"));
  formulas(s.getRange("B28:B31"));
  s.getRange("B21:B31").format.numberFormat = pctFmt;
  s.getRange("B32").format.numberFormat = crFmt;
  s.getRange("B33").format.numberFormat = numFmt;
  s.getRange("B34").format.numberFormat = priceFmt;
  s.getRange("B35").format.numberFormat = pctFmt;
  s.getRange("B36").format.numberFormat = multFmt;
  s.getRange("D21:H27").values = [
    ["Sensitivity Inputs", "Low", "Mid", "High", "Notes"],
    ["WACC range", 0.10, 0.12, 0.14, "Used in DCF sensitivity table"],
    ["Terminal growth range", 0.03, 0.04, 0.05, "Must stay below WACC"],
    ["Exit multiple range", 20.0, 25.0, 30.0, "Cross-check against peers"],
    ["Comps low haircut", 0.15, null, null, "Low case = median less haircut"],
    ["Comps high premium", 0.15, null, null, "High case = median plus premium"],
    ["Precedent premium", 0.20, null, null, "Typical control premium placeholder"],
  ];
  header(s.getRange("D21:H21"));
  inputs(s.getRange("E22:G24"));
  inputs(s.getRange("E25:E27"));
  s.getRange("E22:G23").format.numberFormat = pctFmt;
  s.getRange("E24:G24").format.numberFormat = multFmt;
  s.getRange("E25:E27").format.numberFormat = pctFmt;
  setWidths(s, [26, 16, 4, 24, 14, 14, 14, 46]);
  freeze(s, 4, 1);
}

// Raw Data
{
  const s = sheets["Raw Data"];
  title(s, "Raw Data Entry - Historical Actuals", "A1:J1");
  s.getRange("A3:J3").values = [["Line item", "FY22A", "FY23A", "FY24A", "FY25A", "FY26A", "Units", "Source ID", "Source detail / page", "Entry notes"]];
  header(s.getRange("A3:J3"));
  s.getRange("A4:J15").values = [
    ["Revenue", null, null, null, null, null, "INR cr", "S2/S3", "Annual report / latest result", "Revenue from operations or comparable consolidated revenue line."],
    ["COGS / operating costs", null, null, null, null, null, "INR cr", "S2/S3", "Annual report / latest result", "Enter as positive expense."],
    ["SG&A / other operating expenses", null, null, null, null, null, "INR cr", "S2/S3", "Annual report / latest result", "Enter as positive expense."],
    ["D&A", null, null, null, null, null, "INR cr", "S2/S3", "Cash flow / notes", "Depreciation and amortization expense."],
    ["Cash taxes", null, null, null, null, null, "INR cr", "S2/S3", "Cash flow / tax note", "Use cash taxes where available; otherwise normalized tax proxy."],
    ["Net income", null, null, null, null, null, "INR cr", "S2/S3", "P&L", "Profit / loss attributable to shareholders if available."],
    ["Cash and equivalents", null, null, null, null, null, "INR cr", "S2/S3", "Balance sheet", "Cash, cash equivalents, and liquid investments if treated as cash."],
    ["Debt", null, null, null, null, null, "INR cr", "S2/S3", "Balance sheet", "Borrowings and lease debt if included in net debt bridge."],
    ["Current assets ex-cash", null, null, null, null, null, "INR cr", "S2/S3", "Balance sheet", "Exclude cash and cash equivalents."],
    ["Current liabilities ex-debt", null, null, null, null, null, "INR cr", "S2/S3", "Balance sheet", "Exclude debt / borrowings where separately modeled."],
    ["Capex", null, null, null, null, null, "INR cr", "S2/S3", "Cash flow", "Purchase of PPE/intangibles; enter cash outflow as positive."],
    ["Diluted shares", null, null, null, null, null, "mm", "S9", "Filing / exchange data", "Diluted weighted average or latest diluted share count."],
  ];
  inputs(s.getRange("B4:F15"));
  inputs(s.getRange("H4:J15"));
  s.getRange("B4:F14").format.numberFormat = crFmt;
  s.getRange("B15:F15").format.numberFormat = numFmt;
  s.getRange("I4:J15").format.wrapText = true;
  s.getRange("A18:J18").values = [["Raw Data Rules", "", "", "", "", "", "", "", "", ""]];
  section(s, "Raw Data Rules", "A18:J18");
  s.getRange("A19:J23").values = [
    ["1", "Do not calculate on this sheet. Enter only source-backed values and source details.", "", "", "", "", "", "", "", ""],
    ["2", "Use INR crores consistently. If a source reports INR millions, convert before entry and note it.", "", "", "", "", "", "", "", ""],
    ["3", "Expenses should be positive. Historical FS handles signs and calculations downstream.", "", "", "", "", "", "", "", ""],
    ["4", "Every entered value needs a source ID. Use Sources Audit for URLs and filing notes.", "", "", "", "", "", "", "", ""],
    ["5", "When unsure, leave blank. A blank audited source is better than a guessed value.", "", "", "", "", "", "", "", ""],
  ];
  s.getRange("B19:J23").merge(true);
  s.getRange("B19:J23").format.wrapText = true;
  setWidths(s, [30, 14, 14, 14, 14, 14, 12, 12, 32, 44]);
  freeze(s, 3, 1);
}

// Historical FS
{
  const s = sheets["Historical FS"];
  title(s, "Historical Financial Statements Input", "A1:H1");
  s.getRange("A3:H3").values = [["Line item", "FY22A", "FY23A", "FY24A", "FY25A", "FY26A", "Formula / check", "Source ID"]];
  header(s.getRange("A3:H3"));
  const rows = [
    ["Revenue", null, null, null, null, null, "Input from annual report / results", "S2/S3"],
    ["Revenue growth", null, null, null, null, null, "'=Revenue / prior year - 1", null],
    ["COGS / operating costs", null, null, null, null, null, "Input; expenses shown as positive values", "S2/S3"],
    ["Gross profit", null, null, null, null, null, "'=Revenue - COGS", null],
    ["Gross margin", null, null, null, null, null, "'=Gross profit / Revenue", null],
    ["SG&A / other operating expenses", null, null, null, null, null, "Input; expenses shown as positive values", "S2/S3"],
    ["EBITDA", null, null, null, null, null, "'=Gross profit - SG&A", null],
    ["EBITDA margin", null, null, null, null, null, "'=EBITDA / Revenue", null],
    ["D&A", null, null, null, null, null, "Input; expense shown as positive value", "S2/S3"],
    ["EBIT", null, null, null, null, null, "'=EBITDA - D&A", null],
    ["Cash taxes", null, null, null, null, null, "Input or normalized tax estimate", "S2/S3"],
    ["Net income", null, null, null, null, null, "Input / tie to P&L", "S2/S3"],
    ["Cash and equivalents", null, null, null, null, null, "Input from balance sheet", "S2/S3"],
    ["Debt", null, null, null, null, null, "Input from balance sheet", "S2/S3"],
    ["Net debt / (cash)", null, null, null, null, null, "'=Debt - Cash", null],
    ["Current assets ex-cash", null, null, null, null, null, "Input for NWC", "S2/S3"],
    ["Current liabilities ex-debt", null, null, null, null, null, "Input for NWC", "S2/S3"],
    ["Net working capital", null, null, null, null, null, "'=Current assets ex-cash - current liabilities ex-debt", null],
    ["Change in NWC", null, null, null, null, null, "'=NWC - prior year NWC", null],
    ["Capex", null, null, null, null, null, "Input from cash flow statement; cash outflow as positive", "S2/S3"],
    ["Diluted shares (mm)", null, null, null, null, null, "Input from filings / exchange", "S9"],
  ];
  s.getRange(`A4:H${3 + rows.length}`).values = rows;
  linked(s.getRange("B4:F24"));
  s.getRange("B4:F4").formulas = [["='03_Raw_Data'!B4", "='03_Raw_Data'!C4", "='03_Raw_Data'!D4", "='03_Raw_Data'!E4", "='03_Raw_Data'!F4"]];
  s.getRange("B6:F6").formulas = [["='03_Raw_Data'!B5", "='03_Raw_Data'!C5", "='03_Raw_Data'!D5", "='03_Raw_Data'!E5", "='03_Raw_Data'!F5"]];
  s.getRange("B9:F9").formulas = [["='03_Raw_Data'!B6", "='03_Raw_Data'!C6", "='03_Raw_Data'!D6", "='03_Raw_Data'!E6", "='03_Raw_Data'!F6"]];
  s.getRange("B12:F12").formulas = [["='03_Raw_Data'!B7", "='03_Raw_Data'!C7", "='03_Raw_Data'!D7", "='03_Raw_Data'!E7", "='03_Raw_Data'!F7"]];
  s.getRange("B14:F14").formulas = [["='03_Raw_Data'!B8", "='03_Raw_Data'!C8", "='03_Raw_Data'!D8", "='03_Raw_Data'!E8", "='03_Raw_Data'!F8"]];
  s.getRange("B15:F15").formulas = [["='03_Raw_Data'!B9", "='03_Raw_Data'!C9", "='03_Raw_Data'!D9", "='03_Raw_Data'!E9", "='03_Raw_Data'!F9"]];
  s.getRange("B16:F16").formulas = [["='03_Raw_Data'!B10", "='03_Raw_Data'!C10", "='03_Raw_Data'!D10", "='03_Raw_Data'!E10", "='03_Raw_Data'!F10"]];
  s.getRange("B17:F17").formulas = [["='03_Raw_Data'!B11", "='03_Raw_Data'!C11", "='03_Raw_Data'!D11", "='03_Raw_Data'!E11", "='03_Raw_Data'!F11"]];
  s.getRange("B19:F19").formulas = [["='03_Raw_Data'!B12", "='03_Raw_Data'!C12", "='03_Raw_Data'!D12", "='03_Raw_Data'!E12", "='03_Raw_Data'!F12"]];
  s.getRange("B20:F20").formulas = [["='03_Raw_Data'!B13", "='03_Raw_Data'!C13", "='03_Raw_Data'!D13", "='03_Raw_Data'!E13", "='03_Raw_Data'!F13"]];
  s.getRange("B23:F23").formulas = [["='03_Raw_Data'!B14", "='03_Raw_Data'!C14", "='03_Raw_Data'!D14", "='03_Raw_Data'!E14", "='03_Raw_Data'!F14"]];
  s.getRange("B24:F24").formulas = [["='03_Raw_Data'!B15", "='03_Raw_Data'!C15", "='03_Raw_Data'!D15", "='03_Raw_Data'!E15", "='03_Raw_Data'!F15"]];
  s.getRange("B5").values = [[null]];
  s.getRange("C5:F5").formulas = [["=IFERROR(C4/B4-1,\"\")", "=IFERROR(D4/C4-1,\"\")", "=IFERROR(E4/D4-1,\"\")", "=IFERROR(F4/E4-1,\"\")"]];
  s.getRange("B7:F7").formulas = [["=IFERROR(B4-B6,\"\")", "=IFERROR(C4-C6,\"\")", "=IFERROR(D4-D6,\"\")", "=IFERROR(E4-E6,\"\")", "=IFERROR(F4-F6,\"\")"]];
  s.getRange("B8:F8").formulas = [["=IFERROR(B7/B4,\"\")", "=IFERROR(C7/C4,\"\")", "=IFERROR(D7/D4,\"\")", "=IFERROR(E7/E4,\"\")", "=IFERROR(F7/F4,\"\")"]];
  s.getRange("B10:F10").formulas = [["=IFERROR(B7-B9,\"\")", "=IFERROR(C7-C9,\"\")", "=IFERROR(D7-D9,\"\")", "=IFERROR(E7-E9,\"\")", "=IFERROR(F7-F9,\"\")"]];
  s.getRange("B11:F11").formulas = [["=IFERROR(B10/B4,\"\")", "=IFERROR(C10/C4,\"\")", "=IFERROR(D10/D4,\"\")", "=IFERROR(E10/E4,\"\")", "=IFERROR(F10/F4,\"\")"]];
  s.getRange("B13:F13").formulas = [["=IFERROR(B10-B12,\"\")", "=IFERROR(C10-C12,\"\")", "=IFERROR(D10-D12,\"\")", "=IFERROR(E10-E12,\"\")", "=IFERROR(F10-F12,\"\")"]];
  s.getRange("B18:F18").formulas = [["=IFERROR(B17-B16,\"\")", "=IFERROR(C17-C16,\"\")", "=IFERROR(D17-D16,\"\")", "=IFERROR(E17-E16,\"\")", "=IFERROR(F17-F16,\"\")"]];
  s.getRange("B21:F21").formulas = [["=IFERROR(B19-B20,\"\")", "=IFERROR(C19-C20,\"\")", "=IFERROR(D19-D20,\"\")", "=IFERROR(E19-E20,\"\")", "=IFERROR(F19-F20,\"\")"]];
  s.getRange("B22").values = [[null]];
  s.getRange("C22:F22").formulas = [["=IFERROR(C21-B21,\"\")", "=IFERROR(D21-C21,\"\")", "=IFERROR(E21-D21,\"\")", "=IFERROR(F21-E21,\"\")"]];
  s.getRange("B4:F24").format.numberFormat = numFmt;
  s.getRange("B5:F5").format.numberFormat = pctFmt;
  s.getRange("B8:F8").format.numberFormat = pctFmt;
  s.getRange("B11:F11").format.numberFormat = pctFmt;
  formulas(s.getRange("B5:F5"));
  formulas(s.getRange("B7:F8"));
  formulas(s.getRange("B10:F11"));
  formulas(s.getRange("B13:F13"));
  formulas(s.getRange("B18:F18"));
  formulas(s.getRange("B21:F22"));
  total(s.getRange("A7:F8"));
  total(s.getRange("A10:F11"));
  total(s.getRange("A13:F13"));
  total(s.getRange("A18:F22"));
  linked(s.getRange("H4:H24"));
  setWidths(s, [30, 14, 14, 14, 14, 14, 34, 12]);
  freeze(s, 3, 1);
}

// Operating Model
{
  const s = sheets["Operating Model"];
  title(s, "Operating Forecast Model", "A1:G1");
  s.getRange("A3:G3").values = [["Line item", "FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]];
  header(s.getRange("A3:G3"));
  const rows = [
    ["Revenue", "='04_Historical_FS'!F4", "=B4*(1+'05_Assumptions'!B5)", "=C4*(1+'05_Assumptions'!C5)", "=D4*(1+'05_Assumptions'!D5)", "=E4*(1+'05_Assumptions'!E5)", "=F4*(1+'05_Assumptions'!F5)"],
    ["Revenue growth", "='04_Historical_FS'!F5", "=IFERROR(C4/B4-1,\"\")", "=IFERROR(D4/C4-1,\"\")", "=IFERROR(E4/D4-1,\"\")", "=IFERROR(F4/E4-1,\"\")", "=IFERROR(G4/F4-1,\"\")"],
    ["EBITDA margin", "='04_Historical_FS'!F11", "='05_Assumptions'!B6", "='05_Assumptions'!C6", "='05_Assumptions'!D6", "='05_Assumptions'!E6", "='05_Assumptions'!F6"],
    ["EBITDA", "='04_Historical_FS'!F10", "=C4*C6", "=D4*D6", "=E4*E6", "=F4*F6", "=G4*G6"],
    ["D&A / revenue", "=IFERROR('04_Historical_FS'!F12/'04_Historical_FS'!F4,\"\")", "='05_Assumptions'!B7", "='05_Assumptions'!C7", "='05_Assumptions'!D7", "='05_Assumptions'!E7", "='05_Assumptions'!F7"],
    ["D&A", "='04_Historical_FS'!F12", "=C4*C8", "=D4*D8", "=E4*E8", "=F4*F8", "=G4*G8"],
    ["EBIT", "='04_Historical_FS'!F13", "=C7-C9", "=D7-D9", "=E7-E9", "=F7-F9", "=G7-G9"],
    ["Cash tax rate", "=IFERROR('04_Historical_FS'!F14/'04_Historical_FS'!F13,\"\")", "='05_Assumptions'!B10", "='05_Assumptions'!C10", "='05_Assumptions'!D10", "='05_Assumptions'!E10", "='05_Assumptions'!F10"],
    ["Cash taxes", "='04_Historical_FS'!F14", "=MAX(0,C10*C11)", "=MAX(0,D10*D11)", "=MAX(0,E10*E11)", "=MAX(0,F10*F11)", "=MAX(0,G10*G11)"],
    ["NOPAT", "=IFERROR(B10-B12,\"\")", "=C10-C12", "=D10-D12", "=E10-E12", "=F10-F12", "=G10-G12"],
    ["Capex / revenue", "=IFERROR('04_Historical_FS'!F23/'04_Historical_FS'!F4,\"\")", "='05_Assumptions'!B8", "='05_Assumptions'!C8", "='05_Assumptions'!D8", "='05_Assumptions'!E8", "='05_Assumptions'!F8"],
    ["Capex", "='04_Historical_FS'!F23", "=C4*C14", "=D4*D14", "=E4*E14", "=F4*F14", "=G4*G14"],
    ["NWC / revenue", "=IFERROR('04_Historical_FS'!F20/'04_Historical_FS'!F4,\"\")", "='05_Assumptions'!B9", "='05_Assumptions'!C9", "='05_Assumptions'!D9", "='05_Assumptions'!E9", "='05_Assumptions'!F9"],
    ["NWC", "='04_Historical_FS'!F21", "=C4*C16", "=D4*D16", "=E4*E16", "=F4*F16", "=G4*G16"],
    ["Change in NWC", "='04_Historical_FS'!F22", "=C17-B17", "=D17-C17", "=E17-D17", "=F17-E17", "=G17-F17"],
    ["Unlevered FCF", "=B13+B9-B15-B18", "=C13+C9-C15-C18", "=D13+D9-D15-D18", "=E13+E9-E15-E18", "=F13+F9-F15-F18", "=G13+G9-G15-G18"],
    ["Diluted shares (mm)", "='04_Historical_FS'!F24", "=IF(B20>0,B20*(1+'05_Assumptions'!B11),'05_Assumptions'!B33)", "=C20*(1+'05_Assumptions'!C11)", "=D20*(1+'05_Assumptions'!D11)", "=E20*(1+'05_Assumptions'!E11)", "=F20*(1+'05_Assumptions'!F11)"],
  ];
  s.getRange(`A4:G${3 + rows.length}`).values = rows.map(r => [r[0], null, null, null, null, null, null]);
  s.getRange(`B4:G${3 + rows.length}`).formulas = rows.map(r => r.slice(1));
  s.getRange("B4:G20").format.numberFormat = numFmt;
  for (const row of [5, 6, 8, 11, 14, 16]) s.getRange(`B${row}:G${row}`).format.numberFormat = pctFmt;
  for (const row of [7, 10, 13, 15, 17, 19]) total(s.getRange(`A${row}:G${row}`));
  linked(s.getRange("B4:B20"));
  formulas(s.getRange("C4:G20"));
  setWidths(s, [28, 14, 14, 14, 14, 14, 14]);
  freeze(s, 3, 1);
}

// Statements
{
  const s = sheets.Statements;
  title(s, "Simplified Linked Statements", "A1:G1");
  section(s, "Income Statement", "A3:G3");
  s.getRange("A4:G4").values = [["Line item", "FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]];
  header(s.getRange("A4:G4"));
  s.getRange("A5:G11").values = [
    ["Revenue", null, null, null, null, null, null],
    ["EBITDA", null, null, null, null, null, null],
    ["D&A", null, null, null, null, null, null],
    ["EBIT", null, null, null, null, null, null],
    ["Cash taxes", null, null, null, null, null, null],
    ["NOPAT", null, null, null, null, null, null],
    ["Net income proxy", null, null, null, null, null, null],
  ];
  s.getRange("B5:G11").formulas = [
    ["='06_Operating_Model'!B4", "='06_Operating_Model'!C4", "='06_Operating_Model'!D4", "='06_Operating_Model'!E4", "='06_Operating_Model'!F4", "='06_Operating_Model'!G4"],
    ["='06_Operating_Model'!B7", "='06_Operating_Model'!C7", "='06_Operating_Model'!D7", "='06_Operating_Model'!E7", "='06_Operating_Model'!F7", "='06_Operating_Model'!G7"],
    ["='06_Operating_Model'!B9", "='06_Operating_Model'!C9", "='06_Operating_Model'!D9", "='06_Operating_Model'!E9", "='06_Operating_Model'!F9", "='06_Operating_Model'!G9"],
    ["='06_Operating_Model'!B10", "='06_Operating_Model'!C10", "='06_Operating_Model'!D10", "='06_Operating_Model'!E10", "='06_Operating_Model'!F10", "='06_Operating_Model'!G10"],
    ["='06_Operating_Model'!B12", "='06_Operating_Model'!C12", "='06_Operating_Model'!D12", "='06_Operating_Model'!E12", "='06_Operating_Model'!F12", "='06_Operating_Model'!G12"],
    ["='06_Operating_Model'!B13", "='06_Operating_Model'!C13", "='06_Operating_Model'!D13", "='06_Operating_Model'!E13", "='06_Operating_Model'!F13", "='06_Operating_Model'!G13"],
    ["=B10", "=C10", "=D10", "=E10", "=F10", "=G10"],
  ];
  section(s, "Cash Flow", "A14:G14");
  s.getRange("A15:G15").values = [["Line item", "FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]];
  header(s.getRange("A15:G15"));
  s.getRange("A16:G21").values = [
    ["NOPAT", null, null, null, null, null, null],
    ["D&A add-back", null, null, null, null, null, null],
    ["Capex", null, null, null, null, null, null],
    ["Change in NWC", null, null, null, null, null, null],
    ["Unlevered FCF", null, null, null, null, null, null],
    ["Ending cash proxy", null, null, null, null, null, null],
  ];
  s.getRange("B16:G21").formulas = [
    ["='06_Operating_Model'!B13", "='06_Operating_Model'!C13", "='06_Operating_Model'!D13", "='06_Operating_Model'!E13", "='06_Operating_Model'!F13", "='06_Operating_Model'!G13"],
    ["='06_Operating_Model'!B9", "='06_Operating_Model'!C9", "='06_Operating_Model'!D9", "='06_Operating_Model'!E9", "='06_Operating_Model'!F9", "='06_Operating_Model'!G9"],
    ["='06_Operating_Model'!B15", "='06_Operating_Model'!C15", "='06_Operating_Model'!D15", "='06_Operating_Model'!E15", "='06_Operating_Model'!F15", "='06_Operating_Model'!G15"],
    ["='06_Operating_Model'!B18", "='06_Operating_Model'!C18", "='06_Operating_Model'!D18", "='06_Operating_Model'!E18", "='06_Operating_Model'!F18", "='06_Operating_Model'!G18"],
    ["='06_Operating_Model'!B19", "='06_Operating_Model'!C19", "='06_Operating_Model'!D19", "='06_Operating_Model'!E19", "='06_Operating_Model'!F19", "='06_Operating_Model'!G19"],
    ["='04_Historical_FS'!F16", "=B21+C20", "=C21+D20", "=D21+E20", "=E21+F20", "=F21+G20"],
  ];
  section(s, "Balance Sheet Drivers", "A24:G24");
  s.getRange("A25:G25").values = [["Line item", "FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]];
  header(s.getRange("A25:G25"));
  s.getRange("A26:G32").values = [
    ["Cash", null, null, null, null, null, null],
    ["Debt", null, null, null, null, null, null],
    ["Net debt / (cash)", null, null, null, null, null, null],
    ["Net working capital", null, null, null, null, null, null],
    ["Equity proxy", null, null, null, null, null, null],
    ["Check: cash roll-forward", null, null, null, null, null, null],
    ["Check: net debt bridge", null, null, null, null, null, null],
  ];
  s.getRange("B26:G32").formulas = [
    ["='04_Historical_FS'!F16", "=C21", "=D21", "=E21", "=F21", "=G21"],
    ["='04_Historical_FS'!F17", "=B27", "=C27", "=D27", "=E27", "=F27"],
    ["=B27-B26", "=C27-C26", "=D27-D26", "=E27-E26", "=F27-F26", "=G27-G26"],
    ["='06_Operating_Model'!B17", "='06_Operating_Model'!C17", "='06_Operating_Model'!D17", "='06_Operating_Model'!E17", "='06_Operating_Model'!F17", "='06_Operating_Model'!G17"],
    ["='04_Historical_FS'!F15", "=B30+C11", "=C30+D11", "=D30+E11", "=E30+F11", "=F30+G11"],
    ["=B26-'04_Historical_FS'!F16", "=C26-C21", "=D26-D21", "=E26-E21", "=F26-F21", "=G26-G21"],
    ["=B28-('04_Historical_FS'!F17-'04_Historical_FS'!F16)", "=C28-(C27-C26)", "=D28-(D27-D26)", "=E28-(E27-E26)", "=F28-(F27-F26)", "=G28-(G27-G26)"],
  ];
  s.getRange("B5:G32").format.numberFormat = numFmt;
  for (const row of [10, 11, 20, 21, 28, 30, 31, 32]) total(s.getRange(`A${row}:G${row}`));
  formulas(s.getRange("B5:G32"));
  setWidths(s, [28, 14, 14, 14, 14, 14, 14]);
  freeze(s, 4, 1);
}

// DCF
{
  const s = sheets.DCF;
  title(s, "Discounted Cash Flow Valuation", "A1:H1");
  s.getRange("A3:G3").values = [["Line item", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E", "Notes"]];
  header(s.getRange("A3:G3"));
  s.getRange("A4:G15").values = [
    ["Revenue", null, null, null, null, null, "From operating model"],
    ["EBITDA", null, null, null, null, null, "From operating model"],
    ["EBIT", null, null, null, null, null, "From operating model"],
    ["Cash taxes", null, null, null, null, null, "From operating model"],
    ["NOPAT", null, null, null, null, null, "EBIT less taxes"],
    ["D&A", null, null, null, null, null, "Add-back"],
    ["Capex", null, null, null, null, null, "Cash outflow as positive number"],
    ["Change in NWC", null, null, null, null, null, "Cash outflow as positive number"],
    ["Unlevered FCF", null, null, null, null, null, "NOPAT + D&A - capex - change in NWC"],
    ["Discount period", 0.5, 1.5, 2.5, 3.5, 4.5, "Mid-year convention"],
    ["WACC", null, null, null, null, null, "From assumptions"],
    ["Discount factor", null, null, null, null, null, "'=1/(1+WACC)^period"],
  ];
  s.getRange("B4:F12").formulas = [
    ["='06_Operating_Model'!C4", "='06_Operating_Model'!D4", "='06_Operating_Model'!E4", "='06_Operating_Model'!F4", "='06_Operating_Model'!G4"],
    ["='06_Operating_Model'!C7", "='06_Operating_Model'!D7", "='06_Operating_Model'!E7", "='06_Operating_Model'!F7", "='06_Operating_Model'!G7"],
    ["='06_Operating_Model'!C10", "='06_Operating_Model'!D10", "='06_Operating_Model'!E10", "='06_Operating_Model'!F10", "='06_Operating_Model'!G10"],
    ["='06_Operating_Model'!C12", "='06_Operating_Model'!D12", "='06_Operating_Model'!E12", "='06_Operating_Model'!F12", "='06_Operating_Model'!G12"],
    ["='06_Operating_Model'!C13", "='06_Operating_Model'!D13", "='06_Operating_Model'!E13", "='06_Operating_Model'!F13", "='06_Operating_Model'!G13"],
    ["='06_Operating_Model'!C9", "='06_Operating_Model'!D9", "='06_Operating_Model'!E9", "='06_Operating_Model'!F9", "='06_Operating_Model'!G9"],
    ["='06_Operating_Model'!C15", "='06_Operating_Model'!D15", "='06_Operating_Model'!E15", "='06_Operating_Model'!F15", "='06_Operating_Model'!G15"],
    ["='06_Operating_Model'!C18", "='06_Operating_Model'!D18", "='06_Operating_Model'!E18", "='06_Operating_Model'!F18", "='06_Operating_Model'!G18"],
    ["='06_Operating_Model'!C19", "='06_Operating_Model'!D19", "='06_Operating_Model'!E19", "='06_Operating_Model'!F19", "='06_Operating_Model'!G19"],
  ];
  s.getRange("B14:F14").formulas = [["='05_Assumptions'!B31", "='05_Assumptions'!B31", "='05_Assumptions'!B31", "='05_Assumptions'!B31", "='05_Assumptions'!B31"]];
  s.getRange("B15:F15").formulas = [["=1/(1+B14)^B13", "=1/(1+C14)^C13", "=1/(1+D14)^D13", "=1/(1+E14)^E13", "=1/(1+F14)^F13"]];
  section(s, "Valuation Bridge", "A18:E18");
  s.getRange("A19:E32").values = [
    ["Metric", "Gordon Growth", "Exit Multiple", "Selected", "Notes"],
    ["PV of forecast FCF", null, null, null, "Sum of discounted FY27E-FY31E FCF"],
    ["Terminal value", null, null, null, "Gordon and exit multiple shown separately"],
    ["PV of terminal value", null, null, null, "Discounted using FY31E discount factor"],
    ["Enterprise value", null, null, null, "PV FCF + PV terminal value"],
    ["Net debt / (cash)", null, null, null, "From assumptions until actual bridge is locked"],
    ["Equity value", null, null, null, "Enterprise value - net debt"],
    ["Diluted shares (mm)", null, null, null, "From operating model / assumptions"],
    ["Implied value / share", null, null, null, "Equity value / diluted shares"],
    ["Selected low", null, null, null, "Lower of Gordon / exit method"],
    ["Selected mid", null, null, null, "Average of Gordon / exit method"],
    ["Selected high", null, null, null, "Higher of Gordon / exit method"],
    ["Terminal growth", null, null, null, "From assumptions"],
    ["Exit EBITDA multiple", null, null, null, "From assumptions"],
  ];
  s.getRange("B20:D32").formulas = [
    ["=SUMPRODUCT(B12:F12,B15:F15)", "=B20", "=AVERAGE(B20:C20)"],
    ["=IF('05_Assumptions'!B31>'05_Assumptions'!B35,F12*(1+'05_Assumptions'!B35)/('05_Assumptions'!B31-'05_Assumptions'!B35),0)", "=F5*'05_Assumptions'!B36", "=AVERAGE(B21:C21)"],
    ["=B21*F15", "=C21*F15", "=AVERAGE(B22:C22)"],
    ["=B20+B22", "=C20+C22", "=AVERAGE(B23:C23)"],
    ["='05_Assumptions'!B32", "='05_Assumptions'!B32", "='05_Assumptions'!B32"],
    ["=B23-B24", "=C23-C24", "=D23-D24"],
    ["='06_Operating_Model'!G20", "='06_Operating_Model'!G20", "='06_Operating_Model'!G20"],
    ["=IFERROR(B25/B26,0)", "=IFERROR(C25/C26,0)", "=IFERROR(D25/D26,0)"],
    ["=MIN(B27,C27)", "=B28", "=MIN(B27,C27)"],
    ["=AVERAGE(B27,C27)", "=B29", "=AVERAGE(B27,C27)"],
    ["=MAX(B27,C27)", "=B30", "=MAX(B27,C27)"],
    ["='05_Assumptions'!B35", "='05_Assumptions'!B35", "='05_Assumptions'!B35"],
    ["='05_Assumptions'!B36", "='05_Assumptions'!B36", "='05_Assumptions'!B36"],
  ];
  s.getRange("B4:F12").format.numberFormat = numFmt;
  s.getRange("B13:F15").format.numberFormat = "0.0";
  s.getRange("B14:F14").format.numberFormat = pctFmt;
  s.getRange("B19:D26").format.numberFormat = crFmt;
  s.getRange("B27:D30").format.numberFormat = priceFmt;
  s.getRange("B31:D31").format.numberFormat = pctFmt;
  s.getRange("B32:D32").format.numberFormat = multFmt;
  total(s.getRange("A12:F12"));
  total(s.getRange("A20:D30"));
  section(s, "Sensitivity: Implied Value / Share - Gordon Growth", "A35:F35");
  s.getRange("A36:F41").values = [
    ["WACC \\ TGR", 0.03, 0.035, 0.04, 0.045, 0.05],
    [0.10, null, null, null, null, null],
    [0.11, null, null, null, null, null],
    [0.12, null, null, null, null, null],
    [0.13, null, null, null, null, null],
    [0.14, null, null, null, null, null],
  ];
  s.getRange("B37:F41").formulas = Array.from({ length: 5 }, (_, r) =>
    Array.from({ length: 5 }, (_, c) => {
      const row = 37 + r;
      const col = String.fromCharCode("B".charCodeAt(0) + c);
      return `=IFERROR(((SUMPRODUCT($B$12:$F$12,1/(1+$A${row})^$B$13:$F$13)+($F$12*(1+${col}$36)/($A${row}-${col}$36))*$F$15)-'05_Assumptions'!$B$32)/'06_Operating_Model'!$G$20,0)`;
    })
  );
  header(s.getRange("A36:F36"));
  inputs(s.getRange("B36:F36"));
  inputs(s.getRange("A37:A41"));
  s.getRange("A37:A41").format.numberFormat = pctFmt;
  s.getRange("B36:F36").format.numberFormat = pctFmt;
  s.getRange("B37:F41").format.numberFormat = priceFmt;
  setWidths(s, [28, 16, 16, 16, 16, 16, 34, 14]);
  freeze(s, 3, 1);
}

// Trading Comps
{
  const s = sheets["Trading Comps"];
  title(s, "Comparable Company Analysis", "A1:J1");
  s.getRange("A3:J3").values = [["Company", "Ticker", "Business fit", "Market cap", "Enterprise value", "Revenue", "EBITDA", "EV/Revenue", "EV/EBITDA", "Source / notes"]];
  header(s.getRange("A3:J3"));
  s.getRange("A4:J11").values = [
    ["Swiggy", "SWIGGY.NS", "India food delivery / quick commerce", null, null, null, null, null, null, "Closest listed India peer if current market data is available."],
    ["DoorDash", "DASH", "Global food delivery / local commerce", null, null, null, null, null, null, "Useful global benchmark; currency consistency required."],
    ["Deliveroo", "ROO.L", "UK food delivery", null, null, null, null, null, null, "Check current listing status and latest market data."],
    ["Delivery Hero", "DHER.DE", "Global food delivery", null, null, null, null, null, null, "Peer for scale and margin cross-check."],
    ["Meituan", "3690.HK", "China local services", null, null, null, null, null, null, "Strategic peer but scale/geography differences are material."],
    ["Maplebear / Instacart", "CART", "Grocery / quick-commerce adjacent", null, null, null, null, null, null, "Optional quick-commerce comp."],
    ["Jubilant FoodWorks", "JUBLFOOD.NS", "India food services", null, null, null, null, null, null, "Useful domestic consumer/food benchmark, less direct digitally."],
    ["Median", "", "", null, null, null, null, null, null, ""],
  ];
  s.getRange("H4:H10").formulas = Array.from({ length: 7 }, (_, i) => [`=IFERROR(E${4 + i}/F${4 + i},"")`]);
  s.getRange("I4:I10").formulas = Array.from({ length: 7 }, (_, i) => [`=IFERROR(E${4 + i}/G${4 + i},"")`]);
  s.getRange("H11").formulas = [["=IFERROR(MEDIAN(H4:H10),0)"]];
  s.getRange("I11").formulas = [["=IFERROR(MEDIAN(I4:I10),0)"]];
  inputs(s.getRange("D4:G10"));
  formulas(s.getRange("H4:I11"));
  total(s.getRange("A11:J11"));
  section(s, "Implied Valuation", "A14:H14");
  s.getRange("A15:H18").values = [
    ["Method", "Metric", "Median multiple", "Low", "Mid", "High", "Implied EV", "Implied value / share"],
    ["EV / Revenue", null, null, null, null, null, null, null],
    ["EV / EBITDA", null, null, null, null, null, null, null],
    ["Selected comps", null, null, null, null, null, null, null],
  ];
  s.getRange("B16:B18").formulas = [["='06_Operating_Model'!G4"], ["='06_Operating_Model'!G7"], ["=AVERAGE(B16:B17)"]];
  s.getRange("C16:C17").formulas = [["=H11"], ["=I11"]];
  s.getRange("D16:F17").formulas = [
    ["=IFERROR(C16*(1-'05_Assumptions'!E25),0)", "=IFERROR(C16,0)", "=IFERROR(C16*(1+'05_Assumptions'!E26),0)"],
    ["=IFERROR(C17*(1-'05_Assumptions'!E25),0)", "=IFERROR(C17,0)", "=IFERROR(C17*(1+'05_Assumptions'!E26),0)"],
  ];
  s.getRange("G16:G17").formulas = [["=IFERROR(B16*E16,0)"], ["=IFERROR(B17*E17,0)"]];
  s.getRange("H16:H17").formulas = [["=IFERROR((G16-'05_Assumptions'!B32)/'06_Operating_Model'!G20,0)"], ["=IFERROR((G17-'05_Assumptions'!B32)/'06_Operating_Model'!G20,0)"]];
  s.getRange("D18:F18").formulas = [["=IFERROR(AVERAGE(D16:D17),0)", "=IFERROR(AVERAGE(E16:E17),0)", "=IFERROR(AVERAGE(F16:F17),0)"]];
  s.getRange("G18").formulas = [["=IFERROR(AVERAGE(G16:G17),0)"]];
  s.getRange("H18").formulas = [["=IFERROR(AVERAGE(H16:H17),0)"]];
  header(s.getRange("A15:H15"));
  s.getRange("D4:G10").format.numberFormat = crFmt;
  s.getRange("H4:I11").format.numberFormat = multFmt;
  s.getRange("B16:B18").format.numberFormat = crFmt;
  s.getRange("C16:F18").format.numberFormat = multFmt;
  s.getRange("G16:G18").format.numberFormat = crFmt;
  s.getRange("H16:H18").format.numberFormat = priceFmt;
  total(s.getRange("A18:H18"));
  setWidths(s, [24, 14, 34, 16, 16, 16, 16, 14, 14, 44]);
  freeze(s, 3, 1);
}

// Precedents
{
  const s = sheets.Precedents;
  title(s, "Precedent Transactions", "A1:J1");
  s.getRange("A3:J3").values = [["Target / asset", "Acquirer", "Announced", "Business fit", "Transaction value", "Revenue", "EBITDA", "EV/Revenue", "EV/EBITDA", "Source / notes"]];
  header(s.getRange("A3:J3"));
  s.getRange("A4:J9").values = [
    ["Blinkit", "Zomato", "2022", "Quick commerce", null, null, null, null, null, "Useful strategic precedent; verify transaction value and revenue base."],
    ["Wolt", "DoorDash", "2021", "Food delivery / local commerce", null, null, null, null, null, "Global delivery precedent; check deal currency/date."],
    ["Glovo stake / acquisition", "Delivery Hero", "2021", "Food delivery", null, null, null, null, null, "European food delivery precedent."],
    ["Careem food / grocery assets", "Uber / regional buyers", "TBD", "Local commerce", null, null, null, null, null, "Use only if data is auditable."],
    ["Quick-commerce asset", "TBD", "TBD", "Quick commerce", null, null, null, null, null, "Optional if relevant Indian deal data is available."],
    ["Median", "", "", "", null, null, null, null, null, ""],
  ];
  s.getRange("H4:H8").formulas = Array.from({ length: 5 }, (_, i) => [`=IFERROR(E${4 + i}/F${4 + i},"")`]);
  s.getRange("I4:I8").formulas = Array.from({ length: 5 }, (_, i) => [`=IFERROR(E${4 + i}/G${4 + i},"")`]);
  s.getRange("H9").formulas = [["=IFERROR(MEDIAN(H4:H8),0)"]];
  s.getRange("I9").formulas = [["=IFERROR(MEDIAN(I4:I8),0)"]];
  inputs(s.getRange("E4:G8"));
  formulas(s.getRange("H4:I9"));
  total(s.getRange("A9:J9"));
  section(s, "Implied Precedent Valuation", "A12:H12");
  s.getRange("A13:H15").values = [
    ["Method", "Metric", "Median multiple", "Control premium", "Selected multiple", "Implied EV", "Implied value / share", "Notes"],
    ["EV / Revenue", null, null, null, null, null, null, "Use only if transaction set is credible."],
    ["EV / EBITDA", null, null, null, null, null, null, "Use only if EBITDA bases are clean."],
  ];
  s.getRange("B14:B15").formulas = [["='06_Operating_Model'!G4"], ["='06_Operating_Model'!G7"]];
  s.getRange("C14:C15").formulas = [["=H9"], ["=I9"]];
  s.getRange("D14:D15").formulas = [["='05_Assumptions'!E27"], ["='05_Assumptions'!E27"]];
  s.getRange("E14:E15").formulas = [["=IFERROR(C14*(1+D14),0)"], ["=IFERROR(C15*(1+D15),0)"]];
  s.getRange("F14:F15").formulas = [["=IFERROR(B14*E14,0)"], ["=IFERROR(B15*E15,0)"]];
  s.getRange("G14:G15").formulas = [["=IFERROR((F14-'05_Assumptions'!B32)/'06_Operating_Model'!G20,0)"], ["=IFERROR((F15-'05_Assumptions'!B32)/'06_Operating_Model'!G20,0)"]];
  header(s.getRange("A13:H13"));
  s.getRange("E4:G8").format.numberFormat = crFmt;
  s.getRange("H4:I9").format.numberFormat = multFmt;
  s.getRange("B14:B15").format.numberFormat = crFmt;
  s.getRange("C14:E15").format.numberFormat = multFmt;
  s.getRange("D14:D15").format.numberFormat = pctFmt;
  s.getRange("F14:F15").format.numberFormat = crFmt;
  s.getRange("G14:G15").format.numberFormat = priceFmt;
  setWidths(s, [24, 18, 14, 28, 18, 16, 16, 14, 14, 44]);
  freeze(s, 3, 1);
}

// Football Field
{
  const s = sheets["Football Field"];
  title(s, "Football Field Valuation Summary", "A1:H1");
  s.getRange("A3:E3").values = [["Method", "Low", "Mid", "High", "Notes"]];
  header(s.getRange("A3:E3"));
  s.getRange("A4:E8").values = [
    ["DCF - Gordon Growth", null, null, null, "From DCF sensitivity and bridge"],
    ["DCF - Exit Multiple", null, null, null, "From DCF bridge"],
    ["Trading Comps", null, null, null, "Peer multiples; fill live data before relying on range"],
    ["Precedents", null, null, null, "Optional; use only if transaction set is credible"],
    ["Selected range", null, null, null, "Average of available methods"],
  ];
  s.getRange("B4:D8").formulas = [
    ["='08_DCF'!B28", "='08_DCF'!B29", "='08_DCF'!B30"],
    ["='08_DCF'!C28", "='08_DCF'!C29", "='08_DCF'!C30"],
    ["='09_Trading_Comps'!H16*(1-'05_Assumptions'!E25)", "='09_Trading_Comps'!H18", "='09_Trading_Comps'!H16*(1+'05_Assumptions'!E26)"],
    ["=MIN('10_Precedents'!G14:G15)", "=AVERAGE('10_Precedents'!G14:G15)", "=MAX('10_Precedents'!G14:G15)"],
    ["=AVERAGE(B4:B7)", "=AVERAGE(C4:C7)", "=AVERAGE(D4:D7)"],
  ];
  s.getRange("B4:D8").format.numberFormat = priceFmt;
  total(s.getRange("A8:D8"));
  section(s, "Chart Helper", "G3:J3");
  s.getRange("G4:J8").values = [
    ["Method", "Low", "Range", "Mid"],
    ["DCF - Gordon", null, null, null],
    ["DCF - Exit", null, null, null],
    ["Trading Comps", null, null, null],
    ["Precedents", null, null, null],
  ];
  s.getRange("H5:J8").formulas = [
    ["=B4", "=D4-B4", "=C4"],
    ["=B5", "=D5-B5", "=C5"],
    ["=B6", "=D6-B6", "=C6"],
    ["=B7", "=D7-B7", "=C7"],
  ];
  s.getRange("H5:J8").format.numberFormat = priceFmt;
  section(s, "Football Field Chart Status", "A11:H11");
  s.getRange("A12:H14").merge();
  s.getRange("A12:H14").values = [[
    "The chart is intentionally deferred in Phase 1. Once historical actuals, market data, and peer / transaction multiples are populated, this sheet will convert the valuation range table into the final IB-style football field visual."
  ]];
  s.getRange("A12:H14").format = { wrapText: true, fill: "#F8FAFC", verticalAlignment: "top" };
  s.getRange("A12:H14").format.rowHeight = 56;
  setWidths(s, [24, 14, 14, 14, 46, 4, 18, 14, 14, 14]);
  freeze(s, 3, 1);
}

// Thesis
{
  const s = sheets.Thesis;
  title(s, "One-Page Investment Thesis Draft", "A1:D1");
  s.getRange("A3:B10").values = [
    ["Company", "Eternal Limited"],
    ["Ticker", "ETERNAL.NS / ETERNAL"],
    ["Recommendation", null],
    ["Target price range", null],
    ["Valuation basis", "DCF cross-checked with trading comps and precedent transactions."],
    ["Core question", "Can Eternal compound quick-commerce growth while expanding food delivery margins enough to justify premium local-commerce multiples?"],
    ["Model status", null],
    ["Final recommendation", null],
  ];
  s.getRange("B5").formulas = [["='05_Assumptions'!B14"]];
  s.getRange("B6").formulas = [["=TEXT('11_Football_Field'!B8,\"0.00\")&\" - \"&TEXT('11_Football_Field'!D8,\"0.00\")"]];
  s.getRange("B9").formulas = [["='13_Checks'!F14"]];
  s.getRange("B10").formulas = [["=IF('11_Football_Field'!C8>'05_Assumptions'!B34*1.15,\"Buy\",IF('11_Football_Field'!C8<'05_Assumptions'!B34*0.85,\"Sell\",\"Hold\"))"]];
  section(s, "Investment Thesis", "A12:D12");
  s.getRange("A13:D18").values = [
    ["Overview", "Eternal is a consumer internet / local-commerce platform with food delivery as the mature profit engine and quick commerce as the high-growth strategic swing factor.", "", ""],
    ["Upside driver 1", "Blinkit / quick-commerce growth can expand the addressable market and may warrant a higher revenue multiple if unit economics continue improving.", "", ""],
    ["Upside driver 2", "Food delivery scale and ad monetization can support margin expansion if competitive intensity remains rational.", "", ""],
    ["Downside risk 1", "Quick-commerce competition may require elevated discounts, dark-store capex, and delivery subsidies, delaying consolidated margin expansion.", "", ""],
    ["Downside risk 2", "Regulatory pressure, platform-fee sensitivity, or gig-worker cost inflation could compress contribution margins.", "", ""],
    ["What must be proven", "The final analyst note needs source-backed segment assumptions, peer-multiple sanity checks, and a clear bridge from valuation output to recommendation.", "", ""],
  ];
  section(s, "Interview Defense Prompts", "A21:D21");
  s.getRange("A22:D27").values = [
    ["Walk me through the DCF.", "Start with revenue/margin drivers, bridge to UFCF, explain WACC, terminal value, and sensitivity.", "", ""],
    ["Why Eternal?", "It fills the finance modeling gap while letting you discuss a current strategic story: food delivery cash generation plus quick-commerce growth.", "", ""],
    ["Why not a bank?", "Bank valuations are valuable but use different mechanics; this project specifically targets DCF, EV multiples, precedents, and a football field.", "", ""],
    ["What is the key risk?", "Unit economics and competitive intensity in quick commerce are the major swing factors.", "", ""],
    ["How would you improve it?", "Move from company-level growth assumptions to a segment build with GOV/order frequency/AOV/take-rate/contribution-margin drivers.", "", ""],
    ["What is defensible today?", "The architecture and method; the conclusion becomes defensible only after actuals and market data are sourced.", "", ""],
  ];
  s.getRange("A3:A10").format = { fill: steel, font: { bold: true } };
  s.getRange("A13:A18").format = { fill: steel, font: { bold: true } };
  s.getRange("A22:A27").format = { fill: steel, font: { bold: true } };
  s.getRange("B13:D27").format = { wrapText: true };
  setWidths(s, [26, 64, 18, 18]);
  freeze(s, 1, 0);
}

// Checks
{
  const s = sheets.Checks;
  title(s, "Model Checks", "A1:G1");
  s.getRange("A3:G3").values = [["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Fix hint"]];
  header(s.getRange("A3:G3"));
  s.getRange("A4:G14").values = [
    ["Historical revenue entered", null, 1, null, 0, null, "Enter FY26A revenue in Historical FS."],
    ["Share count available", null, 1, null, 0, null, "Enter diluted shares or use assumption default."],
    ["WACC greater than terminal growth", null, 1, null, 0, null, "Lower terminal growth or raise WACC."],
    ["DCF denominator valid", null, 1, null, 0, null, "WACC must exceed terminal growth."],
    ["Operating FCF is formula-backed", null, 1, null, 0, null, "Check Operating Model row 19."],
    ["Comps median available", null, 1, null, 0, null, "Enter EV, revenue, EBITDA for at least one peer."],
    ["Precedents optional", null, 1, null, 0, null, "Can remain warning if no credible transaction set is used."],
    ["No circular cash roll-forward issue", null, 0, null, 0, null, "Statements cash check should be zero."],
    ["No net debt bridge issue", null, 0, null, 0, null, "Statements net debt check should be zero."],
    ["Football field range valid", null, 1, null, 0, null, "Low should be <= mid <= high."],
    ["Overall model status", null, null, null, null, null, "OK only after core required checks pass."],
  ];
  s.getRange("B4:B14").formulas = [
    ["=IF('04_Historical_FS'!F4>0,1,0)"],
    ["=IF('06_Operating_Model'!G20>0,1,0)"],
    ["=IF('05_Assumptions'!B31>'05_Assumptions'!B35,1,0)"],
    ["=IF('08_DCF'!B21>=0,1,0)"],
    ["=1"],
    ["=IF('09_Trading_Comps'!H11>0,1,0)"],
    ["=IF(OR('10_Precedents'!H9>0,'10_Precedents'!I9>0),1,0)"],
    ["='07_Statements'!G31"],
    ["='07_Statements'!G32"],
    ["=IF(AND('11_Football_Field'!B8<='11_Football_Field'!C8,'11_Football_Field'!C8<='11_Football_Field'!D8),1,0)"],
    ["=COUNTIF(F4:F13,\"OK\")"],
  ];
  s.getRange("D4:D14").formulas = [
    ["=B4-C4"],
    ["=B5-C5"],
    ["=B6-C6"],
    ["=B7-C7"],
    ["=B8-C8"],
    ["=B9-C9"],
    ["=B10-C10"],
    ["=B11-C11"],
    ["=B12-C12"],
    ["=B13-C13"],
    ["=B14-9"],
  ];
  s.getRange("F4:F14").formulas = [
    ["=IF(ABS(D4)<=E4,\"OK\",\"Needs input\")"],
    ["=IF(ABS(D5)<=E5,\"OK\",\"Needs input\")"],
    ["=IF(ABS(D6)<=E6,\"OK\",\"Review\")"],
    ["=IF(ABS(D7)<=E7,\"OK\",\"Review\")"],
    ["=IF(ABS(D8)<=E8,\"OK\",\"Review\")"],
    ["=IF(ABS(D9)<=E9,\"OK\",\"Needs input\")"],
    ["=IF(ABS(D10)<=E10,\"OK\",\"Optional\")"],
    ["=IF(ABS(D11)<=E11,\"OK\",\"Review\")"],
    ["=IF(ABS(D12)<=E12,\"OK\",\"Review\")"],
    ["=IF(ABS(D13)<=E13,\"OK\",\"Review\")"],
    ["=IF(B14>=9,\"OK\",\"Phase 1.1 - needs source inputs\")"],
  ];
  s.getRange("B4:E14").format.numberFormat = numFmt;
  total(s.getRange("A14:G14"));
  setWidths(s, [34, 14, 14, 14, 14, 18, 48]);
  freeze(s, 3, 1);
}

// Add source comments to key editable assumptions.
await addComment(sheets.Assumptions.getRange("B21"), "Source required: current India 10-year government bond yield as of valuation date.");
await addComment(sheets.Assumptions.getRange("B22"), "Source required: levered beta from market data provider; note if adjusted or raw.");
await addComment(sheets.Assumptions.getRange("B23"), "Source required: equity risk premium, preferably Damodaran ERP data with date/version.");
await addComment(sheets.Assumptions.getRange("B33"), "Source required: diluted share count from latest filing / exchange data.");
await addComment(sheets["Raw Data"].getRange("B4:F15"), "Populate from official annual reports and latest exchange-filed results. Keep expenses as positive values.");
await addComment(sheets["Historical FS"].getRange("B4:F24"), "Linked from Raw_Data. Do not enter source values here.");

// Compact visual formatting pass.
for (const sheet of Object.values(sheets)) {
  const used = sheet.getUsedRange();
  used.format.font = { name: "Aptos", size: 10 };
  used.format.wrapText = false;
}
sheets.Cover.getRange("B3:B18").format.wrapText = true;
sheets.Dashboard.getRange("A16:J23").format.wrapText = false;
sheets.Dashboard.getRange("A16:J23").format.rowHeight = 24;
sheets["Raw Data"].getRange("I4:J15").format.wrapText = true;
sheets["Raw Data"].getRange("B19:C23").format.wrapText = true;
sheets["Sources Audit"].getRange("G4:I16").format.wrapText = true;
sheets.Assumptions.getRange("G5:H14").format.wrapText = true;
sheets["Trading Comps"].getRange("J4:J11").format.wrapText = true;
sheets.Precedents.getRange("J4:J9").format.wrapText = true;
sheets.Thesis.getRange("B13:D27").format.wrapText = true;

// Verification artifacts.
const summary = await workbook.inspect({
  kind: "table",
  sheetId: sheetNames.Cover,
  range: "A1:I18",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 10,
  maxChars: 6000,
});
await fs.writeFile(path.join(outputDir, "cover_inspect.ndjson"), summary.ndjson);

const dashboardSummary = await workbook.inspect({
  kind: "table",
  sheetId: sheetNames.Dashboard,
  range: "A1:J23",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 10,
  maxChars: 8000,
});
await fs.writeFile(path.join(outputDir, "dashboard_inspect.ndjson"), dashboardSummary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 6000,
});
await fs.writeFile(path.join(outputDir, "formula_error_scan.ndjson"), errors.ndjson);

for (const [key, displayName] of Object.entries(sheetNames)) {
  const preview = await workbook.render({
    sheetName: displayName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(path.join(outputDir, `${displayName.toLowerCase()}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(path.join(outputDir, "Eternal_Valuation_Model_v1.1_Professional_Structure.xlsx"));
await xlsx.save(path.join(excelDir, "Eternal_Valuation_Model.xlsx"));
