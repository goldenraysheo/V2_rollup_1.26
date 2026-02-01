/*************************************************************
 * CONFIGURATION
 *************************************************************/
const ROLLUP_FOLDER_NAME = '2026 Branch Reporting Workbooks';
const MASTER_SHEET_NAME = 'Monthly Variances - Master';
const DEPT_NAMES_SHEET = 'Dept_Names';
const MONTH_ABBRS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/*************************************************************
 * MENU
 *************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Roll-up Tools')
    .addItem('Refresh Monthly Details', 'refreshAllData')
    .addItem('Refresh Executive Summaries', 'refreshExecutiveSummaries')
    .addItem('Refresh AR Collections', 'refreshARCollections')
    .addItem('Refresh Membership', 'refreshMembershipTrends')
    .addToUi();
}

/*************************************************************
 * MAIN WRAPPER
 *************************************************************/
function refreshAllData() {
  const ss = SpreadsheetApp.getActive();

  // Show start message
  ss.toast('Starting data refresh...', 'Roll-up Progress', -1);

  rebuildMasterData();

  // Completion handled in rebuildMasterData
}

/*************************************************************
 * BUILD MASTER DATA (WITH PROGRESS TRACKING)
 *************************************************************/
function rebuildMasterData() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Preparing master sheet...', 'Progress', -1);

  const master = upsertSheet_(ss, MASTER_SHEET_NAME);
  master.clear();
  master.showSheet();
  writeMasterHeader_(master);

  const header = [
    'Year','Month','Period',
    'Branch','Dept','Dept Name',
    'Account','Account Name','Type',
    'Actual','Budget','Variance','Last Year','Comments'
  ];
  master.getRange(4, 1, 1, header.length).setValues([header]);

  ss.toast('Loading department names...', 'Progress', -1);
  const deptMap = loadDeptNames_(ss);

  ss.toast('Finding branch workbooks...', 'Progress', -1);
  const books = getBranchSpreadsheets_();

  ss.toast(`Processing ${books.length} branch workbook${books.length !== 1 ? 's' : ''}...`, 'Progress', -1);

  const rows = [];
  let processedCount = 0;

  books.forEach(book => {
    processedCount++;
    const bookName = book.getName();
    ss.toast(`Processing ${bookName} (${processedCount} of ${books.length})...`, 'Progress', -1);

    MONTH_ABBRS.forEach((mon, idx) => {
      const period = idx + 1;
      const archive = book.getSheetByName(`Archive_P${period}-${mon}`);
      if (!archive) return;

      const headerRow = findArchiveHeaderRow_(archive);
      if (!headerRow) return;

      const year = extractYearFromArchive_(archive, headerRow);
      const archiveRows = extractArchiveRows_(archive, headerRow);
      const comments = extractCommentsFromPTab_(book, `P${period}-${mon}`);

      archiveRows.forEach(r => {
        const variance = r.type === 'Revenue'
          ? r.actual - r.budget
          : r.budget - r.actual;

        rows.push([
          year,
          mon,
          period,
          "'" + r.branch,
          "'" + r.dept,
          deptMap[r.dept] || '',
          "'" + r.acct,
          r.acctName,
          r.type,
          r.actual,
          r.budget,
          variance,
          r.lastYear,
          comments[`${r.branch}|${r.dept}|${r.acct}`] || ''
        ]);
      });
    });
  });

  ss.toast('Writing data to master sheet...', 'Progress', -1);

  if (rows.length) {
    master.getRange(5, 1, rows.length, header.length).setValues(rows);
  }

  ss.toast('Formatting master sheet...', 'Progress', -1);
  formatMasterSheet_(master, header.length);

  // Final success message with summary
  const branchWord = books.length !== 1 ? 'branches' : 'branch';
  const rowWord = rows.length !== 1 ? 'rows' : 'row';
  ss.toast(
    `Complete! Processed ${rows.length} ${rowWord} from ${books.length} ${branchWord}.`,
    'Success ✓',
    5
  );
}

/*************************************************************
 * FORMAT MASTER SHEET
 *************************************************************/
function formatMasterSheet_(master, colCount) {
  const lastRow = master.getLastRow();
  if (lastRow < 4) return;

  master.setFrozenRows(4);
  master.getDataRange().setFontFamily('Verdana').setFontSize(9);
  master.getRange(4, 1, 1, colCount)
        .setFontWeight('bold')
        .setBorder(false, false, true, false, false, false, null, SpreadsheetApp.BorderStyle.MEDIUM);

  const existingFilter = master.getFilter();
  if (existingFilter) existingFilter.remove();
  master.getRange(4, 1, lastRow - 3, colCount).createFilter();

  // Force TEXT formatting for Branch (4), Dept (5), Account (7)
  if (lastRow > 4) {
    master.getRange(5, 4, lastRow - 4, 1).setNumberFormat('@');
    master.getRange(5, 5, lastRow - 4, 1).setNumberFormat('@');
    master.getRange(5, 7, lastRow - 4, 1).setNumberFormat('@');

    const moneyFmt = '#,##0.00';
    master.getRange(5,10,lastRow-4,1).setNumberFormat(moneyFmt);
    master.getRange(5,11,lastRow-4,1).setNumberFormat(moneyFmt);
    master.getRange(5,12,lastRow-4,1).setNumberFormat(moneyFmt);
    master.getRange(5,13,lastRow-4,1).setNumberFormat(moneyFmt);

    master.getRange(5,1,lastRow-4,colCount).setBorder(
      true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.DOTTED
    );
  }

  master.getRange(1,14,lastRow,1).setWrap(true);
  master.setColumnWidth(14, 320);

  // Manual column widths for better performance (replaces autoResizeColumns)
  master.setColumnWidth(1, 50);   // Year
  master.setColumnWidth(2, 50);   // Month
  master.setColumnWidth(3, 60);   // Period
  master.setColumnWidth(4, 70);   // Branch
  master.setColumnWidth(5, 60);   // Dept
  master.setColumnWidth(6, 150);  // Dept Name
  master.setColumnWidth(7, 80);   // Account
  master.setColumnWidth(8, 200);  // Account Name
  master.setColumnWidth(9, 80);   // Type
  master.setColumnWidth(10, 100); // Actual
  master.setColumnWidth(11, 100); // Budget
  master.setColumnWidth(12, 100); // Variance
  master.setColumnWidth(13, 100); // Last Year
  // Column 14 (Comments) already set to 320 above

  master.setHiddenGridlines(true);
}

/*************************************************************
 * WRITE MASTER HEADER (Rows 1-3)
 *************************************************************/
function writeMasterHeader_(sheet) {
  // Row 1: Title
  sheet.getRange(1, 1).setValue('Monthly Variances - Master');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Row 2: Instructions
  const instructions = 'This sheet shows monthly variances from all branch workbooks. Click "Refresh Monthly Variances" to update. Use filters to view specific months, branches, or departments.';
  sheet.getRange(2, 1).setValue(instructions);
  sheet.getRange(2, 1, 1, 14).merge();
  sheet.getRange(2, 1).setWrap(true).setFontSize(9).setFontStyle('italic');

  // Row 3: Blank (space for button)
}

/*************************************************************
 * DRIVE + SHEET HELPERS
 *************************************************************/
function getBranchSpreadsheets_() {
  const ss = SpreadsheetApp.getActive();
  const thisId = ss.getId();

  const folders = DriveApp.getFoldersByName(ROLLUP_FOLDER_NAME);
  if (!folders.hasNext()) {
    throw new Error(`Folder "${ROLLUP_FOLDER_NAME}" not found.`);
  }
  const folder = folders.next();

  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const list = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.getId() === thisId) continue;
    list.push(SpreadsheetApp.open(file));
  }
  return list;
}

function upsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/*************************************************************
 * LOOKUPS
 *************************************************************/
function loadDeptNames_(ss) {
  const map = {};
  const sh = ss.getSheetByName(DEPT_NAMES_SHEET);
  if (!sh) return map;

  const lr = sh.getLastRow();
  if (lr < 1) return map;

  const vals = sh.getRange(1, 1, lr, 2).getDisplayValues();
  vals.forEach(r => {
    const code = String(r[0] || '').replace(/^'/, '').trim().padStart(2, '0');
    const name = String(r[1] || '').trim();
    if (code && name) map[code] = name;
  });
  return map;
}

/*************************************************************
 * ARCHIVE EXTRACTION
 *************************************************************/
function findArchiveHeaderRow_(sheet) {
  const maxCheck = Math.min(10, sheet.getLastRow()); // Reduced from 25 to 10
  const maxCols = sheet.getLastColumn();
  for (let r = 1; r <= maxCheck; r++) {
    const row = sheet.getRange(r, 1, 1, maxCols).getDisplayValues()[0];
    if (row.indexOf('Account Number') !== -1 && row.indexOf('Account Name') !== -1) {
      return r;
    }
  }
  return null;
}

function extractArchiveRows_(sheet, headerRow) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr <= headerRow) return [];

  const header = sheet.getRange(headerRow,1,1,lc).getDisplayValues()[0];
  const idxAcctNum = header.indexOf('Account Number');
  const idxAcctNam = header.indexOf('Account Name');
  const idxActual  = header.indexOf('Actual');
  const idxBudget  = header.indexOf('Budget');
  const idxLastYr  = header.indexOf('Last Year');

  const vals = sheet.getRange(headerRow+1,1,lr-headerRow,lc).getDisplayValues();
  const structure = detectRevenueExpenseStructure_(vals, headerRow, idxAcctNam);
  const out = [];

  vals.forEach((row, i) => {
    const acctCodeRaw = String(row[idxAcctNum] || '').trim();
    if (!/^\d{2}-\d{2}-\d{2}-\d{4}$/.test(acctCodeRaw)) return;

    const parts = acctCodeRaw.split('-');
    const branch = normalizeBranchCode_(parts[1]);
    const dept   = normalizeDeptCode_(parts[2]);
    const acct   = normalizeAcctCode_(parts[3]);

    const sheetRow = headerRow + 1 + i;
    const type = classifyType_(sheetRow, structure);
    if (!type) return;

    const acctName = String(row[idxAcctNam] || '').trim();
    const actual   = numFromCell_(row[idxActual]);
    const budget   = numFromCell_(row[idxBudget]);
    const lastYr   = idxLastYr !== -1 ? numFromCell_(row[idxLastYr]) : 0;

    if (actual === 0 && budget === 0) return;

    out.push({ branch, dept, acct, acctName, type, actual, budget, lastYear: lastYr });
  });

  return out;
}

function detectRevenueExpenseStructure_(values, headerRow, idxAcctName) {
  let rowRevenue = null, rowTotalRevenue = null;
  let rowExpenses = null, rowTotalExpenses = null;

  values.forEach((r, i) => {
    const label = String(r[0] || r[idxAcctName] || '').trim();
    const n = normalizeLabel_(label);
    const sheetRow = headerRow + 1 + i;
    if (n === 'revenue' && !rowRevenue) rowRevenue = sheetRow;
    if (n === 'totalrevenue' && !rowTotalRevenue) rowTotalRevenue = sheetRow;
    if (n === 'expenses' && !rowExpenses) rowExpenses = sheetRow;
    if (n === 'totalexpenses' && !rowTotalExpenses) rowTotalExpenses = sheetRow;
  });

  return { rowRevenue, rowTotalRevenue, rowExpenses, rowTotalExpenses };
}

function classifyType_(sheetRow, s) {
  if (!s.rowRevenue || !s.rowTotalRevenue || !s.rowExpenses || !s.rowTotalExpenses) return null;
  if (sheetRow > s.rowRevenue && sheetRow < s.rowTotalRevenue) return 'Revenue';
  if (sheetRow > s.rowExpenses && sheetRow < s.rowTotalExpenses) return 'Expense';
  return null;
}

/*************************************************************
 * COMMENTS + YEAR LOOKUP
 *************************************************************/
function extractCommentsFromPTab_(book, ptabName) {
  const sh = book.getSheetByName(ptabName);
  const out = {};
  if (!sh) return out;

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 3) return out;

  const header = sh.getRange(2,1,1,lc).getDisplayValues()[0];
  const idxBranch = header.indexOf('Branch');
  const idxDept   = header.indexOf('Dept');
  const idxAcct   = header.indexOf('Account');
  const idxComm   = header.indexOf('Comments');

  if (idxBranch === -1 || idxDept === -1 || idxAcct === -1 || idxComm === -1) return out;

  const vals = sh.getRange(3,1,lr-2,lc).getDisplayValues();
  vals.forEach(r => {
    const b = normalizeBranchCode_(r[idxBranch]);
    const d = normalizeDeptCode_(r[idxDept]);
    const a = normalizeAcctCode_(r[idxAcct]);
    const c = String(r[idxComm] || '').trim();
    if (c) out[`${b}|${d}|${a}`] = c;
  });
  return out;
}

function extractYearFromArchive_(sheet, headerRow) {
  const values = sheet.getRange(1,1,Math.min(headerRow,10),5).getDisplayValues();
  for (let row of values) {
    for (let cell of row) {
      const m = String(cell || '').match(/\b(20\d{2})\b/);
      if (m) return Number(m[1]);
    }
  }
  return new Date().getFullYear();
}

/*************************************************************
 * NORMALIZATION HELPERS
 *************************************************************/
function normalizeBranchCode_(v) {
  const s = String(v || '').replace(/^'/,'').trim();
  return s.padStart(2,'0');
}
function normalizeDeptCode_(v) {
  const s = String(v || '').replace(/^'/,'').trim();
  return s.padStart(2,'0');
}
function normalizeAcctCode_(v) {
  const s = String(v || '').replace(/^'/,'').trim();
  return s.padStart(4,'0');
}
function normalizeLabel_(s) {
  return String(s||'')
    .replace(/\u00A0/g,' ')
    .replace(/[^\w]/g,'')
    .trim()
    .toLowerCase();
}
function numFromCell_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[$,]/g,'').trim());
  return isNaN(n) ? 0 : n;
}
/*************************************************************
 * REFRESH EXECUTIVE SUMMARIES
 *************************************************************/
function refreshExecutiveSummaries() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Starting executive summaries refresh...', 'Progress', -1);

  const summarySheet = upsertSheet_(ss, 'Branch Executive Summaries');

  ss.toast('Finding branch workbooks...', 'Progress', -1);
  const books = getBranchSpreadsheets_();

  ss.toast(`Processing ${books.length} branch workbook${books.length !== 1 ? 's' : ''}...`, 'Progress', -1);

  const rows = [];
  let processedCount = 0;

  books.forEach(book => {
    processedCount++;
    const bookName = book.getName();
    ss.toast(`Processing ${bookName} (${processedCount} of ${books.length})...`, 'Progress', -1);

    const branchOverview = book.getSheetByName('Branch Overview');
    if (!branchOverview) return;

    const lastRow = branchOverview.getLastRow();
    if (lastRow < 2) return; // No data beyond header

    // Get all data from row 2 onwards
    const data = branchOverview.getRange(2, 1, lastRow - 1, 15).getDisplayValues(); // Columns A-O

    data.forEach(row => {
      const branchRaw = String(row[0] || '').replace(/^'/, '').trim(); // Remove leading apostrophe
      const month = String(row[1] || '').trim();  // Column B

      if (!branchRaw || !month) return; // Skip if missing branch or month

      // Normalize branch to 2-digit text
      const branch = "'" + branchRaw.padStart(2, '0');

      const revActual = row[2];      // Column C
      const revPlan = row[3];        // Column D
      const expActual = row[5];      // Column F
      const expPlan = row[6];        // Column G
      const netActual = row[8];      // Column I
      const netPlan = row[9];        // Column J
      const netVariance = row[10];   // Column K
      const execSummary = String(row[14] || '').trim(); // Column O

      // Parse net variance as number for sorting
      const netVarNum = parseFloat(String(netVariance).replace(/[$,]/g, '')) || 0;

      rows.push({
        branch: branch,
        month: month,
        revActual: revActual,
        revPlan: revPlan,
        expActual: expActual,
        expPlan: expPlan,
        netActual: netActual,
        netPlan: netPlan,
        netVariance: netVariance,
        netVarNum: netVarNum,
        execSummary: execSummary,
        monthOrder: getMonthOrder_(month)
      });
    });
  });

  ss.toast('Writing data to Executive Summaries sheet...', 'Progress', -1);

  if (rows.length === 0) {
    summarySheet.clear();
    writeExecutiveSummaryHeader_(summarySheet);
    ss.toast('No data found in branch workbooks.', 'Complete', 3);
    return;
  }

  // Sort by month (reverse chronological), then branch
  rows.sort((a, b) => {
    if (a.monthOrder !== b.monthOrder) return b.monthOrder - a.monthOrder; // Reverse order (newest first)
    return a.branch.localeCompare(b.branch);
  });

  // Clear and write
  summarySheet.clear();
  writeExecutiveSummaryHeader_(summarySheet);

  const header = [
    'Branch', 'Month',
    'Revenue (actual)', 'Revenue (plan)',
    'Expenses (actual)', 'Expenses (plan)',
    'Net (actual)', 'Net (plan)', 'Net variance',
    'Executive Summary'
  ];

  // Write header at row 4
  summarySheet.getRange(4, 1, 1, header.length).setValues([header]);

  // Write data starting at row 5
  const outRows = rows.map(r => [
    r.branch,
    r.month,
    r.revActual,
    r.revPlan,
    r.expActual,
    r.expPlan,
    r.netActual,
    r.netPlan,
    r.netVariance,
    r.execSummary
  ]);

  summarySheet.getRange(5, 1, outRows.length, header.length).setValues(outRows);

  ss.toast('Formatting Executive Summaries sheet...', 'Progress', -1);
  formatExecutiveSummarySheet_(summarySheet, header.length, rows.length);

  ss.toast(`Complete! Processed ${rows.length} rows from ${books.length} branches.`, 'Success ✓', 5);
}

/*************************************************************
 * HELPER: Get month order for sorting (1-12, with Dec=12)
 *************************************************************/
function getMonthOrder_(monthName) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];
  const index = months.indexOf(monthName);
  return index !== -1 ? index + 1 : 0;
}

/*************************************************************
 * WRITE EXECUTIVE SUMMARY HEADER (Rows 1-3)
 *************************************************************/
function writeExecutiveSummaryHeader_(sheet) {
  // Row 1: Title
  sheet.getRange(1, 1).setValue('Branch Executive Summaries');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Row 2: Instructions
  const instructions = 'This sheet shows executive summaries from all branch workbooks. Click "Refresh Executive Summaries" to update. Use filters to view specific months or branches.';
  sheet.getRange(2, 1).setValue(instructions);
  sheet.getRange(2, 1, 1, 10).merge();
  sheet.getRange(2, 1).setWrap(true).setFontSize(9).setFontStyle('italic');

  // Row 3: Blank (space for button)
}

/*************************************************************
 * FORMAT EXECUTIVE SUMMARY SHEET
 *************************************************************/
function formatExecutiveSummarySheet_(sheet, colCount, rowCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // Freeze header rows (1-4)
  sheet.setFrozenRows(4);

  // Font for entire sheet
  sheet.getDataRange().setFontFamily('Verdana').setFontSize(9);

  // Bold header row (row 4) - DON'T wrap yet, we'll do it manually
  sheet.getRange(4, 1, 1, colCount).setFontWeight('bold');

  // Manually set header text with line breaks for better control
  const headerWithBreaks = [
    'Branch',
    'Month',
    'Revenue\n(actual)',
    'Revenue\n(plan)',
    'Expenses\n(actual)',
    'Expenses\n(plan)',
    'Net\n(actual)',
    'Net\n(plan)',
    'Net\nvariance',
    'Executive Summary'
  ];
  sheet.getRange(4, 1, 1, headerWithBreaks.length).setValues([headerWithBreaks]);
  sheet.getRange(4, 1, 1, colCount).setWrap(true);

  // Add filter to data
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, colCount).createFilter();
  }

  // Format Branch column (A) as text to preserve leading zeros
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, 1).setNumberFormat('@');
  }

  // Format money columns (C, D, E, F, G, H, I) - columns 3-9
  if (lastRow >= 5) {
    const moneyFmt = '$#,##0.00';
    for (let col = 3; col <= 9; col++) {
      sheet.getRange(5, col, lastRow - 4, 1).setNumberFormat(moneyFmt);
    }
  }

  // Light dotted borders around all data cells
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, colCount).setBorder(
      true, true, true, true, true, true,  // all borders
      null,                                 // color
      SpreadsheetApp.BorderStyle.DOTTED     // dotted style
    );
  }

  // Thick vertical borders after columns B, D, F, I (columns 2, 4, 6, 9)
  const thickBorderCols = [2, 4, 6, 9]; // B, D, F, I
  thickBorderCols.forEach(col => {
    if (col > colCount) return;
    sheet.getRange(4, col, lastRow - 3, 1).setBorder(
      null,                                   // top
      null,                                   // left
      null,                                   // bottom
      true,                                   // right border ON
      null,                                   // vertical
      null,                                   // horizontal
      null,                                   // color
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM // medium thickness
    );
  });

  // Conditional formatting for Net Variance (column 9)
  if (lastRow >= 5) {
    const varRange = sheet.getRange(5, 9, lastRow - 4, 1);

    const rules = [];

    // Positive variance: light green background, dark green text
    const posRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#1a7f37')      // dark green text
      .setBackground('#e6f4ea')     // light green background
      .setRanges([varRange])
      .build();

    // Negative variance: light red background, dark red text
    const negRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#a61b1b')      // dark red text
      .setBackground('#fde8e8')     // light red background
      .setRanges([varRange])
      .build();

    rules.push(posRule, negRule);
    sheet.setConditionalFormatRules(rules);
  }

  // Set specific column widths BEFORE wrapping Executive Summary
  sheet.setColumnWidth(1, 70);   // Branch
  sheet.setColumnWidth(2, 90);   // Month
  sheet.setColumnWidth(3, 110);  // Revenue (actual)
  sheet.setColumnWidth(4, 110);  // Revenue (plan)
  sheet.setColumnWidth(5, 110);  // Expenses (actual)
  sheet.setColumnWidth(6, 110);  // Expenses (plan)
  sheet.setColumnWidth(7, 100);  // Net (actual)
  sheet.setColumnWidth(8, 100);  // Net (plan)
  sheet.setColumnWidth(9, 100);  // Net variance
  sheet.setColumnWidth(10, 500); // Executive Summary

  // Wrap Executive Summary column (J, column 10)
  sheet.getRange(1, 10, lastRow).setWrap(true);

  // Hide gridlines
  sheet.setHiddenGridlines(true);
}

/*************************************************************
 * REFRESH AR COLLECTIONS
 *************************************************************/
function refreshARCollections() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Starting AR Collections refresh...', 'Progress', -1);

  // Force flush to ensure all changes are saved before reading
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Small delay to ensure Drive sync

  const arSheet = upsertSheet_(ss, 'ARs (Collections)');

  ss.toast('Finding branch workbooks...', 'Progress', -1);
  const books = getBranchSpreadsheets_();

  ss.toast(`Processing ${books.length} branch workbook${books.length !== 1 ? 's' : ''}...`, 'Progress', -1);

  const rows = [];
  let processedCount = 0;
  const debugLog = []; // Track processing details

  books.forEach(book => {
    processedCount++;
    const bookName = book.getName();
    ss.toast(`Processing ${bookName} (${processedCount} of ${books.length})...`, 'Progress', -1);

    // Get all sheets in this workbook
    const sheets = book.getSheets();
    let foundKPISheets = 0;
    let totalARRows = 0;

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      // Look for sheets matching pattern "## Weekly KPIs"
      const match = sheetName.match(/^(\d{2})\s+Weekly KPIs$/i);
      if (!match) return;

      foundKPISheets++;
      const branch = match[1]; // Extract branch number

      // Find AR header row
      const arHeaderRow = findARHeaderRow_(sheet);
      if (!arHeaderRow) {
        debugLog.push(`${bookName} - ${sheetName}: No AR header found`);
        return; // No AR section found, skip
      }

      // Extract AR data
      const arData = extractARData_(sheet, arHeaderRow, branch);
      totalARRows += arData.length;
      rows.push(...arData);

      debugLog.push(`${bookName} - ${sheetName}: Found ${arData.length} AR records`);
    });

    if (foundKPISheets === 0) {
      debugLog.push(`${bookName}: No Weekly KPI sheets found`);
    }
  });

  ss.toast('Writing data to AR Collections sheet...', 'Progress', -1);

  if (rows.length === 0) {
    arSheet.clear();
    writeARHeader_(arSheet);
    ss.toast('No AR data found in branch workbooks.', 'Complete', 3);
    return;
  }

  // Sort by week (descending), then branch
  rows.sort((a, b) => {
    if (a.week !== b.week) return b.week - a.week; // Newest week first
    return a.branch.localeCompare(b.branch);
  });

  // Clear and write
  arSheet.clear();
  writeARHeader_(arSheet);

  const header = [
    'Week', 'Date', 'Branch', '0-30 Days', '31-60 Days', '61-90 Days', '91-120 Days', 'Over 120 Days', 'Total'
  ];

  // Write header at row 4
  arSheet.getRange(4, 1, 1, header.length).setValues([header]);

  // Write data starting at row 5
  const outRows = rows.map(r => [
    r.week,
    r.date,
    r.branch,
    r.days0_30,
    r.days31_60,
    r.days61_90,
    r.days91_120,
    r.daysOver120,
    r.total
  ]);

  arSheet.getRange(5, 1, outRows.length, header.length).setValues(outRows);

  ss.toast('Formatting AR Collections sheet...', 'Progress', -1);
  formatARSheet_(arSheet, header.length, rows.length);

  // Log debug information to console
  Logger.log('=== AR Collections Debug Log ===');
  debugLog.forEach(entry => Logger.log(entry));
  Logger.log(`Total: ${rows.length} AR records from ${books.length} workbooks`);

  ss.toast(`Complete! Processed ${rows.length} week records from ${books.length} branches.`, 'Success ✓', 5);
}

/*************************************************************
 * FIND AR HEADER ROW
 *************************************************************/
function findARHeaderRow_(sheet) {
  const maxCheck = Math.min(50, sheet.getLastRow());
  const maxCols = sheet.getLastColumn();

  for (let r = 1; r <= maxCheck; r++) {
    const row = sheet.getRange(r, 1, 1, maxCols).getDisplayValues()[0];

    // Look for all aging bucket headers in the same row
    const has0_30 = row.some(cell => String(cell).trim() === '0-30 Days');
    const has31_60 = row.some(cell => String(cell).trim() === '31-60 Days');
    const has61_90 = row.some(cell => String(cell).trim() === '61-90 Days');
    const has91_120 = row.some(cell => String(cell).trim() === '91-120 Days');
    const hasOver120 = row.some(cell => String(cell).trim() === 'Over 120 Days');

    if (has0_30 && has31_60 && has61_90 && has91_120 && hasOver120) {
      return r;
    }
  }

  return null;
}

/*************************************************************
 * EXTRACT AR DATA
 *************************************************************/
function extractARData_(sheet, headerRow, branch) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr <= headerRow) return [];

  const header = sheet.getRange(headerRow, 1, 1, lc).getDisplayValues()[0];

  // Find column indices for each aging bucket
  const idx0_30 = header.findIndex(cell => String(cell).trim() === '0-30 Days');
  const idx31_60 = header.findIndex(cell => String(cell).trim() === '31-60 Days');
  const idx61_90 = header.findIndex(cell => String(cell).trim() === '61-90 Days');
  const idx91_120 = header.findIndex(cell => String(cell).trim() === '91-120 Days');
  const idxOver120 = header.findIndex(cell => String(cell).trim() === 'Over 120 Days');

  if (idx0_30 === -1 || idx31_60 === -1 || idx61_90 === -1 || idx91_120 === -1 || idxOver120 === -1) {
    return []; // Missing required columns
  }

  // Get all rows below header
  const vals = sheet.getRange(headerRow + 1, 1, lr - headerRow, lc).getDisplayValues();
  const out = [];

  vals.forEach((row, i) => {
    // Column A should have week number
    const weekRaw = String(row[0] || '').trim();
    const week = parseInt(weekRaw);
    if (isNaN(week) || week <= 0) return; // Skip if not a valid week number

    // Column B should have date
    const date = String(row[1] || '').trim();
    if (!date) return; // Skip if no date

    // Extract aging bucket values
    const days0_30 = numFromCell_(row[idx0_30]);
    const days31_60 = numFromCell_(row[idx31_60]);
    const days61_90 = numFromCell_(row[idx61_90]);
    const days91_120 = numFromCell_(row[idx91_120]);
    const daysOver120 = numFromCell_(row[idxOver120]);

    // Calculate total
    const total = days0_30 + days31_60 + days61_90 + days91_120 + daysOver120;

    // Skip rows with no AR data
    if (total === 0) return;

    out.push({
      week: week,
      date: date,
      branch: "'" + branch.padStart(2, '0'),
      days0_30: days0_30,
      days31_60: days31_60,
      days61_90: days61_90,
      days91_120: days91_120,
      daysOver120: daysOver120,
      total: total
    });
  });

  return out;
}

/*************************************************************
 * WRITE AR HEADER (Rows 1-3)
 *************************************************************/
function writeARHeader_(sheet) {
  // Row 1: Title
  sheet.getRange(1, 1).setValue('AR Collections Rollup');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Row 2: Instructions
  const instructions = 'This sheet shows accounts receivable aging data from all branch workbooks. Click "Refresh ARs" to update. Use filters to view specific weeks or branches.';
  sheet.getRange(2, 1).setValue(instructions);
  sheet.getRange(2, 1, 1, 9).merge();
  sheet.getRange(2, 1).setWrap(true).setFontSize(9).setFontStyle('italic');

  // Row 3: Blank (space for notes/links)
}

/*************************************************************
 * FORMAT AR SHEET
 *************************************************************/
function formatARSheet_(sheet, colCount, rowCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // Freeze header rows (1-4)
  sheet.setFrozenRows(4);

  // Font for entire sheet - Verdana, size 9, lighter black
  sheet.getDataRange().setFontFamily('Verdana').setFontSize(9);
  sheet.getDataRange().setFontColor('#434343'); // Lighter shade of black

  // Bold header row (row 4)
  sheet.getRange(4, 1, 1, colCount).setFontWeight('bold');

  // Thick border under header row
  sheet.getRange(4, 1, 1, colCount).setBorder(
    false, false, true, false, false, false,
    '#999999', // Light border color
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // Add filter to data
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, colCount).createFilter();
  }

  // Format Branch column (C, column 3) as text
  if (lastRow >= 5) {
    sheet.getRange(5, 3, lastRow - 4, 1).setNumberFormat('@');
  }

  // Format money columns (D-I, columns 4-9)
  if (lastRow >= 5) {
    const moneyFmt = '$#,##0.00';
    for (let col = 4; col <= 9; col++) {
      sheet.getRange(5, col, lastRow - 4, 1).setNumberFormat(moneyFmt);
    }
  }

  // Light grey dotted borders between cells (data rows)
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, colCount).setBorder(
      true, true, true, true, true, true,
      '#d9d9d9', // Light grey
      SpreadsheetApp.BorderStyle.DOTTED
    );
  }

  // Thin border around whole table (header + data)
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, colCount).setBorder(
      true, true, true, true, false, false,
      '#999999', // Light border color
      SpreadsheetApp.BorderStyle.SOLID
    );
  }

  // Set column widths
  sheet.setColumnWidth(1, 60);   // Week
  sheet.setColumnWidth(2, 90);   // Date
  sheet.setColumnWidth(3, 70);   // Branch
  sheet.setColumnWidth(4, 100);  // 0-30 Days
  sheet.setColumnWidth(5, 100);  // 31-60 Days
  sheet.setColumnWidth(6, 100);  // 61-90 Days
  sheet.setColumnWidth(7, 100);  // 91-120 Days
  sheet.setColumnWidth(8, 110);  // Over 120 Days
  sheet.setColumnWidth(9, 100);  // Total

  // Hide gridlines
  sheet.setHiddenGridlines(true);
}

/*************************************************************
 * REFRESH MEMBERSHIP TRENDS
 *************************************************************/
function refreshMembershipTrends() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Starting Membership Trends refresh...', 'Progress', -1);

  SpreadsheetApp.flush();
  Utilities.sleep(500);

  const memSheet = upsertSheet_(ss, 'Membership Trends');

  ss.toast('Finding branch workbooks...', 'Progress', -1);
  const books = getBranchSpreadsheets_();

  ss.toast(`Processing ${books.length} branch workbook${books.length !== 1 ? 's' : ''}...`, 'Progress', -1);

  const rows = [];
  let processedCount = 0;
  const debugLog = [];

  books.forEach(book => {
    processedCount++;
    const bookName = book.getName();
    ss.toast(`Processing ${bookName} (${processedCount} of ${books.length})...`, 'Progress', -1);

    const sheets = book.getSheets();
    let foundKPISheets = 0;

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      // Look for sheets matching pattern "## Weekly KPIs"
      const match = sheetName.match(/^(\d{2})\s+Weekly KPIs$/i);
      if (!match) return;

      foundKPISheets++;
      const branch = match[1];

      // Check row 1 for "Membership Units" to identify membership branches
      if (!isMembershipKPISheet_(sheet)) {
        debugLog.push(`${bookName} - ${sheetName}: Not a membership KPI sheet, skipping`);
        return;
      }

      // Extract membership data
      const memData = extractMembershipData_(sheet, branch);
      rows.push(...memData);

      debugLog.push(`${bookName} - ${sheetName}: Found ${memData.length} membership records`);
    });

    if (foundKPISheets === 0) {
      debugLog.push(`${bookName}: No Weekly KPI sheets found`);
    }
  });

  ss.toast('Writing data to Membership Trends sheet...', 'Progress', -1);

  if (rows.length === 0) {
    memSheet.clear();
    writeMembershipHeader_(memSheet);
    ss.toast('No membership data found in branch workbooks.', 'Complete', 3);
    return;
  }

  // Sort by week (descending), then branch
  rows.sort((a, b) => {
    if (a.week !== b.week) return b.week - a.week;
    return a.branch.localeCompare(b.branch);
  });

  // Clear and write
  memSheet.clear();
  writeMembershipHeader_(memSheet);

  const header = [
    'Branch', 'Week', 'Date',
    'Membership Units', 'Joins', 'Renews', 'Terms',
    'Cumulative Net Joins', 'Retention'
  ];

  // Write header at row 4
  memSheet.getRange(4, 1, 1, header.length).setValues([header]);

  // Write data starting at row 5
  const outRows = rows.map(r => [
    r.branch,
    r.week,
    r.date,
    r.membershipUnits,
    r.joins,
    r.renews,
    r.terms,
    r.cumulativeNetJoins,
    r.retention
  ]);

  memSheet.getRange(5, 1, outRows.length, header.length).setValues(outRows);

  ss.toast('Formatting Membership Trends sheet...', 'Progress', -1);
  formatMembershipSheet_(memSheet, header.length, rows.length);

  // Log debug information to console
  Logger.log('=== Membership Trends Debug Log ===');
  debugLog.forEach(entry => Logger.log(entry));
  Logger.log(`Total: ${rows.length} membership records from ${books.length} workbooks`);

  ss.toast(`Complete! Processed ${rows.length} week records from ${books.length} branches.`, 'Success ✓', 5);
}

/*************************************************************
 * DETECT MEMBERSHIP KPI SHEET
 *************************************************************/
function isMembershipKPISheet_(sheet) {
  const lc = sheet.getLastColumn();
  if (lc < 1) return false;

  const row1 = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0];
  return row1.some(cell => String(cell).trim() === 'Membership Units');
}

/*************************************************************
 * EXTRACT MEMBERSHIP DATA
 *************************************************************/
function extractMembershipData_(sheet, branch) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 4) return []; // Data starts at row 4

  const row1 = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0];

  // Find column indices from row 1 headers
  const idxMemUnits = row1.findIndex(cell => String(cell).trim() === 'Membership Units');
  const idxJoins = row1.findIndex(cell => String(cell).trim() === 'Joins');
  const idxRenews = row1.findIndex(cell => String(cell).trim() === 'Renews');
  const idxTerms = row1.findIndex(cell => String(cell).trim() === 'Terms');
  const idxCumNetJoins = row1.findIndex(cell => String(cell).trim().startsWith('Cumulative'));
  const idxRetention = row1.findIndex(cell => String(cell).trim().startsWith('Retention'));

  if (idxMemUnits === -1 || idxJoins === -1 || idxTerms === -1) {
    return []; // Missing required columns
  }

  // Get all rows from row 4 onwards (row 3 is the "Week" header row)
  const vals = sheet.getRange(4, 1, lr - 3, lc).getDisplayValues();
  const out = [];

  vals.forEach(row => {
    // Column A has week number
    const weekRaw = String(row[0] || '').trim();
    const week = parseInt(weekRaw);
    if (isNaN(week) || week <= 0) return;

    // Column B has date
    const date = String(row[1] || '').trim();
    if (!date) return;

    const membershipUnits = numFromCell_(row[idxMemUnits]);
    const joins = numFromCell_(row[idxJoins]);
    const renews = idxRenews !== -1 ? numFromCell_(row[idxRenews]) : 0;
    const terms = numFromCell_(row[idxTerms]);
    const cumulativeNetJoins = idxCumNetJoins !== -1 ? numFromCell_(row[idxCumNetJoins]) : 0;
    const retention = idxRetention !== -1 ? pctFromCell_(row[idxRetention]) : 0;

    // Skip rows with no data entered
    if (membershipUnits === 0 && joins === 0 && terms === 0) return;

    out.push({
      branch: "'" + branch.padStart(2, '0'),
      week: week,
      date: date,
      membershipUnits: membershipUnits,
      joins: joins,
      renews: renews,
      terms: terms,
      cumulativeNetJoins: cumulativeNetJoins,
      retention: retention
    });
  });

  return out;
}

/*************************************************************
 * PERCENTAGE HELPER
 *************************************************************/
function pctFromCell_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const s = String(v).trim();
  if (s.indexOf('%') !== -1) {
    const n = Number(s.replace(/%/g, '').trim());
    return isNaN(n) ? 0 : n / 100;
  }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

/*************************************************************
 * WRITE MEMBERSHIP HEADER (Rows 1-3)
 *************************************************************/
function writeMembershipHeader_(sheet) {
  // Row 1: Title
  sheet.getRange(1, 1).setValue('Membership Trends');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Row 2: Instructions
  const instructions = 'This sheet shows weekly membership KPI data from all membership branch workbooks. Click "Refresh Membership" to update. Use filters to view specific weeks or branches.';
  sheet.getRange(2, 1).setValue(instructions);
  sheet.getRange(2, 1, 1, 9).merge();
  sheet.getRange(2, 1).setWrap(true).setFontSize(9).setFontStyle('italic');

  // Row 3: Blank (space for button)
}

/*************************************************************
 * FORMAT MEMBERSHIP SHEET
 *************************************************************/
function formatMembershipSheet_(sheet, colCount, rowCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // Freeze header rows (1-4)
  sheet.setFrozenRows(4);

  // Font for entire sheet
  sheet.getDataRange().setFontFamily('Verdana').setFontSize(9);
  sheet.getDataRange().setFontColor('#434343');

  // Bold header row (row 4)
  sheet.getRange(4, 1, 1, colCount).setFontWeight('bold').setWrap(true);

  // Thick border under header row
  sheet.getRange(4, 1, 1, colCount).setBorder(
    false, false, true, false, false, false,
    '#999999',
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // Add filter to data
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, colCount).createFilter();
  }

  // Format Branch column (A, column 1) as text
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, 1).setNumberFormat('@');
  }

  // Format number columns (D-H, columns 4-8) as whole numbers with commas
  if (lastRow >= 5) {
    const numFmt = '#,##0';
    for (let col = 4; col <= 8; col++) {
      sheet.getRange(5, col, lastRow - 4, 1).setNumberFormat(numFmt);
    }
  }

  // Format Retention column (I, column 9) as percentage
  if (lastRow >= 5) {
    sheet.getRange(5, 9, lastRow - 4, 1).setNumberFormat('0%');
  }

  // Light grey dotted borders between cells (data rows)
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, colCount).setBorder(
      true, true, true, true, true, true,
      '#d9d9d9',
      SpreadsheetApp.BorderStyle.DOTTED
    );
  }

  // Thin border around whole table (header + data)
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, colCount).setBorder(
      true, true, true, true, false, false,
      '#999999',
      SpreadsheetApp.BorderStyle.SOLID
    );
  }

  // Set column widths
  sheet.setColumnWidth(1, 70);   // Branch
  sheet.setColumnWidth(2, 60);   // Week
  sheet.setColumnWidth(3, 90);   // Date
  sheet.setColumnWidth(4, 140);  // Membership Units
  sheet.setColumnWidth(5, 70);   // Joins
  sheet.setColumnWidth(6, 80);   // Renews
  sheet.setColumnWidth(7, 70);   // Terms
  sheet.setColumnWidth(8, 160);  // Cumulative Net Joins
  sheet.setColumnWidth(9, 90);   // Retention

  // Hide gridlines
  sheet.setHiddenGridlines(true);
}
