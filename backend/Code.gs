/**
 * Spending Tracker — Apps Script backend
 * ---------------------------------------
 * Deploy this as a Web App (Deploy > New deployment > type: Web app).
 *   Execute as: Me
 *   Who has access: Anyone
 * (It's a private URL only you know — "Anyone" just means the phone browser
 * doesn't need to be logged into a Google session for the fetch() call to work.
 * Nobody can do anything without the URL, and the URL isn't guessable.)
 *
 * After deploying, copy the Web App URL into frontend/index.html (APP_SCRIPT_URL).
 *
 * ONE-TIME SETUP: run the `setup` function once from the Apps Script editor
 * (select it in the function dropdown, click Run). It will:
 *   - find your "CURRENT Monthly Spending Template" file
 *   - create a "Shared Expense Ledger" spreadsheet if one doesn't exist
 * You'll be asked to authorize the script the first time — that's normal.
 */

const TEMPLATE_NAME = 'CURRENT Monthly Spending Template';
const LEDGER_NAME = 'Shared Expense Ledger';
const LEDGER_HEADERS = ['ID', 'Date', 'Description', 'Amount Paid', 'Reimbursed So Far', 'Status', 'Who', 'Notes'];
const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// ---------- Setup ----------

function setup() {
  const templateId = findFileIdByName_(TEMPLATE_NAME);
  if (!templateId) {
    throw new Error('Could not find a file named "' + TEMPLATE_NAME + '" in your Drive. Rename your template to match, or edit TEMPLATE_NAME at the top of Code.gs.');
  }
  const ledgerId = getOrCreateLedger_();
  Logger.log('Template found: ' + templateId);
  Logger.log('Ledger ready: ' + ledgerId);
  Logger.log('Setup complete.');
}

function getOrCreateLedger_() {
  let id = findFileIdByName_(LEDGER_NAME);
  if (id) return id;
  const ss = SpreadsheetApp.create(LEDGER_NAME);
  const sheet = ss.getSheets()[0];
  sheet.setName('Ledger');
  sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return ss.getId();
}

function findFileIdByName_(name) {
  const it = DriveApp.getFilesByName(name);
  if (it.hasNext()) return it.next().getId();
  return null;
}

// ---------- Web app entry points ----------

function doPost(e) {
  return handleRequest_(e);
}
function doGet(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter;
    }
    const action = params.action;
    const fn = ACTIONS[action];
    if (!fn) throw new Error('Unknown action: ' + action);
    const result = fn(params);
    return jsonOut_({ ok: true, data: result });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  ping: () => ({ pong: true }),
  suggestDestinations: (p) => suggestDestinations(p.date),
  getSheetOptions: (p) => getCategoriesAndAccounts_(p.destinationId),
  createDestination: (p) => createDestination(p.name),
  addCategory: (p) => addCategory(p.destinationId, p.categoryName, !!p.alsoTemplate),
  checkDuplicate: (p) => checkDuplicate(p.destinationIds, p.date, p.amount, p.description),
  addEntry: (p) => addEntry(p),
  getOpenLedgerEntries: () => getOpenLedgerEntries(),
  settleLedgerEntry: (p) => settleLedgerEntry(p.ledgerId, p.settled),
};

// ---------- Destinations ----------

function expectedMonthlyName_(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear();
}

function suggestDestinations(dateStr) {
  const autoName = expectedMonthlyName_(dateStr);
  const autoId = findFileIdByName_(autoName);

  // Recent spreadsheets across Drive (excluding template/ledger/config), for
  // picking seasonal/trip sheets as extra or override destinations.
  const templateId = findFileIdByName_(TEMPLATE_NAME);
  const ledgerId = findFileIdByName_(LEDGER_NAME);
  const files = [];
  const it = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    const f = it.next();
    const id = f.getId();
    if (id === templateId || id === ledgerId) continue;
    files.push({ id: id, name: f.getName(), updated: f.getLastUpdated().getTime() });
  }
  files.sort((a, b) => b.updated - a.updated);
  const recent = files.slice(0, 12);

  return {
    autoSuggested: autoId ? { id: autoId, name: autoName } : { id: null, name: autoName, missing: true },
    recent: recent,
  };
}

function createDestination(name) {
  const templateId = findFileIdByName_(TEMPLATE_NAME);
  if (!templateId) throw new Error('Template file not found — run setup() first.');
  const existing = findFileIdByName_(name);
  if (existing) return { id: existing, name: name, alreadyExisted: true };
  const copy = DriveApp.getFileById(templateId).makeCopy(name);
  return { id: copy.getId(), name: name };
}

// ---------- Categories & accounts ----------

/**
 * Scans a spreadsheet's Summary-ish sheet for the Expenses category table
 * and the account totals table, generically (by header text, not fixed
 * coordinates) so it adapts if your layout shifts slightly.
 */
function getCategoriesAndAccounts_(destinationId) {
  const ss = SpreadsheetApp.openById(destinationId);
  const sheet = findSheetContaining_(ss, ['Totals', 'Planned', 'Actual']);
  if (!sheet) return { categories: [], accounts: [] };
  const data = sheet.sheet.getDataRange().getValues();

  const categories = extractLabelColumn_(data, 'Groceries') || extractAnyLabelledBlock_(data, 'Totals', 'Planned');
  const accounts = extractLabelColumn_(data, 'Chase') || [];

  return { categories: categories, accounts: accounts };
}

// Finds a block of category-style rows: a column of text labels sitting to
// the left of numeric Planned/Actual/Diff columns, starting after a row
// whose first non-empty cell in that column reads "Totals".
function extractAnyLabelledBlock_(data, totalsMarker, plannedHeader) {
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim() === totalsMarker) {
        const labels = [];
        for (let rr = r + 1; rr < data.length; rr++) {
          const label = data[rr][c + 1] !== undefined ? String(data[rr][c + 1]).trim() : '';
          if (!label) break;
          labels.push(label);
        }
        if (labels.length) return labels;
      }
    }
  }
  return [];
}

// Fallback: find the column containing a known sample label (e.g.
// "Groceries" for categories, "Chase" for an account name) and collect the
// contiguous non-empty text cells around it in that column.
function extractLabelColumn_(data, sampleContains) {
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const v = String(data[r][c]);
      if (v.indexOf(sampleContains) !== -1) {
        // walk up and down from this row in this column collecting labels
        const labels = [];
        let rr = r;
        while (rr >= 0 && String(data[rr][c]).trim() !== '' && String(data[rr][c]).trim() !== 'Totals') {
          labels.unshift(String(data[rr][c]).trim());
          rr--;
        }
        rr = r + 1;
        while (rr < data.length && String(data[rr][c]).trim() !== '') {
          labels.push(String(data[rr][c]).trim());
          rr++;
        }
        return labels;
      }
    }
  }
  return null;
}

function findSheetContaining_(ss, headerTexts) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const data = sheet.getDataRange().getValues();
    for (let r = 0; r < Math.min(data.length, 40); r++) {
      const row = data[r].map(String);
      if (headerTexts.every((h) => row.some((cell) => cell.trim() === h))) {
        return { sheet: sheet, headerRow: r };
      }
    }
  }
  return null;
}

function addCategory(destinationId, categoryName, alsoTemplate) {
  addCategoryToSpreadsheet_(destinationId, categoryName);
  if (alsoTemplate) {
    const templateId = findFileIdByName_(TEMPLATE_NAME);
    if (templateId && templateId !== destinationId) {
      addCategoryToSpreadsheet_(templateId, categoryName);
    }
  }
  return { added: categoryName };
}

function addCategoryToSpreadsheet_(spreadsheetId, categoryName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const found = findSheetContaining_(ss, ['Totals', 'Planned', 'Actual']);
  if (!found) throw new Error('Could not find the category table in this spreadsheet — add "' + categoryName + '" manually.');
  const sheet = found.sheet;
  const data = sheet.getDataRange().getValues();
  // Locate the "Totals" anchor for the Expenses block specifically (left-most one).
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim() === 'Totals') {
        // Find the last populated row of this category list (first blank in c+1 after r).
        let lastRow = r;
        for (let rr = r + 1; rr < data.length; rr++) {
          const label = data[rr][c + 1] !== undefined ? String(data[rr][c + 1]).trim() : '';
          if (!label) break;
          lastRow = rr;
        }
        // Insert a new row right after the last category row, copy formatting from it.
        sheet.insertRowAfter(lastRow + 1);
        const srcRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());
        const destRange = sheet.getRange(lastRow + 2, 1, 1, sheet.getLastColumn());
        srcRange.copyTo(destRange);
        sheet.getRange(lastRow + 2, c + 2).setValue(categoryName);
        sheet.getRange(lastRow + 2, c + 3).setValue(0); // Planned = 0
        return;
      }
    }
  }
  throw new Error('Could not locate the Expenses category table — add "' + categoryName + '" manually.');
}

// ---------- Transactions ----------

/**
 * Finds the transaction-log tab (not the Summary tab) and returns where the
 * Expense block and Income block headers are, by scanning for "Date" /
 * "Amount" / "Description" sequences.
 */
function findTransactionBlocks_(ss) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const data = sheet.getDataRange().getValues();
    for (let r = 0; r < Math.min(data.length, 20); r++) {
      const row = data[r].map((v) => String(v).trim());
      for (let c = 0; c < row.length; c++) {
        if (row[c] === 'Date' && row[c + 1] === 'Amount' && row[c + 2] === 'Description') {
          const hasCategory = row[c + 3] === 'Category';
          const hasAccount = row[c + 4] === 'Account';
          const block = {
            sheet: sheet,
            headerRow: r,
            dateCol: c,
            amountCol: c + 1,
            descCol: c + 2,
            categoryCol: hasCategory ? c + 3 : null,
            accountCol: hasAccount ? c + 4 : null,
            type: hasAccount ? 'expense' : 'income',
          };
          if (block.type === 'expense') return { sheet: sheet, expense: block, income: findIncomeBlockNear_(data, r, sheet) };
        }
      }
    }
  }
  return null;
}

function findIncomeBlockNear_(data, headerRow, sheet) {
  const row = data[headerRow].map((v) => String(v).trim());
  for (let c = 0; c < row.length; c++) {
    if (row[c] === 'Date' && row[c + 1] === 'Amount' && row[c + 2] === 'Description' && row[c + 4] !== 'Account') {
      return { sheet: sheet, headerRow: headerRow, dateCol: c, amountCol: c + 1, descCol: c + 2, categoryCol: row[c + 3] === 'Category' ? c + 3 : null, accountCol: null, type: 'income' };
    }
  }
  return null;
}

function lastRowInColumn_(sheet, col, headerRow) {
  const values = sheet.getRange(headerRow + 2, col + 1, sheet.getMaxRows() - headerRow - 1, 1).getValues();
  let last = headerRow; // 0-indexed header row; data starts headerRow+1 (0-indexed) => headerRow+2 in 1-indexed
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() !== '') last = headerRow + 1 + i;
  }
  return last; // 0-indexed row of last populated data row (or header row if empty)
}

function appendTransaction_(destinationId, block, isExpense, entry, tagId) {
  const b = isExpense ? block.expense : block.income;
  if (!b) throw new Error('Could not find the ' + (isExpense ? 'expense' : 'income') + ' transaction table in this spreadsheet.');
  const lastRow0 = lastRowInColumn_(b.sheet, b.dateCol, b.headerRow);
  const writeRow0 = lastRow0 + 1; // 0-indexed
  const writeRow1 = writeRow0 + 1; // 1-indexed for getRange

  b.sheet.getRange(writeRow1, b.dateCol + 1).setValue(new Date(entry.date + 'T00:00:00'));
  b.sheet.getRange(writeRow1, b.amountCol + 1).setValue(entry.amount);
  b.sheet.getRange(writeRow1, b.descCol + 1).setValue(entry.description);
  if (b.categoryCol !== null) b.sheet.getRange(writeRow1, b.categoryCol + 1).setValue(entry.category || '');
  if (b.accountCol !== null) b.sheet.getRange(writeRow1, b.accountCol + 1).setValue(entry.account || '');

  // Tag column: first column of the sheet, same row — used to hold the
  // ledger ID for shared-expense entries (successor to the old letter code).
  if (tagId) {
    b.sheet.getRange(writeRow1, 1).setValue(tagId);
  }
  return { row: writeRow1, sheetName: b.sheet.getName() };
}

function checkDuplicate(destinationIds, date, amount, description) {
  const matches = [];
  destinationIds.forEach((id) => {
    const ss = SpreadsheetApp.openById(id);
    const blocks = findTransactionBlocks_(ss);
    if (!blocks) return;
    [blocks.expense, blocks.income].forEach((b) => {
      if (!b) return;
      const lastRow0 = lastRowInColumn_(b.sheet, b.dateCol, b.headerRow);
      if (lastRow0 <= b.headerRow) return;
      const numRows = lastRow0 - b.headerRow;
      const data = b.sheet.getRange(b.headerRow + 2, b.dateCol + 1, numRows, b.descCol - b.dateCol + 1).getValues();
      const targetDate = new Date(date + 'T00:00:00').getTime();
      data.forEach((row) => {
        const rowDate = row[0] instanceof Date ? row[0].getTime() : null;
        const rowAmount = Number(row[1]);
        const rowDesc = String(row[2] || '').toLowerCase();
        if (rowDate === null) return;
        const dayGap = Math.abs(rowDate - targetDate) / 86400000;
        const amountMatch = Math.abs(rowAmount - Number(amount)) < 0.01;
        const descMatch = rowDesc && description && (rowDesc.indexOf(description.toLowerCase()) !== -1 || description.toLowerCase().indexOf(rowDesc) !== -1);
        if (dayGap <= 2 && amountMatch && descMatch) {
          matches.push({ spreadsheetId: id, date: row[0], amount: rowAmount, description: row[2] });
        }
      });
    });
  });
  return { duplicates: matches };
}

function addEntry(p) {
  // p: { date, amount, description, category, account, isExpense (bool),
  //      destinationIds: [id,...], sharedTag: null | {mode:'open', who}
  //                                          | {mode:'payback', ledgerId} }
  const results = [];
  let tagId = null;

  if (p.sharedTag && p.sharedTag.mode === 'open') {
    tagId = createLedgerEntry_(p.date, p.description, p.amount, p.sharedTag.who || '');
  } else if (p.sharedTag && p.sharedTag.mode === 'payback') {
    tagId = p.sharedTag.ledgerId;
  }

  p.destinationIds.forEach((id) => {
    const ss = SpreadsheetApp.openById(id);
    const blocks = findTransactionBlocks_(ss);
    if (!blocks) throw new Error('Could not find the transactions table in spreadsheet ' + id);
    const r = appendTransaction_(id, blocks, p.isExpense, p, tagId);
    results.push({ destinationId: id, row: r.row, sheetName: r.sheetName });
  });

  if (p.sharedTag && p.sharedTag.mode === 'payback') {
    recordReimbursement_(p.sharedTag.ledgerId, p.amount, p.date, p.sharedTag.who || '');
  }

  return { written: results, ledgerId: tagId };
}

// ---------- Shared Expense Ledger ----------

function ledgerSheet_() {
  const id = getOrCreateLedger_();
  return SpreadsheetApp.openById(id).getSheetByName('Ledger');
}

function createLedgerEntry_(date, description, amount, who) {
  const sheet = ledgerSheet_();
  const id = Utilities.formatDate(new Date(date + 'T00:00:00'), Session.getScriptTimeZone(), 'M/d') + '-' + description.substring(0, 12).replace(/\s+/g, '').toLowerCase();
  let finalId = id;
  let n = 2;
  const existingIds = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues().flat();
  while (existingIds.indexOf(finalId) !== -1) {
    finalId = id + '-' + n;
    n++;
  }
  sheet.appendRow([finalId, date, description, amount, 0, 'Open', who, '']);
  return finalId;
}

function recordReimbursement_(ledgerId, amount, date, who) {
  const sheet = ledgerSheet_();
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === ledgerId) {
      const newReimbursed = Number(data[r][4] || 0) + Number(amount);
      sheet.getRange(r + 1, 5).setValue(newReimbursed);
      const note = data[r][7] ? data[r][7] + '; ' : '';
      sheet.getRange(r + 1, 8).setValue(note + date + ' $' + amount + ' from ' + (who || '?'));
      return;
    }
  }
  throw new Error('Ledger entry not found: ' + ledgerId);
}

function getOpenLedgerEntries() {
  const sheet = ledgerSheet_();
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][5]).trim() === 'Open') {
      rows.push({
        id: data[r][0],
        date: data[r][1],
        description: data[r][2],
        amountPaid: data[r][3],
        reimbursedSoFar: data[r][4],
        who: data[r][6],
      });
    }
  }
  return { entries: rows };
}

function settleLedgerEntry(ledgerId, settled) {
  const sheet = ledgerSheet_();
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === ledgerId) {
      sheet.getRange(r + 1, 6).setValue(settled ? 'Settled' : 'Open');
      return { ok: true };
    }
  }
  throw new Error('Ledger entry not found: ' + ledgerId);
}
