// ============================================================
// GUMROAD SALES TRACKER
// Pulls your Gumroad sales into a Google Sheet automatically.
// Run fetchGumroadSales() manually or set a time trigger
// to run every hour/day.
//
// HOW TO SET UP:
//   1. script.google.com → New project
//   2. Paste this file
//   3. Add your Gumroad Access Token to GUMROAD_CONFIG below
//      (Gumroad → Settings → Advanced → Applications → Access Token)
//   4. Click Run → setupTracker  (creates the sheet + headers)
//   5. Click Run → fetchGumroadSales  (pulls all your sales)
//   6. Optional: Run → setDailyTrigger  (auto-refreshes every day)
// ============================================================

const GUMROAD_CONFIG = {
  accessToken: 'YOUR_GUMROAD_ACCESS_TOKEN_HERE',  // ← paste your token
  sheetName:   'Sales Log',
  dashSheet:   'Dashboard',
};

// ─────────────────────────────────────────────────────────────
// Entry points
// ─────────────────────────────────────────────────────────────

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
    || SpreadsheetApp.create('Gumroad Sales Tracker');

  _getOrCreateSheet(ss, GUMROAD_CONFIG.sheetName, _buildSalesHeaders);
  _getOrCreateSheet(ss, GUMROAD_CONFIG.dashSheet,  _buildDashboard);

  SpreadsheetApp.getUi().alert(
    '✅ Gumroad Sales Tracker ready!\n\n' +
    'Next: Run fetchGumroadSales() to pull your sales data.'
  );
}

function fetchGumroadSales() {
  if (GUMROAD_CONFIG.accessToken === 'YOUR_GUMROAD_ACCESS_TOKEN_HERE') {
    SpreadsheetApp.getUi().alert(
      '⚠️ Add your Gumroad Access Token to GUMROAD_CONFIG first.\n\n' +
      'Gumroad → Settings → Advanced → Applications → Access Token'
    );
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GUMROAD_CONFIG.sheetName)
    || _getOrCreateSheet(ss, GUMROAD_CONFIG.sheetName, _buildSalesHeaders);

  const sales = _fetchAllSales();
  if (!sales.length) {
    Logger.log('No sales found.');
    return;
  }

  _writeSales(sheet, sales);
  _refreshDashboard(ss, sales);

  Logger.log('✅ Pulled ' + sales.length + ' sales.');
  SpreadsheetApp.getUi().alert('✅ Done! Pulled ' + sales.length + ' sales.');
}

function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('fetchGumroadSales')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert('✅ Daily auto-refresh set for 8 AM.');
}

// ─────────────────────────────────────────────────────────────
// Gumroad API
// ─────────────────────────────────────────────────────────────

function _fetchAllSales() {
  const sales = [];
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    const url = 'https://api.gumroad.com/v2/sales?page=' + page + '&page_key=' + page;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + GUMROAD_CONFIG.accessToken },
      muteHttpExceptions: true,
    });

    const data = JSON.parse(response.getContentText());

    if (!data.success) {
      Logger.log('API error: ' + JSON.stringify(data));
      break;
    }

    const batch = data.sales || [];
    sales.push(...batch);

    hasMore = batch.length === 10;
    page++;

    if (page > 100) break; // safety limit
  }

  return sales;
}

// ─────────────────────────────────────────────────────────────
// Sheet writers
// ─────────────────────────────────────────────────────────────

function _buildSalesHeaders(sheet) {
  const headers = [
    'Sale ID', 'Date', 'Product Name', 'Buyer Email',
    'Price ($)', 'Fee ($)', 'Net ($)', 'Currency',
    'Refunded', 'Country', 'Referrer'
  ];
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers])
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 130);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 220);
}

function _writeSales(sheet, sales) {
  // Clear old data (keep headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 11).clearContent();

  const rows = sales.map(s => [
    s.id            || '',
    s.created_at    ? new Date(s.created_at) : '',
    s.product_name  || '',
    s.email         || '',
    _cents(s.price),
    _cents(s.fee),
    _cents(s.price) - _cents(s.fee),
    s.currency      || 'USD',
    s.refunded      ? 'Yes' : 'No',
    s.country       || '',
    s.referrer      || '',
  ]);

  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, 11).setValues(rows);

  // Format money columns
  sheet.getRange(2, 5, rows.length, 3).setNumberFormat('$0.00');

  // Format date column
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('yyyy-mm-dd');

  // Highlight refunded rows
  const cfRules = [];
  const refundedRange = sheet.getRange(2, 1, rows.length, 11);
  cfRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="Yes"')
      .setBackground('#fce8e6')
      .setFontColor('#c0392b')
      .setRanges([refundedRange])
      .build()
  );
  sheet.setConditionalFormatRules(cfRules);
}

// ─────────────────────────────────────────────────────────────
// Dashboard
// ─────────────────────────────────────────────────────────────

function _buildDashboard(sheet) {
  sheet.getRange('A1').setValue('📊 Gumroad Dashboard')
    .setFontSize(16).setFontWeight('bold')
    .setBackground('#1a1a2e').setFontColor('#ffffff');
  sheet.getRange('A1:E1').merge();
  sheet.setColumnWidths(1, 5, 160);
}

function _refreshDashboard(ss, sales) {
  const sheet = ss.getSheetByName(GUMROAD_CONFIG.dashSheet)
    || _getOrCreateSheet(ss, GUMROAD_CONFIG.dashSheet, _buildDashboard);

  sheet.clearContents();

  const C = { dark: '#1a1a2e', mid: '#0f3460', light: '#e8f0fe', white: '#ffffff', green: '#34a853' };

  // Title
  sheet.setRowHeight(1, 48);
  sheet.getRange('A1:E1').merge()
    .setValue('📊 Gumroad Sales Dashboard — Last updated: ' + new Date().toLocaleString())
    .setBackground(C.dark).setFontColor(C.white).setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Summary stats ──────────────────────────────────────────
  const notRefunded = sales.filter(s => !s.refunded);
  const totalRevenue   = notRefunded.reduce((sum, s) => sum + _cents(s.price), 0);
  const totalFees      = notRefunded.reduce((sum, s) => sum + _cents(s.fee),   0);
  const totalNet       = totalRevenue - totalFees;
  const totalSales     = notRefunded.length;
  const totalRefunds   = sales.filter(s => s.refunded).length;

  const stats = [
    ['Total Sales',       totalSales,           ''],
    ['Gross Revenue',     totalRevenue,          '$0.00'],
    ['Fees Paid',         totalFees,             '$0.00'],
    ['Net Revenue',       totalNet,              '$0.00'],
    ['Refunds',           totalRefunds,          ''],
  ];

  sheet.setRowHeight(3, 26);
  sheet.getRange('A3:E3').merge().setValue('SUMMARY')
    .setBackground(C.mid).setFontColor(C.white).setFontWeight('bold')
    .setHorizontalAlignment('center');

  stats.forEach(([label, value, fmt], i) => {
    const row = 4 + i;
    sheet.setRowHeight(row, 28);
    const bg = i % 2 === 0 ? C.white : C.light;
    sheet.getRange(row, 1).setValue(label).setFontWeight('bold').setBackground(bg);
    const cell = sheet.getRange(row, 2).setValue(value).setBackground(bg).setHorizontalAlignment('center');
    if (fmt) cell.setNumberFormat(fmt);
  });

  // ── Revenue by product ─────────────────────────────────────
  const byProduct = {};
  notRefunded.forEach(s => {
    const name = s.product_name || 'Unknown';
    if (!byProduct[name]) byProduct[name] = { count: 0, gross: 0, net: 0 };
    byProduct[name].count++;
    byProduct[name].gross += _cents(s.price);
    byProduct[name].net   += _cents(s.price) - _cents(s.fee);
  });

  const productRows = Object.entries(byProduct)
    .sort((a, b) => b[1].net - a[1].net);

  const startRow = 11;
  sheet.setRowHeight(startRow - 1, 26);
  sheet.getRange(startRow - 1, 1, 1, 4).merge()
    .setValue('REVENUE BY PRODUCT')
    .setBackground(C.mid).setFontColor(C.white).setFontWeight('bold')
    .setHorizontalAlignment('center');

  ['Product', 'Units Sold', 'Gross ($)', 'Net ($)'].forEach((h, ci) => {
    sheet.getRange(startRow, ci + 1)
      .setValue(h).setFontWeight('bold').setBackground(C.mid).setFontColor(C.white)
      .setHorizontalAlignment('center');
  });

  productRows.forEach(([name, data], i) => {
    const row = startRow + 1 + i;
    const bg  = i % 2 === 0 ? C.white : C.light;
    sheet.getRange(row, 1).setValue(name).setBackground(bg);
    sheet.getRange(row, 2).setValue(data.count).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 3).setValue(data.gross).setNumberFormat('$0.00').setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 4).setValue(data.net).setNumberFormat('$0.00').setBackground(bg).setHorizontalAlignment('center').setFontWeight('bold');
  });

  sheet.setColumnWidths(1, 4, 180);
  sheet.hideGridlines();
}

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────

function _cents(val) {
  if (!val) return 0;
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n / 100;
}

function _getOrCreateSheet(ss, name, setupFn) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (setupFn) setupFn(sheet);
  }
  return sheet;
}
