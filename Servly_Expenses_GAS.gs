/**
 * ============================================================
 * SERVLY EXPENSES — Google Apps Script
 * ============================================================
 *
 * WHAT THIS DOES:
 *   1. Receives orders from your sales page
 *   2. Logs every order to your Orders sheet
 *   3. Emails YOU a notification immediately
 *   4. Emails THE CLIENT their files automatically
 *      (attaches the right files for their plan from Drive)
 *
 * ONE-TIME SETUP — follow every step carefully:
 *
 * STEP 1 — Upload your files to Google Drive
 *   Create a folder in Drive called "Servly Expenses Files"
 *   Upload these files to it:
 *     - Servly_Expenses.html
 *     - Setup_Guide.html
 *     - Servly_Expenses_GAS.gs  (Business plan only)
 *   For each file: right-click → Get link → copy the file ID
 *   (the ID is the long string in the URL between /d/ and /view)
 *
 * STEP 2 — Paste your file IDs below
 *
 * STEP 3 — Set your email address below
 *
 * STEP 4 — Deploy as Web App
 *   Extensions → Apps Script → Deploy → New deployment
 *   Type: Web app
 *   Execute as: Me
 *   Who has access: Anyone
 *   Click Deploy → authorize → copy the URL
 *
 * STEP 5 — Paste the URL into sales.html
 *   Find: var fromFile = '';
 *   Replace: var fromFile = 'YOUR_URL_HERE';
 *   Re-upload sales.html to GitHub
 *
 * ============================================================
 */

// ── ✏️  YOUR SETTINGS — edit these ──────────────────────────

// Your Gmail address — you'll get a notification for every order
var NOTIFY_EMAIL = 'YOUR_EMAIL@gmail.com';

// Your business name — shown in emails to clients
var BUSINESS_NAME = 'Servly by Marianne Marave';

// Your Facebook Messenger link — shown in client email
var MESSENGER_URL = 'https://m.me/mclubcreo';

// Google Drive file IDs for your deliverable files
// How to get a file ID: open the file in Drive → copy the URL
// The ID is between /d/ and /view  e.g. /d/THIS_PART/view
var FILE_IDS = {
  app:       'PASTE_APP_FILE_ID_HERE',       // Servly_Expenses.html
  guide:     'PASTE_GUIDE_FILE_ID_HERE',     // Setup_Guide.html
  gas:       'PASTE_GAS_FILE_ID_HERE',       // Servly_Expenses_GAS.gs (Business only)
};

// File names as the client will receive them
var FILE_NAMES = {
  app:   'Servly_Expenses.html',
  guide: 'Servly_Expenses_Setup_Guide.html',
  gas:   'Servly_Expenses_GAS.gs',
};

// What to attach per plan
var PLAN_FILES = {
  basic:    ['app', 'guide'],
  pro:      ['app', 'guide'],
  business: ['app', 'guide', 'gas'],
};

// Plan display names
var PLAN_NAMES = {
  basic:    'Basic',
  pro:      'Professional',
  business: 'Business',
};

// Promo prices
var PLAN_PRICES = {
  basic:    299,
  pro:      599,
  business: 999,
};

// ── Sheet tab names ──────────────────────────────────────────
var SHEET_ORDERS   = 'Orders';
var SHEET_EXPENSES = 'Expenses';
var SHEET_SUMMARY  = 'Monthly Summary';
var SHEET_CATEGORIES = 'Categories';

// ============================================================
//  WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  return _jsonResponse({ status: 'ok', service: 'Servly Expenses GAS v2' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'newOrder':        return _handleOrder(body);
      case 'syncExpenses':    return _syncExpenses(body);
      case 'appendExpense':   return _appendExpense(body);
      case 'getExpenses':     return _getExpenses(body);
      case 'deleteExpense':   return _deleteExpense(body);
      case 'generateSummary': return _generateMonthlySummary(body);
      default: return _jsonResponse({ success: false, error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return _jsonResponse({ success: false, error: err.message });
  }
}

// ============================================================
//  ORDER HANDLER
//  1. Log to Orders sheet
//  2. Email YOU a notification
//  3. Email CLIENT their files
// ============================================================
function _handleOrder(body) {
  var tier   = (body.tier || 'basic').toLowerCase();
  var amount = PLAN_PRICES[tier] || body.amount || 0;
  var name   = body.name   || 'Client';
  var email  = body.email  || '';
  var ref    = body.ref    || '';

  // ── 1. Log to Orders sheet ──────────────────────────────
  _logOrder(body, tier, amount);

  // ── 2. Email YOU ────────────────────────────────────────
  _notifyOwner(body, tier, amount);

  // ── 3. Email CLIENT their files ─────────────────────────
  if (email) {
    _sendFilesToClient(body, tier, amount);
  }

  return _jsonResponse({
    success: true,
    ref:     ref,
    message: 'Order logged. Notification sent to ' + NOTIFY_EMAIL + '. Files sent to ' + email + '.'
  });
}

// ── Log order row to sheet ───────────────────────────────────
function _logOrder(body, tier, amount) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _getOrCreateSheet(ss, SHEET_ORDERS);

  if (sheet.getLastRow() < 1) {
    var h = ['Ref No.','Date','Plan','Amount (₱)','Name','Email',
             'Business','Branches','Payment Method','Payment Ref','Notes',
             'Files Sent?','Status'];
    sheet.getRange(1,1,1,h.length).setValues([h]);
    _styleHeader(sheet, h.length, '#c8824a', '#ffffff');
    sheet.setFrozenRows(1);
  }

  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sheet.appendRow([
    body.ref || '', dateStr,
    PLAN_NAMES[tier] || tier.toUpperCase(),
    amount,
    body.name || '', body.email || '', body.biz || '', body.branches || '',
    body.payment || '', body.paymentRef || '', body.notes || '',
    body.email ? 'Sent ✓' : 'No email',
    'Pending verification'
  ]);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 4).setNumberFormat('₱#,##0.00');
  SpreadsheetApp.flush();
}

// ── Email YOU a notification ─────────────────────────────────
function _notifyOwner(body, tier, amount) {
  try {
    var subject = '🛒 New Order — ' + (PLAN_NAMES[tier]||tier) + ' — ₱' + amount;
    var html =
      '<div style="font-family:sans-serif;max-width:520px;color:#1a2626;">' +
      '<div style="background:#1e2d2d;padding:20px 24px;border-radius:10px 10px 0 0;">' +
      '<h2 style="color:#fff;margin:0;font-size:18px;">🛒 New Servly Order</h2>' +
      '<p style="color:rgba(255,255,255,.5);font-size:12px;margin:4px 0 0;">' + _dateStr() + '</p>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #e8e0d4;border-top:none;border-radius:0 0 10px 10px;padding:20px 24px;">' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
      _eRow('Order ref',     body.ref || '—', true) +
      _eRow('Plan',          PLAN_NAMES[tier] || tier, true) +
      _eRow('Amount',        '₱' + amount) +
      _eRow('Name',          body.name || '—') +
      _eRow('Email',         body.email || '—') +
      _eRow('Business',      body.biz || '—') +
      _eRow('Payment via',   body.payment || '—') +
      _eRow('Reference no.', body.paymentRef || '—', true) +
      _eRow('Notes',         body.notes || '—') +
      '</table>' +
      '<div style="margin-top:18px;background:#f5f0e8;border-radius:8px;padding:14px 16px;font-size:13px;">' +
      '<p style="margin:0 0 6px;font-weight:600;">✅ Files already sent to client automatically.</p>' +
      '<p style="margin:0;color:#888;">Verify payment in GCash/Maya matches reference <strong>' + (body.paymentRef||'—') + '</strong>.<br>' +
      'Then update status in your <a href="https://docs.google.com/spreadsheets" style="color:#c8824a;">Orders sheet</a> to "Verified".</p>' +
      '</div></div></div>';

    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, htmlBody: html });
  } catch(err) {
    Logger.log('Owner email failed: ' + err.message);
  }
}

// ── Email CLIENT their files ─────────────────────────────────
function _sendFilesToClient(body, tier, amount) {
  var email     = body.email || '';
  var firstName = (body.name || 'there').split(' ')[0];
  var planName  = PLAN_NAMES[tier] || tier;
  var fileKeys  = PLAN_FILES[tier] || PLAN_FILES['basic'];

  // ── Build file attachments from Drive ──
  var attachments = [];
  var attachedNames = [];
  var missingFiles  = [];

  fileKeys.forEach(function(key) {
    var fileId   = FILE_IDS[key];
    var fileName = FILE_NAMES[key];
    if (!fileId || fileId.indexOf('PASTE') !== -1) {
      missingFiles.push(fileName);
      return;
    }
    try {
      var file = DriveApp.getFileById(fileId);
      attachments.push({
        fileName:    fileName,
        content:     file.getBlob().getBytes(),
        mimeType:    file.getMimeType() || 'application/octet-stream',
      });
      attachedNames.push(fileName);
    } catch(err) {
      Logger.log('File fetch failed for ' + key + ': ' + err.message);
      missingFiles.push(fileName);
    }
  });

  // ── Build client email HTML ──
  var filesListHtml = attachedNames.map(function(f){
    return '<li style="padding:4px 0;">📎 ' + f + '</li>';
  }).join('');

  var missingNote = missingFiles.length > 0
    ? '<p style="color:#c0544a;font-size:12px;">⚠ Some files could not be attached (' + missingFiles.join(', ') + '). Please reply to this email and we will send them manually.</p>'
    : '';

  var subject = '✅ Your Servly Expenses ' + planName + ' files are here!';
  var html =
    '<div style="font-family:sans-serif;max-width:560px;color:#1a2626;">' +

    // Header
    '<div style="background:#1e2d2d;padding:24px 28px;border-radius:12px 12px 0 0;">' +
    '<h1 style="color:#fff;font-size:22px;margin:0 0 4px;">Hi ' + firstName + '! 🎉</h1>' +
    '<p style="color:rgba(255,255,255,.5);font-size:13px;margin:0;">Your Servly Expenses files are ready.</p>' +
    '</div>' +

    // Body
    '<div style="background:#fff;border:1px solid #e8e0d4;border-top:none;padding:24px 28px;">' +

    '<p style="font-size:14px;line-height:1.8;margin:0 0 20px;">Thank you for purchasing <strong>Servly Expenses ' + planName + '</strong>! ' +
    'Your files are attached to this email. Here\'s what you received:</p>' +

    '<ul style="font-size:14px;line-height:1.8;margin:0 0 20px;padding-left:16px;">' +
    filesListHtml +
    '</ul>' +

    missingNote +

    // Order summary box
    '<div style="background:#f5f0e8;border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;">' +
    '<p style="font-weight:600;margin:0 0 8px;font-size:14px;">📋 Order summary</p>' +
    '<table style="width:100%;font-size:13px;border-collapse:collapse;">' +
    _eRow('Plan',      planName) +
    _eRow('Amount',    '₱' + amount) +
    _eRow('Order ref', body.ref || '—') +
    '</table>' +
    '</div>' +

    // Getting started steps
    '<div style="border-left:3px solid #c8824a;padding-left:16px;margin-bottom:20px;">' +
    '<p style="font-weight:600;margin:0 0 8px;font-size:14px;">🚀 Getting started</p>' +
    '<ol style="font-size:13px;line-height:1.9;margin:0;padding-left:16px;">' +
    '<li>Save <strong>Servly_Expenses.html</strong> to your device</li>' +
    '<li>Double-click it — it opens in your browser instantly</li>' +
    '<li>On your phone: tap Share → <strong>Add to Home Screen</strong></li>' +
    '<li>Open the Setup Guide if you need step-by-step help</li>' +
    '</ol>' +
    '</div>' +

    // Plan-specific note
    _planNote(tier) +

    // Support
    '<p style="font-size:13px;line-height:1.8;margin:0;">' +
    'Questions or need help setting up? Reply to this email or message us on ' +
    '<a href="' + MESSENGER_URL + '" style="color:#c8824a;font-weight:500;">Facebook Messenger</a>.' +
    (tier === 'pro' ? ' You have <strong>14-day email support</strong> included.' : '') +
    (tier === 'business' ? ' You have <strong>30-day priority support</strong> included.' : '') +
    '</p>' +

    '</div>' +

    // Footer
    '<div style="background:#f5f0e8;border-radius:0 0 12px 12px;padding:14px 24px;text-align:center;">' +
    '<p style="font-size:11px;color:#888;margin:0;">' + BUSINESS_NAME + ' &nbsp;·&nbsp; <a href="' + MESSENGER_URL + '" style="color:#c8824a;">' + MESSENGER_URL + '</a></p>' +
    '</div>' +
    '</div>';

  try {
    var emailOptions = {
      to:       email,
      subject:  subject,
      htmlBody: html,
      name:     BUSINESS_NAME,
      replyTo:  NOTIFY_EMAIL,
    };

    if (attachments.length > 0) {
      emailOptions.attachments = attachments.map(function(a){
        return Utilities.newBlob(a.content, a.mimeType, a.fileName);
      });
    }

    MailApp.sendEmail(emailOptions);
    Logger.log('Client email sent to ' + email + ' with ' + attachments.length + ' file(s)');
  } catch(err) {
    Logger.log('Client email failed: ' + err.message);
    // Fallback — notify owner that manual send is needed
    try {
      MailApp.sendEmail({
        to:      NOTIFY_EMAIL,
        subject: '⚠ Manual delivery needed — ' + email,
        body:    'Automatic file delivery failed for order ' + (body.ref||'') + '.\nError: ' + err.message + '\nPlease send files manually to: ' + email
      });
    } catch(e2) {}
  }
}

// ── Plan-specific email note ─────────────────────────────────
function _planNote(tier) {
  if (tier === 'business') {
    return '<div style="background:#f0f7f5;border:1px solid rgba(61,107,97,.2);border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;">' +
      '<p style="font-weight:600;color:#3d6b61;margin:0 0 6px;">🏢 Business plan — Google Sheets sync setup</p>' +
      '<p style="margin:0;line-height:1.7;">The attached <strong>Servly_Expenses_GAS.gs</strong> file connects your app to Google Sheets. ' +
      'Open the Setup Guide for the full 6-step walkthrough. Reply to this email if you need help — we\'ll guide you through it.</p>' +
      '</div>';
  }
  if (tier === 'pro') {
    return '<div style="background:#fdf5f0;border:1px solid rgba(200,130,74,.2);border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;">' +
      '<p style="font-weight:600;color:#c8824a;margin:0 0 6px;">📊 Pro features to explore first</p>' +
      '<p style="margin:0;line-height:1.7;">Go to <strong>Categories</strong> to set budget limits per category. ' +
      'When logging an expense, toggle <strong>Recurring</strong> for monthly fixed costs. ' +
      'Add revenue on the Dashboard to see your <strong>Profit & Loss</strong> automatically.</p>' +
      '</div>';
  }
  return '<div style="background:#f5f0e8;border:1px solid #e8e0d4;border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;">' +
    '<p style="font-weight:600;margin:0 0 6px;">💡 Quick tip</p>' +
    '<p style="margin:0;line-height:1.7;">Export a CSV at the end of every month from the sidebar — ' +
    'this is your backup and accountant-ready file. Want more features? ' +
    '<a href="' + MESSENGER_URL + '" style="color:#c8824a;">Message us to upgrade anytime.</a></p>' +
    '</div>';
}

// ============================================================
//  EXPENSE SYNC FUNCTIONS (Business plan Sheets sync)
// ============================================================
function _syncExpenses(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expSheet = _getOrCreateSheet(ss, SHEET_EXPENSES);
  expSheet.clearContents();

  var expHeaders = ['ID','Date','Description','Category','Category ID',
    'Payment Method','Amount (₱)','Status','Staff','Branch','Notes','Synced At'];
  expSheet.getRange(1,1,1,expHeaders.length).setValues([expHeaders]);
  _styleHeader(expSheet, expHeaders.length, '#1e2d2d', '#ffffff');

  var expenses   = body.expenses   || [];
  var categories = _buildCatMap(body.categories || []);
  var syncedAt   = new Date().toISOString();

  if (expenses.length > 0) {
    var rows = expenses.map(function(e){
      return [e.id||'', e.date||'', e.desc||'', categories[e.catId]||e.catId||'',
        e.catId||'', e.payment||'Cash', parseFloat(e.amount)||0,
        e.status||'approved', e.staffName||'', e.branch||body.branch||'Main',
        e.notes||'', syncedAt];
    });
    expSheet.getRange(2,1,rows.length,expHeaders.length).setValues(rows);
    expSheet.getRange(2,7,rows.length,1).setNumberFormat('₱#,##0.00');
    expSheet.getRange(2,2,rows.length,1).setNumberFormat('@');
  }
  expSheet.autoResizeColumns(1, expHeaders.length);

  _syncCategories(ss, body.categories || []);
  _generateMonthlySummary({ expenses: expenses, categories: body.categories || [] });
  SpreadsheetApp.flush();

  return _jsonResponse({ success: true, synced: expenses.length,
    message: 'Synced ' + expenses.length + ' expenses to Google Sheets' });
}

function _syncCategories(ss, categories) {
  var s = _getOrCreateSheet(ss, SHEET_CATEGORIES);
  s.clearContents();
  var h = ['ID','Name','Color','Budget Limit (₱)'];
  s.getRange(1,1,1,h.length).setValues([h]);
  _styleHeader(s, h.length, '#c8824a', '#ffffff');
  if (categories.length > 0) {
    var rows = categories.map(function(c){ return [c.id||'',c.name||'',c.color||'',c.budget||'']; });
    s.getRange(2,1,rows.length,h.length).setValues(rows);
    rows.forEach(function(r,i){
      if(r[2]&&/^#[0-9a-fA-F]{6}$/.test(r[2])){
        try{ s.getRange(i+2,3).setBackground(r[2]).setFontColor('#ffffff'); }catch(e){}
      }
    });
  }
  s.autoResizeColumns(1,h.length);
}

function _generateMonthlySummary(body) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet = _getOrCreateSheet(ss, SHEET_SUMMARY);
  sumSheet.clearContents();
  var h = ['Month','Year','Total Expenses (₱)','Entry Count','Avg per Entry (₱)','Largest Expense (₱)','Top Category','Generated At'];
  sumSheet.getRange(1,1,1,h.length).setValues([h]);
  _styleHeader(sumSheet, h.length, '#3d6b61', '#ffffff');

  var expenses   = body.expenses || [];
  var categories = _buildCatMap(body.categories || []);
  var byMonth    = {};
  expenses.forEach(function(e){
    var ym = String(e.date||'').slice(0,7);
    if(!ym) return;
    if(!byMonth[ym]) byMonth[ym]=[];
    byMonth[ym].push(e);
  });

  var months      = Object.keys(byMonth).sort();
  var generatedAt = new Date().toISOString();
  if (months.length > 0) {
    var rows = months.map(function(ym){
      var exps  = byMonth[ym];
      var total = exps.reduce(function(s,e){ return s+(parseFloat(e.amount)||0); },0);
      var max   = exps.reduce(function(a,e){ return (parseFloat(e.amount)||0)>a?(parseFloat(e.amount)||0):a; },0);
      var byCat = {};
      exps.forEach(function(e){ byCat[e.catId]=(byCat[e.catId]||0)+(parseFloat(e.amount)||0); });
      var topCatId = Object.keys(byCat).sort(function(a,b){ return byCat[b]-byCat[a]; })[0]||'';
      var parts = ym.split('-');
      return [ym, parts[0]||'', parseFloat(total.toFixed(2)), exps.length,
        exps.length?parseFloat((total/exps.length).toFixed(2)):0,
        parseFloat(max.toFixed(2)), categories[topCatId]||topCatId, generatedAt];
    });
    sumSheet.getRange(2,1,rows.length,h.length).setValues(rows);
    [3,5,6].forEach(function(c){ sumSheet.getRange(2,c,rows.length,1).setNumberFormat('₱#,##0.00'); });
  }
  sumSheet.autoResizeColumns(1,h.length);
  return _jsonResponse({ success: true, months: months.length });
}

function _appendExpense(body) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _getOrCreateSheet(ss, SHEET_EXPENSES);
  if (sheet.getLastRow() < 1) {
    var h = ['ID','Date','Description','Category','Category ID','Payment','Amount (₱)','Status','Staff','Branch','Notes','Synced At'];
    sheet.getRange(1,1,1,h.length).setValues([h]);
    _styleHeader(sheet, h.length, '#1e2d2d', '#ffffff');
  }
  var e = body.expense || {};
  var cats = _buildCatMap(body.categories || []);
  sheet.appendRow([e.id||'',e.date||'',e.desc||'',cats[e.catId]||e.catId||'',
    e.catId||'',e.payment||'Cash',parseFloat(e.amount)||0,
    e.status||'approved',e.staffName||'',e.branch||'Main',e.notes||'',new Date().toISOString()]);
  SpreadsheetApp.flush();
  return _jsonResponse({ success: true, message: 'Expense appended' });
}

function _getExpenses(body) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet || sheet.getLastRow() < 2) return _jsonResponse({ success: true, data: [] });
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows    = data.slice(1).map(function(row){
    var obj = {};
    headers.forEach(function(h,i){ obj[h]=row[i]; });
    return obj;
  });
  return _jsonResponse({ success: true, data: rows, count: rows.length });
}

function _deleteExpense(body) {
  var id    = String(body.id||'').trim();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet||!id) return _jsonResponse({ success: false, error: 'No ID' });
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i+1);
      SpreadsheetApp.flush();
      return _jsonResponse({ success: true, message: 'Deleted' });
    }
  }
  return _jsonResponse({ success: false, error: 'ID not found: ' + id });
}

// ============================================================
//  HELPERS
// ============================================================
function _getOrCreateSheet(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); Logger.log('Created sheet: ' + name); }
  return s;
}

function _styleHeader(sheet, colCount, bg, fg) {
  sheet.getRange(1,1,1,colCount).setFontWeight('bold').setBackground(bg).setFontColor(fg).setFontSize(11);
  sheet.setFrozenRows(1);
}

function _buildCatMap(categories) {
  var map = {};
  (categories||[]).forEach(function(c){ map[c.id]=c.name; });
  return map;
}

function _dateStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function _eRow(label, value, bold) {
  return '<tr style="border-bottom:1px solid #f0e8dc;">' +
    '<td style="padding:7px 10px;color:#888;font-size:12px;width:130px;">' + label + '</td>' +
    '<td style="padding:7px 10px;' + (bold?'font-weight:600;':'') + '">' + (value||'—') + '</td>' +
    '</tr>';
}

function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  SHEETS MENU (for manual use)
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Servly Expenses')
    .addItem('Generate monthly summary', 'menuGenerateSummary')
    .addItem('Clear expense data', 'menuClearExpenses')
    .addToUi();
}

function menuGenerateSummary() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No expense data. Sync from Servly Expenses first.');
    return;
  }
  var data = sheet.getDataRange().getValues();
  var expenses = data.slice(1).map(function(r){ return { date:r[1], amount:r[6], catId:r[4] }; });
  _generateMonthlySummary({ expenses: expenses, categories: [] });
  SpreadsheetApp.getUi().alert('Monthly summary generated in the "' + SHEET_SUMMARY + '" tab.');
}

function menuClearExpenses() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear all expense data?','This cannot be undone.',ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_EXPENSES, SHEET_CATEGORIES, SHEET_SUMMARY].forEach(function(n){
    var s = ss.getSheetByName(n);
    if(s) s.clearContents();
  });
  ui.alert('Done.');
}
