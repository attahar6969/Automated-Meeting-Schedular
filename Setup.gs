/**
 * File: Setup.gs   (Part 3 of 7)
 * Creates every sheet, header, format, dropdown, colour rule and the Dashboard
 * layout. Safe to run again: it never deletes data, never overwrites values
 * already typed in Settings, and only adds what is missing.
 *
 * Functions you run: setupWorkbook(), verifySetup()
 */

var ROWS_PREPARED = 3000;      // rows pre-formatted in each data sheet
var KPI_RANGE_END = 5000;      // KPI formulas look at rows 2..5000

var DASH_BUTTONS = [
  { label: 'NEW MEETING',               fn: 'openNewMeetingForm',      color: '#13315C' },
  { label: 'SYNC GOOGLE CALENDAR',      fn: 'syncCalendarNow',         color: '#1F5FBF' },
  { label: 'ADD MEETING MINUTES',       fn: 'openMinutesForm',         color: '#13315C' },
  { label: 'VIEW MEETING REGISTER',     fn: 'goToMeetingRegister',     color: '#13315C' },
  { label: 'VIEW ACTION ITEMS',         fn: 'goToActionItems',         color: '#13315C' },
  { label: 'VIEW FOLLOW-UPS',           fn: 'goToFollowUps',           color: '#13315C' },
  { label: 'REQUEST CUSTOMER LOCATION', fn: 'openRequestLocationDialog', color: '#B9770E' },
  { label: 'GENERATE MEETING REPORT',   fn: 'openReportDialog',        color: '#13315C' },
  { label: 'REFRESH DASHBOARD',         fn: 'refreshDashboard',        color: '#1E8449' }
];

var DASH_LIST_HEADERS = ['Meeting ID', 'Date', 'Time', 'Meeting Title', 'Customer', 'Type', 'Type of Customer', 'Status', 'Location Type', 'Location Status', 'Calendar Sync', 'Remarks'];
var DASH_ACTION_HEADERS = ['Action ID', 'Meeting ID', 'Task', 'Assigned To', 'Type of Customer', 'Assigned', 'Deadline', 'Status', 'Days Overdue', 'Meeting Title', 'Remarks', 'Last Updated'];
var DASH_ACTION_TITLE_ROW = 67;
var DASH_ACTION_HEADER_ROW = 68;
var DASH_ACTION_FIRST_ROW = 69;
var DASH_ACTION_MAX_ROWS = 30;

/* ============================== MAIN SETUP ============================== */

function setupWorkbook() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Set up KISAN ENTERPRISES system',
    'This will create any missing sheets (Dashboard, Meeting Register, Meeting Minutes, Action Items, Follow-Up Tracker, Customers, Calendar Sync Log, Audit Log, Settings), ' +
    'headers, dropdowns and formatting.\n\nExisting data is NOT deleted.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var report = withLock_(function () { return runSetup_(); });

  logAudit_('System setup', 'Workbook', 'SETUP', 'Version', '', KE.VERSION);
  ui.alert('Setup complete',
    report.join('\n') +
    '\n\nNEXT STEPS:\n1. Open the Settings sheet and fill Owner/EA e-mail and WhatsApp numbers, Calendar ID, Office address.\n' +
    '2. Run menu: KISAN ENTERPRISES > Verify Setup.',
    ui.ButtonSet.OK);
}

function runSetup_() {
  var ss = getSS_();
  var report = [];
  ss.setSpreadsheetTimeZone(KE.TZ);
  try { ss.setSpreadsheetLocale('en_IN'); } catch (e) { /* locale not critical */ }

  // 1. Sheets in the required order
  var order = [KE.SHEET.DASHBOARD, KE.SHEET.MEETINGS, KE.SHEET.MINUTES, KE.SHEET.ACTIONS,
    KE.SHEET.FOLLOWUPS, KE.SHEET.CUSTOMERS, KE.SHEET.SYNCLOG, KE.SHEET.AUDIT, KE.SHEET.SETTINGS];
  var sheets = {};
  for (var i = 0; i < order.length; i++) {
    var existed = !!ss.getSheetByName(order[i]);
    sheets[order[i]] = ensureSheet_(order[i]);
    ss.setActiveSheet(sheets[order[i]]);
    ss.moveActiveSheet(i + 1);
    report.push((existed ? 'Found:   ' : 'Created: ') + order[i]);
  }

  // 2. Remove the empty default "Sheet1" only if it is completely blank
  ['Sheet1', 'Sheet 1'].forEach(function (n) {
    var d = ss.getSheetByName(n);
    if (d && ss.getSheets().length > order.length && d.getLastRow() === 0 && d.getLastColumn() === 0) ss.deleteSheet(d);
  });

  // 3. Settings first (dropdown lists are needed by validations)
  buildSettingsSheet_(sheets[KE.SHEET.SETTINGS]);

  // 4. Data sheets
  buildDataSheet_(sheets[KE.SHEET.MEETINGS], MEETING_FIELDS, {
    date: KE.SHEET_DATE_FMT, start: '@', end: '@', custMobile: '@', custWa: '@', ownerMobile: '@',
    createdAt: 'dd-mmm-yyyy hh:mm', updatedAt: 'dd-mmm-yyyy hh:mm'
  }, MC, [120, 200, 220, 105, 85, 85, 140, 160, 160, 120, 130, 120, 100, 160, 300, 220, 220, 200, 300, 80, 110, 140, 130, 140, 120, 190, 130, 130, 250]);
  buildDataSheet_(sheets[KE.SHEET.MINUTES], MINUTES_FIELDS, {
    date: KE.SHEET_DATE_FMT, createdAt: 'dd-mmm-yyyy hh:mm', updatedAt: 'dd-mmm-yyyy hh:mm'
  }, MM, [120, 120, 220, 105, 170, 380, 260, 300, 300, 240, 105, 130, 130]);
  buildDataSheet_(sheets[KE.SHEET.ACTIONS], ACTION_FIELDS, {
    assignedDate: KE.SHEET_DATE_FMT, deadline: KE.SHEET_DATE_FMT, completedDate: KE.SHEET_DATE_FMT, updatedAt: 'dd-mmm-yyyy hh:mm'
  }, AC, [120, 120, 220, 380, 150, 150, 80, 105, 105, 100, 110, 260, 130]);
  buildDataSheet_(sheets[KE.SHEET.FOLLOWUPS], FOLLOWUP_FIELDS, {
    date: KE.SHEET_DATE_FMT, completedDate: KE.SHEET_DATE_FMT
  }, FC, [120, 120, 170, 300, 105, 160, 130, 100, 260, 110]);
  buildDataSheet_(sheets[KE.SHEET.CUSTOMERS], CUSTOMER_FIELDS, {
    mobile: '@', wa: '@', createdAt: 'dd-mmm-yyyy hh:mm'
  }, CC, [110, 170, 190, 125, 135, 210, 340, 240, 260, 130]);
  buildDataSheet_(sheets[KE.SHEET.SYNCLOG], SYNCLOG_FIELDS, { ts: 'dd-mmm-yyyy hh:mm:ss' }, SC, [155, 120, 220, 130, 150, 90, 400, 200]);
  buildDataSheet_(sheets[KE.SHEET.AUDIT], AUDIT_FIELDS, { ts: 'dd-mmm-yyyy hh:mm:ss' }, UC, [155, 220, 190, 150, 130, 190, 300, 300]);

  // 5. Dropdowns, colour rules, dashboard
  applyValidations_();
  applyConditionalFormats_();
  buildDashboardLayout_(sheets[KE.SHEET.DASHBOARD]);

  // 6. Land on the Dashboard
  ss.setActiveSheet(sheets[KE.SHEET.DASHBOARD]);
  return report;
}

function ensureSheet_(name) {
  var ss = getSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureColumns_(sh, n) {
  var have = sh.getMaxColumns();
  if (have < n) sh.insertColumnsAfter(have, n - have);
}

function ensureRows_(sh, n) {
  var have = sh.getMaxRows();
  if (have < n) sh.insertRowsAfter(have, n - have);
}

function colLetter_(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function styleHeader_(range) {
  range.setBackground(KE.COLOR.NAVY).setFontColor(KE.COLOR.WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

/* ============================ DATA SHEET BUILDER ======================== */

function buildDataSheet_(sh, fields, formatsByKey, colMap, widths) {
  var n = fields.length;
  ensureColumns_(sh, n);
  ensureRows_(sh, ROWS_PREPARED + 1);
  var hdr = sh.getRange(1, 1, 1, n);
  // Write headers only if the sheet is new (row 1 empty); never overwrite a user's data.
  if (trim_(sh.getRange(1, 1).getValue()) === '') hdr.setValues([headersOf_(fields)]);
  styleHeader_(hdr);
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(Math.min(3, n) >= 3 && sh.getName() === KE.SHEET.MEETINGS ? 3 : 1);
  if (widths) for (var c = 0; c < widths.length && c < n; c++) sh.setColumnWidth(c + 1, widths[c]);

  var body = sh.getRange(2, 1, ROWS_PREPARED, n);
  body.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  var keys = Object.keys(formatsByKey);
  for (var k = 0; k < keys.length; k++) {
    sh.getRange(2, colMap[keys[k]], ROWS_PREPARED, 1).setNumberFormat(formatsByKey[keys[k]]);
  }
  // Wrap long text columns
  var wrapKeys = ['title', 'address', 'agenda', 'summary', 'topics', 'decisions', 'instructions', 'remarks', 'task', 'subject', 'participants', 'notes', 'error', 'oldV', 'newV'];
  wrapKeys.forEach(function (wk) {
    if (colMap[wk]) sh.getRange(2, colMap[wk], ROWS_PREPARED, 1).setWrap(true);
  });
  // Light banding line
  if (sh.getBandings().length === 0) {
    try {
      sh.getRange(1, 1, ROWS_PREPARED + 1, n)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
      var b = sh.getBandings()[0];
      b.setHeaderRowColor(KE.COLOR.NAVY).setFirstRowColor('#FFFFFF').setSecondRowColor(KE.COLOR.GREY_BG);
    } catch (e) { Logger.log('Banding skipped on ' + sh.getName() + ': ' + e); }
  }
}

/* ============================ SETTINGS SHEET ============================ */

function buildSettingsSheet_(sh) {
  var listCount = Object.keys(LISTS).length;
  ensureColumns_(sh, LISTS_FIRST_COL + listCount);
  ensureRows_(sh, 120);
  sh.setHiddenGridlines(true);

  sh.getRange('A1:C1').merge().setValue('KISAN ENTERPRISES — SETTINGS')
    .setBackground(KE.COLOR.NAVY).setFontColor(KE.COLOR.WHITE).setFontWeight('bold').setFontSize(16).setHorizontalAlignment('left');
  sh.getRange('A2:C2').merge()
    .setValue('Edit only the VALUE column (B). Dropdown lists on the right (columns E onward) control every dropdown in the system.')
    .setFontStyle('italic').setFontColor('#555555');
  sh.setRowHeight(1, 36);

  var head = sh.getRange(3, 1, 1, 3);
  head.setValues([['Setting', 'Value', 'Notes']]);
  styleHeader_(head);

  // Settings table: add only missing keys
  var existing = {};
  var lastA = sh.getLastRow();
  var nextRow = SETTINGS_FIRST_ROW;
  if (lastA >= SETTINGS_FIRST_ROW) {
    var colA = sh.getRange(SETTINGS_FIRST_ROW, 1, lastA - SETTINGS_FIRST_ROW + 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var key = trim_(colA[i][0]);
      if (key) { existing[key] = true; nextRow = SETTINGS_FIRST_ROW + i + 1; }
    }
  }
  // nextRow only moves past the settings block when no list data sits in column A (lists are in E+)
  SETTINGS_DEFAULTS.forEach(function (d) {
    if (!existing[d[0]]) {
      sh.getRange(nextRow, 1, 1, 3).setValues([d]);
      nextRow++;
    }
  });

  var tableLast = nextRow - 1;
  var table = sh.getRange(SETTINGS_FIRST_ROW, 1, tableLast - SETTINGS_FIRST_ROW + 1, 3);
  table.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, KE.COLOR.GREY_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(SETTINGS_FIRST_ROW, 1, tableLast - SETTINGS_FIRST_ROW + 1, 1).setFontWeight('bold').setBackground(KE.COLOR.GREY_BG);
  sh.getRange(SETTINGS_FIRST_ROW, 2, tableLast - SETTINGS_FIRST_ROW + 1, 1).setBackground('#FFFBE6');
  sh.getRange(SETTINGS_FIRST_ROW, 3, tableLast - SETTINGS_FIRST_ROW + 1, 1).setFontColor('#555555');

  // Phone / ID style cells as text so leading digits are never altered
  var rowOf = {};
  var keysNow = sh.getRange(SETTINGS_FIRST_ROW, 1, tableLast - SETTINGS_FIRST_ROW + 1, 1).getValues();
  for (var r = 0; r < keysNow.length; r++) rowOf[trim_(keysNow[r][0])] = SETTINGS_FIRST_ROW + r;
  [SK.OWNER_WA, SK.EA_WA].forEach(function (k) { if (rowOf[k]) sh.getRange(rowOf[k], 2).setNumberFormat('@'); });

  // Small value dropdowns
  function setList(key, values) {
    if (!rowOf[key]) return;
    sh.getRange(rowOf[key], 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build());
  }
  setList(SK.SYNC_MIN, ['1', '5', '10', '15', '30']);
  setList(SK.REM_ENABLED, ['Yes', 'No']);
  setList(SK.REM_TO, ['Owner', 'EA', 'Both']);
  var hrs = []; for (var h = 0; h < 24; h++) hrs.push(String(h));
  setList(SK.REM_HOUR, hrs);

  // Central dropdown lists (written only if that list column is still empty)
  var li = 0;
  Object.keys(LISTS).forEach(function (lk) {
    var col = LISTS_FIRST_COL + li;
    var titleCell = sh.getRange(3, col);
    if (trim_(titleCell.getValue()) === '') {
      titleCell.setValue(LISTS[lk].title);
      var items = LISTS[lk].items.map(function (x) { return [x]; });
      sh.getRange(SETTINGS_FIRST_ROW, col, items.length, 1).setValues(items);
    }
    styleHeader_(titleCell);
    sh.getRange(SETTINGS_FIRST_ROW, col, 100, 1).setBackground('#F7FAFF').setFontFamily('Arial').setFontSize(10);
    sh.setColumnWidth(col, 170);
    li++;
  });

  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 330);
  sh.setColumnWidth(3, 470);
  sh.setColumnWidth(4, 30);
  sh.setFrozenRows(3);
}

/* ============================ VALIDATIONS =============================== */

function listValidation_(listKey) {
  var sh = getSheet_(KE.SHEET.SETTINGS);
  var idx = Object.keys(LISTS).indexOf(listKey);
  var rng = sh.getRange(SETTINGS_FIRST_ROW, LISTS_FIRST_COL + idx, 100, 1);
  return SpreadsheetApp.newDataValidation().requireValueInRange(rng, true)
    .setAllowInvalid(false).setHelpText('Please choose a value from the list.').build();
}

function applyValidations_() {
  function put(sheetName, colNum, listKey) {
    getSheet_(sheetName).getRange(2, colNum, ROWS_PREPARED, 1).setDataValidation(listValidation_(listKey));
  }
  var M = KE.SHEET.MEETINGS;
  put(M, MC.type, 'meetingType');
  put(M, MC.mode, 'meetingMode');
  put(M, MC.locType, 'locationType');
  put(M, MC.priority, 'priority');
  put(M, MC.status, 'meetingStatus');
  put(M, MC.locStatus, 'locationStatus');
  put(M, MC.ownerNotif, 'notifStatus');
  put(M, MC.custNotif, 'notifStatus');
  put(M, MC.syncStatus, 'syncStatus');
  put(KE.SHEET.MINUTES, MM.status, 'minutesStatus');
  put(KE.SHEET.ACTIONS, AC.priority, 'priority');
  put(KE.SHEET.ACTIONS, AC.status, 'actionStatus');
  put(KE.SHEET.FOLLOWUPS, FC.type, 'followType');
  put(KE.SHEET.FOLLOWUPS, FC.status, 'followStatus');
}

/* ========================= CONDITIONAL FORMATTING ======================= */

function textRule_(range, text, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text)
    .setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build();
}

function formulaRule_(range, formula, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

function applyConditionalFormats_() {
  var C = KE.COLOR;
  var n = ROWS_PREPARED;

  // Meeting Register (only if the sheet has no rules yet — keeps any manual rules)
  var reg = getSheet_(KE.SHEET.MEETINGS);
  if (reg.getConditionalFormatRules().length === 0) {
    var st = reg.getRange(2, MC.status, n, 1);
    var ls = reg.getRange(2, MC.locStatus, n, 1);
    var sy = reg.getRange(2, MC.syncStatus, n, 1);
    var ad = reg.getRange(2, MC.address, n, 1);
    var pr = reg.getRange(2, MC.priority, n, 1);
    var L = colLetter_(MC.locStatus), I = colLetter_(MC.id);
    var rules = [
      textRule_(st, 'Scheduled', C.BLUE_BG, C.BLUE),
      textRule_(st, 'Rescheduled', C.BLUE_BG, C.BLUE),
      textRule_(st, 'Completed', C.GREEN_BG, C.GREEN),
      textRule_(st, 'Cancelled', C.RED_BG, C.RED),
      textRule_(st, 'Deleted Externally', C.RED_BG, C.RED),
      textRule_(ls, 'Pending', C.AMBER_BG, C.AMBER),
      textRule_(ls, 'Request Sent', C.AMBER_BG, C.AMBER),
      textRule_(ls, 'Received', C.BLUE_BG, C.BLUE),
      textRule_(ls, 'Confirmed', C.GREEN_BG, C.GREEN),
      // Visible warning: address cell turns amber while the location is unresolved
      formulaRule_(ad, '=AND($' + I + '2<>"",OR($' + L + '2="Pending",$' + L + '2="Request Sent"))', C.AMBER_BG, C.AMBER),
      textRule_(sy, 'Synced', C.GREEN_BG, C.GREEN),
      textRule_(sy, 'SYNC CONFLICT', C.RED_BG, C.RED),
      textRule_(sy, 'Error', C.RED_BG, C.RED),
      textRule_(sy, 'Pending Update', C.AMBER_BG, C.AMBER),
      textRule_(sy, 'Deleted Externally', C.RED_BG, C.RED),
      textRule_(pr, 'High', C.RED_BG, C.RED),
      textRule_(pr, 'Medium', C.AMBER_BG, C.AMBER),
      textRule_(pr, 'Low', C.GREEN_BG, C.GREEN)
    ];
    reg.setConditionalFormatRules(rules);
  }

  // Action Items
  var act = getSheet_(KE.SHEET.ACTIONS);
  if (act.getConditionalFormatRules().length === 0) {
    var as = act.getRange(2, AC.status, n, 1);
    var dl = act.getRange(2, AC.deadline, n, 1);
    var S = colLetter_(AC.status), D = colLetter_(AC.deadline), A = colLetter_(AC.id);
    act.setConditionalFormatRules([
      textRule_(as, 'Completed', C.GREEN_BG, C.GREEN),
      textRule_(as, 'Pending', C.AMBER_BG, C.AMBER),
      textRule_(as, 'In Progress', C.BLUE_BG, C.BLUE),
      textRule_(as, 'On Hold', C.GREY_BG, '#555555'),
      textRule_(as, 'Overdue', C.RED_BG, C.RED),
      formulaRule_(dl, '=AND($' + A + '2<>"",$' + D + '2<>"",$' + D + '2<TODAY(),$' + S + '2<>"Completed")', C.RED_BG, C.RED)
    ]);
  }

  // Minutes
  var mom = getSheet_(KE.SHEET.MINUTES);
  if (mom.getConditionalFormatRules().length === 0) {
    var ms = mom.getRange(2, MM.status, n, 1);
    mom.setConditionalFormatRules([
      textRule_(ms, 'Draft', C.AMBER_BG, C.AMBER),
      textRule_(ms, 'Under Review', C.BLUE_BG, C.BLUE),
      textRule_(ms, 'Finalized', C.GREEN_BG, C.GREEN)
    ]);
  }

  // Follow-ups
  var fu = getSheet_(KE.SHEET.FOLLOWUPS);
  if (fu.getConditionalFormatRules().length === 0) {
    var fs = fu.getRange(2, FC.status, n, 1);
    var fd = fu.getRange(2, FC.date, n, 1);
    var FS = colLetter_(FC.status), FD = colLetter_(FC.date), FI = colLetter_(FC.id);
    fu.setConditionalFormatRules([
      textRule_(fs, 'Completed', C.GREEN_BG, C.GREEN),
      textRule_(fs, 'Pending', C.AMBER_BG, C.AMBER),
      textRule_(fs, 'Rescheduled', C.BLUE_BG, C.BLUE),
      textRule_(fs, 'Cancelled', C.RED_BG, C.RED),
      formulaRule_(fd, '=AND($' + FI + '2<>"",$' + FD + '2<>"",$' + FD + '2<TODAY(),$' + FS + '2="Pending")', C.RED_BG, C.RED)
    ]);
  }

  // Sync log: failures in red
  var sl = getSheet_(KE.SHEET.SYNCLOG);
  if (sl.getConditionalFormatRules().length === 0) {
    var ss2 = sl.getRange(2, SC.status, n, 1);
    sl.setConditionalFormatRules([
      textRule_(ss2, 'Success', C.GREEN_BG, C.GREEN),
      textRule_(ss2, 'Failed', C.RED_BG, C.RED),
      textRule_(ss2, 'Conflict', C.RED_BG, C.RED),
      textRule_(ss2, 'Skipped', C.AMBER_BG, C.AMBER)
    ]);
  }
}

/* ============================ DASHBOARD LAYOUT ========================== */

function kpiRange_(sheetName, colNum) {
  var L = colLetter_(colNum);
  return "'" + sheetName + "'!" + L + '2:' + L + KPI_RANGE_END;
}

function buildDashboardLayout_(sh) {
  var C = KE.COLOR;
  ensureColumns_(sh, 14);
  ensureRows_(sh, 120);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, 120, 14).setFontFamily('Arial').setVerticalAlignment('middle');
  sh.setColumnWidth(1, 18);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 104);
  sh.setColumnWidth(14, 18);

  // Heading
  sh.getRange('B1:M1').merge().setValue('KISAN ENTERPRISES').setBackground(C.NAVY).setFontColor(C.WHITE)
    .setFontSize(26).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1, 52);
  sh.getRange('B2:M2').merge().setValue("OWNER'S MEETING MANAGEMENT DASHBOARD").setBackground(C.NAVY_LIGHT).setFontColor(C.WHITE)
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(2, 30);
  sh.getRange('B3:F3').merge().setFormula('="Today:  "&TEXT(TODAY(),"dd-mmm-yyyy")')
    .setFontColor(C.NAVY).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('G3:M3').merge().setHorizontalAlignment('right').setFontColor('#555555').setFontStyle('italic');
  if (trim_(sh.getRange('G3').getValue()) === '') sh.getRange('G3').setValue('Lists last refreshed: not yet');
  sh.setRowHeight(3, 24);

  // KPI section
  sectionHeader_(sh, 5, 'KEY INDICATORS');
  var M = KE.SHEET.MEETINGS, A = KE.SHEET.ACTIONS, N = KE.SHEET.MINUTES;
  var mDate = kpiRange_(M, MC.date), mStat = kpiRange_(M, MC.status), mLoc = kpiRange_(M, MC.locStatus), mId = kpiRange_(M, MC.id);
  var aStat = kpiRange_(A, AC.status), aDead = kpiRange_(A, AC.deadline);
  var nMid = kpiRange_(N, MM.meetingId), nStat = kpiRange_(N, MM.status);

  var kpis = [
    ["TODAY'S MEETINGS",
      '=COUNTIFS(' + mDate + ',TODAY(),' + mStat + ',"<>Cancelled",' + mStat + ',"<>Deleted Externally")', C.BLUE, C.BLUE_BG],
    ['UPCOMING MEETINGS',
      '=COUNTIFS(' + mDate + ',">"&TODAY(),' + mStat + ',"Scheduled")+COUNTIFS(' + mDate + ',">"&TODAY(),' + mStat + ',"Rescheduled")', C.BLUE, C.BLUE_BG],
    ['COMPLETED MEETINGS',
      '=COUNTIF(' + mStat + ',"Completed")', C.GREEN, C.GREEN_BG],
    ['CANCELLED MEETINGS',
      '=COUNTIF(' + mStat + ',"Cancelled")+COUNTIF(' + mStat + ',"Deleted Externally")', C.RED, C.RED_BG],
    ['PENDING ACTION ITEMS',
      '=COUNTIFS(' + aStat + ',"Pending",' + aDead + ',">="&TODAY())+COUNTIFS(' + aStat + ',"Pending",' + aDead + ',"")' +
      '+COUNTIFS(' + aStat + ',"In Progress",' + aDead + ',">="&TODAY())+COUNTIFS(' + aStat + ',"In Progress",' + aDead + ',"")', C.AMBER, C.AMBER_BG],
    ['OVERDUE ACTION ITEMS',
      '=COUNTIF(' + aStat + ',"Overdue")+COUNTIFS(' + aStat + ',"Pending",' + aDead + ',"<"&TODAY())+COUNTIFS(' + aStat + ',"In Progress",' + aDead + ',"<"&TODAY())', C.RED, C.RED_BG],
    ['MEETING MINUTES PENDING',
      '=SUMPRODUCT((' + mId + '<>"")*(((' + mDate + '<TODAY())+(' + mStat + '="Completed"))>0)*(' + mStat + '<>"Cancelled")*(' + mStat + '<>"Deleted Externally")' +
      '*(COUNTIFS(' + nMid + ',' + mId + ',' + nStat + ',"Finalized")=0))', C.AMBER, C.AMBER_BG],
    ['CUSTOMER LOCATIONS PENDING',
      '=COUNTIFS(' + mLoc + ',"Pending",' + mStat + ',"<>Cancelled",' + mStat + ',"<>Deleted Externally")' +
      '+COUNTIFS(' + mLoc + ',"Request Sent",' + mStat + ',"<>Cancelled",' + mStat + ',"<>Deleted Externally")', C.AMBER, C.AMBER_BG]
  ];
  for (var i = 0; i < kpis.length; i++) {
    var group = Math.floor(i / 4);
    var labelRow = 6 + group * 3, valueRow = labelRow + 1;
    var col = 2 + (i % 4) * 3;
    var lab = sh.getRange(labelRow, col, 1, 3).merge();
    lab.setValue(kpis[i][0]).setBackground(kpis[i][2]).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
    var val = sh.getRange(valueRow, col, 1, 3).merge();
    val.setFormula(kpis[i][1]).setBackground(kpis[i][3]).setFontColor(kpis[i][2]).setFontWeight('bold').setFontSize(28)
      .setHorizontalAlignment('center').setNumberFormat('0');
    sh.getRange(labelRow, col, 2, 3).setBorder(true, true, true, true, false, false, C.WHITE, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.setRowHeight(labelRow, 24);
    sh.setRowHeight(valueRow, 50);
  }
  sh.setRowHeight(8, 10);
  sh.setRowHeight(11, 10);

  // Filters
  sectionHeader_(sh, 12, 'FILTERS  (leave blank to show all)');
  var fl = [
    ['B13:C13', 'From Date'], ['D13:E13', 'To Date'], ['F13:G13', 'Meeting Type'],
    ['H13:I13', 'Meeting Status'], ['J13:K13', 'Customer'], ['L13:M13', 'Type of Customer'],
    ['B15:C15', 'Action Item Status']
  ];
  fl.forEach(function (f) {
    sh.getRange(f[0]).merge().setValue(f[1]).setFontSize(9).setFontColor('#555555').setFontWeight('bold').setHorizontalAlignment('left');
  });
  ['B14:C14', 'D14:E14', 'F14:G14', 'H14:I14', 'J14:K14', 'L14:M14', 'B16:C16'].forEach(function (a) {
    sh.getRange(a).merge().setBackground('#FFFBE6').setHorizontalAlignment('center').setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, C.GREY_LINE, SpreadsheetApp.BorderStyle.SOLID);
  });
  var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Enter a valid date.').build();
  sh.getRange(KE.DASH.FILTER.FROM).setDataValidation(dateRule).setNumberFormat(KE.SHEET_DATE_FMT);
  sh.getRange(KE.DASH.FILTER.TO).setDataValidation(dateRule).setNumberFormat(KE.SHEET_DATE_FMT);
  sh.getRange(KE.DASH.FILTER.TYPE).setDataValidation(listValidation_('meetingType'));
  sh.getRange(KE.DASH.FILTER.STATUS).setDataValidation(listValidation_('meetingStatus'));
  sh.getRange(KE.DASH.FILTER.PRIORITY).setDataValidation(listValidation_('priority'));
  sh.getRange(KE.DASH.FILTER.ACTION_STATUS).setDataValidation(listValidation_('actionStatus'));
  sh.getRange(KE.DASH.FILTER.CUSTOMER).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(getSheet_(KE.SHEET.CUSTOMERS).getRange('B2:B3000'), true).setAllowInvalid(true).build());
  sh.getRange('D16:M16').merge().setValue('Choose filters, then tick  REFRESH DASHBOARD  to update the two lists below.')
    .setFontStyle('italic').setFontColor('#555555').setHorizontalAlignment('left');
  sh.setRowHeight(14, 28);
  sh.setRowHeight(16, 28);
  sh.setRowHeight(17, 10);

  // Action buttons (tick-box buttons — see note in Part 6 about the installable trigger)
  sectionHeader_(sh, 19, 'ACTIONS  (tick the box beside a button to run it)');
  for (var b = 0; b < DASH_BUTTONS.length; b++) {
    var row = KE.DASH.BUTTON_ROWS[Math.floor(b / 3)];
    var chkCol = KE.DASH.BUTTON_CHECK_COLS[b % 3];
    var chk = sh.getRange(row, chkCol);
    chk.insertCheckboxes().setValue(false).setHorizontalAlignment('center').setBackground(C.GREY_BG);
    sh.getRange(row, chkCol + 1, 1, 3).merge().setValue(DASH_BUTTONS[b].label).setBackground(DASH_BUTTONS[b].color)
      .setFontColor(C.WHITE).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10)
      .setBorder(true, true, true, true, false, false, C.WHITE, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.setRowHeight(row, 34);
  }
  sh.setRowHeight(23, 10);

  // Filtered meeting list (filled by refreshDashboard in Part 5)
  sectionHeader_(sh, 24, 'MEETINGS  (filtered)');
  var lh = sh.getRange(KE.DASH.LIST_HEADER_ROW, 2, 1, DASH_LIST_HEADERS.length);
  lh.setValues([DASH_LIST_HEADERS]);
  styleHeader_(lh);
  sh.setRowHeight(KE.DASH.LIST_HEADER_ROW, 30);
  sh.getRange(KE.DASH.LIST_FIRST_ROW, 2, KE.DASH.LIST_MAX_ROWS, DASH_LIST_HEADERS.length)
    .setFontSize(9).setWrap(true).setVerticalAlignment('middle')
    .setBorder(null, null, null, null, false, true, C.GREY_LINE, SpreadsheetApp.BorderStyle.SOLID);

  // Action item list
  sectionHeader_(sh, DASH_ACTION_TITLE_ROW, 'ACTION ITEMS  (filtered)');
  var ah = sh.getRange(DASH_ACTION_HEADER_ROW, 2, 1, DASH_ACTION_HEADERS.length);
  ah.setValues([DASH_ACTION_HEADERS]);
  styleHeader_(ah);
  sh.setRowHeight(DASH_ACTION_HEADER_ROW, 30);
  sh.getRange(DASH_ACTION_FIRST_ROW, 2, DASH_ACTION_MAX_ROWS, DASH_ACTION_HEADERS.length)
    .setFontSize(9).setWrap(true).setVerticalAlignment('middle')
    .setBorder(null, null, null, null, false, true, C.GREY_LINE, SpreadsheetApp.BorderStyle.SOLID);

  sh.setTabColor(C.NAVY);
  sh.setFrozenRows(3);
}

function sectionHeader_(sh, row, text) {
  sh.getRange(row, 2, 1, 12).merge().setValue(text).setBackground(KE.COLOR.GREY_BG).setFontColor(KE.COLOR.NAVY)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left')
    .setBorder(false, false, true, false, false, false, KE.COLOR.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(row, 26);
}

/* =============================== VERIFY ================================= */

/**
 * Checks everything Part 1 created and runs self-tests on the helper functions.
 * Also asks Google for Calendar permission (first-time authorisation).
 */
function verifySetup() {
  var lines = [];
  var okAll = true;
  function ok(msg) { lines.push('✔ ' + msg); }
  function bad(msg) { lines.push('✘ ' + msg); okAll = false; }

  // 1. Sheets and headers
  var defs = [
    [KE.SHEET.MEETINGS, MEETING_FIELDS], [KE.SHEET.MINUTES, MINUTES_FIELDS], [KE.SHEET.ACTIONS, ACTION_FIELDS],
    [KE.SHEET.FOLLOWUPS, FOLLOWUP_FIELDS], [KE.SHEET.CUSTOMERS, CUSTOMER_FIELDS],
    [KE.SHEET.SYNCLOG, SYNCLOG_FIELDS], [KE.SHEET.AUDIT, AUDIT_FIELDS]
  ];
  if (getSS_().getSheetByName(KE.SHEET.DASHBOARD)) ok('Sheet "Dashboard" exists'); else bad('Sheet "Dashboard" missing');
  if (getSS_().getSheetByName(KE.SHEET.SETTINGS)) ok('Sheet "Settings" exists'); else bad('Sheet "Settings" missing');
  defs.forEach(function (d) {
    var sh = getSS_().getSheetByName(d[0]);
    if (!sh) { bad('Sheet "' + d[0] + '" missing'); return; }
    var want = headersOf_(d[1]);
    var have = sh.getRange(1, 1, 1, want.length).getValues()[0];
    var diff = [];
    for (var i = 0; i < want.length; i++) if (String(have[i]).trim() !== want[i]) diff.push(colLetter_(i + 1));
    if (diff.length) bad('"' + d[0] + '": headers differ in column(s) ' + diff.join(', ')); else ok('"' + d[0] + '": ' + want.length + ' columns correct');
  });

  // 2. Settings
  var s = getSettings_();
  Object.keys(SK).forEach(function (k) {
    if (s[SK[k]] === undefined) bad('Setting missing: ' + SK[k]);
  });
  ok('Settings table present');
  var needFill = [SK.OWNER_EMAIL, SK.EA_EMAIL, SK.OWNER_WA, SK.EA_WA, SK.OFFICE_ADDR].filter(function (k) { return trim_(s[k]) === ''; });
  if (needFill.length) lines.push('⚠ Please fill in Settings: ' + needFill.join(', '));
  Object.keys(LISTS).forEach(function (lk) {
    if (getList_(lk).length === 0) bad('Dropdown list empty: ' + LISTS[lk].title);
  });
  ok('Dropdown lists present (' + Object.keys(LISTS).length + ')');

  // 3. Helper self-tests
  var t = [];
  function check(name, cond) { if (cond) t.push(name); else bad('Self-test failed: ' + name); }
  check('phone 10-digit', normalizePhone_('98765 43210').value === '919876543210');
  check('phone +91', normalizePhone_('+91-9876543210').value === '919876543210');
  check('phone leading 0', normalizePhone_('09876543210').value === '919876543210');
  check('phone invalid short', normalizePhone_('12345').ok === false);
  check('phone invalid start digit', normalizePhone_('5876543210').ok === false);
  check('time 24h', timeToMinutes_('14:30') === 870);
  check('time 12h', timeToMinutes_('02:30 PM') === 870);
  check('time 12 AM', timeToMinutes_('12:15 AM') === 15);
  check('time text out', minutesToText_(870) === '02:30 PM');
  check('date iso', fmtDate_('2026-03-05') === '05-Mar-2026');
  check('date DD/MM/YYYY', fmtDate_('05/03/2026') === '05-Mar-2026');
  check('whatsapp url', whatsappUrl_('9876543210', 'Hi & bye').indexOf('https://wa.me/919876543210?text=Hi%20%26%20bye') === 0);
  if (t.length === 12) ok('All 12 helper self-tests passed');

  // 4. Calendar service + permission
  try {
    if (typeof Calendar === 'undefined') {
      bad('Advanced Calendar service is not enabled (check appsscript.json or Services > Calendar API)');
    } else {
      var calId = trim_(s[SK.CAL_ID]) || 'primary';
      var cal = Calendar.Calendars.get(calId);
      ok('Calendar API reachable: "' + (cal.summary || calId) + '" (' + (cal.timeZone || '') + ')');
    }
  } catch (e) {
    bad('Calendar API error: ' + e.message + ' — check "Owner Calendar ID" in Settings and that you granted permission.');
  }

  // 5. Spreadsheet time zone
  if (getSS_().getSpreadsheetTimeZone() === KE.TZ) ok('Spreadsheet time zone = Asia/Kolkata'); else bad('Spreadsheet time zone is not Asia/Kolkata');

  var msg = lines.join('\n') + '\n\n' + (okAll ? 'RESULT: Part 1 is complete and working.' : 'RESULT: Some checks failed — see ✘ lines above.');
  showAlert_('Verify Setup', msg);
  Logger.log(msg);
  return okAll;
}
