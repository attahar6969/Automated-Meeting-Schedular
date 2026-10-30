/**
 * File: Utils.gs   (Part 2 of 7)
 * Shared helpers: settings access, locking, ID generation, phone
 * normalisation, date/time handling, audit + sync logging, access control.
 */

/* ============================ SHEET ACCESS ============================== */

var SS_CACHE_ = null;
function getSS_() {
  if (SS_CACHE_) return SS_CACHE_;
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* web app */ }
  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('SSID');
    if (!id) throw new Error('Spreadsheet not linked yet. Open the spreadsheet once, then run authorizeAll in the Apps Script editor.');
    ss = SpreadsheetApp.openById(id);
  }
  SS_CACHE_ = ss;
  return ss;
}

var SHEET_MEMO_ = {};
function getSheet_(name) {
  if (SHEET_MEMO_[name]) return SHEET_MEMO_[name];
  var s = getSS_().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" was not found. Run menu: ATTAHAR > Setup (first time) first.');
  SHEET_MEMO_[name] = s;
  return s;
}

/** Safe UI alert (does nothing when running from a trigger with no UI). */
function showAlert_(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(title + ': ' + message);
  }
}

function toast_(message, title) {
  try { getSS_().toast(message, title || 'ATTAHAR', 6); } catch (e) { Logger.log(message); }
}

/* ============================== SETTINGS ================================ */
/* ---------- fast cache helpers ---------- */
function cacheGet_(key) {
  try { var v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function cachePut_(key, val, ttl) {
  try { var s = JSON.stringify(val); if (s.length < 90000) CacheService.getScriptCache().put(key, s, ttl || 300); } catch (e) {}
}
function cacheDrop_(keys) {
  try { CacheService.getScriptCache().removeAll(keys); } catch (e) {}
}
function clearCfgCache_() {
  SETTINGS_MEMO_ = null; LISTS_ALL_ = null;
  cacheDrop_(['KE_SETTINGS', 'KE_LISTS']);
}
/** Returns {settingName: value} read from the Settings sheet (one batch read). */
var SETTINGS_MEMO_ = null;
function getSettings_() {
  if (SETTINGS_MEMO_) return SETTINGS_MEMO_;
  var cached = cacheGet_('KE_SETTINGS');
  if (cached) { SETTINGS_MEMO_ = cached; return cached; }
  var sh = getSS_().getSheetByName(KE.SHEET.SETTINGS);
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < SETTINGS_FIRST_ROW) return out;
  var vals = sh.getRange(SETTINGS_FIRST_ROW, 1, last - SETTINGS_FIRST_ROW + 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim();
    if (!k) continue;
    var v = vals[i][1];
    if (v === null || v === undefined) v = '';
    if (Object.prototype.toString.call(v) === '[object Date]') v = minutesToText_(timeToMinutes_(v));
    out[k] = v;
  }
  SETTINGS_MEMO_ = out;
  cachePut_('KE_SETTINGS', out, 300);
  return out;
}

function getSetting_(key, fallback) {
  var s = getSettings_();
  var v = s[key];
  if (v === undefined || v === '') return (fallback === undefined ? '' : fallback);
  return v;
}

function setSetting_(key, value) {
  clearCfgCache_();
  var sh = getSheet_(KE.SHEET.SETTINGS);
  var last = sh.getLastRow();
  if (last >= SETTINGS_FIRST_ROW) {
    var keys = sh.getRange(SETTINGS_FIRST_ROW, 1, last - SETTINGS_FIRST_ROW + 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        sh.getRange(SETTINGS_FIRST_ROW + i, 2).setValue(value);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, 2).setValues([[key, value]]);
}

/** Writes many settings with ONE sheet read. Only cells whose value really changed are written. */
function setSettingsBatch_(pairs) {
  var sh = getSheet_(KE.SHEET.SETTINGS);
  var last = sh.getLastRow(), rowOf = {}, cur = {}, lastKeyRow = SETTINGS_FIRST_ROW - 1;
  if (last >= SETTINGS_FIRST_ROW) {
    var vals = sh.getRange(SETTINGS_FIRST_ROW, 1, last - SETTINGS_FIRST_ROW + 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var k = String(vals[i][0]).trim();
      if (k) { rowOf[k] = SETTINGS_FIRST_ROW + i; cur[k] = vals[i][1]; lastKeyRow = SETTINGS_FIRST_ROW + i; }
    }
  }
  var adds = [];
  pairs.forEach(function (p) {
    var key = p[0], val = p[1];
    var want = String(val === null || val === undefined ? '' : val).replace(/^'/, '');
    if (rowOf[key]) {
      var have = cur[key];
      if (Object.prototype.toString.call(have) === '[object Date]') have = minutesToText_(timeToMinutes_(have));
      if (String(have === null || have === undefined ? '' : have) === want) return;   // unchanged
      sh.getRange(rowOf[key], 2).setValue(val);
    } else {
      adds.push([key, val]);
    }
  });
  if (adds.length) sh.getRange(lastKeyRow + 1, 1, adds.length, 2).setValues(adds);
  clearCfgCache_();
}

/** Returns the items of a central dropdown list (see LISTS in Config.gs). */
var LISTS_ALL_ = null;
function getAllLists_() {
  if (LISTS_ALL_) return LISTS_ALL_;
  var hit = cacheGet_('KE_LISTS');
  if (hit) { LISTS_ALL_ = hit; return hit; }
  var keys = Object.keys(LISTS), out = {}, vals = [];
  var sh = getSS_().getSheetByName(KE.SHEET.SETTINGS);
  if (sh && sh.getLastRow() >= SETTINGS_FIRST_ROW) {
    vals = sh.getRange(SETTINGS_FIRST_ROW, LISTS_FIRST_COL, sh.getLastRow() - SETTINGS_FIRST_ROW + 1, keys.length).getValues();
  }
  keys.forEach(function (k, ci) {
    var items = [];
    for (var r = 0; r < vals.length; r++) { var t = String(vals[r][ci]).trim(); if (t) items.push(t); }
    out[k] = items.length ? items : LISTS[k].items.slice();
  });
  if (sh) { LISTS_ALL_ = out; cachePut_('KE_LISTS', out, 300); }
  return out;
}
function getList_(listKey) {
  if (!LISTS[listKey]) return [];
  return getAllLists_()[listKey].slice();
}

/* ============================== LOCKING ================================= */

var LOCK_DEPTH_ = 0;
var LOCK_OBJ_ = null;

/**
 * Runs fn() while holding the script lock. Re-entrant inside one execution.
 * Prevents two users/triggers from writing IDs or syncing at the same moment.
 *
 * Every function that CHANGES data runs through here, so a VIEWER (Open To All
 * mode) is stopped here on the server, whatever the browser sends.
 */
var VIEWER_OK_ = false;
/** Call at the top of functions that a VIEWER may use to add or edit records. */
function allowViewerWrite_() { VIEWER_OK_ = true; }
function withLock_(fn, timeoutMs) {
  var acquired = false;
  if (LOCK_DEPTH_ === 0) {
    if (!VIEWER_OK_) blockViewerWrites_();
    LOCK_OBJ_ = LockService.getScriptLock();
    if (!LOCK_OBJ_.tryLock(timeoutMs || 30000)) {
      LOCK_OBJ_ = null;
      throw new Error('The system is busy (another update or calendar sync is running). Please wait a few seconds and try again.');
    }
    acquired = true;
  }
  LOCK_DEPTH_++;
  try {
    return fn();
  } finally {
    LOCK_DEPTH_--;
    if (acquired) {
      SpreadsheetApp.flush();
      LOCK_OBJ_.releaseLock();
      LOCK_OBJ_ = null;
    }
  }
}

/** Throws when a signed-in VIEWER tries to change anything. Time-driven triggers are not affected. */
function blockViewerWrites_() {
  var real = '';
  try { real = Session.getActiveUser().getEmail(); } catch (e) { real = ''; }
  if (real && getUserRole_() === 'VIEWER') {
    throw new Error('View only: you can look at everything, but only the Owner or the Executive Assistant can make changes.');
  }
}

/** Use at the top of functions that change something but do not use withLock_. */
function denyViewer_() {
  if (getUserRole_() === 'VIEWER') {
    throw new Error('View only: you can look at everything, but only the Owner or the Executive Assistant can make changes.');
  }
}

/* ============================ ID GENERATION ============================= */

/**
 * Returns the next ID for kind = MTG | MOM | ACT | FUP | CUS.
 * Counters are stored in Document Properties, so an ID is NEVER reused,
 * even if a row is deleted. Call inside withLock_().
 */
function nextId_(kind, knownIds) {
  var def = KE.ID[kind];
  if (!def) throw new Error('Unknown ID type: ' + kind);
  var year = Utilities.formatDate(new Date(), KE.TZ, 'yyyy');
  var prefix = def.yearly ? def.prefix + '-' + year + '-' : def.prefix + '-';
  var key = 'CTR_' + kind + (def.yearly ? '_' + year : '');
  var props = PropertiesService.getDocumentProperties();
  var stored = parseInt(props.getProperty(key), 10);
  if (isNaN(stored)) stored = 0;
  var inSheet = knownIds ? maxSeqIn_(knownIds, prefix) : maxExistingSeq_(KE.SHEET[def.sheet], def.col, prefix);
  var n = Math.max(stored, inSheet) + 1;      // never go below what is already in the sheet
  props.setProperty(key, String(n));
  return prefix + Utilities.formatString('%0' + def.pad + 'd', n);
}

function maxSeqIn_(ids, prefix) {
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var v = String(ids[i]);
    if (v.indexOf(prefix) === 0) {
      var n = parseInt(v.substring(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return max;
}

/** Highest sequence number already present in a sheet column for the prefix. */
function maxExistingSeq_(sheetName, col, prefix) {
  var sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  return maxSeqIn_(vals.map(function (r) { return r[0]; }), prefix);
}

/* ============================ PHONE NUMBERS ============================= */

/**
 * Validates and normalises an Indian mobile number.
 * Accepts: 9876543210, 09876543210, +91 98765 43210, 91-9876543210 ...
 * Returns {ok:true, value:'919876543210', display:'+91 98765 43210'}
 *      or {ok:false, error:'...'}.
 */
function normalizePhone_(raw) {
  var s = String(raw === null || raw === undefined ? '' : raw).replace(/[^\d]/g, '');
  if (s.length === 11 && s.charAt(0) === '0') s = s.substring(1);
  if (s.length === 12 && s.indexOf('91') === 0) s = s.substring(2);
  if (s.length === 10 && /^[6-9]\d{9}$/.test(s)) {
    return { ok: true, value: '91' + s, display: '+91 ' + s.substring(0, 5) + ' ' + s.substring(5) };
  }
  return { ok: false, error: 'Invalid Indian mobile number. Enter 10 digits starting with 6-9 (optionally with +91).' };
}

/** Digits-only normalised number, or '' if invalid/blank. */
function waDigits_(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === '') return '';
  var r = normalizePhone_(raw);
  return r.ok ? r.value : '';
}

/** Builds a WhatsApp click-to-chat URL. number must already be normalised. */
function whatsappUrl_(number, message) {
  var n = waDigits_(number);
  if (!n) throw new Error('A valid WhatsApp number is required to open WhatsApp.');
  return 'https://wa.me/' + n + '?text=' + encodeURIComponent(message);
}

/* ============================ DATE AND TIME ============================= */

/** Accepts Date, 'yyyy-MM-dd' or 'dd-MMM-yyyy' text; returns a Date at 00:00 India time, or null. */
function toDate_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[-\/ ](\d{4})$/);
  if (m) {
    var mon = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());
    if (mon >= 0) return new Date(+m[3], mon, +m[1]);
  }
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);   // Indian order: DD/MM/YYYY
  return null;
}

function fmtDate_(v) {
  var d = toDate_(v);
  return d ? Utilities.formatDate(d, KE.TZ, KE.DATE_FMT) : '';
}

function todayDate_() {
  var n = new Date();
  var s = Utilities.formatDate(n, KE.TZ, 'yyyy-MM-dd').split('-');
  return new Date(+s[0], +s[1] - 1, +s[2]);
}

/**
 * Converts a time value to minutes after midnight.
 * Accepts '14:30', '2:30 PM', '02:30 pm', a number 0-1 (sheet time) or a Date.
 * Returns null when it cannot be understood.
 */
function timeToMinutes_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    // Sheet "time only" cells come back as 30-Dec-1899 + time-of-day.
    var base = new Date(1899, 11, 30, 0, 0, 0);
    var mins = Math.round((v.getTime() - base.getTime()) / 60000);
    if (mins >= 0 && mins < 1440) return mins;
    return v.getHours() * 60 + v.getMinutes();
  }
  if (typeof v === 'number') {
    if (v >= 0 && v < 1) return Math.round(v * 1440);
    return null;
  }
  var s = String(v).trim().toUpperCase();
  var m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/);
  if (!m) return null;
  var h = +m[1], mi = +m[2];
  if (mi > 59) return null;
  if (m[3]) {
    if (h < 1 || h > 12) return null;
    if (m[3] === 'AM' && h === 12) h = 0;
    if (m[3] === 'PM' && h !== 12) h += 12;
  } else if (h > 23) {
    return null;
  }
  return h * 60 + mi;
}

/** minutes after midnight -> 'hh:mm AM/PM' (this text format is what we store in the sheet). */
function minutesToText_(mins) {
  if (mins === null || mins === undefined) return '';
  var h = Math.floor(mins / 60) % 24, m = mins % 60;
  var ap = h >= 12 ? 'PM' : 'AM';
  var h12 = h % 12; if (h12 === 0) h12 = 12;
  return (h12 < 10 ? '0' : '') + h12 + ':' + (m < 10 ? '0' : '') + m + ' ' + ap;
}

function fmtTime_(v) {
  return minutesToText_(timeToMinutes_(v));
}

/** Combines a date and a time value into a real Date (India time). */
function combineDateTime_(dateVal, timeVal) {
  var d = toDate_(dateVal);
  var mins = timeToMinutes_(timeVal);
  if (!d || mins === null) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(mins / 60), mins % 60, 0);
}

function nowStamp_() {
  return new Date();
}

/* ============================== USERS / ACCESS ========================== */

function currentUserEmail_() {
  var e = '';
  try { e = Session.getActiveUser().getEmail(); } catch (x) { /* ignore */ }
  if (!e) { try { e = Session.getEffectiveUser().getEmail(); } catch (y) { /* ignore */ } }
  return String(e || '').toLowerCase();
}

/** 'a@x.in, B@x.in' -> ['a@x.in','b@x.in'] (lower-case, no blanks, no duplicates). */
function emailList_(v) {
  var seen = {}, out = [];
  String(v === null || v === undefined ? '' : v).split(/[,;\s]+/).forEach(function (x) {
    x = x.trim().toLowerCase();
    if (x && !seen[x]) { seen[x] = 1; out.push(x); }
  });
  return out;
}

/**
 * Returns 'OWNER', 'EA', 'VIEWER' (Open To All = Yes: read-only),
 * 'SETUP' (no emails configured yet) or '' (not authorised).
 */
function getUserRole_() {
  var s = getSettings_();
  var owners = emailList_(s[SK.OWNER_EMAIL]);
  var eas = emailList_(s[SK.EA_EMAIL]);
  if (!owners.length && !eas.length) return 'SETUP';
  var me = currentUserEmail_();
  if (me && owners.indexOf(me) !== -1) return 'OWNER';
  if (me && eas.indexOf(me) !== -1) return 'EA';
  if (me && String(s[SK.OPEN_ALL] || '').toLowerCase() === 'yes') return 'VIEWER';
  return '';
}

/**
 * Call at the start of every menu / button / form action.
 * Default: Owner, EA and (when Open To All is on) VIEWER may call it.
 * Anything that changes data is stopped for a VIEWER inside withLock_ / denyViewer_.
 * Time-driven triggers do not call this (they run as the installing user).
 */
function requireRole_(allowed) {
  var role = getUserRole_();
  var ok = (allowed || ['OWNER', 'EA', 'VIEWER']).concat(['SETUP']);
  if (ok.indexOf(role) === -1) {
    throw new Error('Access denied. Only the Owner and the Executive Assistant (as set in the Settings sheet) can use this function.');
  }
  return role;
}

/* ============================ LOGGING =================================== */
var AUDIT_BUF_ = null;

function logAudit_(action, sheetName, recordId, field, oldV, newV) {
  try {
    var row = [new Date(), currentUserEmail_() || 'system', action, sheetName, recordId || '', field || '',
      auditVal_(oldV), auditVal_(newV)];
    if (AUDIT_BUF_) { AUDIT_BUF_.push(row); return; }
    var sh = getSS_().getSheetByName(KE.SHEET.AUDIT);
    if (!sh) return;
    sh.appendRow(row);
  } catch (e) {
    Logger.log('Audit log failed: ' + e);
  }
}
function auditVal_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, KE.TZ, 'dd-MMM-yyyy HH:mm');
  var s = String(v);
  return s.length > 1000 ? s.substring(0, 1000) + '…' : s;
}

/** Logs the differences between two meeting objects (keys from MEETING_FIELDS). */
function logAuditDiff_(action, sheetName, recordId, fields, oldObj, newObj, skipKeys) {
  skipKeys = skipKeys || ['updatedAt'];
  var rows = [], now = new Date(), who = currentUserEmail_() || 'system';
  for (var i = 0; i < fields.length; i++) {
    var k = fields[i][0];
    if (skipKeys.indexOf(k) !== -1) continue;
    var a = auditVal_(oldObj[k]), b = auditVal_(newObj[k]);
    if (a !== b) rows.push([now, who, action, sheetName, recordId || '', fields[i][1], a, b]);
  }
  if (!rows.length) return;
  if (AUDIT_BUF_) { rows.forEach(function (r) { AUDIT_BUF_.push(r); }); return; }
  try {
    var sh = getSS_().getSheetByName(KE.SHEET.AUDIT);
    if (!sh) return;
    var start = sh.getLastRow() + 1;
    if (start + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length);
    sh.getRange(start, 1, rows.length, 8).setValues(rows);
  } catch (e) {
    Logger.log('Audit log failed: ' + e);
  }
}
/** Writes all collected audit rows with one sheet write. */
function flushAudit_() {
  var rows = AUDIT_BUF_; AUDIT_BUF_ = null;
  if (!rows || !rows.length) return;
  try {
    var sh = getSheet_(KE.SHEET.AUDIT);
    var start = sh.getLastRow() + 1;
    try { sh.getRange(start, 1, rows.length, 8).setValues(rows); }
    catch (e1) {
      sh.insertRowsAfter(sh.getMaxRows(), rows.length + 100);
      sh.getRange(start, 1, rows.length, 8).setValues(rows);
    }
  } catch (e) {
    Logger.log('Audit flush failed: ' + e);
  }
}


function logSync_(meetingId, eventId, direction, operation, status, errorText) {
  try {
    var sh = getSS_().getSheetByName(KE.SHEET.SYNCLOG);
    if (!sh) return;
    sh.appendRow([new Date(), meetingId || '', eventId || '', direction, operation, status,
      errorText ? String(errorText).substring(0, 1000) : '', currentUserEmail_() || 'system (trigger)']);
  } catch (e) {
    Logger.log('Sync log failed: ' + e);
  }
}

/* ============================ MISC HELPERS ============================== */

function trim_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

/** Reads every data row of a sheet as an array of arrays (batch read). */
function readRows_(sheet, numCols) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, numCols).getValues();
}

/** Converts a row array into an object using a *_FIELDS definition. */
function rowToObj_(row, fields) {
  var o = {};
  for (var i = 0; i < fields.length; i++) o[fields[i][0]] = row[i];
  return o;
}

function objToRow_(obj, fields) {
  var r = [];
  for (var i = 0; i < fields.length; i++) {
    var v = obj[fields[i][0]];
    r.push(v === undefined || v === null ? '' : v);
  }
  return r;
}

/** Simple retry wrapper for Google service calls that can fail transiently. */
function retry_(fn, attempts, label) {
  var lastErr;
  attempts = attempts || 3;
  for (var i = 0; i < attempts; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      var msg = String(e && e.message ? e.message : e);
      // Do not retry permanent errors
      if (/Not Found|404|410|Forbidden|403|Invalid|400|Unauthorized/i.test(msg) && !/rate|quota|limit/i.test(msg)) throw e;
      Utilities.sleep(500 * Math.pow(2, i));
    }
  }
  throw lastErr;
}
