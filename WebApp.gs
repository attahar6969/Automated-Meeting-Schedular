/**
 * WebApp.gs
 * Serves the web app and holds every dynamic setting API:
 * branding + logo, contacts, automation, meeting defaults, dropdown lists,
 * first-time setup, daily maintenance and "meeting starts soon" alerts.
 */

var XK = {
  TAGLINE: 'Tagline',
  C1: 'Brand Primary Color',
  C2: 'Brand Accent Color',
  DUR: 'Default Meeting Duration (minutes)',
  DEF_START: 'Default Start Time',
  DEF_TYPE: 'Default Meeting Type',
  DEF_PRI: 'Default Priority',
  DEF_MODE: 'Default Meeting Mode',
  DEF_LOC: 'Default Location Type',
  ALERT_MIN: 'Meeting Alert Minutes Before (0 = off)',
  AUTO_OVERDUE: 'Auto Mark Overdue Actions (Yes/No)',
  AUTO_COMPLETE: 'Auto Complete Past Meetings (Yes/No)',
  FOLDER: 'Report Folder Name'
};

function yes_(v, d) {
  if (v === undefined || v === null || v === '') return !!d;
  return String(v).toLowerCase() === 'yes';
}

/* ============================== SERVE =================================== */

function doGet() {
  var b = brand_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(b.company + ' — Meeting Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================== BRANDING ================================ */

function brand_() {
  var s = {};
  try { s = getSettings_(); } catch (e) { /* sheet not ready */ }
  function g(k, d) { var v = s[k]; return (v === undefined || v === '') ? d : String(v); }
  return {
    company: g(SK.COMPANY, 'KISAN ENTERPRISES'),
    tagline: g(XK.TAGLINE, 'Hardware & Plumbing Solutions'),
    c1: g(XK.C1, '#18202E'),
    c2: g(XK.C2, '#D62828'),
    logo: getLogo_()
  };
}

function appBranding() { return brand_(); }

function appWho() { return currentUserEmail_(); }

function getLogo_() {
  try {
    var p = PropertiesService.getScriptProperties();
    var n = parseInt(p.getProperty('LOGO_N'), 10);
    if (n) {
      var all = p.getProperties(), s = '';
      for (var i = 0; i < n; i++) s += all['LOGO_' + i] || '';
      if (s.indexOf('data:image/') === 0) return s;
    }
  } catch (e) { /* use default */ }
  return DEFAULT_LOGO_;
}

/** dataUrl = '' resets to the default logo. Returns the logo now in use. */
function appSaveLogo(dataUrl) {
  requireRole_();
  var p = PropertiesService.getScriptProperties();
  var n = parseInt(p.getProperty('LOGO_N'), 10) || 0;
  for (var i = 0; i < n; i++) p.deleteProperty('LOGO_' + i);
  p.deleteProperty('LOGO_N');
  dataUrl = String(dataUrl || '');
  if (dataUrl) {
    if (!/^data:image\/(png|jpeg|webp);base64,/.test(dataUrl)) throw new Error('Logo must be a PNG, JPG or WEBP image.');
    if (dataUrl.length > 46000) throw new Error('Logo is too large. Use a smaller image.');
    var chunks = Math.ceil(dataUrl.length / 8000), o = {};
    for (var c = 0; c < chunks; c++) o['LOGO_' + c] = dataUrl.substr(c * 8000, 8000);
    o.LOGO_N = String(chunks);
    p.setProperties(o);
  }
  logAudit_('Update logo', KE.SHEET.SETTINGS, 'BRANDING', 'Logo', '', dataUrl ? 'Custom logo' : 'Default logo');
  return getLogo_();
}

/* ============================== BOOTSTRAP =============================== */

function cfg_(s) {
  function g(k, d) { var v = s[k]; return (v === undefined || v === '') ? d : v; }
  var m = timeToMinutes_(s[XK.DEF_START]);
  if (m === null) m = 600;
  return {
    dur: parseInt(g(XK.DUR, 60), 10) || 60,
    defStart: ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2),
    defType: String(g(XK.DEF_TYPE, 'Customer Meeting')),
    defPri: String(g(XK.DEF_PRI, 'Medium')),
    defMode: String(g(XK.DEF_MODE, 'In-Person')),
    defLoc: String(g(XK.DEF_LOC, "Customer's Office"))
  };
}

function appBootstrap() {
  var ss = getSS_();
  if (!ss.getSheetByName(KE.SHEET.MEETINGS) || !ss.getSheetByName(KE.SHEET.SETTINGS)) {
    requireRole_();
    return { needsSetup: true, user: currentUserEmail_(), brand: brand_() };
  }
  var b = dashBootstrap();
  b.user = currentUserEmail_();
  b.brand = brand_();
  b.cfg = cfg_(getSettings_());
  return b;
}

/** First-time setup straight from the web app (creates all sheets). */
function appRunSetup() {
  requireRole_();
  var rep = withLock_(function () { return runSetup_(); });
  logAudit_('System setup', 'Workbook', 'SETUP', 'Version', '', KE.VERSION);
  return rep;
}

/* ============================== SETTINGS ================================ */

function appGetSettings() {
  requireRole_();
  var s = getSettings_(), c = cfg_(s);
  function g(k, d) { var v = s[k]; return (v === undefined || v === '') ? (d === undefined ? '' : String(d)) : String(v); }
  return {
    company: g(SK.COMPANY, 'KISAN ENTERPRISES'), tagline: g(XK.TAGLINE, 'Hardware & Plumbing Solutions'),
    c1: g(XK.C1, '#18202E'), c2: g(XK.C2, '#D62828'),
    ownerName: g(SK.OWNER_NAME), ownerEmail: g(SK.OWNER_EMAIL), ownerWa: g(SK.OWNER_WA),
    eaEmail: g(SK.EA_EMAIL), eaWa: g(SK.EA_WA), officeAddr: g(SK.OFFICE_ADDR), officeMap: g(SK.OFFICE_MAP),
    calId: g(SK.CAL_ID, 'primary'), syncMin: g(SK.SYNC_MIN, 5),
    remEnabled: g(SK.REM_ENABLED, 'Yes'), remHour: g(SK.REM_HOUR, 9), remTo: g(SK.REM_TO, 'Both'),
    alertMin: g(XK.ALERT_MIN, 30), autoOverdue: yes_(s[XK.AUTO_OVERDUE], true) ? 'Yes' : 'No',
    autoComplete: yes_(s[XK.AUTO_COMPLETE], false) ? 'Yes' : 'No', folder: g(XK.FOLDER, 'KISAN ENTERPRISES Reports'),
    dur: String(c.dur), defStart: c.defStart, defType: c.defType, defPri: c.defPri, defMode: c.defMode, defLoc: c.defLoc
  };
}

function appSaveSettings(d) {
  requireRole_();
  function oneOf(v, arr, label) { if (arr.indexOf(String(v)) === -1) throw new Error(label + ' is not valid.'); }
  function num(v, lo, hi, label) {
    var n = parseInt(v, 10);
    if (isNaN(n) || n < lo || n > hi) throw new Error(label + ' must be between ' + lo + ' and ' + hi + '.');
    return n;
  }
  oneOf(d.syncMin, ['1', '5', '10', '15', '30'], 'Sync interval');
  oneOf(d.remEnabled, ['Yes', 'No'], 'Daily reminder');
  oneOf(d.remTo, ['Owner', 'EA', 'Both'], 'Reminder recipients');
  oneOf(d.autoOverdue, ['Yes', 'No'], 'Auto mark overdue');
  oneOf(d.autoComplete, ['Yes', 'No'], 'Auto complete meetings');
  var hour = num(d.remHour, 0, 23, 'Reminder hour');
  var alert = num(d.alertMin, 0, 720, 'Meeting alert minutes');
  var dur = num(d.dur, 5, 720, 'Default duration');
  var st = timeToMinutes_(d.defStart);
  if (st === null) throw new Error('Default start time is not valid.');
  if (!/^#[0-9a-fA-F]{6}$/.test(d.c1) || !/^#[0-9a-fA-F]{6}$/.test(d.c2)) throw new Error('Brand colours must look like #18202E.');
  if (getList_('meetingType').indexOf(d.defType) === -1) throw new Error('Default meeting type is not in the list.');
  if (getList_('priority').indexOf(d.defPri) === -1) throw new Error('Default priority is not in the list.');
  if (getList_('meetingMode').indexOf(d.defMode) === -1) throw new Error('Default meeting mode is not in the list.');
  if (getList_('locationType').indexOf(d.defLoc) === -1) throw new Error('Default location type is not in the list.');
  var folder = trim_(d.folder) || 'KISAN ENTERPRISES Reports';

  dashSaveSettings(d);   // company, owner / EA contacts, office (validated + lock-out protected)
  withLock_(function () {
    setSetting_(XK.TAGLINE, trim_(d.tagline));
    setSetting_(XK.C1, d.c1);
    setSetting_(XK.C2, d.c2);
    setSetting_(SK.CAL_ID, trim_(d.calId) || 'primary');
    setSetting_(SK.SYNC_MIN, parseInt(d.syncMin, 10));
    setSetting_(SK.REM_ENABLED, d.remEnabled);
    setSetting_(SK.REM_HOUR, hour);
    setSetting_(SK.REM_TO, d.remTo);
    setSetting_(XK.ALERT_MIN, alert);
    setSetting_(XK.AUTO_OVERDUE, d.autoOverdue);
    setSetting_(XK.AUTO_COMPLETE, d.autoComplete);
    setSetting_(XK.FOLDER, folder);
    setSetting_(XK.DUR, dur);
    setSetting_(XK.DEF_START, "'" + minutesToText_(st));
    setSetting_(XK.DEF_TYPE, d.defType);
    setSetting_(XK.DEF_PRI, d.defPri);
    setSetting_(XK.DEF_MODE, d.defMode);
    setSetting_(XK.DEF_LOC, d.defLoc);
    var t = triggerInfo_();
    if (t.sync || t.daily) { try { installTriggers_(); } catch (e) { Logger.log('Trigger refresh failed: ' + e); } }
  });
  return true;
}

/* ============================== DROPDOWN LISTS ========================== */

/** Values the code relies on. These stay in the list; everything else is editable. */
function listKeep_(k) {
  if (k === 'meetingType') return ['Customer Meeting'];
  if (k === 'followType') return ['Customer Call'];
  return LISTS[k].items.slice();
}

function appGetLists() {
  requireRole_();
  return Object.keys(LISTS).map(function (k) {
    return { key: k, title: LISTS[k].title, items: getList_(k), keep: listKeep_(k) };
  });
}

function appSaveList(key, items) {
  requireRole_();
  if (!LISTS[key]) throw new Error('Unknown list.');
  var seen = {}, clean = [];
  (items || []).forEach(function (x) {
    x = trim_(x);
    if (x && !seen[x.toLowerCase()]) { seen[x.toLowerCase()] = 1; clean.push(x); }
  });
  listKeep_(key).forEach(function (k) {
    if (clean.indexOf(k) === -1) throw new Error('"' + k + '" is used by the system and cannot be removed.');
  });
  if (clean.length > 100) throw new Error('A list can hold at most 100 values.');
  withLock_(function () {
    var sh = getSheet_(KE.SHEET.SETTINGS);
    var col = LISTS_FIRST_COL + Object.keys(LISTS).indexOf(key);
    sh.getRange(SETTINGS_FIRST_ROW, col, 100, 1).clearContent();
    sh.getRange(SETTINGS_FIRST_ROW, col, clean.length, 1).setValues(clean.map(function (x) { return [x]; }));
    logAudit_('Update list', KE.SHEET.SETTINGS, key, LISTS[key].title, '', clean.join(', '));
  });
  return clean;
}

/* ============================== AUTOMATION EXTRAS ======================= */

/** Daily: mark late actions Overdue, optionally complete past meetings. */
function runMaintenance_() {
  var s = getSettings_(), out = { overdue: 0, completed: 0 }, today = todayDate_().getTime();
  if (yes_(s[XK.AUTO_OVERDUE], true)) {
    var ash = getSheet_(KE.SHEET.ACTIONS);
    readRows_(ash, ACTION_FIELDS.length).forEach(function (r, i) {
      var st = trim_(r[AC.status - 1]), d = toDate_(r[AC.deadline - 1]);
      if (trim_(r[0]) && (st === 'Pending' || st === 'In Progress') && d && d.getTime() < today) {
        ash.getRange(i + 2, AC.status).setValue('Overdue');
        ash.getRange(i + 2, AC.updatedAt).setValue(new Date());
        logAudit_('Auto overdue', KE.SHEET.ACTIONS, r[0], 'Status', st, 'Overdue');
        out.overdue++;
      }
    });
  }
  if (yes_(s[XK.AUTO_COMPLETE], false)) {
    var msh = getSheet_(KE.SHEET.MEETINGS);
    readRows_(msh, MEETING_FIELDS.length).forEach(function (r, i) {
      var st = trim_(r[MC.status - 1]), d = toDate_(r[MC.date - 1]);
      if (trim_(r[0]) && (st === 'Scheduled' || st === 'Rescheduled') && d && d.getTime() < today) {
        msh.getRange(i + 2, MC.status).setValue('Completed');
        msh.getRange(i + 2, MC.updatedAt).setValue(new Date());
        logAudit_('Auto complete', KE.SHEET.MEETINGS, r[0], 'Status', st, 'Completed');
        out.completed++;
      }
    });
  }
  return out;
}

/** Called by the sync timer: e-mails owner/EA when a meeting starts within N minutes. */
function meetingAlerts_() {
  var mins = parseInt(getSetting_(XK.ALERT_MIN, 30), 10);
  if (isNaN(mins) || mins < 1) return;
  var to = recipients_(getSettings_());
  if (!to.length) return;
  var now = new Date(), props = PropertiesService.getDocumentProperties();
  var sent = {};
  try { sent = JSON.parse(props.getProperty('ALERTS') || '{}'); } catch (e) { sent = {}; }
  var due = [];
  loadObjs_(KE.SHEET.MEETINGS, MEETING_FIELDS).forEach(function (m) {
    if (m.status === 'Cancelled' || m.status === 'Deleted Externally' || m.status === 'Completed') return;
    var st = combineDateTime_(m.date, m.start);
    if (!st) return;
    var diff = (st.getTime() - now.getTime()) / 60000;
    var key = m.id + '|' + isoDate_(st) + '|' + fmtTime_(m.start);
    if (diff > 0 && diff <= mins && !sent[key]) { due.push(m); sent[key] = now.getTime(); }
  });
  Object.keys(sent).forEach(function (k) { if (now.getTime() - sent[k] > 2 * DAY_MS_) delete sent[k]; });
  props.setProperty('ALERTS', JSON.stringify(sent));
  if (!due.length) return;
  var b = brand_();
  due.forEach(function (m) {
    var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1b2437"><div style="background:' + b.c1 + ';color:#fff;padding:12px 16px;border-radius:8px"><b>' +
      escH_(b.company) + '</b><br>Meeting starting soon</div><p style="font-size:16px"><b>' + escH_(m.title) + '</b></p><p>' +
      escH_(fmtDate_(m.date)) + ', ' + escH_(fmtTime_(m.start)) + '<br>With: ' + escH_(m.custName || m.company || '—') +
      '<br>Place: ' + escH_(m.mode === 'Virtual' ? (m.meetLink || 'Google Meet') : (m.address || 'to be confirmed')) +
      (m.mapLink ? '<br>Map: ' + escH_(m.mapLink) : '') + (m.agenda ? '<br>Agenda: ' + escH_(m.agenda) : '') + '</p></div>';
    MailApp.sendEmail({ to: to.join(','), subject: 'Starting soon: ' + m.title + ' (' + fmtTime_(m.start) + ')', htmlBody: html, name: b.company });
  });
}
function dashDelete(name, id) {
  requireRole_();
  if (['meeting', 'customer', 'minutes', 'action', 'followup'].indexOf(name) === -1) {
    throw new Error('Unknown record type.');
  }
  return withLock_(function () {
    var e = ent_(name), sh = getSheet_(e.sheet), row = findRow_(sh, id);
    if (!row) throw new Error('Record not found.');
    if (name === 'meeting') {
      removeMeetingEvents_(id, trim_(sh.getRange(row, MC.eventId).getValue()));
      deleteLinkedFollowups_(id);
    }
    sh.deleteRow(row);
    logAudit_('Delete', e.sheet, id, '', '', 'Deleted from web app');
    return true;
  });
}
function deleteLinkedFollowups_(meetingId) {
  var fsh = getSheet_(KE.SHEET.FOLLOWUPS), last = fsh.getLastRow(), n = 0;
  if (last < 2) return 0;
  var ids = fsh.getRange(2, FC.meetingId, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {      // from the bottom, so row numbers stay correct
    if (trim_(ids[i][0]) === trim_(meetingId)) { fsh.deleteRow(i + 2); n++; }
  }
  if (n) logAudit_('Delete', KE.SHEET.FOLLOWUPS, meetingId, '', '', n + ' follow-up(s) deleted with the meeting');
  return n;
}
