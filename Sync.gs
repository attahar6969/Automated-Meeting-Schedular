/**
 * Sync.gs (Part 7 of 7)
 * Calendar sync, automation triggers, daily brief e-mail, PDF reports, system health.
 * Replaces any earlier Sync.gs.
 */

var DAY_MS_ = 86400000;
var TRIGGER_FNS_ = ['autoSync', 'dailyReminderRun'];

function calId_() { return trim_(getSetting_(SK.CAL_ID)) || 'primary'; }
function isoDate_(d) { return Utilities.formatDate(d, KE.TZ, 'yyyy-MM-dd'); }
function hm_(d) { return Utilities.formatDate(d, KE.TZ, 'HH:mm'); }
function escH_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

function loadObjs_(sheetName, fields) {
  return readRows_(getSheet_(sheetName), fields.length).map(function (r) { return rowToObj_(r, fields); })
    .filter(function (o) { return trim_(o[fields[0][0]]) !== ''; });
}

/* ============================ CALENDAR SYNC ============================= */

function syncCalendarNow() {
  requireRole_();
  var msg = syncSummary_(withLock_(function () { return syncAll_(); }));
  toast_(msg, 'Calendar sync');
  return msg;
}

function dashSyncNow() {
  requireRole_();
  return syncSummary_(withLock_(function () { return syncAll_(); }));
}

function autoSync() {            // time-driven trigger
  try { withLock_(function () { syncAll_(); }, 10000); } catch (e) { Logger.log('autoSync: ' + e); }
  try { meetingAlerts_(); } catch (e2) { Logger.log('meetingAlerts: ' + e2); }
}

function syncSummary_(r) {
  return 'Created ' + r.created + ', updated ' + r.updated + ', pulled from calendar ' + r.pulled +
    ', removed ' + r.removed + ', deleted in calendar ' + r.external + ', conflicts ' + r.conflicts + ', errors ' + r.errors + '.';
}

/** Stops the sync early if this account cannot open the calendar. Nothing is changed in the sheet. */
function assertCalendarAccess_(calId) {
  try {
    resolveCal_(calId);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    // write to the sync log at most once every 6 hours, so the 1-5 minute timer does not flood it
    var props = PropertiesService.getDocumentProperties();
    var last = parseInt(props.getProperty('ACCESS_FAIL_AT'), 10) || 0;
    if (Date.now() - last > 6 * 3600000) {
      logSync_('', '', 'Sheet → Calendar', 'Access check', 'Failed', msg);
      props.setProperty('ACCESS_FAIL_AT', String(Date.now()));
    }
    throw new Error('Calendar sync stopped: ' + (currentUserEmail_() || 'this account') +
      ' cannot open the calendar "' + calId + '". Ask the calendar owner to share it with ' +
      '"Make changes to events" permission. No meetings were changed.');
  }
}



function eventBody_(m) {
  var sMin = timeToMinutes_(m.start), eMin = timeToMinutes_(m.end);
  if (eMin === null || eMin <= sMin) eMin = Math.min(sMin + 60, 1439);
  var day = isoDate_(toDate_(m.date));
  function dt(mins) { return day + 'T' + ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2) + ':00'; }
  var lines = ['Meeting ID: ' + m.id];
  if (trim_(m.custName) || trim_(m.company)) lines.push('Customer: ' + trim_(m.custName) + (trim_(m.company) ? ' (' + trim_(m.company) + ')' : ''));
  if (trim_(m.custMobile)) lines.push('Mobile: ' + trim_(m.custMobile));
  if (trim_(m.type)) lines.push('Type: ' + trim_(m.type) + ' | Customer type: ' + trim_(m.priority));
  if (trim_(m.participants)) lines.push('Participants: ' + trim_(m.participants));
  if (trim_(m.agenda)) lines.push('', 'Agenda: ' + trim_(m.agenda));
  if (trim_(m.mapLink)) lines.push('', 'Map: ' + trim_(m.mapLink));
  if (trim_(m.meetLink)) lines.push('Meet: ' + trim_(m.meetLink));
  return {
    summary: trim_(m.title),
    location: trim_(m.address),
    description: lines.join('\n'),
    start: { dateTime: dt(sMin), timeZone: KE.TZ },
    end: { dateTime: dt(eMin), timeZone: KE.TZ },
    extendedProperties: { 'private': { keMeetingId: String(m.id) } }
  };
}



var EVPRE_ = {};
/** Loads all events of the sync window in 1-2 calls, so we do not call Google once per meeting. */
function prefetchEvents_(calId, rows) {
  EVPRE_[calId] = null;
  var lo = null, hi = null;
  rows.forEach(function (r) {
    var d = toDate_(r[MC.date - 1]);
    if (!d || !trim_(r[MC.id - 1])) return;
    var t = d.getTime();
    if (lo === null || t < lo) lo = t;
    if (hi === null || t > hi) hi = t;
  });
  if (lo === null) return;
  lo -= 2 * DAY_MS_; hi += 3 * DAY_MS_;
  if (hi - lo > 400 * DAY_MS_) hi = lo + 400 * DAY_MS_;
  var map = {}, token = null, pages = 0;
  try {
    do {
      var opts = { timeMin: new Date(lo).toISOString(), timeMax: new Date(hi).toISOString(), maxResults: 2500, showDeleted: false };
      if (token) opts.pageToken = token;
      var res = Calendar.Events.list(calId, opts);
      (res.items || []).forEach(function (ev) { map[ev.id] = ev; });
      token = res.nextPageToken; pages++;
    } while (token && pages < 4);
    EVPRE_[calId] = map;
  } catch (e) { EVPRE_[calId] = null; }
}

function getEvent_(calId, eid) {
  var pre = EVPRE_[calId];
  if (pre && pre[eid]) return pre[eid];          // found in the prefetched list: no API call
  try {                                          // not found: ask Google to confirm (e.g. deleted event)
    return retry_(function () { return Calendar.Events.get(calId, eid); }, 3);
  } catch (e) {
    if (/Not Found|404|410|deleted/i.test(String(e.message || e))) return null;
    throw e;
  }
}

function removeEvent_(calId, eid) {
  try { Calendar.Events.remove(calId, eid); } catch (e) {
    if (!/Not Found|404|410|deleted/i.test(String(e.message || e))) throw e;
  }
}

function eventDiffers_(m, ev) {
  if (!ev.start || !ev.start.dateTime || !ev.end || !ev.end.dateTime) return false;   // all-day events: ignore
  var s = new Date(ev.start.dateTime), e = new Date(ev.end.dateTime);
  if (isoDate_(s) !== isoDate_(toDate_(m.date))) return true;
  if (timeToMinutes_(hm_(s)) !== timeToMinutes_(m.start)) return true;
  var me = timeToMinutes_(m.end);
  if (me !== null && timeToMinutes_(hm_(e)) !== me) return true;
  if (trim_(ev.summary) !== trim_(m.title)) return true;
  if (trim_(ev.location) !== trim_(m.address)) return true;
  return false;
}

function pullFromEvent_(sh, rowNum, m, ev) {
  var s = new Date(ev.start.dateTime), e = new Date(ev.end.dateTime);
  var timeChanged = isoDate_(s) !== isoDate_(toDate_(m.date)) || timeToMinutes_(hm_(s)) !== timeToMinutes_(m.start);
  var patch = { id: m.id, title: trim_(ev.summary) || m.title, date: isoDate_(s), start: hm_(s), end: hm_(e) };
  if (trim_(ev.location)) patch.address = trim_(ev.location);
  if (timeChanged && trim_(m.status) === 'Scheduled') patch.status = 'Rescheduled';
  saveEntity_('meeting', patch);
  sh.getRange(rowNum, MC.syncStatus).setValue('Synced');
}

function dashResolveConflict(id, choice) {
  requireRole_();
  return withLock_(function () {
    var sh = getSheet_(KE.SHEET.MEETINGS), row = findRow_(sh, id);
    if (!row) throw new Error('Meeting not found.');
    var m = rowToObj_(sh.getRange(row, 1, 1, MEETING_FIELDS.length).getValues()[0], MEETING_FIELDS);
    if (choice === 'sheet') {
      sh.getRange(row, MC.updatedAt).setValue(new Date());
      sh.getRange(row, MC.syncStatus).setValue('Pending Update');
    } else if (choice === 'calendar') {
      var ev = getEvent_(calId_(), trim_(m.eventId));
      if (!ev || ev.status === 'cancelled') {
        sh.getRange(row, MC.status).setValue('Deleted Externally');
        sh.getRange(row, MC.syncStatus).setValue('Deleted Externally');
      } else {
        pullFromEvent_(sh, row, m, ev);
      }
    } else {
      throw new Error('Unknown choice.');
    }
    logAudit_('Resolve conflict', KE.SHEET.MEETINGS, id, 'Sync', 'SYNC CONFLICT', choice === 'sheet' ? 'Kept web app version' : 'Used calendar version');
    syncAll_();
    return 'Resolved.';
  });
}

/* ============================ AUTOMATION ================================ */

function installTriggers_() {
  removeTriggers_();
  var s = getSettings_();
  var min = parseInt(s[SK.SYNC_MIN], 10);
  if ([1, 5, 10, 15, 30].indexOf(min) === -1) min = 5;
  ScriptApp.newTrigger('autoSync').timeBased().everyMinutes(min).create();
  var h = parseInt(s[SK.REM_HOUR], 10);
  if (isNaN(h) || h < 0 || h > 23) h = 9;
  ScriptApp.newTrigger('dailyReminderRun').timeBased().everyDays(1).atHour(h).inTimezone(KE.TZ).create();
  cacheDrop_(['KE_AUTO']);
  // Viewers cannot see triggers, so the state is also stored as a property
  PropertiesService.getDocumentProperties().setProperty('AUTO_STATE', JSON.stringify({ sync: true, daily: true }));
}

function removeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_FNS_.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  cacheDrop_(['KE_AUTO']);
  PropertiesService.getDocumentProperties().setProperty('AUTO_STATE', JSON.stringify({ sync: false, daily: false }));
}

function triggerInfo_() {
  var props = PropertiesService.getDocumentProperties();
  // A viewer's account owns no triggers, so it reads the saved state instead
  if (getUserRole_() === 'VIEWER') {
    try { return JSON.parse(props.getProperty('AUTO_STATE') || '{"sync":false,"daily":false}'); }
    catch (e) { return { sync: false, daily: false }; }
  }
  var hit = cacheGet_('KE_AUTO');
  if (hit) return hit;
  var names = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  var out = { sync: names.indexOf('autoSync') >= 0, daily: names.indexOf('dailyReminderRun') >= 0 };
  cachePut_('KE_AUTO', out, 300);
  if (out.sync || out.daily) props.setProperty('AUTO_STATE', JSON.stringify(out));
  return out;
}
function dashToggleAuto(on) {
  requireRole_();
  withLock_(function () { if (on) installTriggers_(); else removeTriggers_(); });
  logAudit_('Automation', 'Workbook', 'TRIGGERS', 'Automation', '', on ? 'ON' : 'OFF');
  return true;
}

/** Run once from the editor to grant Calendar / Drive / Mail permissions and link the spreadsheet. */
function authorizeAll() {
  linkSpreadsheet_();
  Calendar.CalendarList.list({ maxResults: 1 });
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  ScriptApp.getProjectTriggers();
  Logger.log('All permissions granted. Spreadsheet linked.');
}

/* ============================ DAILY BRIEF =============================== */

function dailyReminderRun() {    // time-driven trigger: maintenance first, then the e-mail
  try { runMaintenance_(); } catch (e) { Logger.log('maintenance: ' + e); }
  try {
    if (yes_(getSetting_(SK.REM_ENABLED, 'Yes'), true)) sendReminders_(false);
  } catch (e2) { Logger.log('dailyReminderRun: ' + e2); }
}

function dashSendTestReminder() { requireRole_(); denyViewer_(); return sendReminders_(true); }

function recipients_(s) {
  var to = [], who = String(s[SK.REM_TO] || 'Both');
  if (who !== 'EA') to = to.concat(emailList_(s[SK.OWNER_EMAIL]));
  if (who !== 'Owner') to = to.concat(emailList_(s[SK.EA_EMAIL]));
  var seen = {};
  return to.filter(function (x) { if (seen[x]) return false; seen[x] = 1; return true; });
}

function sendReminders_(force) {
  var s = getSettings_(), to = recipients_(s);
  if (!to.length) throw new Error('No recipient e-mail set. Fill Owner / EA e-mail in Settings.');
  var tMs = todayDate_().getTime();
  function dms(v) { var d = toDate_(v); return d ? d.getTime() : null; }
  function live(m) { return m.status !== 'Cancelled' && m.status !== 'Deleted Externally'; }
  function byStart(a, b) { return (timeToMinutes_(a.start) || 0) - (timeToMinutes_(b.start) || 0); }
  var meetings = loadObjs_(KE.SHEET.MEETINGS, MEETING_FIELDS);
  var actions = loadObjs_(KE.SHEET.ACTIONS, ACTION_FIELDS);
  var follow = loadObjs_(KE.SHEET.FOLLOWUPS, FOLLOWUP_FIELDS);

  var todayM = meetings.filter(function (m) { return live(m) && dms(m.date) === tMs; }).sort(byStart);
  var tomM = meetings.filter(function (m) { return live(m) && dms(m.date) === tMs + DAY_MS_; }).sort(byStart);
  var open = actions.filter(function (a) { return a.status !== 'Completed'; });
  var overdue = open.filter(function (a) { var t = dms(a.deadline); return a.status === 'Overdue' || (t !== null && t < tMs); });
  var dueToday = open.filter(function (a) { return dms(a.deadline) === tMs; });
  var fDue = follow.filter(function (f) { var t = dms(f.date); return f.status === 'Pending' && t !== null && t <= tMs; });
  var loc = meetings.filter(function (m) { var t = dms(m.date); return live(m) && t !== null && t >= tMs && (m.locStatus === 'Pending' || m.locStatus === 'Request Sent'); });

  var total = todayM.length + tomM.length + overdue.length + dueToday.length + fDue.length + loc.length;
  if (!total && !force) return 'Nothing to report.';

  function mrow(m) { return escH_(fmtTime_(m.start)) + ' — <b>' + escH_(m.title) + '</b> (' + escH_(m.custName || m.company || '—') + ')'; }
  function arow(a) { return escH_(a.task) + ' — ' + escH_(a.assignedTo || 'unassigned') + ', due ' + escH_(fmtDate_(a.deadline) || '—'); }
  function section(title, items, color) {
    if (!items.length) return '';
    return '<h3 style="color:' + color + ';margin:18px 0 6px">' + title + ' (' + items.length + ')</h3><ul style="margin:0;padding-left:18px">' +
      items.map(function (x) { return '<li style="margin:3px 0">' + x + '</li>'; }).join('') + '</ul>';
  }
  var b = brand_();
  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1b2437"><div style="background:' + b.c1 + ';color:#fff;padding:14px 18px;border-radius:8px;border-bottom:4px solid ' + b.c2 + '"><b>' + escH_(b.company) +
    '</b><br>Daily brief — ' + Utilities.formatDate(new Date(), KE.TZ, 'EEEE, dd MMM yyyy') + '</div>' +
    section("Today's meetings", todayM.map(mrow), '#1F5FBF') +
    section("Tomorrow's meetings", tomM.map(mrow), '#1F5FBF') +
    section('Overdue action items', overdue.map(arow), '#C0392B') +
    section('Action items due today', dueToday.map(arow), '#B9770E') +
    section('Follow-ups due', fDue.map(function (f) { return escH_(f.subject) + ' — ' + escH_(f.custName || '') + ' (' + escH_(fmtDate_(f.date)) + ')'; }), '#B9770E') +
    section('Customer locations pending', loc.map(function (m) { return escH_(m.title) + ' — ' + escH_(m.custName || '') + ' (' + escH_(fmtDate_(m.date)) + ')'; }), '#B9770E') +
    (total ? '' : '<p>Nothing needs attention today.</p>') + '</div>';

  MailApp.sendEmail({ to: to.join(','), subject: b.company + ' — Daily brief ' + Utilities.formatDate(new Date(), KE.TZ, 'dd-MMM-yyyy'), htmlBody: html, name: b.company });
  logAudit_('Reminder e-mail', 'Workbook', 'REMINDER', 'Recipients', '', to.join(', '));
  return 'Reminder e-mail sent to ' + to.join(', ') + '.';
}

/* ============================ PDF REPORT ================================ */

function dashReport(fromISO, toISO) {
  requireRole_();
  var from = toDate_(fromISO), to = toDate_(toISO);
  if (!from || !to) throw new Error('Choose both dates.');
  if (to.getTime() < from.getTime()) throw new Error('"To" date must be on or after "From" date.');
  var f = from.getTime(), t = to.getTime();
  function inR(v) { var d = toDate_(v); return d && d.getTime() >= f && d.getTime() <= t; }
  var meetings = loadObjs_(KE.SHEET.MEETINGS, MEETING_FIELDS).filter(function (m) { return inR(m.date); })
    .sort(function (a, b) { return toDate_(a.date) - toDate_(b.date) || (timeToMinutes_(a.start) || 0) - (timeToMinutes_(b.start) || 0); });
  var actions = loadObjs_(KE.SHEET.ACTIONS, ACTION_FIELDS).filter(function (a) { return inR(a.deadline) || inR(a.assignedDate); });
  var follow = loadObjs_(KE.SHEET.FOLLOWUPS, FOLLOWUP_FIELDS).filter(function (x) { return inR(x.date); });
  var counts = {};
  meetings.forEach(function (m) { counts[m.status || 'Scheduled'] = (counts[m.status || 'Scheduled'] || 0) + 1; });
  var b = brand_();

  function table(head, rows) {
    if (!rows.length) return '<p style="color:#777">No records.</p>';
    return '<table style="width:100%;border-collapse:collapse;font-size:10px"><tr>' + head.map(function (h) { return '<th style="background:' + b.c1 + ';color:#fff;padding:5px;text-align:left">' + h + '</th>'; }).join('') + '</tr>' +
      rows.map(function (r) { return '<tr>' + r.map(function (c) { return '<td style="border-bottom:1px solid #ddd;padding:5px;vertical-align:top">' + escH_(c) + '</td>'; }).join('') + '</tr>'; }).join('') + '</table>';
  }
  var html = '<div style="font-family:Arial,sans-serif;color:#1b2437"><h1 style="margin:0;color:' + b.c1 + '">' + escH_(b.company) + '</h1>' +
    '<div style="color:' + b.c2 + ';font-weight:bold;margin-bottom:6px">' + escH_(b.tagline) + '</div>' +
    '<h3 style="margin:2px 0 12px">Meeting Report: ' + fmtDate_(from) + ' to ' + fmtDate_(to) + '</h3>' +
    '<p><b>Total meetings:</b> ' + meetings.length + ' &nbsp; ' + Object.keys(counts).map(function (k) { return '<b>' + escH_(k) + ':</b> ' + counts[k]; }).join(' &nbsp; ') + '</p>' +
    '<h3>Meetings</h3>' + table(['Date', 'Time', 'Meeting', 'Customer', 'Type', 'Status'],
      meetings.map(function (m) { return [fmtDate_(m.date), fmtTime_(m.start), m.title, m.custName || m.company, m.type, m.status]; })) +
    '<h3>Action items</h3>' + table(['Task', 'Assigned to', 'Priority', 'Deadline', 'Status'],
      actions.map(function (a) { return [a.task, a.assignedTo, a.priority, fmtDate_(a.deadline), a.status]; })) +
    '<h3>Follow-ups</h3>' + table(['Subject', 'Customer', 'Date', 'Responsible', 'Status'],
      follow.map(function (x) { return [x.subject, x.custName, fmtDate_(x.date), x.person, x.status]; })) +
    '<p style="color:#888;font-size:9px;margin-top:18px">Generated ' + Utilities.formatDate(new Date(), KE.TZ, 'dd-MMM-yyyy HH:mm') + '</p></div>';

  var name = 'Meeting Report ' + isoDate_(from) + ' to ' + isoDate_(to) + '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  var folderName = trim_(getSetting_(XK.FOLDER)) || 'KISAN ENTERPRISES Reports';
  var it = DriveApp.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  var file = folder.createFile(blob);
  logAudit_('Report', 'Workbook', 'REPORT', 'Range', '', isoDate_(from) + ' to ' + isoDate_(to));
  return { url: file.getUrl(), name: name, count: meetings.length, folder: folderName };
}

/* ============================ HEALTH + SYNC STATUS ====================== */

function dashHealth() {
  requireRole_();
  var s = getSettings_(), out = [];
  function add(ok, text, fix) { out.push({ ok: !!ok, text: text, fix: fix || '' }); }
  add(trim_(s[SK.OWNER_NAME]), 'Owner name set', 'Fill "Owner name" in Contacts');
  add(trim_(s[SK.OWNER_EMAIL]), 'Owner e-mail set (access control and reminders)', 'Fill "Owner Google e-mail" in Contacts');
  add(trim_(s[SK.EA_EMAIL]), 'EA e-mail set', 'Fill "EA Google e-mail" in Contacts');
  add(trim_(s[SK.OWNER_WA]), 'Owner WhatsApp number set', 'Fill "Owner WhatsApp" in Contacts');
  add(trim_(s[SK.OFFICE_ADDR]), 'Office address set', 'Fill "Office address" in Contacts');
  if (getUserRole_() === 'VIEWER') {
    // A viewer's own Google account may not reach the owner's calendar, so do not report a false failure
    add(true, 'Google Calendar: ' + calId_() + ' (connection check is shown to Owner / EA only)');
  } else {
    try {
      var cal = resolveCal_(trim_(s[SK.CAL_ID]) || 'primary');
      add(true, 'Google Calendar reachable: ' + cal.name);
    } catch (e) {
      add(false, 'Google Calendar not reachable', 'Check "Calendar ID" in Calendar & Automation, and run authorizeAll once');
    }
  }
  var t = triggerInfo_();
  add(t.sync && t.daily, 'Automation is ON (sync, alerts, daily brief, maintenance)', 'Open Calendar Sync and turn on automation');
  add(getSS_().getSpreadsheetTimeZone() === KE.TZ, 'Spreadsheet time zone is ' + KE.TZ, 'Run Setup workbook from the sheet menu');
  return out;
}

function dashSyncStatus() {
  requireRole_();
  var meetings = loadObjs_(KE.SHEET.MEETINGS, MEETING_FIELDS);
  var c = { synced: 0, pending: 0, conflict: 0, error: 0, deleted: 0 }, conflicts = [];
  meetings.forEach(function (m) {
    if (m.status === 'Cancelled' && !trim_(m.eventId)) return;
    var s = trim_(m.syncStatus);
    if (s === 'Synced') c.synced++;
    else if (s === 'SYNC CONFLICT') { c.conflict++; conflicts.push({ id: m.id, title: m.title, date: fmtDate_(m.date) }); }
    else if (s === 'Error') c.error++;
    else if (s === 'Deleted Externally') c.deleted++;
    else c.pending++;
  });
  var log = [], sl = getSheet_(KE.SHEET.SYNCLOG), last = sl.getLastRow();
  if (last >= 2) {
    var from = Math.max(2, last - 11);
    sl.getRange(from, 1, last - from + 1, SYNCLOG_FIELDS.length).getValues().reverse().forEach(function (r) {
      log.push({
        ts: Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], KE.TZ, 'dd-MMM HH:mm') : String(r[0]),
        meetingId: String(r[1]), dir: String(r[3]), op: String(r[4]), status: String(r[5]), error: String(r[6])
      });
    });
  }
  var t = triggerInfo_();
  return { counts: c, conflicts: conflicts, log: log, auto: t.sync, every: String(getSetting_(SK.SYNC_MIN, 5)), calendar: calId_() };
}
