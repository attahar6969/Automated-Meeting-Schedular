function myFunction() {
  
}
/**
 * CalendarSync.gs   (NEW FILE)
 * Multi-calendar synchronisation:
 *   - Registry sheet "Calendar Sync Registry": ONE row per (Meeting ID + Calendar ID)
 *   - syncAll_()            : replaces the old single-calendar syncAll_/syncOne_ (auto + "Sync now")
 *   - Reset & Re-Sync       : dashResetPreview / dashResetBegin / dashResetBatch / dashResetEnd
 *   - Cancel / delete       : applied to EVERY registered calendar
 *
 * The Meeting Register columns "Calendar Event ID" and "Calendar Sync Status" are kept,
 * but now only MIRROR the currently configured calendar, so the existing UI keeps working.
 * Events are never deleted because the Calendar ID changed.
 */

var REG_ = {
  SHEET: 'Calendar Sync Registry',
  HEAD: ['Meeting ID', 'Calendar ID', 'Google Calendar Event ID', 'Sync Status', 'Last Synced At',
         'Last Sync Error', 'Meeting Last Modified', 'Content Fingerprint'],
  NCOL: 8
};
var RESET_BATCH_ = 15;            // meetings per server call
var RESET_FLAG_ = 'RESET_RUNNING';
var CALRES_ = {};

function errMsg_(e) { return String(e && e.message ? e.message : e); }

/* ============================ REGISTRY SHEET ============================ */

function regSheet_() {
  var ss = getSS_(), sh = ss.getSheetByName(REG_.SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(REG_.SHEET);
  var rows = sh.getMaxRows();
  sh.getRange(1, 1, rows, 4).setNumberFormat('@');          // IDs must stay text
  sh.getRange(1, 8, rows, 1).setNumberFormat('@');
  sh.getRange(2, 5, rows - 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  sh.getRange(2, 7, rows - 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  sh.getRange(1, 1, 1, REG_.NCOL).setValues([REG_.HEAD]);
  styleHeader_(sh.getRange(1, 1, 1, REG_.NCOL));
  sh.setFrozenRows(1);
  [130, 230, 240, 150, 150, 380, 150, 260].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setTabColor('#8b95a7');
  return sh;
}

/** Loads the registry. map["MTG|CAL"] = entry ; by["MTG"] = [entries]. First row wins on duplicates. */
function loadReg_() {
  var sh = regSheet_(), last = sh.getLastRow(), map = {}, by = {};
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, REG_.NCOL).getValues();
    for (var i = 0; i < v.length; i++) {
      var mid = trim_(v[i][0]), cal = trim_(v[i][1]);
      if (!mid || !cal || map[mid + '|' + cal]) continue;
      var e = { row: i + 2, mid: mid, cal: cal, eid: trim_(v[i][2]), status: trim_(v[i][3]),
        at: v[i][4], err: trim_(v[i][5]), mod: v[i][6], hash: trim_(v[i][7]) };
      map[mid + '|' + cal] = e;
      (by[mid] = by[mid] || []).push(e);
    }
  }
  return { sh: sh, map: map, by: by };
}

/** Creates or updates the single registry row of (meeting, calendar) and writes it immediately. */
function putReg_(R, mid, cal, f) {
  var key = mid + '|' + cal, e = R.map[key];
  if (!e) {
    e = { row: Math.max(2, R.sh.getLastRow() + 1), mid: mid, cal: cal, eid: '', status: '', at: '', err: '', mod: '', hash: '' };
    R.map[key] = e;
    (R.by[mid] = R.by[mid] || []).push(e);
  }
  Object.keys(f).forEach(function (k) { e[k] = f[k]; });
  if (e.row > R.sh.getMaxRows()) R.sh.insertRowsAfter(R.sh.getMaxRows(), 200);
  R.sh.getRange(e.row, 1, 1, REG_.NCOL).setValues([[e.mid, e.cal, e.eid, e.status, e.at, e.err, e.mod, e.hash]]);
  return e;
}

/** Called by renumberMeetings(): keeps registry rows attached to the renumbered meetings. */
function remapRegistryIds_(map) {
  var sh = regSheet_(), last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, 1, last - 1, 1), v = rng.getValues(), n = 0;
  for (var i = 0; i < v.length; i++) {
    var id = trim_(v[i][0]);
    if (id && map.hasOwnProperty(id) && map[id] !== id) { v[i][0] = map[id]; n++; }
  }
  if (n) rng.setValues(v);
  return n;
}

/* ============================ CALENDAR HELPERS ========================== */

/** Resolves "primary"/blank/e-mail to the real calendar id, so "primary" and the e-mail never become two calendars. */
function resolveCal_(raw) {
  raw = trim_(raw) || 'primary';
  if (CALRES_[raw]) return CALRES_[raw];
  var c = retry_(function () { return Calendar.Calendars.get(raw); }, 2);
  var out = { id: c.id || raw, name: c.summary || c.id || raw };
  CALRES_[raw] = out;
  return out;
}

/** Throws a friendly error unless the calendar opens and (when known) allows editing events. */
function validateCalendar_(raw) {
  var ck = 'KE_CALOK_' + (trim_(raw) || 'primary') + '_' + currentUserEmail_();
  var hit = cacheGet_(ck);
  if (hit) return hit;
  var info;
  try { info = resolveCal_(raw); }
  catch (e) {
    throw new Error('This account (' + (currentUserEmail_() || 'unknown') + ') cannot open the calendar "' + (trim_(raw) || 'primary') +
      '". Check the Calendar ID in Settings and share the calendar with "Make changes to events". Nothing was changed.');
  }
  var role = '';
  try { role = Calendar.CalendarList.get(info.id).accessRole || ''; } catch (e2) { role = ''; }
  if (role && ['owner', 'writer'].indexOf(role) === -1) {
    throw new Error('This account only has "' + role + '" access to "' + info.name + '". It needs "Make changes to events". Nothing was changed.');
  }
  cachePut_(ck, info, 300);
  return info;
}
function fingerprint_(body) {
  var s = JSON.stringify([body.summary, body.location, body.description, body.start, body.end]);
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/** Finds an event this app created for the meeting in that calendar (private extended property). Never deletes. */
function findByMeetingId_(calId, mid) {
  var res = retry_(function () {
    return Calendar.Events.list(calId, { privateExtendedProperty: 'keMeetingId=' + mid, showDeleted: false, maxResults: 50 });
  }, 3);
  // Double-check every result, so a wrong match can never be returned even if the filter is ignored
  var items = (res.items || []).filter(function (x) {
    var p = x.extendedProperties && x.extendedProperties['private'];
    return x.status !== 'cancelled' && p && String(p.keMeetingId) === String(mid);
  });
  items.sort(function (a, b) { return String(a.created).localeCompare(String(b.created)); });
  return items.length ? items[0] : null;
}




function findSafe_(calId, mid) { try { return findByMeetingId_(calId, mid); } catch (e) { return null; } }

/** True when another meeting's registry row in the same calendar already uses this event id. */
function claimedByOther_(R, calId, mid, eid) {
  var keys = Object.keys(R.map);
  for (var i = 0; i < keys.length; i++) {
    var x = R.map[keys[i]];
    if (x.cal === calId && x.eid === eid && x.mid !== mid) return true;
  }
  return false;
}

function isStopError_(msg) {
  return /403|forbidden|insufficient|do not have permission|writer access|quota|rate limit|usage limit/i.test(msg);
}

function logCalFailOnce_(calId, mid, msg) {          // at most one log row per calendar per 6 hours
  var p = PropertiesService.getDocumentProperties(), map = {};
  try { map = JSON.parse(p.getProperty('CALFAIL') || '{}'); } catch (e) { map = {}; }
  if (Date.now() - (map[calId] || 0) < 6 * 3600000) return;
  map[calId] = Date.now();
  p.setProperty('CALFAIL', JSON.stringify(map));
  logSync_(mid, '', 'Sheet → Calendar', 'Update other calendar', 'Failed', 'Calendar ' + calId + ': ' + msg);
}

/** Mirrors the ACTIVE calendar's event id and status into the Meeting Register (only writes on change). */
function setMirror_(sh, rowNum, m, eid, status) {
  if (trim_(m.eventId) !== trim_(eid)) { sh.getRange(rowNum, MC.eventId).setValue(eid); m.eventId = eid; }
  if (trim_(m.syncStatus) !== status) { sh.getRange(rowNum, MC.syncStatus).setValue(status); m.syncStatus = status; }
}

/* ============================ CORE: ONE MEETING, ONE CALENDAR =========== */

/**
 * Makes sure the meeting has exactly one correct event in this calendar.
 * Order: registry id -> validate event -> recover by private property -> create.
 * Returns {op:'created'|'recovered'|'updated'|'unchanged', ev}. Throws on API errors.
 */
function ensureEvent_(R, m, cal) {
  var body = eventBody_(m), fp = fingerprint_(body);
  var e = R.map[m.id + '|' + cal.id], ev = null, recovered = false;
   if (e && e.eid) {
    ev = getEvent_(cal.id, e.eid);
    if (ev && ev.status === 'cancelled') ev = null;
    if (ev && claimedByOther_(R, cal.id, m.id, e.eid)) ev = null;   // shared event id: do not overwrite it
  }
  if (!ev) { ev = findByMeetingId_(cal.id, m.id); recovered = !!ev; }
  var now = new Date();

  if (ev) {
    var need = recovered || !e || e.hash !== fp || eventDiffers_(m, ev);
    if (need) retry_(function () { return Calendar.Events.patch(body, cal.id, ev.id); }, 3);
    putReg_(R, m.id, cal.id, { eid: ev.id, status: 'Synced', hash: fp, at: now, err: '', mod: m.updatedAt || '' });
    return { op: recovered ? 'recovered' : (need ? 'updated' : 'unchanged'), ev: ev };
  }

  var cbody = eventBody_(m), opts = {};
  if (m.mode === 'Virtual' && !trim_(m.meetLink)) {
    cbody.conferenceData = { createRequest: { requestId: m.id + '-' + now.getTime(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
    opts.conferenceDataVersion = 1;
  }
  var created = null;
  // No automatic retry on insert (a timeout may hide a success); after a failure we look for the event before giving up.
  try { created = Calendar.Events.insert(cbody, cal.id, opts); }
  catch (e1) {
    created = findSafe_(cal.id, m.id);
    if (!created && cbody.conferenceData) {
      delete cbody.conferenceData;
      try { created = Calendar.Events.insert(cbody, cal.id); }
      catch (e2) { created = findSafe_(cal.id, m.id); if (!created) throw e2; }
    } else if (!created) { throw e1; }
  }
  putReg_(R, m.id, cal.id, { eid: created.id, status: 'Synced', hash: fp, at: now, err: '', mod: m.updatedAt || '' });
  logSync_(m.id, created.id, 'Sheet → Calendar', 'Create event', 'Success', 'Calendar ' + cal.id);
  Utilities.sleep(100);
  return { op: 'created', ev: created };
}

/** Applies the cancellation policy (remove event) to every registered calendar. */
function cancelEverywhere_(R, m, active, failed) {
  var out = { removed: 0, errors: 0 }, now = new Date(), entries = (R.by[m.id] || []).slice();
  entries.forEach(function (x) {
    if (!x.eid) {
      if (x.status !== 'Cancelled') putReg_(R, m.id, x.cal, { status: 'Cancelled', at: now, err: '' });
      return;
    }
    if (failed && failed[x.cal]) { out.errors++; return; }
    try {
      removeEvent_(x.cal, x.eid);
      putReg_(R, m.id, x.cal, { eid: '', status: 'Cancelled', hash: '', at: now, err: '' });
      out.removed++;
      logSync_(m.id, x.eid, 'Sheet → Calendar', 'Delete event (cancelled)', 'Success', 'Calendar ' + x.cal);
    } catch (e) {
      var msg = errMsg_(e);
      putReg_(R, m.id, x.cal, { status: 'Error', err: msg.substring(0, 500), at: now });   // eid kept for retry
      out.errors++;
      if (failed) failed[x.cal] = 1;
      logCalFailOnce_(x.cal, m.id, msg);
    }
  });
  if (!entries.length && trim_(m.eventId)) {       // legacy row that never reached the registry
    try { removeEvent_(active.id, trim_(m.eventId)); out.removed++; } catch (e2) { out.errors++; }
  }
  return out;
}

/** Pushes edited meeting details to every NON-active calendar that has an event for it. */
function propagate_(R, m, active, failed, res) {
  var entries = R.by[m.id] || [], fp = null;
  entries.forEach(function (x) {
    if (x.cal === active.id || !x.eid) return;
    if (['Synced', 'Error'].indexOf(x.status) === -1) return;
    if (failed[x.cal]) return;
    if (fp === null) fp = fingerprint_(eventBody_(m));
    if (x.status === 'Synced' && x.hash === fp) return;
    try {
      retry_(function () { return Calendar.Events.patch(eventBody_(m), x.cal, x.eid); }, 3);
      putReg_(R, m.id, x.cal, { status: 'Synced', hash: fp, at: new Date(), err: '', mod: m.updatedAt || '' });
      res.updated++;
      logSync_(m.id, x.eid, 'Sheet → Calendar', 'Update event', 'Success', 'Calendar ' + x.cal);
    } catch (e) {
      var msg = errMsg_(e);
      if (/Not Found|404|410|deleted/i.test(msg)) {
        putReg_(R, m.id, x.cal, { status: 'Deleted Externally', err: 'Event not found in that calendar. Reset & Re-Sync on that calendar will recreate it.', at: new Date() });
      } else {
        putReg_(R, m.id, x.cal, { status: 'Error', err: msg.substring(0, 500), at: new Date() });
        failed[x.cal] = 1;
        res.errors++;
        logCalFailOnce_(x.cal, m.id, msg);
      }
    }
  });
}

/* ============================ AUTOMATIC / "SYNC NOW" ==================== */

function resetRunning_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(RESET_FLAG_);
  if (!raw) return null;
  try { var f = JSON.parse(raw); if (Date.now() - f.t < 600000) return f; } catch (e) { /* stale */ }
  return null;
}

function syncAll_() {
  if (typeof Calendar === 'undefined') throw new Error('Advanced Calendar service is not enabled (appsscript.json).');
  var rr = resetRunning_();
  if (rr) throw new Error('A Reset & Re-Sync started by ' + rr.by + ' is running. Try again in a few minutes.');
  var calRaw = calId_();
  assertCalendarAccess_(calRaw);
  var active = resolveCal_(calRaw);
  var sh = getSheet_(KE.SHEET.MEETINGS);
  var R = loadReg_();
  migrateLegacyIds_(R, sh, active, false);

  var rows = readRows_(sh, MEETING_FIELDS.length);
  var res = { created: 0, updated: 0, pulled: 0, removed: 0, external: 0, conflicts: 0, errors: 0 };
  var todayMs = todayDate_().getTime(), failed = {}, t0 = Date.now();
  prefetchEvents_(active.id, rows.filter(function (r) {
  var d = toDate_(r[MC.date - 1]);
  return d && d.getTime() >= todayMs - DAY_MS_;
}));
  for (var i = 0; i < rows.length; i++) {
    if (Date.now() - t0 > 270000) break;                 // time guard: the next run continues
    var m = rowToObj_(rows[i], MEETING_FIELDS);
    if (!trim_(m.id)) continue;
    try {
      syncMeeting_(sh, i + 2, m, R, active, res, todayMs, failed);
    } catch (e) {
      var em = errMsg_(e);
      if (isStopError_(em)) {
        logSync_(m.id, m.eventId, 'Sheet → Calendar', 'Sync', 'Failed', em);
        throw new Error('Calendar sync stopped: this account cannot change events (or the quota is used up). ' + em.substring(0, 200));
      }
      res.errors++;
      putReg_(R, m.id, active.id, { status: 'Error', err: em.substring(0, 500), at: new Date() });
      setMirror_(sh, i + 2, m, trim_(m.eventId), 'Error');
      logSync_(m.id, m.eventId, 'Sheet → Calendar', 'Sync', 'Failed', em);
    }
  }
  return res;
}

function syncMeeting_(sh, rowNum, m, R, active, res, todayMs, failed) {
  var st = trim_(m.status), sync = trim_(m.syncStatus);
  var d = toDate_(m.date);
  if (!d || timeToMinutes_(m.start) === null) return;
  if (st === 'Deleted Externally') return;
  if (sync === 'SYNC CONFLICT') { res.conflicts++; return; }
  var entries = R.by[m.id] || [];
  var act = R.map[m.id + '|' + active.id];
  var now = new Date();

  if (st === 'Cancelled') {
    var cr = cancelEverywhere_(R, m, active, failed);
    res.removed += cr.removed; res.errors += cr.errors;
    if (cr.errors) setMirror_(sh, rowNum, m, trim_(m.eventId), 'Error');
    else if (cr.removed || trim_(m.eventId)) setMirror_(sh, rowNum, m, '', 'Synced');
    return;
  }

  var aEid = act ? act.eid : '';
  var needsPush = ['Pending Update', 'Not Synced', 'Error', ''].indexOf(sync) >= 0;
  var recent = d.getTime() >= todayMs - DAY_MS_;

  if (!aEid) {
    if (!recent) { setMirror_(sh, rowNum, m, '', 'Not Synced'); return; }     // old meetings: only a Reset & Re-Sync adds them
    var r = ensureEvent_(R, m, active);
    if (r.op === 'created') res.created++; else if (r.op !== 'unchanged') res.updated++;
    setMirror_(sh, rowNum, m, r.ev.id, 'Synced');
    if (r.ev.hangoutLink && !trim_(m.meetLink)) { sh.getRange(rowNum, MC.meetLink).setValue(r.ev.hangoutLink); m.meetLink = r.ev.hangoutLink; }
  } else if (needsPush || (recent && sync === 'Synced' && st !== 'Completed')) {
    var ev = getEvent_(active.id, aEid);
    if (!ev || ev.status === 'cancelled') {
      putReg_(R, m.id, active.id, { status: 'Deleted Externally', err: 'Event deleted in Google Calendar', at: now });
      var alive = entries.some(function (x) { return x.cal !== active.id && x.eid && x.status === 'Synced'; });
      if (!alive) sh.getRange(rowNum, MC.status).setValue('Deleted Externally');   // old behaviour, only when no other calendar still has it
      setMirror_(sh, rowNum, m, aEid, 'Deleted Externally');
      res.external++;
      logSync_(m.id, aEid, 'Calendar → Sheet', 'Event deleted', 'Success', 'Event was deleted in Google Calendar ' + active.id);
      if (!alive) return;
    } else {
      var differs = eventDiffers_(m, ev);
      if (needsPush) {
        var evUpdated = new Date(ev.updated).getTime();
        var rowUpdated = Object.prototype.toString.call(m.updatedAt) === '[object Date]' ? m.updatedAt.getTime() : 0;
        if (differs && evUpdated > rowUpdated + 60000) {
          setMirror_(sh, rowNum, m, aEid, 'SYNC CONFLICT');
          res.conflicts++;
          logSync_(m.id, aEid, 'Both', 'Conflict', 'Conflict', 'Changed in both the web app and Google Calendar.');
          return;
        }
        var body = eventBody_(m);
        retry_(function () { return Calendar.Events.patch(body, active.id, aEid); }, 3);
        putReg_(R, m.id, active.id, { status: 'Synced', hash: fingerprint_(body), at: now, err: '', mod: m.updatedAt || '' });
        setMirror_(sh, rowNum, m, aEid, 'Synced');
        res.updated++;
        logSync_(m.id, aEid, 'Sheet → Calendar', 'Update event', 'Success', '');
      } else if (differs) {
        pullFromEvent_(sh, rowNum, m, ev);
        res.pulled++;
        logSync_(m.id, aEid, 'Calendar → Sheet', 'Pull changes', 'Success', 'Calendar edit copied to the web app');
        m = lookupMeeting_(m.id) || m;             // fresh data for the other calendars
      }
    }
  }
  propagate_(R, m, active, failed, res);           // edits reach every registered calendar
}

/* ============================ MIGRATION ================================= */

/** Registers legacy single "Calendar Event ID" values that really exist in the given calendar. Idempotent. */
function migrateLegacyIds_(R, sh, active, force) {
  var props = PropertiesService.getDocumentProperties(), flag = 'REGMIG_' + active.id;
  if (!force && props.getProperty(flag)) return { migrated: 0, complete: true };
  var rows = readRows_(sh, MEETING_FIELDS.length), n = 0, t0 = Date.now(), complete = true;
  for (var i = 0; i < rows.length; i++) {
    var id = trim_(rows[i][MC.id - 1]), eid = trim_(rows[i][MC.eventId - 1]);
    if (!id || !eid || R.map[id + '|' + active.id]) continue;
    if (Date.now() - t0 > 200000) { complete = false; break; }
    var ev = getEvent_(active.id, eid);
    if (ev && ev.status !== 'cancelled') {
      putReg_(R, id, active.id, { eid: eid, status: 'Synced', at: new Date(), err: '', hash: '', mod: rows[i][MC.updatedAt - 1] || '' });
      n++;
    }
  }
  if (complete) props.setProperty(flag, '1');
  return { migrated: n, complete: complete };
}

/** Run once from the Apps Script editor (also runs automatically). Safe to repeat. */
function migrateCalendarRegistry() {
  requireRole_();
  var msg = withLock_(function () {
    var info = validateCalendar_(calId_());
    var r = migrateLegacyIds_(loadReg_(), getSheet_(KE.SHEET.MEETINGS), info, true);
    return 'Registry ready for "' + info.name + '". Legacy event IDs registered: ' + r.migrated + (r.complete ? '.' : '. Not finished: run again.');
  });
  showAlert_('Calendar registry', msg);
  return msg;
}

/* ============================ RESET & RE-SYNC =========================== */

function dashResetPreview() {
  requireRole_();
  var info = validateCalendar_(calId_());
  var rows = readRows_(getSheet_(KE.SHEET.MEETINGS), MEETING_FIELDS.length);
  var total = 0, eligible = 0, cancelled = 0, deletedExt = 0;
  rows.forEach(function (r) {
    if (!trim_(r[MC.id - 1])) return;
    total++;
    var st = trim_(r[MC.status - 1]);
    if (st === 'Cancelled') cancelled++;
    else if (st === 'Deleted Externally') deletedExt++;
    else if (toDate_(r[MC.date - 1]) && timeToMinutes_(r[MC.start - 1]) !== null) eligible++;
  });
  return { calendarName: info.name, calendarId: info.id, total: total, eligible: eligible, cancelled: cancelled,
    deletedExternally: deletedExt, batch: RESET_BATCH_ };
}

function touchReset_() {
  PropertiesService.getDocumentProperties().setProperty(RESET_FLAG_, JSON.stringify({ t: Date.now(), by: currentUserEmail_() || 'someone' }));
}

function dashResetBegin() {
  requireRole_();
  return withLock_(function () {
    var info = validateCalendar_(calId_());
    var rr = resetRunning_();
    if (rr && rr.by !== (currentUserEmail_() || 'someone')) {
      throw new Error('A Reset & Re-Sync started by ' + rr.by + ' is still running. Try again in a few minutes.');
    }
    touchReset_();
    var sh = getSheet_(KE.SHEET.MEETINGS);
    migrateLegacyIds_(loadReg_(), sh, info, true);
    var total = readRows_(sh, MEETING_FIELDS.length).length;
    logAudit_('Reset & Re-Sync started', KE.SHEET.MEETINGS, 'CALENDAR', 'Calendar', '', info.id);
    return { calId: info.id, total: total, batch: RESET_BATCH_ };
  });
}

function dashResetBatch(calId, cursor, restoreDeleted) {
  requireRole_();
  cursor = Math.max(0, parseInt(cursor, 10) || 0);
  return withLock_(function () {
    var info = validateCalendar_(calId_());
    if (info.id !== trim_(calId)) {
      throw new Error('The Calendar ID in Settings changed while the Reset was running. Stopped safely; nothing was deleted. Start again.');
    }
    touchReset_();
    var sh = getSheet_(KE.SHEET.MEETINGS);
    var rows = readRows_(sh, MEETING_FIELDS.length);
    var R = loadReg_();
    prefetchEvents_(info.id, rows.slice(cursor, cursor + RESET_BATCH_));
    var out = { checked: 0, created: 0, recovered: 0, unchanged: 0, updated: 0, removed: 0, skipped: [], failed: [] };
    var i = cursor, t0 = Date.now(), stopped = '';
    for (; i < rows.length && i < cursor + RESET_BATCH_; i++) {
      if (Date.now() - t0 > 200000) break;
      var m = rowToObj_(rows[i], MEETING_FIELDS);
      if (!trim_(m.id)) continue;
      out.checked++;
      try { resetOne_(sh, i + 2, m, R, info, out, !!restoreDeleted); }
      catch (e) {
        if (e && e.stopRun) { stopped = errMsg_(e); out.checked--; break; }
        throw e;
      }
    }
    out.next = i;
    out.done = !stopped && i >= rows.length;
    out.stopped = stopped;
    return out;
  });
}

function resetOne_(sh, rowNum, m, R, info, out, restoreDeleted) {
  var id = m.id, st = trim_(m.status), title = trim_(m.title);
  function skip(reason) { out.skipped.push({ id: id, title: title, reason: reason }); }
  try {
    if (st === 'Deleted Externally') {
      if (!restoreDeleted) return skip('Status is "Deleted Externally" (not restored)');
      sh.getRange(rowNum, MC.status).setValue('Scheduled');
      logAudit_('Restore meeting', KE.SHEET.MEETINGS, id, 'Meeting Status', 'Deleted Externally', 'Scheduled');
      m.status = 'Scheduled'; st = 'Scheduled';
    }
    if (!toDate_(m.date) || timeToMinutes_(m.start) === null) return skip('Meeting date or start time is missing/invalid');
    if (st === 'Cancelled') {
      var cr = cancelEverywhere_(R, m, info, null);
      out.removed += cr.removed;
      if (cr.errors) throw new Error('Cancelled meeting: could not remove an event from one calendar (kept for retry).');
      return skip('Cancelled: no event created' + (cr.removed ? '; ' + cr.removed + ' existing event(s) removed per the cancellation policy' : ''));
    }
    var r = ensureEvent_(R, m, info);
    out[r.op]++;
    setMirror_(sh, rowNum, m, r.ev.id, 'Synced');
    if (r.ev.hangoutLink && !trim_(m.meetLink)) sh.getRange(rowNum, MC.meetLink).setValue(r.ev.hangoutLink);
    if (r.op !== 'unchanged') logSync_(id, r.ev.id, 'Sheet → Calendar', 'Reset & Re-Sync (' + r.op + ')', 'Success', 'Calendar ' + info.id);
  } catch (e) {
    var msg = errMsg_(e);
    putReg_(R, id, info.id, { status: 'Error', err: msg.substring(0, 500), at: new Date() });
    setMirror_(sh, rowNum, m, trim_(m.eventId), 'Error');
    logSync_(id, '', 'Sheet → Calendar', 'Reset & Re-Sync', 'Failed', msg);
    if (isStopError_(msg)) {
      var se = new Error('Stopped at ' + id + ': ' + msg.substring(0, 250) + ' (permission or quota problem)');
      se.stopRun = true;
      throw se;
    }
    out.failed.push({ id: id, title: title, error: msg.substring(0, 300) });
  }
}

function dashResetEnd(summary) {
  requireRole_();
  PropertiesService.getDocumentProperties().deleteProperty(RESET_FLAG_);
  logAudit_('Reset & Re-Sync finished', KE.SHEET.MEETINGS, 'CALENDAR', 'Result', '', String(summary || '').substring(0, 400));
  return true;
}

/* ============================ DELETE A MEETING ========================== */

/** Used by dashDelete: removes the meeting's events from EVERY registered calendar. */
function removeMeetingEvents_(meetingId, legacyEid) {
  var R = loadReg_(), entries = (R.by[meetingId] || []).slice(), gone = [];
  entries.forEach(function (x) {
    if (x.eid) {
      try { removeEvent_(x.cal, x.eid); }
      catch (e) {
        putReg_(R, meetingId, x.cal, { status: 'Orphan (meeting deleted)', err: errMsg_(e).substring(0, 500), at: new Date() });
        logSync_(meetingId, x.eid, 'Sheet → Calendar', 'Delete event', 'Failed', 'Calendar ' + x.cal + ': ' + errMsg_(e));
        return;                                  // keep the row so an administrator can see the leftover event
      }
    }
    gone.push(x.row);
  });
  gone.sort(function (a, b) { return b - a; }).forEach(function (r) { R.sh.deleteRow(r); });
  if (!entries.length && legacyEid) removeEvent_(calId_(), legacyEid);   // legacy row never registered
}

/* ============================ SYNC ONE MEETING ========================== */

/** Syncs ONE meeting to the configured calendar. Called by the Sync button in the Meeting Register. */
function dashSyncOne(meetingId) {
  requireRole_();
  return withLock_(function () {
    var rr = resetRunning_();
    if (rr) throw new Error('A Reset & Re-Sync started by ' + rr.by + ' is running. Try again in a few minutes.');

    var info = validateCalendar_(calId_());
    var sh = getSheet_(KE.SHEET.MEETINGS);
    var row = findRow_(sh, meetingId);
    if (!row) throw new Error('Meeting not found.');
    var m = rowToObj_(sh.getRange(row, 1, 1, MEETING_FIELDS.length).getValues()[0], MEETING_FIELDS);
    var st = trim_(m.status);

    if (st === 'Deleted Externally') throw new Error('This meeting was deleted in Google Calendar. Use Reset & Re-Sync with the restore option to bring it back.');
    if (trim_(m.syncStatus) === 'SYNC CONFLICT') throw new Error('This meeting has a conflict. Resolve it in Calendar Sync first.');
    if (!toDate_(m.date) || timeToMinutes_(m.start) === null) throw new Error('Meeting date or start time is missing.');

    var R = loadReg_();

    if (st === 'Cancelled') {
      var cr = cancelEverywhere_(R, m, info, null);
      if (cr.errors) throw new Error('Could not remove the event of this cancelled meeting. Try again.');
      setMirror_(sh, row, m, '', 'Synced');
      logSync_(m.id, '', 'Sheet → Calendar', 'Sync single meeting (cancelled)', 'Success', 'Calendar ' + info.id);
      return { id: m.id, status: 'Synced', eventId: '', message: m.id + ': cancelled, event removed from calendar.' };
    }

    var r;
    try {
      r = ensureEvent_(R, m, info);
    } catch (e) {
      var msg = errMsg_(e);
      putReg_(R, m.id, info.id, { status: 'Error', err: msg.substring(0, 500), at: new Date() });
      setMirror_(sh, row, m, trim_(m.eventId), 'Error');
      logSync_(m.id, trim_(m.eventId), 'Sheet → Calendar', 'Sync single meeting', 'Failed', msg);
      throw e;
    }
    setMirror_(sh, row, m, r.ev.id, 'Synced');
    if (r.ev.hangoutLink && !trim_(m.meetLink)) sh.getRange(row, MC.meetLink).setValue(r.ev.hangoutLink);
    propagate_(R, m, info, {}, { updated: 0, errors: 0 });   // keep any other registered calendars current
    if (r.op !== 'created') logSync_(m.id, r.ev.id, 'Sheet → Calendar', 'Sync single meeting (' + r.op + ')', 'Success', 'Calendar ' + info.id);

    var word = { created: 'added to calendar', recovered: 'linked to existing calendar event', updated: 'calendar event updated', unchanged: 'already up to date' }[r.op];
    return { id: m.id, status: 'Synced', eventId: r.ev.id, message: m.id + ': ' + word + '.' };
  });
}
