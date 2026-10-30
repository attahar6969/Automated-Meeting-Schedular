/**
 * Dashboard.gs (Part 5 of 7) (full working dashboard back-end)
 */

var GLOBAL_ = this;
var DASH_ALLOWED_ACTIONS_ = ['syncCalendarNow'];

var DATE_KEYS_  = ['date', 'assignedDate', 'deadline', 'completedDate'];
var TIME_KEYS_  = ['start', 'end'];
var STAMP_KEYS_ = ['createdAt', 'updatedAt'];
var PHONE_KEYS_ = ['custMobile', 'custWa', 'ownerMobile', 'mobile', 'wa'];
var PROTECTED_  = ['id', 'eventId', 'createdBy', 'createdAt'];

function ent_(name) {
  var d = {
    meeting:  { sheet: KE.SHEET.MEETINGS,  fields: MEETING_FIELDS,  kind: 'MTG' },
    minutes:  { sheet: KE.SHEET.MINUTES,   fields: MINUTES_FIELDS,  kind: 'MOM' },
    action:   { sheet: KE.SHEET.ACTIONS,   fields: ACTION_FIELDS,   kind: 'ACT' },
    followup: { sheet: KE.SHEET.FOLLOWUPS, fields: FOLLOWUP_FIELDS, kind: 'FUP' },
    customer: { sheet: KE.SHEET.CUSTOMERS, fields: CUSTOMER_FIELDS, kind: 'CUS' }
  }[name];
  if (!d) throw new Error('Unknown record type: ' + name);
  return d;
}

/* ============================ OPEN / NAV ================================ */

function openDashboardApp() {
  requireRole_();
  var html = HtmlService.createHtmlOutputFromFile('DashboardUI').setWidth(1400).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'ATTAHAR — Dashboard');
}

function runDashAction(fnName) {
  requireRole_();
  if (DASH_ALLOWED_ACTIONS_.indexOf(fnName) === -1) throw new Error('Action not allowed.');
  if (typeof GLOBAL_[fnName] !== 'function') return 'Calendar sync is not installed yet. It comes in the next part of the build.';
  GLOBAL_[fnName]();
  return 'Done.';
}

/* ============================ SERIALISING =============================== */

function ser_(o, fields) {
  var out = {};
  fields.forEach(function (f) {
    var k = f[0], v = o[k];
    if (v === null || v === undefined || v === '') { out[k] = ''; return; }
    if (DATE_KEYS_.indexOf(k) >= 0) {
      var d = toDate_(v);
      out[k] = d ? Utilities.formatDate(d, KE.TZ, 'yyyy-MM-dd') : '';
    } else if (TIME_KEYS_.indexOf(k) >= 0) {
      var m = timeToMinutes_(v);
      out[k] = m === null ? '' : ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
    } else if (STAMP_KEYS_.indexOf(k) >= 0) {
      out[k] = Object.prototype.toString.call(v) === '[object Date]'
        ? Utilities.formatDate(v, KE.TZ, 'yyyy-MM-dd HH:mm') : String(v);
    } else {
      out[k] = String(v);
    }
  });
  return out;
}

function convertIn_(k, v, label) {
  if (v === null || v === undefined) return '';
  if (DATE_KEYS_.indexOf(k) >= 0) {
    if (trim_(v) === '') return '';
    var d = toDate_(v);
    if (!d) throw new Error(label + ': invalid date.');
    return d;
  }
  if (TIME_KEYS_.indexOf(k) >= 0) {
    if (trim_(v) === '') return '';
    var m = timeToMinutes_(v);
    if (m === null) throw new Error(label + ': invalid time.');
    return minutesToText_(m);
  }
  if (PHONE_KEYS_.indexOf(k) >= 0) {
    if (trim_(v) === '') return '';
    var p = normalizePhone_(v);
    if (!p.ok) throw new Error(label + ': ' + p.error);
    return p.display;
  }
  return trim_(v);
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function inList_(listKey, v, label) {
  if (v === '' || v === null || v === undefined) return;
  if (getList_(listKey).indexOf(v) === -1) throw new Error(label + ': "' + v + '" is not an allowed value.');
}

/* ============================ READ APIs ================================= */

function dashBootstrap() {
  var role = requireRole_();
  var lists = {};
  Object.keys(LISTS).forEach(function (k) { lists[k] = getList_(k); });
  var s = getSettings_();
  return {
    role: role,
    lists: lists,
    settings: {
      company: trim_(s[SK.COMPANY]), ownerName: trim_(s[SK.OWNER_NAME]), ownerEmail: trim_(s[SK.OWNER_EMAIL]),
      ownerWa: trim_(s[SK.OWNER_WA]), eaEmail: trim_(s[SK.EA_EMAIL]), eaWa: trim_(s[SK.EA_WA]),
      officeAddr: trim_(s[SK.OFFICE_ADDR]), officeMap: trim_(s[SK.OFFICE_MAP])
    }
  };
}

function dashList(name) {
  requireRole_();
  var e = ent_(name);
  var out = [];
  readRows_(getSheet_(e.sheet), e.fields.length).forEach(function (r) {
    if (trim_(r[0]) === '') return;
    out.push(ser_(rowToObj_(r, e.fields), e.fields));
  });
  return out;
}

/* ============================ SAVE ====================================== */
function dashSave(name, data) {
  requireRole_();
      if (!VIEWER_OK_)  allowViewerWrite_();
  return withLock_(function () {
    AUDIT_BUF_ = [];
    try { return saveEntity_(name, data); }
    finally { flushAudit_(); }
  });
}
var MEET_CACHE_ = null;

function saveEntity_(name, data) {
  var e = ent_(name);
  var sh = getSheet_(e.sheet), n = e.fields.length;
  var rows = readRows_(sh, n);          // ONE read: finds the row, the old values and the next ID
  var row = 0, old = null, obj = {};
  if (data.id) {
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) { row = i + 2; break; }
    }
    if (!row) throw new Error('Record ' + data.id + ' was not found.');
    old = rowToObj_(rows[row - 2], e.fields);
    obj = rowToObj_(rows[row - 2], e.fields);
    if (data.__linkCheck && trim_(old.meetingId) !== trim_(data.__linkCheck)) {
      throw new Error('That follow-up is not linked to this meeting.');
    }
  }
  var isNew = !old;
  e.fields.forEach(function (f) {
    var k = f[0];
    if (PROTECTED_.indexOf(k) >= 0 || data[k] === undefined) return;
    obj[k] = convertIn_(k, data[k], f[1]);
  });

  if (name === 'meeting') normMeeting_(obj, old, isNew);
  else if (name === 'minutes') normMinutes_(obj, isNew);
  else if (name === 'action') normAction_(obj, isNew);
  else if (name === 'followup') normFollowup_(obj, isNew, old);
  else if (name === 'customer') normCustomer_(obj, isNew);

  var now = new Date();
  if (isNew) {
    obj.id = nextId_(e.kind, rows.map(function (r) { return r[0]; }));
    if (obj.createdAt === undefined || obj.createdAt === '') obj.createdAt = now;
    if (e.fields.some(function (f) { return f[0] === 'createdBy'; })) obj.createdBy = currentUserEmail_() || 'system';
    row = Math.max(2, rows.length + 2);
  }
  if (e.fields.some(function (f) { return f[0] === 'updatedAt'; })) obj.updatedAt = now;

  sh.getRange(row, 1, 1, n).setValues([objToRow_(obj, e.fields)]);
  if (name === 'meeting') MEET_CACHE_ = obj;     // lets a follow-up saved right now skip re-reading the register

  if (isNew) logAudit_('Create', e.sheet, obj.id, '', '', obj.title || obj.task || obj.subject || obj.name || '');
  else logAuditDiff_('Update', e.sheet, obj.id, e.fields, old, obj);

  // Optional: remember the customer typed on a new meeting
  if (name === 'meeting' && data.saveCustomer && trim_(obj.custName)) {
    var existing = readRows_(getSheet_(KE.SHEET.CUSTOMERS), CUSTOMER_FIELDS.length)
      .some(function (r) { return trim_(r[CC.name - 1]).toLowerCase() === trim_(obj.custName).toLowerCase(); });
    if (!existing) {
      saveEntity_('customer', { name: obj.custName, company: obj.company, mobile: obj.custMobile,
        wa: obj.custWa, address: obj.address, mapLink: obj.mapLink });
    }
  }

  // Follow-up typed in the meeting form: update the linked one, or create a new one
  var fuErr = '';
  if (name === 'meeting' && trim_(data.fuDate)) {
    try {
      var fuId = trim_(data.fuId);
      if (fuId) {
        var fpatch = {
          id: fuId,
          date: data.fuDate,
          type: trim_(data.fuType) || 'Next Meeting',
          status: trim_(data.fuStatus) || 'Pending',
          __linkCheck: obj.id           // refuses to edit a follow-up that belongs to another meeting
        };
        if (trim_(data.fuSubject)) fpatch.subject = trim_(data.fuSubject);
        saveEntity_('followup', fpatch);          // same ID, so no duplicate
      } else {
        saveEntity_('followup', {
          meetingId: obj.id,
          subject: trim_(data.fuSubject) || ('Follow-up: ' + (trim_(obj.title) || trim_(obj.custName))),
          date: data.fuDate,
          type: trim_(data.fuType) || 'Next Meeting',
          custName: obj.custName,
          status: trim_(data.fuStatus) || 'Pending'
        });
      }
    } catch (fe) {
      fuErr = 'Meeting saved, but the follow-up was not saved: ' + errMsg_(fe);
    }
  }
  return { id: obj.id, record: ser_(obj, e.fields), followupError: fuErr };
}

  function normMeeting_(m, old, isNew) {
  var s = getSettings_();
  if (isNew) {
    if (!trim_(m.custName)) throw new Error('Customer Name is required.');
    if (!trim_(m.custMobile)) throw new Error('Customer Mobile is required.');
  }
  if (!trim_(m.title)) m.title = trim_(m.custName) || trim_(m.company) || 'Meeting';
  if (!trim_(m.custWa)) m.custWa = m.custMobile;
  var sm = timeToMinutes_(m.start);
  var em = timeToMinutes_(m.end);
  if (sm !== null) {
  if (em === null) { m.end = minutesToText_(Math.min(sm + 60, 1439)); em = timeToMinutes_(m.end); }
  if (em <= sm) throw new Error('End Time must be after Start Time.');
}
  m.type = m.type || 'Customer Meeting';
  m.priority = m.priority || 'Medium';
  m.mode = m.mode || 'In-Person';
  m.status = m.status || 'Scheduled';
  inList_('meetingType', m.type, 'Meeting Type');
  inList_('priority', m.priority, 'Type of Customer');
  inList_('meetingMode', m.mode, 'Meeting Mode');
  inList_('meetingStatus', m.status, 'Meeting Status');
  inList_('locationType', m.locType, 'Location Type');
  inList_('locationStatus', m.locStatus, 'Location Status');
  inList_('notifStatus', m.ownerNotif, 'Owner Notification');
  inList_('notifStatus', m.custNotif, 'Customer Notification');

  if (m.mode === 'Virtual') {
    m.locType = 'Google Meet';
    if (!m.locStatus) m.locStatus = 'Not Required';
  } else {
    if (m.locType === 'Google Meet') throw new Error('Location Type "Google Meet" needs Meeting Mode = Virtual.');
    if (!m.locType) m.locType = "Customer's Office";
    if (m.locType === 'ATTAHAR Office') {
      if (!m.address) m.address = trim_(s[SK.OFFICE_ADDR]);
      if (!m.mapLink) m.mapLink = trim_(s[SK.OFFICE_MAP]);
      m.locStatus = m.address ? 'Confirmed' : 'Pending';
    } else if (!m.locStatus) {
      m.locStatus = m.address ? 'Received' : 'Pending';
    } else if (m.address && (m.locStatus === 'Pending' || m.locStatus === 'Request Sent')) {
      m.locStatus = 'Received';
    }
  }
  if (!m.ownerNotif) m.ownerNotif = 'Not Sent';
  if (!m.custNotif) m.custNotif = 'Not Sent';

  if (isNew) {
    if (!m.syncStatus) m.syncStatus = 'Not Synced';
  } else if (old && old.syncStatus === 'Synced') {
    var watch = ['title', 'date', 'start', 'end', 'address', 'mapLink', 'meetLink', 'agenda', 'participants', 'status'];
    var changed = watch.some(function (k) { return auditVal_(old[k]) !== auditVal_(m[k]); });
    if (changed) m.syncStatus = 'Pending Update';
  }
}

function lookupMeeting_(id) {
  if (MEET_CACHE_ && String(MEET_CACHE_.id) === String(id)) return MEET_CACHE_;
  var rows = readRows_(getSheet_(KE.SHEET.MEETINGS), MEETING_FIELDS.length);
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(id)) return rowToObj_(rows[i], MEETING_FIELDS);
  return null;
}

function normMinutes_(o, isNew) {
  if (!o.meetingId) throw new Error('Please choose the meeting.');
  var m = lookupMeeting_(o.meetingId);
  if (!m) throw new Error('Meeting ' + o.meetingId + ' was not found.');
  o.title = m.title; o.date = m.date;
  if (!o.summary) throw new Error('Discussion Summary is required.');
  o.status = o.status || 'Draft';
  inList_('minutesStatus', o.status, 'Minutes Status');
  if (!o.recordedBy) o.recordedBy = currentUserEmail_();
}

function normAction_(o, isNew) {
  if (!o.task) throw new Error('Task Description is required.');
  if (o.meetingId) {
    var m = lookupMeeting_(o.meetingId);
    if (!m) throw new Error('Meeting ' + o.meetingId + ' was not found.');
    o.title = m.title;
  }
  o.priority = o.priority || 'Medium';
  o.status = o.status || 'Pending';
  inList_('priority', o.priority, 'Priority');
  inList_('actionStatus', o.status, 'Status');
  if (isNew && !o.assignedDate) o.assignedDate = todayDate_();
  if (!o.assignedBy) o.assignedBy = trim_(getSetting_(SK.OWNER_NAME)) || currentUserEmail_();
  if (o.status === 'Completed') { if (!toDate_(o.completedDate)) o.completedDate = todayDate_(); }
  else o.completedDate = '';
}

function normFollowup_(o, isNew, old) {
  if (!o.subject) throw new Error('Follow-Up Subject is required.');
  if (!toDate_(o.date)) throw new Error('Follow-Up Date is required.');
  // Only look the meeting up when it is new or has changed
  var sameMeeting = old && trim_(old.meetingId) === trim_(o.meetingId) && trim_(o.custName);
  if (o.meetingId && !sameMeeting) {
    var m = lookupMeeting_(o.meetingId);
    if (!m) throw new Error('Meeting ' + o.meetingId + ' was not found.');
    if (!o.custName) o.custName = m.custName || m.company;
  }
  o.type = o.type || 'Customer Call';
  o.status = o.status || 'Pending';
  inList_('followType', o.type, 'Follow-Up Type');
  inList_('followStatus', o.status, 'Status');
  if (o.status === 'Completed') { if (!toDate_(o.completedDate)) o.completedDate = todayDate_(); }
  else o.completedDate = '';
}

function normCustomer_(o, isNew) {
  if (!o.name) throw new Error('Customer Name is required.');
  if (!o.wa && o.mobile) o.wa = o.mobile;
  if (isNew && o.mobile) {
    var dup = readRows_(getSheet_(KE.SHEET.CUSTOMERS), CUSTOMER_FIELDS.length)
      .some(function (r) { return waDigits_(r[CC.mobile - 1]) === waDigits_(o.mobile); });
    if (dup) throw new Error('A customer with this mobile number already exists.');
  }
}

/* ============================ SETTINGS ================================== */

function contactPairs_(d) {
  function email(v, label) {
    var list = emailList_(v);
    list.forEach(function (x) {
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)) throw new Error(label + ': "' + x + '" is not a valid e-mail address.');
    });
    return list.join(', ');
  }
  function wa(v, label) {
    if (trim_(v) === '') return '';
    var p = normalizePhone_(v);
    if (!p.ok) throw new Error(label + ': ' + p.error);
    return p.display;
  }
  var oe = email(d.ownerEmail, 'Owner Email'), ee = email(d.eaEmail, 'EA Email');
  var me = currentUserEmail_();
  var allowed = emailList_(oe).concat(emailList_(ee));
  if (allowed.length && me && allowed.indexOf(me) === -1) {
    throw new Error('You are signed in as ' + me + '. Your own e-mail must be in the Owner or EA list, otherwise you would lock yourself out.');
  }
  return [
    [SK.COMPANY, trim_(d.company) || 'ATTAHAR'],
    [SK.OWNER_NAME, trim_(d.ownerName)],
    [SK.OWNER_EMAIL, oe],
    [SK.OWNER_WA, wa(d.ownerWa, 'Owner WhatsApp')],
    [SK.EA_EMAIL, ee],
    [SK.EA_WA, wa(d.eaWa, 'EA WhatsApp')],
    [SK.OFFICE_ADDR, trim_(d.officeAddr)],
    [SK.OFFICE_MAP, trim_(d.officeMap)]
  ];
}

function dashSaveSettings(d) {
  requireRole_();
  return withLock_(function () {
    setSettingsBatch_(contactPairs_(d));
    logAudit_('Update settings', KE.SHEET.SETTINGS, 'SETTINGS', '', '', 'Updated from dashboard');
    return true;
  });
}
/* ============================ WHATSAPP ================================== */

function dashWaLink(meetingId, kind) {
  requireRole_();
  allowViewerWrite_();
  return withLock_(function () {
    var sh = getSheet_(KE.SHEET.MEETINGS);
    var row = findRow_(sh, meetingId);
    if (!row) throw new Error('Meeting not found.');
    var m = ser_(rowToObj_(sh.getRange(row, 1, 1, MEETING_FIELDS.length).getValues()[0], MEETING_FIELDS), MEETING_FIELDS);
    var s = getSettings_();
    var company = trim_(s[SK.COMPANY]) || 'ATTAHAR';
    var sign = trim_(s[SK.OWNER_NAME]) || company;
    var when = fmtDate_(m.date) + ', ' + fmtTime_(m.start) + (m.end ? ' – ' + fmtTime_(m.end) : '');
    var where = m.mode === 'Virtual' ? ('Google Meet: ' + (m.meetLink || 'link will be shared')) : (m.address || 'to be confirmed');
    var to, msg, patch = { id: meetingId };

    if (kind === 'customer') {
      to = m.custWa || m.custMobile;
      msg = 'Dear ' + (m.custName || 'Sir/Madam') + ',\n\n' + company + ' confirms the following meeting:\n\n*' + m.title + '*\n' +
        'Date & Time: ' + when + '\nPlace: ' + where + (m.mapLink ? '\nMap: ' + m.mapLink : '') + '\n\nRegards,\n' + sign;
      patch.custNotif = 'WhatsApp Opened';
    } else if (kind === 'owner') {
      to = trim_(s[SK.OWNER_WA]) || m.ownerMobile;
      msg = '*Meeting reminder*\n\n' + m.title + '\nWith: ' + (m.custName || m.company || '—') + '\nDate & Time: ' + when +
        '\nPlace: ' + where + (m.mapLink ? '\nMap: ' + m.mapLink : '') + (m.agenda ? '\nAgenda: ' + m.agenda : '');
      patch.ownerNotif = 'WhatsApp Opened';
    } else if (kind === 'location') {
      to = m.custWa || m.custMobile;
            msg = 'Dear ' + (m.custName || 'Sir/Madam') + ',\n\nFor our meeting "' + m.title + '" on ' + when +
        ', please share your exact location (Google Maps pin) so we can reach on time.\n\nThank you,\n' + sign +
        (sign.toLowerCase() !== company.toLowerCase() ? '\n' + company : ''); 
      if (m.locStatus === 'Pending') patch.locStatus = 'Request Sent';
    } else {
      throw new Error('Unknown message type.');
    }
    var num = waDigits_(to);
    if (!num) throw new Error('No valid WhatsApp number found for this ' + (kind === 'owner' ? 'owner (set it in Settings).' : 'customer (edit the meeting and add it).'));
    saveEntity_('meeting', patch);
    return { url: 'https://wa.me/' + num + '?text=' + encodeURIComponent(msg), to: to };
  });
}

/* ============================ OVERVIEW DATA ============================= */

function weekStart_(d) {
  var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return x;
}

function getDashboardData() {
  requireRole_();
  allowViewerWrite_();
  var DAY = 86400000;
  var today = todayDate_(), tMs = today.getTime();
  var s = getSettings_();
  var hour = +Utilities.formatDate(new Date(), KE.TZ, 'H');
  var greet = hour < 12 ? 'Good morning' : (hour < 17 ? 'Good afternoon' : 'Good evening');

  function load(sheetName, fields) {
    return readRows_(getSheet_(sheetName), fields.length).map(function (r) { return rowToObj_(r, fields); })
      .filter(function (o) { return trim_(o[fields[0][0]]) !== ''; });
  }
  var meetings = load(KE.SHEET.MEETINGS, MEETING_FIELDS);
  var actions = load(KE.SHEET.ACTIONS, ACTION_FIELDS);
  var follow = load(KE.SHEET.FOLLOWUPS, FOLLOWUP_FIELDS);
  var minutes = load(KE.SHEET.MINUTES, MINUTES_FIELDS);
  var finalized = {};
  minutes.forEach(function (m) { if (m.status === 'Finalized') finalized[m.meetingId] = true; });

  function dead(m) { return m.status === 'Cancelled' || m.status === 'Deleted Externally'; }
  function dMs(v) { var d = toDate_(v); return d ? d.getTime() : null; }

  var todays = [], upcoming = 0, minutesPending = 0, locPending = [], statusCount = {};
  meetings.forEach(function (m) {
    var t = dMs(m.date);
    var st = m.status || 'Scheduled';
    statusCount[st] = (statusCount[st] || 0) + 1;
    if (dead(m)) return;
    if (t === tMs) todays.push(m);
    if (t !== null && t > tMs && (st === 'Scheduled' || st === 'Rescheduled')) upcoming++;
    if (((t !== null && t < tMs) || st === 'Completed') && !finalized[m.id]) minutesPending++;
    if (m.locStatus === 'Pending' || m.locStatus === 'Request Sent') locPending.push(m);
  });
  todays.sort(function (a, b) { return (timeToMinutes_(a.start) || 0) - (timeToMinutes_(b.start) || 0); });

  var openActions = actions.filter(function (a) { return a.status !== 'Completed'; });
  var overdue = openActions.filter(function (a) {
    var t = dMs(a.deadline); return a.status === 'Overdue' || (t !== null && t < tMs);
  });
  openActions.sort(function (a, b) { return (dMs(a.deadline) || 9e15) - (dMs(b.deadline) || 9e15); });
  var followDue = follow.filter(function (f) { var t = dMs(f.date); return f.status === 'Pending' && t !== null && t <= tMs; });

  var ws = weekStart_(today).getTime(), weeks = [], idx = {};
  for (var i = 7; i >= 0; i--) {
    var st0 = ws - i * 7 * DAY;
    idx[st0] = weeks.length;
    weeks.push({ label: Utilities.formatDate(new Date(st0), KE.TZ, 'dd MMM'), count: 0 });
  }
  meetings.forEach(function (m) {
    var d = toDate_(m.date); if (!d || dead(m)) return;
    var k = weekStart_(d).getTime();
    if (idx[k] !== undefined) weeks[idx[k]].count++;
  });

  return {
    greeting: greet, owner: trim_(s[SK.OWNER_NAME]),
    todayText: Utilities.formatDate(new Date(), KE.TZ, 'EEE, dd MMM yyyy'),
    kpi: { today: todays.length, upcoming: upcoming, pendingActions: openActions.length - overdue.length,
      overdueActions: overdue.length, followDue: followDue.length, locPending: locPending.length,
      minutesPending: minutesPending, total: meetings.length },
    schedule: todays.slice(0, 8).map(function (m) {
      return { id: m.id, time: fmtTime_(m.start), title: m.title, customer: m.custName || m.company || '—',
        type: m.type, status: m.status };
    }),
    actions: openActions.slice(0, 6).map(function (a) {
      var t = dMs(a.deadline);
      return { id: a.id, task: a.task, who: a.assignedTo, priority: a.priority, deadline: fmtDate_(a.deadline),
        late: t !== null && t < tMs };
    }),
    locations: locPending.slice(0, 5).map(function (m) {
      return { id: m.id, title: m.title, customer: m.custName || m.company || '—', date: fmtDate_(m.date), status: m.locStatus };
    }),
        weeks: weeks,
    statuses: Object.keys(statusCount).map(function (k) { return { name: k, count: statusCount[k] }; })
  };
}
