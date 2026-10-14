/** Code.gs (menu + helpers) */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KISAN ENTERPRISES')
    .addItem('Open Web App', 'openWebApp')
    .addSeparator()
    .addItem('Sync Calendar now', 'syncCalendarNow')
    .addItem('Install / refresh automation', 'installAutomation')
    .addItem('Send test reminder e-mail', 'sendTestReminder')
    .addSeparator()
    .addItem('Setup workbook (first time)', 'setupWorkbook')
    .addItem('Verify Setup', 'verifySetup')
    .addToUi();
}

/** Remembers which spreadsheet the web app must use. */
function linkSpreadsheet_() {
  PropertiesService.getScriptProperties().setProperty('SSID', SpreadsheetApp.getActiveSpreadsheet().getId());
}

function openWebApp() {
  linkSpreadsheet_();
  var url = ScriptApp.getService().getUrl();
  if (!url) { showAlert_('Web app', 'Not deployed yet. In Apps Script choose Deploy > New deployment > Web app.'); return; }
  var html = HtmlService.createHtmlOutput('<p style="font-family:Arial;font-size:14px">Your web app link:<br><br><a href="' + url + '" target="_blank">' + url + '</a></p>')
    .setWidth(560).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, 'KISAN ENTERPRISES Web App');
}

function installAutomation() {
  requireRole_();
  withLock_(function () { installTriggers_(); });
  showAlert_('Automation', 'Auto-sync, meeting alerts, daily reminder and daily maintenance are now installed.');
}

function sendTestReminder() {
  requireRole_();
  showAlert_('Reminder', sendReminders_(true));
}


function copyAgendaToTitle() {
  var sh = getSheet_(KE.SHEET.MEETINGS);
  var rows = readRows_(sh, MEETING_FIELDS.length);
  if (!rows.length) return;
  var out = rows.map(function (r) {
    return [trim_(r[0]) ? trim_(r[MC.agenda - 1]) : r[MC.title - 1]];
  });
  sh.getRange(2, MC.title, out.length, 1).setValues(out);
}

function repairDuplicateIds() {
  requireRole_();
  var report = withLock_(function () {
    var out = [];
    Object.keys(KE.ID).forEach(function (kind) {
      var def = KE.ID[kind], sh = getSheet_(KE.SHEET[def.sheet]);
      var last = sh.getLastRow();
      if (last < 2) return;
      var ids = sh.getRange(2, def.col, last - 1, 1).getValues();
      var seen = {};
      for (var i = 0; i < ids.length; i++) {
        var id = trim_(ids[i][0]);
        if (!id) continue;
        if (!seen[id]) { seen[id] = true; continue; }   // first (older) row keeps its ID

        var row = i + 2, newId = nextId_(kind);
        sh.getRange(row, def.col).setValue(newId);

        // Meetings that already have a calendar event: rewrite the event on next sync
        if (kind === 'MTG' && trim_(sh.getRange(row, MC.eventId).getValue()) !== '' &&
            trim_(sh.getRange(row, MC.syncStatus).getValue()) === 'Synced') {
          sh.getRange(row, MC.syncStatus).setValue('Pending Update');
        }
        logAudit_('Repair duplicate ID', def.sheet ? KE.SHEET[def.sheet] : '', newId, 'ID', id, newId);
        out.push(KE.SHEET[def.sheet] + ' row ' + row + ': ' + id + ' -> ' + newId);
      }
    });
    return out;
  });
  var msg = report.length ? report.join('\n') : 'No duplicate IDs found.';
  Logger.log(msg);
  showAlert_('Repair duplicate IDs', msg);
  return msg;
}

function renumberMeetings() {
  requireRole_();
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Renumber meetings',
    'This renumbers ALL meetings from 001 in the current row order, and updates the Meeting ID in Minutes, Action Items and Follow-Ups.\n\nHave you made a backup copy of this spreadsheet?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var msg = withLock_(function () {
    var def = KE.ID.MTG;
    var year = Utilities.formatDate(new Date(), KE.TZ, 'yyyy');
    var prefix = def.prefix + '-' + year + '-';
    var sh = getSheet_(KE.SHEET.MEETINGS);
    var last = sh.getLastRow();
    if (last < 2) return 'No meetings found.';

    var ids = sh.getRange(2, MC.id, last - 1, 1).getValues();
    var evs = sh.getRange(2, MC.eventId, last - 1, 1).getValues();
    var syn = sh.getRange(2, MC.syncStatus, last - 1, 1).getValues();

    var map = {}, n = 0, newIds = [], newSyn = [];
    for (var i = 0; i < ids.length; i++) {
      var old = trim_(ids[i][0]);
      if (!old || old.indexOf(prefix) !== 0) {      // blank or other-year ID: leave alone
        newIds.push([ids[i][0]]);
        newSyn.push([syn[i][0]]);
        continue;
      }
      n++;
      var nid = prefix + Utilities.formatString('%0' + def.pad + 'd', n);
      if (!map.hasOwnProperty(old)) map[old] = nid;   // if an ID was duplicated, links go to the first row
      newIds.push([nid]);
      // Calendar events carry the Meeting ID in their description: refresh them on next sync
      if (trim_(evs[i][0]) !== '' && trim_(syn[i][0]) === 'Synced') newSyn.push(['Pending Update']);
      else newSyn.push([syn[i][0]]);
    }
    sh.getRange(2, MC.id, newIds.length, 1).setValues(newIds);
    sh.getRange(2, MC.syncStatus, newSyn.length, 1).setValues(newSyn);
           remapRegistryIds_(map);
    // Keep linked records pointing at the right meeting
    var linked = [
      [KE.SHEET.MINUTES, MM.meetingId],
      [KE.SHEET.ACTIONS, AC.meetingId],
      [KE.SHEET.FOLLOWUPS, FC.meetingId]
    ];
    var relinked = 0;
    linked.forEach(function (L) {
      var s = getSheet_(L[0]), lr = s.getLastRow();
      if (lr < 2) return;
      var rng = s.getRange(2, L[1], lr - 1, 1), vals = rng.getValues(), changed = false;
      for (var r = 0; r < vals.length; r++) {
        var v = trim_(vals[r][0]);
        if (v && map.hasOwnProperty(v) && map[v] !== v) { vals[r][0] = map[v]; changed = true; relinked++; }
      }
      if (changed) rng.setValues(vals);
    });

    // Next new meeting continues after the last number
    PropertiesService.getDocumentProperties().setProperty('CTR_MTG_' + year, String(n));
    logAudit_('Renumber meetings', KE.SHEET.MEETINGS, 'ALL', 'Meeting ID', '', n + ' meetings renumbered');
    return 'Renumbered ' + n + ' meetings (001 to ' + Utilities.formatString('%03d', n) + ').\n' +
      'Linked records updated: ' + relinked + '.\nNext new meeting will be ' + prefix + Utilities.formatString('%03d', n + 1) + '.';
  });
  showAlert_('Renumber meetings', msg);
  return msg;
}

/** Gives an ID to every row that has data but a blank ID (use after pasting old data). */
function assignMissingIds() {
  requireRole_();
  var msg = withLock_(function () {
    var sheets = [
      ['MTG', KE.SHEET.MEETINGS, MEETING_FIELDS],
      ['MOM', KE.SHEET.MINUTES, MINUTES_FIELDS],
      ['ACT', KE.SHEET.ACTIONS, ACTION_FIELDS],
      ['FUP', KE.SHEET.FOLLOWUPS, FOLLOWUP_FIELDS],
      ['CUS', KE.SHEET.CUSTOMERS, CUSTOMER_FIELDS]
    ];
    var out = [];
    sheets.forEach(function (s) {
      var sh = getSheet_(s[1]), last = sh.getLastRow();
      if (last < 2) return;
      var data = sh.getRange(2, 1, last - 1, s[2].length).getValues();
      var ids = [], added = 0;
      for (var i = 0; i < data.length; i++) {
        var hasId = trim_(data[i][0]) !== '';
        var hasData = false;
        for (var j = 1; j < data[i].length; j++) {
          if (trim_(data[i][j]) !== '') { hasData = true; break; }
        }
        if (!hasId && hasData) { ids.push([nextId_(s[0])]); added++; }
        else ids.push([data[i][0]]);
      }
      if (added) {
        sh.getRange(2, 1, ids.length, 1).setValues(ids);
        logAudit_('Assign missing IDs', s[1], 'ALL', 'ID', '', added + ' IDs created');
        out.push(s[1] + ': ' + added + ' IDs created');
      }
    });
    return out.length ? out.join('\n') : 'No rows without an ID were found.';
  });
  showAlert_('Assign missing IDs', msg);
  return msg;
}

   function findOrphanEvents_() {
     var calId = calId_();
     var known = {};
     loadObjs_(KE.SHEET.MEETINGS, MEETING_FIELDS).forEach(function (m) { known[trim_(m.id)] = true; });
     var orphans = [], token = null;
     var min = new Date(Date.now() - 365 * DAY_MS_).toISOString();
     var max = new Date(Date.now() + 730 * DAY_MS_).toISOString();
     do {
       var opts = { timeMin: min, timeMax: max, maxResults: 250, showDeleted: false };
       if (token) opts.pageToken = token;
       var res = retry_(function () { return Calendar.Events.list(calId, opts); }, 3);
       (res.items || []).forEach(function (ev) {
         var p = ev.extendedProperties && ev.extendedProperties['private'];
         if (p && p.keMeetingId && !known[p.keMeetingId]) orphans.push(ev);
       });
       token = res.nextPageToken;
     } while (token);
     return orphans;
   }

function cleanupCalendarOrphans() {
  requireRole_();
  var ui = SpreadsheetApp.getUi();
  var orphans = findOrphanEvents_();
  if (!orphans.length) { showAlert_('Calendar cleanup', 'No leftover events found. Calendar and sheet match.'); return; }
  var list = orphans.slice(0, 15).map(function (ev) {
    var when = ev.start ? (ev.start.dateTime || ev.start.date || '') : '';
    return '• ' + (ev.summary || '(no title)') + '  [' + String(when).substring(0, 10) + ']';
  }).join('\n');
  var more = orphans.length > 15 ? '\n…and ' + (orphans.length - 15) + ' more' : '';
  var ok = ui.alert('Delete leftover calendar events',
    orphans.length + ' calendar event(s) have no matching meeting in the register:\n\n' + list + more +
    '\n\nDelete them from Google Calendar?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var calId = calId_(), n = 0;
  orphans.forEach(function (ev) {
    removeEvent_(calId, ev.id);
    var p = ev.extendedProperties['private'];
    logSync_(p.keMeetingId, ev.id, 'Sheet → Calendar', 'Delete orphan event', 'Success', 'No matching row in register');
    n++;
  });
  showAlert_('Calendar cleanup', n + ' leftover event(s) deleted.');
}
   function resetCalendarLinks() {
     requireRole_();
     showAlert_('Reset calendar links', 'This older tool is retired because it erased saved event IDs. Open the web app > Calendar Sync > "Reset & Re-Sync Calendar" instead. It keeps a separate event ID for every calendar and never deletes events.');
   }
