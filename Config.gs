/*
 * File: Config.gs   (Part 1 of 6)
 *
 * Central configuration: sheet names, column maps, dropdown lists, colours,
 * ID formats and the default values of the Settings sheet.
 * Nothing company-specific (phone, address, e-mail) is hardcoded here —
 * those live in the Settings sheet.
 */

var KE = {
  VERSION: '1.0.0',
  TZ: 'Asia/Kolkata',
  DATE_FMT: 'dd-MMM-yyyy',        // Apps Script (Utilities.formatDate) pattern
  SHEET_DATE_FMT: 'dd-mmm-yyyy',  // Google Sheets number-format pattern
  TIME_FMT: 'hh:mm a',

  SHEET: {
    DASHBOARD: 'Dashboard',
    MEETINGS: 'Meeting Register',
    MINUTES: 'Meeting Minutes',
    ACTIONS: 'Action Items',
    FOLLOWUPS: 'Follow-Up Tracker',
    CUSTOMERS: 'Customers',
    SYNCLOG: 'Calendar Sync Log',
    AUDIT: 'Audit Log',
    SETTINGS: 'Settings'
  },

  COLOR: {
    NAVY: '#0B2545',
    NAVY_LIGHT: '#13315C',
    WHITE: '#FFFFFF',
    GREY_BG: '#F3F5F9',
    GREY_LINE: '#D0D5DD',
    GREEN: '#1E8449',
    GREEN_BG: '#E3F4EA',
    RED: '#C0392B',
    RED_BG: '#FBE5E2',
    AMBER: '#B9770E',
    AMBER_BG: '#FDF1D6',
    BLUE: '#1F5FBF',
    BLUE_BG: '#E1ECFB'
  },

  // ID formats. yearly:true => KE-MTG-2026-001 ; yearly:false => KE-CUS-001
  ID: {
    MTG: { prefix: 'KE-MTG', yearly: true,  pad: 3, sheet: 'MEETINGS',  col: 1 },
    MOM: { prefix: 'KE-MOM', yearly: true,  pad: 3, sheet: 'MINUTES',   col: 1 },
    ACT: { prefix: 'KE-ACT', yearly: true,  pad: 3, sheet: 'ACTIONS',   col: 1 },
    FUP: { prefix: 'KE-FUP', yearly: true,  pad: 3, sheet: 'FOLLOWUPS', col: 1 },
    CUS: { prefix: 'KE-CUS', yearly: false, pad: 3, sheet: 'CUSTOMERS', col: 1 }
  },

  // Fixed addresses on the Dashboard sheet (used by Setup, Dashboard and button handler)
  DASH: {
    FILTER: {
      FROM: 'B14', TO: 'D14', TYPE: 'F14', STATUS: 'H14', CUSTOMER: 'J14', PRIORITY: 'L14',
      ACTION_STATUS: 'B16'
    },
    BUTTON_ROWS: [20, 21, 22],
    BUTTON_CHECK_COLS: [2, 6, 10],   // B, F, J  (checkbox cells)
    LIST_HEADER_ROW: 25,
    LIST_FIRST_ROW: 26,
    LIST_MAX_ROWS: 40
  }
};

/* ---------------------------------------------------------------------------
 * Column definitions: [key, header text]. Order = column order in the sheet.
 * ------------------------------------------------------------------------- */
var MEETING_FIELDS = [
  ['id', 'Meeting ID'], ['eventId', 'Calendar Event ID'], ['title', 'Agenda'],
  ['date', 'Meeting Date'], ['start', 'Start Time'], ['end', 'End Time'],
  ['type', 'Meeting Type'], ['custName', 'Customer Name'], ['company', 'Company Name'],
  ['custMobile', 'Customer Mobile'], ['custWa', 'Customer WhatsApp Number'], ['ownerMobile', 'Owner Mobile'],
  ['mode', 'Meeting Mode'], ['locType', 'Meeting Location Type'], ['address', 'Full Meeting Address'],
  ['mapLink', 'Google Maps Link'], ['meetLink', 'Google Meet Link'], ['participants', 'Participants'],
  ['agenda', 'Meeting Agenda'], ['priority', 'Priority'], ['status', 'Meeting Status'],
  ['locStatus', 'Location Confirmation Status'], ['ownerNotif', 'Owner Notification Status'],
  ['custNotif', 'Customer Notification Status'], ['syncStatus', 'Calendar Sync Status'],
  ['createdBy', 'Created By'], ['createdAt', 'Created At'], ['updatedAt', 'Last Updated'],
  ['remarks', 'Remarks']
];

var MINUTES_FIELDS = [
  ['id', 'Minutes ID'], ['meetingId', 'Meeting ID'], ['title', 'Meeting Title'], ['date', 'Meeting Date'],
  ['recordedBy', 'Recorded By'], ['summary', 'Discussion Summary'], ['topics', 'Topics Discussed'],
  ['decisions', 'Decisions Taken'], ['instructions', 'Owner Instructions'], ['remarks', 'Additional Remarks'],
  ['status', 'Minutes Status'], ['createdAt', 'Created At'], ['updatedAt', 'Last Updated']
];

var ACTION_FIELDS = [
  ['id', 'Action ID'], ['meetingId', 'Meeting ID'], ['title', 'Meeting Title'], ['task', 'Task Description'],
  ['assignedTo', 'Assigned To'], ['assignedBy', 'Assigned By'], ['priority', 'Priority'],
  ['assignedDate', 'Assigned Date'], ['deadline', 'Deadline'], ['status', 'Status'],
  ['completedDate', 'Completion Date'], ['remarks', 'Remarks'], ['updatedAt', 'Last Updated']
];

var FOLLOWUP_FIELDS = [
  ['id', 'Follow-Up ID'], ['meetingId', 'Meeting ID'], ['custName', 'Customer Name'],
  ['subject', 'Follow-Up Subject'], ['date', 'Follow-Up Date'], ['person', 'Responsible Person'],
  ['type', 'Follow-Up Type'], ['status', 'Status'], ['remarks', 'Remarks'], ['completedDate', 'Completed Date']
];

var CUSTOMER_FIELDS = [
  ['id', 'Customer ID'], ['name', 'Customer Name'], ['company', 'Company Name'], ['mobile', 'Mobile Number'],
  ['wa', 'WhatsApp Number'], ['email', 'Email'], ['address', 'Address'], ['mapLink', 'Google Maps Link'],
  ['notes', 'Notes'], ['createdAt', 'Created At']
];

var SYNCLOG_FIELDS = [
  ['ts', 'Timestamp'], ['meetingId', 'Meeting ID'], ['eventId', 'Calendar Event ID'],
  ['direction', 'Sync Direction'], ['operation', 'Operation'], ['status', 'Status'],
  ['error', 'Error Details'], ['by', 'Processed By']
];

var AUDIT_FIELDS = [
  ['ts', 'Timestamp'], ['user', 'User Email'], ['action', 'Action'], ['sheet', 'Sheet Name'],
  ['recordId', 'Record ID'], ['field', 'Field Changed'], ['oldV', 'Old Value'], ['newV', 'New Value']
];

/** Builds {key: 1-based column number} from a *_FIELDS array. */
function colMap_(fields) {
  var m = {};
  for (var i = 0; i < fields.length; i++) m[fields[i][0]] = i + 1;
  return m;
}
function headersOf_(fields) {
  return fields.map(function (f) { return f[1]; });
}

var MC = colMap_(MEETING_FIELDS);   // Meeting Register columns
var MM = colMap_(MINUTES_FIELDS);   // Meeting Minutes columns
var AC = colMap_(ACTION_FIELDS);    // Action Items columns
var FC = colMap_(FOLLOWUP_FIELDS);  // Follow-Up Tracker columns
var CC = colMap_(CUSTOMER_FIELDS);  // Customers columns
var SC = colMap_(SYNCLOG_FIELDS);   // Calendar Sync Log columns
var UC = colMap_(AUDIT_FIELDS);     // Audit Log columns

/* ---------------------------------------------------------------------------
 * Central dropdown lists (written to the Settings sheet, columns E onward).
 * ------------------------------------------------------------------------- */
var LISTS = {
  meetingType:  { title: 'Meeting Type',  items: ['Internal Meeting', 'Customer Meeting', 'Vendor Meeting', 'Management Review', 'Business Discussion', 'Personal Appointment', 'Virtual Meeting', 'Other'] },
  priority:     { title: 'Priority',      items: ['High', 'Medium', 'Low'] },
  meetingStatus:{ title: 'Meeting Status',items: ['Scheduled', 'Completed', 'Cancelled', 'Rescheduled', 'Deleted Externally'] },
  meetingMode:  { title: 'Meeting Mode',  items: ['In-Person', 'Virtual'] },
  locationType: { title: 'Location Type', items: ['KISAN ENTERPRISES Office', "Customer's Office", 'Other Location', 'Google Meet'] },
  locationStatus:{title: 'Location Status', items: ['Pending', 'Request Sent', 'Received', 'Confirmed', 'Not Required'] },
  notifStatus:  { title: 'Notification Status', items: ['Not Sent', 'WhatsApp Opened'] },
  syncStatus:   { title: 'Calendar Sync Status', items: ['Not Synced', 'Synced', 'Pending Update', 'SYNC CONFLICT', 'Error', 'Deleted Externally'] },
  minutesStatus:{ title: 'Minutes Status',items: ['Draft', 'Under Review', 'Finalized'] },
  actionStatus: { title: 'Action Status', items: ['Pending', 'In Progress', 'Completed', 'On Hold', 'Overdue'] },
  followType:   { title: 'Follow-Up Type',items: ['Customer Call', 'WhatsApp', 'Email', 'Internal Review', 'Next Meeting', 'Other'] },
  followStatus: { title: 'Follow-Up Status', items: ['Pending', 'Completed', 'Rescheduled', 'Cancelled'] }
};

/* ---------------------------------------------------------------------------
 * Settings keys (column A of the Settings sheet) and defaults.
 * Contact details and the office address are intentionally EMPTY — the owner/EA
 * must type them in. The system never invents them.
 * ------------------------------------------------------------------------- */
var SK = {
  COMPANY: 'Company Name',
  OWNER_NAME: 'Owner Name',
  OWNER_EMAIL: 'Owner Email',
  OWNER_WA: 'Owner WhatsApp Number',
  EA_EMAIL: 'EA Email',
  EA_WA: 'EA WhatsApp Number',
  CAL_ID: 'Owner Calendar ID',
  OFFICE_ADDR: 'Company Office Address',
  OFFICE_MAP: 'Company Google Maps Link',
  TZ: 'Time Zone',
  SYNC_MIN: 'Synchronization Interval (minutes)',
  REM_ENABLED: 'Daily Reminders Enabled (Yes/No)',
  REM_HOUR: 'Daily Reminder Hour (0-23)',
  REM_TO: 'Reminder Recipients (Owner / EA / Both)'
};

var SETTINGS_DEFAULTS = [
  [SK.COMPANY,     'KISAN ENTERPRISES', 'Printed in all WhatsApp messages and reports.'],
  [SK.OWNER_NAME,  '', 'Full name of the Owner.'],
  [SK.OWNER_EMAIL, '', 'Owner\'s Google account e-mail. Used for access control.'],
  [SK.OWNER_WA,    '', 'Owner\'s WhatsApp number (10 digits or with +91).'],
  [SK.EA_EMAIL,    '', 'Executive Assistant\'s Google account e-mail. Used for access control.'],
  [SK.EA_WA,       '', 'EA\'s WhatsApp number.'],
  [SK.CAL_ID,      'primary', 'Calendar to sync. "primary" = the calendar of the Google account that installs the triggers. Or paste the Owner\'s calendar e-mail / ID.'],
  [SK.OFFICE_ADDR, '', 'Full address of KISAN ENTERPRISES office. Used when location = KISAN ENTERPRISES Office.'],
  [SK.OFFICE_MAP,  '', 'Google Maps link of the office.'],
  [SK.TZ,          'Asia/Kolkata', 'Do not change unless required.'],
  [SK.SYNC_MIN,    5, 'Allowed values: 1, 5, 10, 15, 30.'],
  [SK.REM_ENABLED, 'Yes', 'Daily follow-up / overdue-task reminder e-mail.'],
  [SK.REM_HOUR,    9, 'Hour of day (24h, India time) when the reminder e-mail is sent.'],
  [SK.REM_TO,      'Both', 'Who receives the daily reminder e-mail.']
];

/** First data row of the settings table on the Settings sheet. */
var SETTINGS_FIRST_ROW = 4;
/** First column of the dropdown lists block on the Settings sheet (E = 5). */
var LISTS_FIRST_COL = 5;
