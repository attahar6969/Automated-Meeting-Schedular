# Meeting Scheduler & Management System
For Demo of meeting schedular open this link:https://script.google.com/macros/s/AKfycbxGTdiFku-w-8rRcj-JP7q0AqaoshoBIeT-ZJbUAk_WCfVLV5X_a2HvPTBMyqXQ-dhA/exec

A custom meeting-management application built **from scratch** with Google Apps Script, JavaScript, HTML, CSS, Google Sheets, and Google Calendar API v3. It brings meeting scheduling, customer information, minutes, action items, follow-ups, calendar synchronization, reminders, reports, and administration into one web application.

The system uses Google Sheets as its structured data store and Google Calendar as the scheduling integration. A browser-based web app provides the main interface, while Apps Script handles data operations, access checks, automation, reporting, and integrations.

## Highlights

- Built from scratch as a custom application rather than assembled from a ready-made meeting-management template.
- Web-based interface with dashboard, navigation, record forms, filters, and an interactive help guide.
- Meeting, minutes, action-item, follow-up, and customer records stored in Google Sheets.
- Google Calendar synchronization with a multi-calendar registry, conflict handling, and reset/re-sync workflows.
- Configurable scheduled synchronization, meeting alerts, daily maintenance, and daily brief emails.
- Company branding and configurable application settings.
- **Open To All** access option with a view-only experience for viewers.
- PDF reporting with report files saved to Google Drive.
- Audit and synchronization logs to support operational traceability.

## Main Features

### 1. Meeting Register

Create and maintain meeting records with fields such as meeting date, start and end time, agenda/title, customer details, meeting type, priority, mode, location, participants, and status.

The application supports meeting lifecycle statuses such as Scheduled, Rescheduled, Completed, and Cancelled, along with synchronization and location status information.

The Meeting Register includes search and filtering tools for working with the meeting list. Meeting records can be opened to review their details and related minutes, action items, and follow-ups.

### 2. Dashboard and Overview

The Overview page brings key operational information together, including:

- KPI cards for meeting and task status.
- Today's meeting schedule and upcoming meetings.
- Pending and overdue action items.
- Follow-ups due and other work requiring attention.
- Meeting activity and status charts.
- Quick navigation to common application areas.
- Manual refresh and dashboard calendar-sync actions.

The dashboard is intended to provide a quick operational summary; the underlying records remain in the relevant Google Sheets tabs.

### 3. Google Calendar Synchronization

The application integrates with Google Calendar through the Apps Script Advanced Calendar service (Calendar API).

Supported workflows include:

- Manual full synchronization and individual meeting synchronization.
- Automatic synchronization through an installed time-driven trigger.
- Creating and updating calendar events from meeting records.
- Pulling supported calendar-side changes into the application.
- Detecting and presenting synchronization conflicts for resolution.
- Tracking synchronization status, timestamps, errors, and event references.
- Reset and re-sync workflows for rebuilding or reconciling calendar links.
- Handling meeting cancellation and deletion across calendars registered for that meeting.

#### Multi-calendar registry

`CalendarSync.gs` maintains a dedicated **Calendar Sync Registry** with one record per Meeting ID and Calendar ID combination. This allows the application to retain references to events in multiple calendars.

The Meeting Register's existing Calendar Event ID and Calendar Sync Status fields are retained as a mirror of the currently configured calendar for compatibility with the interface.

**Calendar safety note:** Changing the configured Calendar ID does not automatically delete events from the previous calendar. Review the registry and synchronization report before performing administrative cleanup.

Calendar synchronization depends on valid Calendar IDs, granted permissions, enabled Advanced Services, trigger execution, and Google's service availability. Failed operations and conflicts may require administrator review.

### 4. Meeting Minutes

Record discussions, decisions, outcomes, and other meeting notes. Minutes are linked to their parent meeting through the Meeting ID and use their own record identifiers.

The application organizes and stores minutes; users must enter or verify the actual discussion and decisions.

### 5. Action Items

Create tasks arising from meetings and track:

- Assigned person or responsibility.
- Assignment date and deadline.
- Priority and status.
- Completion and overdue state.
- Relationship to the originating meeting.

Action-item status is reflected in the dashboard and relevant record views. Task ownership and progress still require users to maintain accurate information.

### 6. Follow-Up Tracker

Create follow-up records linked to meetings and customers. Track follow-up dates, subjects, statuses, completion information, and notes.

The meeting list can surface the relevant open or most recent follow-up. Users can update follow-up status as customer contact and next steps progress.

### 7. Customer Management

Maintain a reusable customer directory containing supported customer contact and address details. Customer information can be reused when creating meetings, reducing repeated data entry.

The application includes customer-related search and validation. Phone numbers are normalized and validated by the backend according to the application's rules.

### 8. WhatsApp Shortcuts and Email

The interface provides WhatsApp-assisted workflows for supported meeting communications, including meeting confirmation, location requests, and owner reminders.

WhatsApp actions prepare a message and open WhatsApp for the user. **The application does not automatically send WhatsApp messages**; the user reviews and sends the prepared message in WhatsApp.

Email functionality includes scheduled daily brief/reminder messages and a test-email workflow. Recipients and reminder behavior are controlled through Settings. Delivery depends on valid recipient addresses, Apps Script authorization, trigger execution, and Google's email quotas.

### 9. Automation and Daily Maintenance

Apps Script time-driven triggers support the background processes.

| Process | Description |
|---|---|
| Automatic calendar sync | Runs the calendar synchronization workflow at the configured interval. |
| Meeting alerts | Checks for meetings approaching their configured alert window during automatic sync. |
| Daily maintenance | Runs scheduled maintenance tasks. |
| Daily brief email | Sends a summary email to the configured owner/EA recipients when enabled. |

Automation can be installed, refreshed, or turned off through the available administrative controls. Trigger state is also stored so the application can display automation status to users who do not own the triggers.

Time-driven triggers are scheduled by Google and may not execute at the exact minute selected.

### 10. Reports

The Reports page supports generating a PDF report for a selected date range. The report includes relevant meeting, action-item, and follow-up information and status summaries.

Generated report files are saved in the configured Google Drive report folder, and the application provides a link to the resulting file. Drive authorization and folder access are required.

### 11. Branding and Settings

The Settings area provides configuration for the options supported by the current implementation, including:

- Company name and tagline.
- Primary and accent brand colours.
- Company logo, including restoring the default logo.
- Owner and executive-assistant contact information.
- Office address and map link.
- Calendar ID and synchronization preferences.
- Meeting alert and daily reminder options.
- Reminder recipients and report folder name.
- Default meeting duration, start time, type, priority, mode, and location.
- Configurable dropdown lists.
- System health checks.

Settings are saved through the web app. Logo data is stored in Apps Script properties rather than as a public image file.

### 12. Interactive Instructions Guide

The application includes an in-app guide that explains the current interface, tabs, forms, controls, and workflows. The guide can provide English and Hinglish/voice-oriented assistance.

The guide is designed to explain the interface; it does not save records, delete data, or send messages on the user's behalf.

### 13. Audit and Sync Logs

The application maintains operational records for supported changes and calendar synchronization activity.

- Audit Log: supported record and administrative changes.
- Calendar Sync Log: synchronization outcomes, errors, and diagnostic details.
- Calendar Sync Registry: meeting-to-calendar event mappings and synchronization metadata.

Shared locking, validation, and unique record IDs help reduce conflicting writes and invalid data. Logs are intended to assist troubleshooting and are not a guarantee that every external event or failure will be captured.

## Open To All and Access Control

The application includes an **Open to all** option in **Settings → Contacts**.

- **No:** access is limited according to the configured email access list and role checks.
- **Yes:** the application can be made available to people with Google accounts, subject to the web-app deployment configuration and server-side authorization.

The access model distinguishes the Owner/EA administrative role from the VIEWER experience. Viewers receive a **VIEWER** badge and a read-only interface for records. The view-only layer also disables settings fields and hides controls for changing Settings or using Calendar Sync.

The view-only browser code is maintained separately in `ViewOnlyJs.gs`. `WebApp.gs` injects it into the served HTML at runtime, immediately before the application boot code. This means the existing `Index.html` does not need to be edited to add this layer.

**Security is enforced on the server as well as in the browser.** The interface restrictions are a usability safeguard, not the sole security boundary. Backend role checks and viewer-denial checks protect restricted operations. The Open To All setting may also change spreadsheet sharing; review the sharing prompt and the spreadsheet's Share settings before enabling it.

> Before sharing the web app publicly or with a broad audience, verify the deployment's execution identity, access setting, spreadsheet sharing, owner/EA configuration, and server-side role checks. A public URL alone is not a security policy.

## System Architecture

```text
                 User / Google Account
                          |
                          v
                 Web App (Index.html)
                 + ViewOnlyJs.gs layer
                          |
                          v
                  Google Apps Script
             Authorization / Business Logic
                 /          |          \
                v           v           v
         Google Sheets  Google Calendar  Google Drive
         Records/Logs   Calendar API     PDF Reports
                ^
                |
          Time-driven triggers
       Auto sync / Alerts / Daily brief
```

### Technology Stack

| Technology | Purpose |
|---|---|
| Google Apps Script | Server-side logic, access control, data operations, automation, and integrations |
| JavaScript | Browser-side interactions and application logic |
| HTML and CSS | Web application interface and styling |
| Google Sheets | Structured records, settings, registry, and logs |
| Google Calendar API (Advanced Service) | Calendar event synchronization |
| Apps Script Triggers | Scheduled synchronization, alerts, and daily processes |
| MailApp / Gmail services | Reminder and daily brief email workflows |
| Google Drive services | PDF report storage and file access |

## Project Structure

The project is divided into files by responsibility. The table describes the modules represented in the supplied project files; retain the exact filenames used in the Apps Script editor.

| File | Responsibility |
|---|---|
| `Config.gs` | Central configuration, sheet names, field schemas, IDs, lists, and defaults |
| `Utils.gs` | Shared helpers, spreadsheet access, settings, caching, locking, role checks, validation, and logging |
| `Setup.gs` | Initial workbook setup, sheets, headers, formatting, dropdown validation, and setup verification |
| `Code.gs` | Spreadsheet menu, web-app launch, authorization entry point, automation installation, and maintenance utilities |
| `Dashboard.gs` | Dashboard data APIs, record serialization, validation, and create/update operations |
| `WebApp.gs` | Web-app entry point, bootstrap, branding, logo handling, and settings APIs |
| `ViewOnlyJs.gs` | Browser-side Open To All/view-only layer injected into the served web app |
| `Sync.gs` | Calendar sync orchestration, triggers, alerts, daily brief emails, PDF reports, and system health |
| `CalendarSync.gs` | Multi-calendar registry, calendar event reconciliation, reset/re-sync, and calendar-specific workflows |
| `Index.html` | Main web application interface and client-side JavaScript |
| `DashboardUI.html` | Spreadsheet dashboard dialog interface |
| `appsscript.json` | Apps Script runtime, time zone, Advanced Services, deployment settings, and OAuth scopes |

## Google Sheets Data Organization

The application uses Google Sheets as its primary structured data store. Core operational tabs include:

| Sheet | Purpose |
|---|---|
| `Dashboard` | Spreadsheet-side dashboard or summary surface, where created by setup |
| `Meeting Register` | Main meeting records and calendar sync status |
| `Meeting Minutes` | Meeting discussions and decisions |
| `Action Items` | Assigned tasks, deadlines, and progress |
| `Follow-Up Tracker` | Customer and meeting follow-up records |
| `Customers` | Reusable customer details |
| `Calendar Sync Log` | Calendar synchronization history and errors |
| `Calendar Sync Registry` | Event references for each meeting/calendar combination |
| `Audit Log` | Supported administrative and data-change activity |
| `Settings` | Company, access, calendar, reminder, defaults, and dropdown configuration |

The setup module is responsible for creating and formatting the required workbook structure. The exact sheet set should be verified against the deployed version using the project's setup verification function.

### Record IDs and Relationships

The application uses distinct IDs for meetings, minutes, action items, follow-ups, and customers. Minutes, action items, and follow-ups are associated with their parent meeting through the Meeting ID.

The ID prefixes and formats are defined centrally in the configuration module. Use the current `Config.gs` definitions as the source of truth if the formats are changed.

## Installation and Deployment

### Prerequisites

- A Google account with access to Google Apps Script.
- A Google spreadsheet to hold application data.
- Access to the Google Calendars that will be synchronized.
- Permission to authorize the Google Sheets, Calendar, Drive, Mail, and trigger services used by the project.
- The complete Apps Script source files, including the HTML files and `appsscript.json`.

### Setup

1. **Create or open the Apps Script project.** Add the project source files and HTML files using the filenames listed in Project Structure.
2. **Configure the manifest.** Review `appsscript.json`, the project time zone, enabled Advanced Services, and OAuth scopes. Enable the Google Calendar API Advanced Service if it is not already enabled.
3. **Link and authorize the spreadsheet and services.** Run `authorizeAll()` from the Apps Script editor and approve the requested permissions. This function links the active spreadsheet and checks Calendar, Drive, Mail, and trigger access.
4. **Run the initial workbook setup.** Run `setupWorkbook()` and allow the setup process to create the required sheets and formatting.
5. **Verify the workbook.** Run `verifySetup()` and resolve any reported problems.
6. **Configure application settings.** Open the web app and set the company branding, owner/EA details, calendar ID, reminders, defaults, dropdown lists, and report folder.
7. **Deploy the web app.** In Apps Script, select **Deploy → New deployment → Web app**. Choose the execution identity and access policy appropriate for the intended users, then deploy.
8. **Install automation.** Use the Calendar Sync/automation controls or run `installAutomation()` as an authorized administrator. Confirm that the sync and daily triggers are installed.
9. **Test the application.** Create a test meeting, run a manual calendar sync, verify the calendar event, test the reminder email, generate a sample PDF report, and review the logs.
10. **Configure sharing carefully.** If enabling Open To All, read the sharing prompt and verify the spreadsheet's sharing settings and web-app deployment access before distributing the URL.

The first-time setup, authorization, deployment, and trigger installation require administrator action. Scheduled processes run only after the required permissions and triggers are in place.

## Typical Workflow

### Initial administration

1. Authorize the Apps Script project and link the spreadsheet.
2. Create and verify the workbook structure.
3. Configure branding, contacts, calendar, reminders, defaults, and dropdown lists.
4. Deploy the web app and configure access.
5. Install automation and test the integrations.

### Daily use

1. Review the Overview dashboard.
2. Create or update meeting records.
3. Use calendar synchronization to publish or reconcile meeting events.
4. Record meeting minutes after meetings.
5. Create and assign action items.
6. Maintain follow-up dates and outcomes.
7. Update customer details when needed.
8. Generate date-range PDF reports when required.
9. Review synchronization or audit logs when an operation needs attention.

## Customization and Maintenance

Many company-specific options are managed through Settings. For code-level changes, related layers may need to be updated together:

1. `Config.gs` — define or change sheet names, fields, IDs, settings, and dropdown definitions.
2. `Setup.gs` — update sheet headers, formatting, and validation.
3. Backend modules — implement reading, validation, authorization, and saving.
4. `Index.html` — update the web interface when the user-facing workflow changes.
5. `ViewOnlyJs.gs` and `WebApp.gs` — update the injected view-only behavior or its insertion point if access behavior changes.
6. `Sync.gs` and `CalendarSync.gs` — update calendar, automation, alert, and report behavior when those workflows depend on the changed fields.

Test changes with non-production data before deploying them to a live workbook.

## Security and Operational Notes

- Do not commit real customer data, private spreadsheets, credentials, access tokens, or sensitive configuration values to a public repository.
- Review Apps Script OAuth scopes and grant only the permissions needed by the application.
- Restrict web-app deployment access to the intended audience where possible.
- Configure and verify Owner/EA identities and role checks.
- Review spreadsheet sharing when changing Open To All.
- Browser-side hiding and disabling of controls must not replace backend authorization checks.
- Calendar event changes require the appropriate calendar permissions, including permission to make changes to events.
- Email and trigger operations are subject to Google quotas and service availability.
- Time-driven triggers may run later than their nominal scheduled time.
- WhatsApp shortcuts prepare messages for the user; they are not unattended WhatsApp delivery.
- PDF report creation requires Drive access and a valid report destination.

## Troubleshooting

| Problem | Checks |
|---|---|
| Web app does not open | Confirm deployment exists, the URL is current, and the account has access. |
| App cannot find the spreadsheet | Run `authorizeAll()` from the intended bound project/spreadsheet and confirm the spreadsheet link. |
| Setup reports missing sheets or headers | Run `verifySetup()` and review the setup output before using live data. |
| Calendar sync fails | Check Calendar ID, Advanced Calendar Service, calendar sharing permissions, and Calendar Sync Log details. |
| Automatic sync is not running | Check trigger installation, Apps Script execution history, authorization, and automation state. |
| Meeting alert or daily brief is missing | Check alert/reminder settings, recipient email addresses, trigger execution, and Mail quotas. |
| Viewer sees controls that should be restricted | Confirm `ViewOnlyJs.gs` is included, `doGet()` injects the script, and backend role checks are active. |
| Open To All access does not behave as expected | Review the saved Open To All setting, user role, web-app deployment access, and spreadsheet sharing. |
| PDF report generation fails | Check Drive authorization, report folder configuration, and Apps Script execution logs. |
| Dashboard data looks outdated | Refresh the dashboard and verify the underlying sheet records and filters. |

For diagnosis, review **Apps Script → Executions**, the Calendar Sync Log, and the Audit Log.

## Limitations

- Google Apps Script, Calendar API, Drive, email, and trigger services are subject to quotas and availability.
- Synchronization is not guaranteed to be instantaneous.
- Conflicts and failed external operations may require manual review.
- Meeting minutes, task assignments, customer responses, and follow-up outcomes require user input.
- WhatsApp messages are not sent automatically by this application.
- The effective access model depends on both the application's server-side checks and the Apps Script deployment and spreadsheet-sharing configuration.
- Confirm behavior in the deployed version after making code or configuration changes.

## License

Add a `LICENSE` file if you intend to distribute this project under a specific open-source license. Until a license is added, reuse and distribution are subject to the rights holder's applicable terms.

---

**Developed from scratch with Google Apps Script, Google Sheets, JavaScript, HTML, CSS, and Google Calendar API v3.**
