# Meeting Scheduler & Management System

A professional, automation-driven Meeting Scheduler built with **Google Apps Script, JavaScript, HTML, Google Sheets, and Google Calendar API v3**. Designed to simplify meeting coordination, reduce repetitive administrative work, maintain meeting records, and track follow-ups from one centralized system.

## Overview

The Meeting Scheduler combines a web-based interface with Google Sheets for structured data management and Google Calendar for scheduling. It helps teams manage the complete meeting lifecycle — from scheduling and calendar synchronization to meeting minutes, assigned tasks, customer coordination, and follow-up tracking.

The system is designed for minimal manual workload: routine calendar synchronization and scheduled reminder processes can run automatically after initial setup. Users primarily need to enter meeting information, record decisions, and update the status of assigned work.

## Key Features

### 1. Meeting Management
- Create and manage meetings with unique meeting IDs.
- Maintain meeting dates, start and end times, titles, customers, meeting types, priorities, locations, statuses, and remarks.
- Track scheduled, rescheduled, completed, and cancelled meetings through the meeting records and applicable status workflows.
- View meeting information from the dashboard and Meeting Register.
- Filter meeting records by date range, type, status, customer, and priority.

### 2. Google Calendar Synchronization
- Integrates with Google Calendar using the Calendar API v3 Advanced Service.
- Synchronizes meeting records with configured Google Calendars.
- Supports automatic synchronization through time-driven triggers.
- Provides a manual **Sync Google Calendar** action for immediate synchronization.
- Maintains a synchronization registry with a separate record for each meeting and calendar combination.
- Records synchronization status, last synchronization time, errors, and event references.
- Includes reset and re-synchronization workflows.
- Supports synchronization-related error tracking and conflict resolution where available in the implementation.
- Handles registered calendar copies when applicable, including cancellation and deletion workflows across registered calendars.

**Important:** Calendar synchronization is not instantaneous at all times. Automatic updates depend on the configured trigger interval, Google service availability, permissions, and successful execution. Some conflicts or failed operations may require manual review.

### 3. Automated Reminders and Maintenance
The system supports scheduled background tasks through Google Apps Script time-driven triggers.

- **Automatic calendar synchronization:** Runs at the configured interval.
- **Daily reminder process:** Runs at a configured hour in the `Asia/Kolkata` time zone.
- **Daily maintenance:** Executes as part of the daily scheduled process.
- **Configurable reminders:** Reminder settings can be adjusted through the Settings configuration.
- **Test reminders:** Administrators can test reminder email functionality.
- **Trigger management:** Automation can be installed, refreshed, or switched off through the available controls.

The supported synchronization intervals are **1, 5, 10, 15, or 30 minutes**, with a default of 5 minutes. The default daily reminder hour is 9 AM, subject to the saved settings and Apps Script trigger scheduling.

### 4. Meeting Minutes
- Record meeting discussions, decisions, and outcomes.
- Link minutes to the relevant meeting using its Meeting ID.
- Maintain a structured record of meeting history.
- Use unique meeting-minutes IDs for record identification.

Meeting minutes still require a person to enter or verify the actual discussion and decisions. The system helps organize the information rather than automatically generating the contents of a meeting.

### 5. Action Item Tracking
- Create action items associated with meetings.
- Assign responsibilities to individuals.
- Set priorities, assignment dates, and deadlines.
- Track task status and overdue days.
- Review pending work through the dashboard and Action Items sheet.

This reduces the need to maintain separate task lists or manually calculate overdue work.

### 6. Follow-Up Management
- Maintain follow-up records linked to meetings and customers.
- Track follow-up dates, statuses, and relevant remarks.
- Keep outstanding follow-up work organized in a dedicated tracker.

Follow-up records and their outcomes still need to be entered or updated by the responsible person.

### 7. Customer Management and Location Requests
- Maintain customer information in a centralized Customers sheet.
- Link customer records to meetings.
- Provide a customer location request workflow.
- Support meeting coordination using location information and configured contact details.

Customer responses and location sharing depend on the customer and the communication method used. A location request does not guarantee that the customer will respond or share their location.

### 8. WhatsApp and Email Communication
- Supports WhatsApp-related meeting communication shortcuts where configured in the interface.
- Supports email reminders through Google Apps Script.
- Provides a test reminder workflow for checking email functionality.
- Uses configurable reminder settings and recipients.

WhatsApp shortcuts should be understood as communication assistance, not as a claim of unattended WhatsApp message delivery. Email reminders depend on valid recipient settings, granted permissions, successful trigger execution, and applicable Google sending limits.

### 9. Dashboard and Reports
- Centralized dashboard for meeting and action-item visibility.
- Meeting filters for more focused views.
- Quick navigation to the Meeting Register, Minutes, Action Items, Follow-Up Tracker, and related workflows.
- Dashboard refresh and manual calendar synchronization controls.
- Meeting report generation and PDF/report workflows where configured.

### 10. Audit Logs and Synchronization Logs
- Maintains audit records for supported administrative and data-change operations.
- Maintains calendar synchronization logs and error details.
- Uses unique IDs to identify meetings, minutes, action items, follow-ups, and customers.
- Includes locking and validation helpers to reduce conflicting updates and invalid data.

Logs help with troubleshooting and traceability. They should not be interpreted as a guarantee that every possible external action or failure is captured.

## How the Automation Works

The system separates routine background work from actions that require human input.

| Process | How it works | Manual work required |
|---|---|---|
| Calendar synchronization | A time-driven trigger runs the synchronization process at the configured interval. | Normally none after setup; resolve failures when needed. |
| Daily reminders | A daily trigger runs maintenance and then the configured reminder process. | Configure recipients and reminder preferences; review failures. |
| Meeting creation | The user enters meeting details through the available interface. | Enter and verify the meeting information. |
| Meeting records | The backend saves supported data to Google Sheets and maintains related records. | Review the saved information when necessary. |
| Meeting minutes | Minutes are recorded against the meeting. | Enter the actual discussions and decisions. |
| Action items | Tasks are associated with meetings and tracked by status and deadline. | Assign work and update task progress. |
| Follow-ups | Follow-up records organize the next steps. | Record outcomes and update follow-up status. |
| Customer location | A location request can be initiated through the interface. | Initiate the request and follow up if the customer does not respond. |
| Reporting | Available reporting workflows produce meeting reports. | Select the required report options and review the output. |
| Troubleshooting | Sync and audit logs help identify issues. | Investigate permission errors, conflicts, or unsuccessful operations. |

### Typical daily workflow

**Initial setup — once per deployment**
1. Configure the Apps Script project and authorize the required services.
2. Run the workbook setup process.
3. Configure company, owner/EA, calendar, and reminder settings.
4. Deploy the web app.
5. Install or refresh the automation triggers.
6. Test calendar synchronization and email reminders.

**Routine operation**
1. Create a meeting and enter the necessary details.
2. Allow the configured synchronization process to update Google Calendar, or use manual synchronization when needed.
3. Record meeting minutes after the meeting.
4. Assign action items and deadlines.
5. Update follow-up records as work progresses.
6. Review the dashboard and logs when necessary.

Once configured correctly, routine synchronization and scheduled background processing do not require the user to run the same functions manually each time.

## System Architecture

```text
              User / Administrator
                       |
                       v
            Web App / HTML Interface
                       |
                       v
              Google Apps Script
              Backend and Logic
                /           \
               v             v
       Google Sheets     Google Calendar
       Data and Logs      Calendar API v3
               ^
               |
       Time-Driven Triggers
       - Automatic Sync
       - Daily Maintenance
       - Daily Reminders
```

### Technology Stack

| Technology | Purpose |
|---|---|
| Google Apps Script | Backend logic, automation, authorization, and integrations |
| JavaScript | Application logic and client-side interactions |
| HTML and CSS | Web interface and presentation |
| Google Sheets | Structured records, settings, dashboard, and logs |
| Google Calendar API v3 | Calendar event synchronization |
| Apps Script Triggers | Scheduled background execution |
| Google Mail services | Email reminder functionality |
| Google Drive services | Report and file workflows where configured |

## Project Structure

The Apps Script project is organized into separate files by responsibility.

| File | Responsibility |
|---|---|
| `Config.gs` | Central configuration, sheet names, field schemas, IDs, colors, and defaults |
| `Utils.gs` | Shared helpers, settings access, locks, IDs, role checks, and logging |
| `Setup.gs` | Workbook setup, headers, formatting, dropdowns, and dashboard structure |
| `Code.gs` | Spreadsheet menu, web-app launch, automation installation, and maintenance helpers |
| `Dashboard.gs` | Dashboard data, record operations, validation, and entity handling |
| `WebApp.gs` | Web-app entry point, initialization, branding, and settings APIs |
| `Sync.gs` | Calendar synchronization support, triggers, reminders, reporting, and maintenance |
| `CalendarSync.gs` | Multi-calendar registry, re-synchronization, and calendar-specific workflows |
| `Index.html` | User interface and client-side JavaScript |
| `appsscript.json` | Runtime, time zone, Advanced Services, deployment settings, and OAuth scopes |

The exact files in your GitHub repository should match the files present in your current Apps Script project.

## Data Organization

The system uses Google Sheets as its data and operational storage layer.

| Sheet | Purpose |
|---|---|
| `Dashboard` | Meeting overview, filters, quick actions, and action-item visibility |
| `Meeting Register` | Main meeting records and synchronization status |
| `Meeting Minutes` | Meeting discussions, decisions, and minutes |
| `Action Items` | Assigned tasks, deadlines, and progress |
| `Follow-Up Tracker` | Follow-up activities and status |
| `Customers` | Customer records |
| `Calendar Sync Log` | Synchronization history and errors |
| `Calendar Sync Registry` | Meeting-to-calendar event references and synchronization metadata |
| `Audit Log` | Supported change and administrative activity records |
| `Settings` | Configurable application settings |

### Unique ID System

The application uses structured identifiers to make records easier to reference.

- Meetings: `KE-MTG-YYYY-001`
- Meeting minutes: `KE-MOM-YYYY-001`
- Action items: `KE-ACT-YYYY-001`
- Follow-ups: `KE-FUP-YYYY-001`
- Customers: `KE-CUS-001`

Related minutes, tasks, and follow-ups can be associated with their parent meeting through the Meeting ID.

## Customization

The application is designed to centralize many configuration choices instead of scattering company-specific values throughout the code.

Depending on the setting exposed by the current implementation, administrators can configure:

- Company branding and logo.
- Owner and executive-assistant details.
- Office address and map information.
- Calendar configuration.
- Automatic synchronization interval.
- Reminder enablement and reminder hour.
- Reminder recipients and meeting-alert preferences.
- Default meeting options.
- Report storage settings and other supported preferences.

### Adding or changing functionality

For a new feature or field, the relevant parts of the application may need to be updated together:

1. `Config.gs` — define the field, setting, or dropdown option.
2. `Setup.gs` — create or update sheet headers, formatting, and validation.
3. Backend files — implement validation, saving, reading, and business logic.
4. `Index.html` — add or update the corresponding interface.
5. Synchronization and reporting files — update them if the new information must appear in Calendar events or reports.

**Important:** Changing a configuration value alone does not automatically implement a new feature. Any dependent backend, interface, data, and integration logic must also be updated and tested.

## Installation and Deployment

### Prerequisites

- A Google account with access to Google Apps Script.
- A Google spreadsheet for application data.
- Access to the Google Calendars the application must manage.
- Permission to authorize Calendar, Sheets, Drive, email, and trigger operations used by the project.
- The project files listed in the repository.

### Setup steps

1. **Create or open the Apps Script project.** Add the project files and ensure the required functions and file names match the implementation.
2. **Configure `appsscript.json`.** Set the appropriate time zone, enable the Calendar Advanced Service, and retain only the OAuth scopes needed by the application.
3. **Run the initial setup.** Run `setupWorkbook()` and confirm the creation of the required sheets and formatting. Existing data should be reviewed before any migration or structural changes.
4. **Verify the setup.** Run `verifySetup()` and resolve any reported issues.
5. **Authorize the services.** Run the project's authorization function, such as `authorizeAll()` if it is present in the deployed version, and grant the requested permissions.
6. **Configure Settings.** Enter the appropriate calendar, owner/EA, reminder, branding, and report details.
7. **Deploy the web app.** In Apps Script, select **Deploy → New deployment → Web app**, configure the execution identity and access restrictions, and deploy.
8. **Install automation.** Run `installAutomation()` or use the application's equivalent control. Approve the required permissions and verify that the triggers exist.
9. **Test the workflows.** Create a test meeting, run a manual calendar sync, verify the calendar event, test a reminder email, and check the logs.
10. **Use the deployed application.** Share the web-app URL only with the intended users and according to the chosen access policy.

The first deployment and authorization require manual work. Automation begins only after setup, permissions, configuration, and trigger installation are complete.

## Security and Permissions

The application includes role-checking and audit-logging helpers. The current configuration uses the Apps Script web-app deployment options `USER_ACCESSING` and `ANYONE`; actual access and authorization behavior depends on the deployment, Google account context, Workspace policies, and backend role checks.

Before using the application with real business data:

- Restrict deployment access to the intended users wherever possible.
- Configure and verify owner/EA authorization settings.
- Review every operation that reads or modifies meetings, settings, customers, and logs.
- Avoid publishing confidential information, private customer details, access tokens, passwords, or credentials to GitHub.
- Keep private business data out of screenshots and example records.
- Review OAuth scopes and grant only the permissions the application needs.

A publicly accessible web-app URL should not be treated as proof that the backend is secure. Verify the effective access controls before sharing it.

## Troubleshooting

| Issue | What to check |
|---|---|
| Calendar events are not updating | Calendar ID, calendar permissions, Advanced Calendar Service, trigger installation, and synchronization logs |
| Automatic sync is not running | Apps Script triggers, execution history, authorization, and configured interval |
| Reminder email is not received | Reminder settings, recipient addresses, mail permissions, trigger execution, and sending limits |
| Dashboard information appears outdated | Use the refresh action and verify the backend data and filters |
| Calendar event was edited externally | Run synchronization and review the available conflict-resolution workflow |
| Report generation fails | Drive permissions, report settings, execution logs, and required input |
| Web app cannot access data | Spreadsheet linking, deployment configuration, account permissions, and role settings |

For troubleshooting, review **Apps Script → Executions**, the Calendar Sync Log, and the Audit Log where applicable.

## Limitations and Operational Notes

- Google Apps Script execution, trigger scheduling, email sending, and API usage are subject to Google quotas and service availability.
- Time-driven triggers are scheduled by Google and may not execute at the exact minute requested.
- Calendar synchronization and spreadsheet writes are separate operations; one may succeed while the other fails.
- Automated synchronization does not eliminate the need to review failed operations or conflicts.
- Meeting minutes, action-item assignments, customer responses, and follow-up outcomes require human input unless a separate integration is implemented.
- WhatsApp-related controls do not imply an independently authorized WhatsApp Business API integration.
- Actual behavior should be verified against the deployed version and its execution logs.

## Intended Use

This project is suitable for small teams, business owners, executive assistants, and administrative teams that need a centralized way to manage meetings, calendar events, minutes, responsibilities, and follow-ups while reducing repetitive administrative work.

## License

Add a `LICENSE` file to the repository if you intend to distribute this project under a specific open-source license. Until then, the repository's reuse and distribution terms are determined by the rights holder and any applicable third-party terms.

---

**Built with Google Apps Script, Google Sheets, JavaScript, HTML, and Google Calendar API v3.**
