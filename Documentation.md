# Documentation

## 1. Project Overview

Recurring Task Tracker is a Google Sheets and Google Apps Script prototype for scheduling recurring operational tasks, recording completion evidence, monitoring deadlines, and escalating unresolved tasks.

The spreadsheet acts as the data store and user interface. Apps Script handles recurring cycle generation, completion validation, status updates, email notifications, escalation, and audit logging.

## 2. Requirements

- A Google account with access to Google Sheets and Google Apps Script.
- A new Google Spreadsheet.
- The complete `Code.gs` source file from this repository.
- Permission to authorise spreadsheet access, installable triggers, and email sending if email features are enabled.

## 3. Installation

1. Create a new Google Spreadsheet.
2. Open **Extensions → Apps Script**.
3. Replace the editor's default code with the complete contents of `Code.gs`.
4. Save the Apps Script project.
5. In Apps Script, open **Project Settings** and set the time zone to `Asia/Kolkata`.
6. Select `setupSystem` in the function selector and run it once.
7. Review the requested permissions and authorise the script.
8. Return to the spreadsheet and refresh it.
9. Open the **Task System** menu.

The setup function creates or prepares these tabs: `Cycles`, `Dashboard`, `Task_Master`, `Config`, `Completions`, `Audit_Log`, and `Test_Cases`. It also installs the edit and time-driven triggers.

## 4. Initial Configuration

Open the `Config` tab and review the settings.

| Key | Default | Purpose |
|---|---|---|
| TIMEZONE | Asia/Kolkata | Informational; the source code uses IST |
| WORKING_DAYS | Mon,Tue,Wed,Thu,Fri,Sat | Working days |
| HOLIDAYS | Blank | Optional comma-separated dates in `yyyy-MM-dd` format |
| DEFAULT_DUE_TIME | 17:00 | Fallback task deadline |
| MANAGER_NAME | Manager | Manager display/fallback name |
| MANAGER_EMAIL | manager@example.com | Manager notification recipient; replace before use |
| LEADERSHIP_EMAIL | leadership@example.com | Leadership escalation recipient; replace before use |
| TEST_EMAIL | testmail@example.com | Destination for TEST-mode and demo emails |
| EMAIL_MODE | OFF | Email sending mode |
| MANAGER_ALERT_DELAY_MIN | 30 | Delay after task deadline before manager alert |
| ACK_DEADLINE_TIME | 10:30 | Manager acknowledgement deadline on next working day |
| RESOLUTION_DEADLINE_TIME | 17:00 | Resolution deadline on next working day |
| ENGINE_INTERVAL_MIN | 15 | Time-driven engine interval |
| DEMO_NOW_OVERRIDE | Blank | Optional manual demonstration clock override |

### Email modes

- `OFF`: sends no email. This is the default and recommended initial setting.
- `TEST`: redirects notifications to `TEST_EMAIL`. Set `TEST_EMAIL` to an address you control before using this mode.
- `LIVE`: sends notifications to the configured manager and leadership recipients.

Do not enable `LIVE` until recipient addresses and the spreadsheet's access settings have been checked.

## 5. Configure Tasks

Use the `Task_Master` tab. The setup function inserts four example tasks when the tab is empty. Replace the sample task details with your own authorised demo or operational data.

Important fields:

| Field | Purpose |
|---|---|
| Task_ID | Unique task identifier |
| Task_Name | Task description |
| Assigned_To | Assigned employee or role |
| Employee_Email | Email used for conditional identity matching |
| Frequency_Days | Recurrence interval in days |
| Schedule_Mode | `CALENDAR_DAILY` or `ROLLING` |
| Proof_Type | `WHATSAPP_CHECKBOX` or `TICKET_ID` |
| Due_Time | Deadline in `HH:mm` format |
| Start_Date | First eligible date |
| Active | Whether the task is active |

Daily tasks use `CALENDAR_DAILY`. Weekly, fortnightly, and monthly examples use `ROLLING` with frequency values of 7, 15, and 30 days.

## 6. How the Workflow Operates

### Recurring cycles

- Daily tasks receive a new cycle on each working day.
- Rolling tasks receive their next cycle after completion, based on the actual completion date plus the configured frequency.
- Rolling due dates that fall on a non-working day are moved to the next working day.
- The system prevents creation of a duplicate cycle with the same cycle identifier and prevents multiple open rolling cycles for the same task.

### Completion

Employees enter the required proof in `Ticket_ID_or_Proof` where applicable and tick `Done` in `Cycles`.

For `TICKET_ID`, the script checks the submitted value against a format pattern. It does not connect to or verify the external portal.

For `WHATSAPP_CHECKBOX`, completion is an employee self-attestation. The script does not inspect WhatsApp or verify that a message was posted.

The script records completion details in `Cycles`, adds a record to `Completions`, and writes events to `Audit_Log`.

### Manager acknowledgement and escalation

Manager acknowledgement is separate from task completion. The manager ticks `Manager_Acknowledged` on an alerted cycle and may add a `Manager_Note`.

The configured workflow is:

| Event | Time |
|---|---|
| Task deadline | 5:00 PM |
| Manager alert | 5:30 PM, based on a 30-minute delay |
| Manager acknowledgement deadline | Next working day, 10:30 AM |
| Leadership escalation | If the manager has not acknowledged by the acknowledgement deadline |
| Unresolved task escalation | If acknowledged but still unresolved by 5:00 PM on the next working day |

The engine runs at the configured interval. With the default 15-minute interval, processing can occur up to approximately 15 minutes after a threshold.

## 7. Triggers

`setupSystem()` installs two triggers:

- `onEditHandler`: processes changes to the `Done` and `Manager_Acknowledged` checkboxes in `Cycles`.
- `engineTrigger`: runs the scheduling, status-refresh, notification, and escalation engine on a time-driven schedule.

Installable triggers execute under the Google account that created them. That account's authorisation is used for operations such as sending email. Running setup again replaces the project's existing triggers for these two handler functions.

## 8. Menu Actions

The spreadsheet menu **Task System** provides:

- **Setup / Initialise System**: prepares sheets, formatting, validations, dashboard, and triggers.
- **Run Engine Now**: manually runs the operational engine.
- **Run Demo / Test**: adds isolated `DEMO_` records and runs simulated scenarios.

The demo runner requires `EMAIL_MODE` to be `TEST` or `LIVE` and requires `TEST_EMAIL`. It asks for confirmation and sends approximately three test emails to `TEST_EMAIL`. Demo records are marked inactive and excluded from the real engine and dashboard. The runner updates the Actual_Result, Status, and Evidence/Notes fields in `Test_Cases`.

## 9. Test Coverage and Verification Limits

The included test scenarios cover recurrence, proof validation, daily completion, manager alerts, acknowledgement, leadership escalation, SLA thresholds, and duplicate prevention.

The test runner invokes handlers directly with simulated dates. Passing these tests does not establish complete live end-to-end behaviour through actual spreadsheet checkbox edits and scheduled triggers. Validate those paths separately in a controlled spreadsheet before operational use.

## 10. Data Access and Security Limitations

- Do not commit real employee names, email addresses, task records, spreadsheet URLs, spreadsheet IDs, or operational logs to a public repository.
- `example.com` addresses in the sample data are placeholders.
- Warning-only protected ranges are not a security boundary.
- Spreadsheet viewers can still see data in the spreadsheet; this prototype does not provide employee-level row confidentiality.
- The identity recorded as `GOOGLE_EMAIL` is used only when the editor's email is available and matches the configured employee or manager email. Otherwise, the code records a fallback identity that is explicitly marked as unverified.
- Review sharing permissions before adding employees or external users.
- Keep email mode `OFF` until recipients and access have been checked.

## 11. Setup Behaviour and Existing Data

`setupSystem()` prepares the required sheets and formatting, installs triggers, and writes a setup event to `Audit_Log`.

The setup code preserves existing rows in `Cycles`, `Completions`, and `Audit_Log`. However, it rebuilds the `Dashboard` using `clear()`. It may also delete an empty default `Sheet1` when the specified sheet-count condition is met. Review the spreadsheet before rerunning setup.

## 12. Repository Files

- `Code.gs`: Apps Script implementation.
- `README.md`: project overview and feature summary.
- `.gitignore`: excludes local configuration and credential files.
- `appsscript.json`: Apps Script project manifest.
- `Documentation.md`: installation, configuration, workflow, and limitations.

## 13. Disclaimer

This is a prototype intended for demonstration and further development. Test it with non-sensitive sample data and controlled email recipients before using it in a live operational environment.
