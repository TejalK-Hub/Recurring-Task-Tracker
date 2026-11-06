# Recurring Task Tracker

A Google Sheets and Google Apps Script-based automation system for managing recurring operational tasks, validating completion, monitoring deadlines, and escalating unresolved work.

## Problem Statement

Small operations teams often manage recurring responsibilities through separate WhatsApp updates, portal checks, spreadsheets, and manual reminders.

This creates fragmented task visibility, inconsistent follow-ups, unclear ownership, and delays in identifying overdue work. Managers may lack a central view of pending tasks, completion evidence, acknowledgement status, and unresolved issues.

**Recurring Task Tracker** addresses this problem by centralising recurring task management and automating task scheduling, completion validation, SLA monitoring, audit logging, and escalation workflows.

## Key Features

- Automated recurring task generation
- Daily, weekly, fortnightly, and monthly schedules
- Configurable working days and holidays
- Completion validation using task-specific proof requirements
- Ticket ID validation for portal-based tasks
- Automatic task status updates
- Overdue task monitoring
- Grouped manager email notifications
- Manager acknowledgement tracking
- Leadership escalation for unresolved tasks
- Completion ledger and audit trail
- Formula-driven manager dashboard
- Configurable notification modes and SLA settings
- Duplicate prevention for cycles, completions, and escalations

## Technology Stack

| Component | Technology |
|---|---|
| Spreadsheet interface | Google Sheets |
| Automation | Google Apps Script |
| Event handling | Installable onEdit trigger |
| Scheduled execution | Time-driven trigger |
| Data storage | Google Sheets |
| Reporting | Spreadsheet formulas |
| Notifications | Google Apps Script MailApp |

## Task Scheduling

The prototype supports four recurring task patterns.

| Frequency | Example | Scheduling Logic |
|---|---|---|
| Daily | WhatsApp group operational update | New cycle on every working day |
| Weekly | Portal verification | Next due date = completion date + 7 days |
| Fortnightly | Data quality check | Next due date = completion date + 15 days |
| Monthly | MIS report preparation | Next due date = completion date + 30 days |

**Working calendar:** Monday to Saturday.

Sundays and configured holidays are treated as non-working days. Rolling task deadlines falling on non-working days move to the next working day.

Daily tasks follow calendar-based scheduling, while weekly, fortnightly, and monthly tasks follow rolling scheduling based on actual completion dates.

## SLA and Escalation Workflow

| Event | Configured Time |
|---|---|
| Task deadline | 5:00 PM |
| Manager notification | 5:30 PM |
| Manager acknowledgement deadline | Next working day, 10:30 AM |
| Leadership escalation | If manager acknowledgement is missing |
| Unresolved task escalation | Next working day, 5:00 PM |

The automation engine runs at configurable intervals, with 15 minutes as the default.

Notifications may therefore be processed up to one engine interval after the configured threshold.

Manager acknowledgement and task completion are treated as separate actions.

## Spreadsheet Architecture

The system uses seven spreadsheet tabs.

| Sheet | Responsibility |
|---|---|
| Task_Master | Task definitions, ownership, recurrence, and proof requirements |
| Cycles | Generated task instances, deadlines, statuses, and escalation tracking |
| Dashboard | Manager-level operational overview |
| Config | Working calendar, notification settings, and SLA configuration |
| Completions | Completion records and submitted proof |
| Audit_Log | System events, actions, and escalation history |
| Test_Cases | Test scenarios and execution results |

## Email Configuration

Three notification modes are supported:

| Mode | Behaviour |
|---|---|
| OFF | No emails are sent |
| TEST | All emails are redirected to the configured TEST_EMAIL |
| LIVE | Emails are sent to configured recipients |

The source code defaults to `OFF`. A test recipient must be configured manually before running email-based tests.

## Setup Instructions

1. Create a new Google Spreadsheet.
2. Open **Extensions → Apps Script**.
3. Replace the default script content with the contents of `Code.gs`.
4. Save the project.
5. Set the project time zone to `Asia/Kolkata`.
6. Run `setupSystem()` once.
7. Review and approve the requested Google permissions.
8. Return to the spreadsheet and refresh it.
9. Open the **Task System** menu.
10. Review the generated sheets and configuration.

Keep `EMAIL_MODE` set to `OFF` during initial setup. Configure test recipients and deliberately switch to `TEST` before running email-based demonstrations.

Installable triggers execute under the Google account that created them.

## Testing

The project includes eight simulated test scenarios covering:

1. Seven-day recurring cycle generation
2. Completion validation and ticket ID requirements
3. Daily task completion
4. Manager notification
5. Manager acknowledgement
6. Leadership escalation under both conditions
7. SLA threshold handling
8. Duplicate prevention

The test runner uses simulated dates and direct handler calls. These results do not establish complete live end-to-end verification through actual spreadsheet edits and scheduled triggers.

## Important Implementation Notes

- WhatsApp completion is recorded through employee self-attestation. The system does not independently verify WhatsApp messages.
- Ticket ID validation checks the submitted format, not the external portal.
- Google account identity is recorded only when the editor's email is available and matches the configured assignee email. Otherwise, the system uses an explicitly marked fallback identity.
- Spreadsheet protection in this prototype is warning-only for system-managed ranges.
- Spreadsheet viewers may still see the underlying data. The prototype does not implement employee-level row confidentiality.
- Email delivery depends on Google permissions, recipient configuration, and available sending quotas.

## Future Scope

- WhatsApp API integration
- External portal verification
- Role-based access control
- Employee-specific authenticated interfaces
- Advanced operational analytics
- Integration with external reporting systems

## Project Objective

To demonstrate how lightweight spreadsheet-based automation can replace fragmented recurring task tracking with a structured, auditable, and configurable operational workflow.
