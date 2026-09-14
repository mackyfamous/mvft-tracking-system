# MVFT Tracking System

Google Sheets task tracker with Apps Script email notifications.

## What This Provides

- A structured task tracker for MVFT tasks and requests
- Auto-generated task IDs such as `MVFT-0001`
- Dropdowns for priority, status, and assignee
- Email notification when a task is assigned
- Creator tracking with `Created By` and `Created By Email`, prefilled from the current Google user when available
- Email notification to the creator when an existing task is updated
- Duplicate-send protection using the last notified assignee email
- Grouped menu actions for tasks, members, and settings inside Google Sheets

## Repository Structure

```text
mvft-tracking-system/
├── docs/
│   ├── assets/
│   ├── google-sheet-template.md
│   ├── setup.md
│   └── user-manual.md
├── src/
│   ├── Code.gs
│   ├── MemberDialog.html
│   ├── SettingDialog.html
│   ├── TaskDialog.html
│   └── appsscript.json
├── .gitignore
└── README.md
```

## Google Sheet Tabs

The script creates and manages three tabs:

- `Tasks`: main tracker for tasks
- `Members`: list of team members and their email addresses
- `Settings`: basic configuration values

If an older sheet has `Tickets` or `Assignees`, running `setupTracker` renames them to `Tasks` and `Members`.

## Core Workflow

1. Create a Google Sheet in the Google account that will own the tracker.
2. Open `Extensions > Apps Script`.
3. Paste the contents of `src/Code.gs`.
4. Add the HTML files from `src/TaskDialog.html`, `src/MemberDialog.html`, and `src/SettingDialog.html`.
5. Add the `src/appsscript.json` manifest settings if using the Apps Script editor manifest view.
6. Run `setupTracker`.
7. Run `installTriggers`.
8. Add members in the `Members` tab.
9. Create tasks in the `Tasks` tab or use `MVFT Tracker > Task > Add task`.

When a manually edited row has the required task fields filled, the assigned person receives an email. Selecting a known member as assignee fills `Assignee Email`.

When an existing task changes, the person in `Created By Email` receives an update email with the changed fields.

For direct sheet entry, notifications wait until `Created By`, `Created By Email`, `Title`, `Priority`, `Status`, `Assignee`, `Assignee Email`, and `Due Date` are filled. `Description` and `Notes` are optional. Script metadata columns do not count toward task readiness.

The task form preselects the current Google user in the `Created By` dropdown when Apps Script can read the user's email. If Google does not provide the current user, the dropdown starts with the first available member and still lets you choose another existing member.

`Notes` sits after `Due Date`. The columns after `Notes` are notification metadata used by the script and are hidden by setup.

Date columns are formatted in the sheet as `yyyy-MM-dd`.

## Google Sheets Menu

The script adds a `MVFT Tracker` menu with these groups:

- `Task`: add a task, update the selected task, and send pending assignment emails
- `Member`: add a member and update the selected member
- `Settings`: run setup, update a setting, and install the email trigger

The add and update actions open one form dialog with all fields, then save with one submit.

## Important Note About Email Permissions

Google requires the account that owns/runs the Apps Script to authorize email sending. The first time `installTriggers` or `sendPendingAssignmentNotifications` is run, Google will ask for permissions.

The sender will be the Google account that authorized the script.

Google may hide the editor email for some consumer accounts, shared files, or domain settings. In that case, the form still requires the creator name and email before saving.

## User Manual

See `docs/user-manual.md` for the screenshot-based staff/admin guide covering setup, roles, task creation, direct sheet entry, notifications, statuses, and dashboard suggestions.
