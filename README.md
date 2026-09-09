# MVFT Tracking System

Google Sheets ticket and task tracker with Apps Script email notifications.

## What This Provides

- A structured ticket tracker for MVFT tasks and requests
- Auto-generated ticket IDs such as `MVFT-0001`
- Dropdowns for priority, status, and assignee
- Email notification when a ticket is assigned
- Duplicate-send protection using the last notified assignee email
- Manual menu actions inside Google Sheets

## Repository Structure

```text
mvft-tracking-system/
├── docs/
│   ├── google-sheet-template.md
│   └── setup.md
├── src/
│   ├── Code.gs
│   └── appsscript.json
├── .gitignore
└── README.md
```

## Google Sheet Tabs

The script creates and manages three tabs:

- `Tickets`: main tracker for tasks and tickets
- `Assignees`: list of team members and their email addresses
- `Settings`: basic configuration values

## Core Workflow

1. Create a Google Sheet in the Google account that will own the tracker.
2. Open `Extensions > Apps Script`.
3. Paste the contents of `src/Code.gs`.
4. Add the `src/appsscript.json` manifest settings if using the Apps Script editor manifest view.
5. Run `setupTracker`.
6. Run `installTriggers`.
7. Add assignees in the `Assignees` tab.
8. Create tickets in the `Tickets` tab.

When a row has a ticket title and assignee email, the assigned person receives an email.

## Important Note About Email Permissions

Google requires the account that owns/runs the Apps Script to authorize email sending. The first time `installTriggers` or `sendPendingAssignmentNotifications` is run, Google will ask for permissions.

The sender will be the Google account that authorized the script.

