# Setup Guide

## 1. Create The Google Sheet

In the Google account that should own the tracker, create a new Google Sheet named:

```text
MVFT Tracking System
```

## 2. Open Apps Script

From the Google Sheet, go to:

```text
Extensions > Apps Script
```

Rename the Apps Script project to:

```text
MVFT Tracking System
```

## 3. Add The Script

Open `src/Code.gs` from this repository and paste its contents into the Apps Script `Code.gs` file.

If the Apps Script editor shows the manifest file `appsscript.json`, replace it with the contents of `src/appsscript.json`. If the manifest is hidden, go to:

```text
Project Settings > Show "appsscript.json" manifest file in editor
```

## 4. Run Initial Setup

In Apps Script, select the function:

```text
setupTracker
```

Click `Run`.

Google will ask for authorization. Approve the permissions using the Google account that should send assignment emails.

## 5. Install The Edit Trigger

In Apps Script, select:

```text
installTriggers
```

Click `Run`.

This installs an edit trigger so emails can be sent when new tickets are created or reassigned.

## 6. Add Members

Return to the Google Sheet and open the `Assignees` tab.

Add names and email addresses:

```text
Name	Email	Active
Jane Doe	jane@example.com	Yes
John Smith	john@example.com	Yes
```

You can also use:

```text
MVFT Tracker > Member > Add member
```

## 7. Create Tickets

Open the `Tickets` tab and add a new row with at least:

```text
Title
Assignee Email
Created By
Created By Email
```

The script will fill in ticket metadata and send the assignment notification email.

You can also use:

```text
MVFT Tracker > Task > Add task
```

The menu version prompts for the creator, task details, assignee, due date, and notes.

## Update Tickets

To update an existing ticket through the menu, select any cell in the ticket row and use:

```text
MVFT Tracker > Task > Update selected task
```

The creator listed in `Created By Email` receives an update email when meaningful fields change, such as title, description, priority, status, assignee, due date, or notes.

If editing directly in the sheet, make sure `Created By Email` is filled so creator update notifications can send.

## Manual Notification Run

If a ticket was added before the trigger was installed, use:

```text
MVFT Tracker > Task > Send pending assignments
```

This sends assignment emails for tickets that have not yet notified their current assignee.

## Reassignment Behavior

If the assignee email changes, the script sends a new notification to the new email address. It will not resend to the same email unless the `Assignment Notified Email` value is cleared or changed.

## Settings

Use the `Settings` tab or this menu item to update defaults:

```text
MVFT Tracker > Settings > Update setting
```

Available default settings include:

```text
Tracker Name
Notification Prefix
Default Priority
Default Status
Enable Assignment Emails
Enable Creator Update Emails
```

Set either email toggle to `No` if that notification type should be paused.
