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

Then add three HTML files in Apps Script:

```text
TaskDialog
MemberDialog
SettingDialog
```

Paste the matching repository files into them:

```text
src/TaskDialog.html
src/MemberDialog.html
src/SettingDialog.html
```

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

This installs an edit trigger so emails can be sent when new tasks are created or reassigned.

## 6. Add Members

Return to the Google Sheet and open the `Members` tab.

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

## 7. Create Tasks

Open the `Tasks` tab and add a new row with at least:

```text
Title
Assignee
Assignee Email
Due Date
```

The script will fill in task metadata and send the assignment notification email after the required task fields are complete. If you select an assignee from `Members`, the script fills `Assignee Email`. If `Created By` and `Created By Email` are blank on a direct sheet edit, the edit trigger fills them from the editor's Google email when Google makes that email available.

For direct sheet entry, notifications wait until `Created By`, `Created By Email`, `Title`, `Priority`, `Status`, `Assignee`, `Assignee Email`, and `Due Date` are filled. `Description` and `Notes` are optional.

You can also use:

```text
MVFT Tracker > Task > Add task
```

The menu version opens one form dialog for the creator, task details, assignee, due date, and notes. The `Created by` dropdown is preselected with the current Google user when possible. If Google does not provide the current user, the dropdown starts with the first available member and still lets you choose another existing member.

Dates display in the sheet as `yyyy-MM-dd`.

## Update Tasks

To update an existing task through the menu, select any cell in the task row and use:

```text
MVFT Tracker > Task > Update selected task
```

This opens one form dialog with the selected task prefilled. The creator listed in `Created By Email` receives an update email when meaningful fields change, such as title, description, priority, status, assignee, due date, or notes.

If editing directly in the sheet and Google does not expose the editor email, make sure `Created By Email` is filled so creator update notifications can send.

## Manual Notification Run

If a task was added before the trigger was installed, use:

```text
MVFT Tracker > Task > Send pending assignments
```

This sends assignment emails for tasks that have not yet notified their current assignee.

## Reassignment Behavior

If the assignee email changes, the script sends a new notification to the new email address. It will not resend to the same email unless the `Assignment Notified Email` value is cleared or changed.

## Troubleshooting

If saving a task says a title or other field violates a priority/status validation rule, paste the latest `Code.gs`, save the Apps Script project, and run `setupTracker` once. This clears old validations from earlier tracker versions and reapplies them to the current columns.

If your sheet still has the old `Tickets` or `Assignees` tabs, run `setupTracker` once after updating the script. It renames them to `Tasks` and `Members`, changes `Ticket ID` to `Task ID`, and moves `Notes` after `Due Date`.

## Settings

Use the `Settings` tab or this menu item to update defaults:

```text
MVFT Tracker > Settings > Update setting
```

This opens one form dialog where you select a setting and update its value.

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
