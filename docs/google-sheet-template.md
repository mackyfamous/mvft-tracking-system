# Google Sheet Template

You can let the script create these tabs automatically by running `setupTracker`, or create them manually first.

## Tasks Tab

Create a tab named:

```text
Tasks
```

Use this header row:

```text
Task ID	Created At	Updated At	Created By	Created By Email	Title	Description	Priority	Status	Assignee	Assignee Email	Due Date	Notes	Assignment Notification Status	Assignment Notified Email	Assignment Notified At	Update Notification Status	Update Notified At
```

`Created By` should be the requester or creator name. `Created By Email` is used for update notifications when the task changes.

When adding a task from the menu, the `Created by` dropdown is preselected from the current Google user when Apps Script can read the user's email. If Google does not provide the user and the creator is not in `Members`, choose `Enter manually`. When a task is created by typing directly into the sheet, blank creator fields are filled from the editor email when Google makes it available.

Use names from `Members` for `Created By` and `Assignee` when possible. The script fills the matching email fields from the member list.

`Notes` is the last normal task-entry field. The columns after `Notes` are notification metadata used by the script.

Date columns display as `yyyy-MM-dd` and remain real date values for sorting/filtering.

Recommended dropdown values:

Priority:

```text
Low
Medium
High
Urgent
```

Status:

```text
Open
In Progress
Waiting
Completed
Cancelled
```

## Members Tab

Create a tab named:

```text
Members
```

Use this header row:

```text
Name	Email	Active
```

The `Active` column should use:

```text
Yes
No
```

## Settings Tab

Create a tab named:

```text
Settings
```

Use these starting values:

```text
Setting	Value
Tracker Name	MVFT Tracking System
Notification Prefix	[MVFT]
Default Priority	Medium
Default Status	Open
Enable Assignment Emails	Yes
Enable Creator Update Emails	Yes
```
