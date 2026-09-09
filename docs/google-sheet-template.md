# Google Sheet Template

You can let the script create these tabs automatically by running `setupTracker`, or create them manually first.

## Tickets Tab

Create a tab named:

```text
Tickets
```

Use this header row:

```text
Ticket ID	Created At	Updated At	Created By	Created By Email	Title	Description	Priority	Status	Assignee	Assignee Email	Due Date	Assignment Notification Status	Assignment Notified Email	Assignment Notified At	Update Notification Status	Update Notified At	Notes
```

`Created By` should be the requester or creator name. `Created By Email` is used for update notifications when the ticket changes.

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

## Assignees Tab

Create a tab named:

```text
Assignees
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
