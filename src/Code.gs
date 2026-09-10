const CONFIG = {
  trackerName: 'MVFT Tracking System',
  notificationPrefix: '[MVFT]',
  sheets: {
    tasks: 'Tasks',
    members: 'Members',
    settings: 'Settings',
  },
  priorities: ['Low', 'Medium', 'High', 'Urgent'],
  statuses: ['Open', 'In Progress', 'Waiting', 'Completed', 'Cancelled'],
  activeOptions: ['Yes', 'No'],
  defaultPriority: 'Medium',
  defaultStatus: 'Open',
  taskPrefix: 'MVFT',
};

const TASK_HEADERS = [
  'Task ID',
  'Created At',
  'Updated At',
  'Created By',
  'Created By Email',
  'Title',
  'Description',
  'Priority',
  'Status',
  'Assignee',
  'Assignee Email',
  'Due Date',
  'Notes',
  'Assignment Notification Status',
  'Assignment Notified Email',
  'Assignment Notified At',
  'Update Notification Status',
  'Update Notified At',
];

const MEMBER_HEADERS = ['Name', 'Email', 'Active'];
const SETTINGS_HEADERS = ['Setting', 'Value'];

const COL = {
  TASK_ID: 1,
  CREATED_AT: 2,
  UPDATED_AT: 3,
  CREATED_BY: 4,
  CREATED_BY_EMAIL: 5,
  TITLE: 6,
  DESCRIPTION: 7,
  PRIORITY: 8,
  STATUS: 9,
  ASSIGNEE: 10,
  ASSIGNEE_EMAIL: 11,
  DUE_DATE: 12,
  NOTES: 13,
  ASSIGNMENT_NOTIFICATION_STATUS: 14,
  ASSIGNMENT_NOTIFIED_EMAIL: 15,
  ASSIGNMENT_NOTIFIED_AT: 16,
  UPDATE_NOTIFICATION_STATUS: 17,
  UPDATE_NOTIFIED_AT: 18,
};

const TASK_HEADER_ALIASES = {
  'Ticket ID': 'Task ID',
  'Notification Status': 'Assignment Notification Status',
  'Notified Email': 'Assignment Notified Email',
  'Notified At': 'Assignment Notified At',
};

const DEFAULT_SETTINGS = [
  ['Tracker Name', CONFIG.trackerName],
  ['Notification Prefix', CONFIG.notificationPrefix],
  ['Default Priority', CONFIG.defaultPriority],
  ['Default Status', CONFIG.defaultStatus],
  ['Enable Assignment Emails', 'Yes'],
  ['Enable Creator Update Emails', 'Yes'],
];

const UPDATE_NOTIFICATION_FIELDS = [
  ['title', 'Title'],
  ['description', 'Description'],
  ['priority', 'Priority'],
  ['status', 'Status'],
  ['assignee', 'Assignee'],
  ['assigneeEmail', 'Assignee Email'],
  ['dueDate', 'Due Date'],
  ['notes', 'Notes'],
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('MVFT Tracker')
    .addSubMenu(
      ui.createMenu('Task')
        .addItem('Add task', 'addTask')
        .addItem('Update selected task', 'updateSelectedTask')
        .addItem('Send pending assignments', 'sendPendingAssignmentNotifications')
    )
    .addSubMenu(
      ui.createMenu('Member')
        .addItem('Add member', 'addMember')
        .addItem('Update selected member', 'updateSelectedMember')
    )
    .addSubMenu(
      ui.createMenu('Settings')
        .addItem('Set up tracker', 'setupTracker')
        .addItem('Update setting', 'updateSetting')
        .addItem('Install email trigger', 'installTriggers')
    )
    .addToUi();
}

function setupTracker() {
  ensureTrackerReady_();
  alertUser_('MVFT Tracking System setup is complete.');
}

function ensureTrackerReady_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketsSheet = getOrCreateSheet_(ss, CONFIG.sheets.tasks, ['Tickets']);
  const assigneesSheet = getOrCreateSheet_(ss, CONFIG.sheets.members, ['Assignees']);
  const settingsSheet = getOrCreateSheet_(ss, CONFIG.sheets.settings);

  setHeaderRow_(ticketsSheet, TASK_HEADERS, TASK_HEADER_ALIASES);
  setHeaderRow_(assigneesSheet, MEMBER_HEADERS);
  setHeaderRow_(settingsSheet, SETTINGS_HEADERS);

  seedSettings_(settingsSheet);
  formatTicketsSheet_(ticketsSheet);
  formatAssigneesSheet_(assigneesSheet);
  formatSettingsSheet_(settingsSheet);
  applyValidations_(ticketsSheet, assigneesSheet);
  syncTicketCounter_(ticketsSheet);
  syncTicketSnapshots_(ticketsSheet);

  return { ss, ticketsSheet, assigneesSheet, settingsSheet };
}

function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers()
    .filter((trigger) => ['handleTaskEdit', 'handleTicketEdit'].includes(trigger.getHandlerFunction()))
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('handleTaskEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  alertUser_('Email notification trigger installed.');
}

function handleTicketEdit(e) {
  handleTaskEdit(e);
}

function handleTaskEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.sheets.tasks) {
    return;
  }

  const firstRow = e.range.getRow();
  const lastRow = firstRow + e.range.getNumRows() - 1;
  if (lastRow <= 1) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const editorEmail = getEditActorEmail_(e);
    for (let row = Math.max(firstRow, 2); row <= lastRow; row += 1) {
      const previousTicket = getStoredTicketSnapshotForRow_(sheet, row);
      normalizeTicketRow_(sheet, row, editorEmail);
      maybeHydrateMemberEmails_(sheet, row);
      maybeSendAssignmentEmail_(sheet, row);
      const currentTicket = mapTicketRow_(getTicketRow_(sheet, row));
      maybeSendCreatorUpdateEmail_(sheet, row, previousTicket, currentTicket, editorEmail);
      saveTicketSnapshot_(currentTicket);
    }
  } finally {
    lock.releaseLock();
  }
}

function sendPendingAssignmentNotifications() {
  const { ticketsSheet: sheet } = ensureTrackerReady_();

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    alertUser_('No tasks found.');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let sentCount = 0;
  try {
    for (let row = 2; row <= lastRow; row += 1) {
      normalizeTicketRow_(sheet, row);
      maybeHydrateMemberEmails_(sheet, row);
      if (maybeSendAssignmentEmail_(sheet, row)) {
        sentCount += 1;
      }
      saveTicketSnapshot_(mapTicketRow_(getTicketRow_(sheet, row)));
    }
  } finally {
    lock.releaseLock();
  }

  alertUser_(`Pending notification run complete. Emails sent: ${sentCount}.`);
}

function addTask() {
  showTaskDialog_('add');
}

function updateSelectedTask() {
  showTaskDialog_('update');
}

function addMember() {
  showMemberDialog_('add');
}

function updateSelectedMember() {
  showMemberDialog_('update');
}

function updateSetting() {
  showSettingDialog_();
}

function showTaskDialog_(mode) {
  const { ticketsSheet, assigneesSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();
  const currentUser = getCurrentUserProfile_(assigneesSheet);
  let rowNumber = '';
  let ticket = defaultTaskDialogTicket_(currentUser);

  if (mode === 'update') {
    if (!activeRange || activeRange.getSheet().getName() !== CONFIG.sheets.tasks || activeRange.getRow() <= 1) {
      alertUser_('Select a task row in the Tasks tab before using Update selected task.');
      return;
    }

    rowNumber = activeRange.getRow();
    const row = getTicketRow_(ticketsSheet, rowNumber);
    if (isBlankTicketRow_(row)) {
      alertUser_('The selected task row is blank.');
      return;
    }

    ticket = serializeTicketForDialog_(mapTicketRow_(row));
  }

  showDialog_('TaskDialog', mode === 'add' ? 'Add Task' : 'Update Task', {
    mode,
    rowNumber,
    task: ticket,
    priorities: CONFIG.priorities,
    statuses: CONFIG.statuses,
    members: getMembers_(assigneesSheet),
    currentUser,
  });
}

function submitTaskForm(form) {
  const { ticketsSheet, assigneesSheet } = ensureTrackerReady_();
  const mode = normalizeMode_(form.mode, ['add', 'update']);
  const task = normalizeTaskForm_(form, assigneesSheet, getCurrentUserProfile_(assigneesSheet));

  if (mode === 'add') {
    const now = new Date();
    const rowNumber = ticketsSheet.getLastRow() + 1;
    const taskId = nextTaskId_();
    const row = emptyTicketRow_();

    row[COL.TASK_ID - 1] = taskId;
    row[COL.CREATED_AT - 1] = now;
    row[COL.UPDATED_AT - 1] = now;
    row[COL.CREATED_BY - 1] = task.createdBy;
    row[COL.CREATED_BY_EMAIL - 1] = task.createdByEmail;
    row[COL.TITLE - 1] = task.title;
    row[COL.DESCRIPTION - 1] = task.description;
    row[COL.PRIORITY - 1] = task.priority;
    row[COL.STATUS - 1] = task.status;
    row[COL.ASSIGNEE - 1] = task.assignee;
    row[COL.ASSIGNEE_EMAIL - 1] = task.assigneeEmail;
    row[COL.DUE_DATE - 1] = task.dueDate || '';
    row[COL.NOTES - 1] = task.notes;

    ticketsSheet.getRange(rowNumber, 1, 1, TASK_HEADERS.length).setValues([row]);
    maybeSendAssignmentEmail_(ticketsSheet, rowNumber);
    saveTicketSnapshot_(mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber)));
    ticketsSheet.setActiveRange(ticketsSheet.getRange(rowNumber, COL.TITLE));

    return { message: `Task ${taskId} was created.` };
  }

  const rowNumber = Number(form.rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber <= 1 || rowNumber > ticketsSheet.getLastRow()) {
    throw new Error('Select a valid task row before updating.');
  }

  const previousTicket = mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber));
  if (isBlankTicketRow_(getTicketRow_(ticketsSheet, rowNumber))) {
    throw new Error('The selected task row is blank.');
  }

  const row = getTicketRow_(ticketsSheet, rowNumber);
  row[COL.CREATED_BY - 1] = task.createdBy;
  row[COL.CREATED_BY_EMAIL - 1] = task.createdByEmail;
  row[COL.TITLE - 1] = task.title;
  row[COL.DESCRIPTION - 1] = task.description;
  row[COL.PRIORITY - 1] = task.priority;
  row[COL.STATUS - 1] = task.status;
  row[COL.ASSIGNEE - 1] = task.assignee;
  row[COL.ASSIGNEE_EMAIL - 1] = task.assigneeEmail;
  row[COL.DUE_DATE - 1] = task.dueDate || '';
  row[COL.NOTES - 1] = task.notes;

  ticketsSheet.getRange(rowNumber, 1, 1, TASK_HEADERS.length).setValues([row]);
  normalizeTicketRow_(ticketsSheet, rowNumber);
  maybeHydrateMemberEmails_(ticketsSheet, rowNumber);
  maybeSendAssignmentEmail_(ticketsSheet, rowNumber);

  const currentTicket = mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber));
  maybeSendCreatorUpdateEmail_(ticketsSheet, rowNumber, snapshotTicket_(previousTicket), currentTicket, getActiveUserEmail_());
  saveTicketSnapshot_(currentTicket);

  return { message: `Task ${currentTicket.taskId || rowNumber} was updated.` };
}

function showMemberDialog_(mode) {
  const { assigneesSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();
  let rowNumber = '';
  let member = { name: '', email: '', active: 'Yes' };

  if (mode === 'update') {
    if (!activeRange || activeRange.getSheet().getName() !== CONFIG.sheets.members || activeRange.getRow() <= 1) {
      alertUser_('Select a member row in the Members tab before using Update selected member.');
      return;
    }

    rowNumber = activeRange.getRow();
    const row = assigneesSheet.getRange(rowNumber, 1, 1, MEMBER_HEADERS.length).getValues()[0];
    member = {
      name: String(row[0] || '').trim(),
      email: String(row[1] || '').trim(),
      active: String(row[2] || 'Yes').trim(),
    };

    if (!member.name && !member.email) {
      alertUser_('The selected member row is blank.');
      return;
    }
  }

  showDialog_('MemberDialog', mode === 'add' ? 'Add Member' : 'Update Member', {
    mode,
    rowNumber,
    member,
    activeOptions: CONFIG.activeOptions,
  });
}

function submitMemberForm(form) {
  const { assigneesSheet, ticketsSheet } = ensureTrackerReady_();
  const mode = normalizeMode_(form.mode, ['add', 'update']);
  const name = requireString_(form.name, 'Member name');
  const email = requireEmail_(form.email, 'Member email');
  const active = requireChoice_(form.active || 'Yes', CONFIG.activeOptions, 'Active');
  const rowNumber = mode === 'update' ? Number(form.rowNumber) : null;
  const duplicate = getAllMembers_(assigneesSheet).find(
    (member) => normalizeLookupValue_(member.email) === normalizeLookupValue_(email) && member.rowNumber !== rowNumber
  );

  if (duplicate) {
    throw new Error('That email already exists in the Members tab.');
  }

  if (mode === 'update') {
    if (!Number.isInteger(rowNumber) || rowNumber <= 1 || rowNumber > assigneesSheet.getLastRow()) {
      throw new Error('Select a valid member row before updating.');
    }

    assigneesSheet.getRange(rowNumber, 1, 1, MEMBER_HEADERS.length).setValues([[name, email, active]]);
    applyValidations_(ticketsSheet, assigneesSheet);
    return { message: `Member ${name} was updated.` };
  }

  const newRowNumber = assigneesSheet.getLastRow() + 1;
  assigneesSheet.getRange(newRowNumber, 1, 1, MEMBER_HEADERS.length).setValues([[name, email, active]]);
  applyValidations_(ticketsSheet, assigneesSheet);
  assigneesSheet.setActiveRange(assigneesSheet.getRange(newRowNumber, 1));

  return { message: `Member ${name} was added.` };
}

function showSettingDialog_() {
  const { settingsSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();
  let selectedSetting = '';

  if (activeRange && activeRange.getSheet().getName() === CONFIG.sheets.settings && activeRange.getRow() > 1) {
    selectedSetting = String(settingsSheet.getRange(activeRange.getRow(), 1).getValue() || '').trim();
  }

  showDialog_('SettingDialog', 'Update Setting', {
    selectedSetting,
    settings: getSettingsForDialog_(settingsSheet),
    defaultSettingNames: DEFAULT_SETTINGS.map((row) => row[0]),
  });
}

function submitSettingForm(form) {
  const { settingsSheet } = ensureTrackerReady_();
  const selectedSettingName = String(form.settingName || '').trim();
  const settingName =
    selectedSettingName === '__custom__'
      ? requireString_(form.customSettingName, 'Custom setting name')
      : requireString_(selectedSettingName, 'Setting name');
  const value = String(form.value || '').trim();

  setSettingValue_(settingsSheet, settingName, value);

  return { message: `Setting "${settingName}" was updated.` };
}

function normalizeTicketRow_(sheet, rowNumber, editorEmail) {
  const row = getTicketRow_(sheet, rowNumber);
  if (isBlankTicketRow_(row)) {
    return;
  }

  const now = new Date();
  const createdBy = String(row[COL.CREATED_BY - 1] || '').trim();
  const createdByEmail = String(row[COL.CREATED_BY_EMAIL - 1] || '').trim();
  const shouldFillCreatorName = !createdBy && (createdByEmail || editorEmail);
  const shouldFillCreatorEmail = !createdBy && !createdByEmail && editorEmail;

  if (!row[COL.TASK_ID - 1]) {
    sheet.getRange(rowNumber, COL.TASK_ID).setValue(nextTaskId_());
  }

  if (!row[COL.CREATED_AT - 1]) {
    sheet.getRange(rowNumber, COL.CREATED_AT).setValue(now);
  }

  if (shouldFillCreatorName || shouldFillCreatorEmail) {
    const assigneesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.members);
    const currentUser = getCurrentUserProfile_(assigneesSheet, createdByEmail || editorEmail);

    if (shouldFillCreatorName && currentUser.name) {
      sheet.getRange(rowNumber, COL.CREATED_BY).setValue(currentUser.name);
    }

    if (shouldFillCreatorEmail && currentUser.email) {
      sheet.getRange(rowNumber, COL.CREATED_BY_EMAIL).setValue(currentUser.email);
    }
  }

  sheet.getRange(rowNumber, COL.UPDATED_AT).setValue(now);

  if (!row[COL.PRIORITY - 1]) {
    sheet.getRange(rowNumber, COL.PRIORITY).setValue(getSetting_('Default Priority', CONFIG.defaultPriority));
  }

  if (!row[COL.STATUS - 1]) {
    sheet.getRange(rowNumber, COL.STATUS).setValue(getSetting_('Default Status', CONFIG.defaultStatus));
  }
}

function maybeSendAssignmentEmail_(sheet, rowNumber) {
  if (!isSettingEnabled_('Enable Assignment Emails', true)) {
    return false;
  }

  const row = getTicketRow_(sheet, rowNumber);
  const ticket = mapTicketRow_(row);

  if (!ticket.title) {
    return false;
  }

  if (isClosedStatus_(ticket.status)) {
    return false;
  }

  if (!ticket.assigneeEmail) {
    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFICATION_STATUS).setValue('Missing Email');
    return false;
  }

  if (!isValidEmail_(ticket.assigneeEmail)) {
    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFICATION_STATUS).setValue('Invalid Email');
    return false;
  }

  if (
    String(ticket.assignmentNotificationStatus).toLowerCase() === 'sent' &&
    String(ticket.assignmentNotifiedEmail).toLowerCase() === String(ticket.assigneeEmail).toLowerCase()
  ) {
    return false;
  }

  const email = buildAssignmentEmail_(ticket);

  try {
    MailApp.sendEmail({
      to: ticket.assigneeEmail,
      subject: email.subject,
      body: email.body,
      htmlBody: email.htmlBody,
    });

    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFICATION_STATUS).setValue('Sent');
    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFIED_EMAIL).setValue(ticket.assigneeEmail);
    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFIED_AT).setValue(new Date());
    return true;
  } catch (error) {
    sheet.getRange(rowNumber, COL.ASSIGNMENT_NOTIFICATION_STATUS).setValue(`Error: ${error.message}`);
    console.error(error);
    return false;
  }
}

function buildAssignmentEmail_(ticket) {
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const trackerName = getSetting_('Tracker Name', CONFIG.trackerName);
  const prefix = getSetting_('Notification Prefix', CONFIG.notificationPrefix);
  const subject = `${prefix} New assignment: ${ticket.taskId || ticket.title}`;
  const dueDate = formatDateForEmail_(ticket.dueDate);
  const assigneeName = ticket.assignee || 'there';

  const body = [
    `Hi ${assigneeName},`,
    '',
    `You have been assigned a new ${trackerName} task.`,
    '',
    `Task ID: ${ticket.taskId || 'Pending ID'}`,
    `Title: ${ticket.title}`,
    `Priority: ${ticket.priority || CONFIG.defaultPriority}`,
    `Status: ${ticket.status || CONFIG.defaultStatus}`,
    `Created By: ${ticket.createdBy || 'Not set'}`,
    `Due Date: ${dueDate}`,
    '',
    ticket.description ? `Description: ${ticket.description}` : '',
    '',
    `Open tracker: ${sheetUrl}`,
  ]
    .filter((line) => line !== '')
    .join('\n');

  const htmlBody = `
    <p>Hi ${escapeHtml_(assigneeName)},</p>
    <p>You have been assigned a new <strong>${escapeHtml_(trackerName)}</strong> task.</p>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#d9d9d9;">
      <tr><td><strong>Task ID</strong></td><td>${escapeHtml_(ticket.taskId || 'Pending ID')}</td></tr>
      <tr><td><strong>Title</strong></td><td>${escapeHtml_(ticket.title)}</td></tr>
      <tr><td><strong>Priority</strong></td><td>${escapeHtml_(ticket.priority || CONFIG.defaultPriority)}</td></tr>
      <tr><td><strong>Status</strong></td><td>${escapeHtml_(ticket.status || CONFIG.defaultStatus)}</td></tr>
      <tr><td><strong>Created By</strong></td><td>${escapeHtml_(ticket.createdBy || 'Not set')}</td></tr>
      <tr><td><strong>Due Date</strong></td><td>${escapeHtml_(dueDate)}</td></tr>
      ${
        ticket.description
          ? `<tr><td><strong>Description</strong></td><td>${escapeHtml_(ticket.description)}</td></tr>`
          : ''
      }
    </table>
    <p><a href="${escapeHtml_(sheetUrl)}">Open the tracker</a></p>
  `;

  return { subject, body, htmlBody };
}

function maybeSendCreatorUpdateEmail_(sheet, rowNumber, previousTicket, currentTicket, editorEmail) {
  if (!previousTicket || !isSettingEnabled_('Enable Creator Update Emails', true)) {
    return false;
  }

  const changes = getTicketChanges_(previousTicket, snapshotTicket_(currentTicket));
  if (changes.length === 0) {
    return false;
  }

  if (!currentTicket.createdByEmail) {
    sheet.getRange(rowNumber, COL.UPDATE_NOTIFICATION_STATUS).setValue('Missing Creator Email');
    return false;
  }

  if (!isValidEmail_(currentTicket.createdByEmail)) {
    sheet.getRange(rowNumber, COL.UPDATE_NOTIFICATION_STATUS).setValue('Invalid Creator Email');
    return false;
  }

  const email = buildCreatorUpdateEmail_(currentTicket, changes, editorEmail);

  try {
    MailApp.sendEmail({
      to: currentTicket.createdByEmail,
      subject: email.subject,
      body: email.body,
      htmlBody: email.htmlBody,
    });

    sheet.getRange(rowNumber, COL.UPDATE_NOTIFICATION_STATUS).setValue('Sent');
    sheet.getRange(rowNumber, COL.UPDATE_NOTIFIED_AT).setValue(new Date());
    return true;
  } catch (error) {
    sheet.getRange(rowNumber, COL.UPDATE_NOTIFICATION_STATUS).setValue(`Error: ${error.message}`);
    console.error(error);
    return false;
  }
}

function buildCreatorUpdateEmail_(ticket, changes, editorEmail) {
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const trackerName = getSetting_('Tracker Name', CONFIG.trackerName);
  const prefix = getSetting_('Notification Prefix', CONFIG.notificationPrefix);
  const subject = `${prefix} Task updated: ${ticket.taskId || ticket.title}`;
  const dueDate = formatDateForEmail_(ticket.dueDate);
  const creatorName = ticket.createdBy || 'there';
  const updatedBy = editorEmail || 'A team member';
  const changeLines = changes.map((change) => `${change.label}: ${change.before || 'blank'} -> ${change.after || 'blank'}`);

  const body = [
    `Hi ${creatorName},`,
    '',
    `A ${trackerName} task you created was updated.`,
    '',
    `Task ID: ${ticket.taskId || 'Pending ID'}`,
    `Title: ${ticket.title}`,
    `Updated By: ${updatedBy}`,
    `Priority: ${ticket.priority || CONFIG.defaultPriority}`,
    `Status: ${ticket.status || CONFIG.defaultStatus}`,
    `Assignee: ${ticket.assignee || 'Not assigned'}`,
    `Due Date: ${dueDate}`,
    '',
    'Changes:',
    ...changeLines,
    '',
    `Open tracker: ${sheetUrl}`,
  ].join('\n');

  const htmlBody = `
    <p>Hi ${escapeHtml_(creatorName)},</p>
    <p>A <strong>${escapeHtml_(trackerName)}</strong> task you created was updated.</p>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#d9d9d9;">
      <tr><td><strong>Task ID</strong></td><td>${escapeHtml_(ticket.taskId || 'Pending ID')}</td></tr>
      <tr><td><strong>Title</strong></td><td>${escapeHtml_(ticket.title)}</td></tr>
      <tr><td><strong>Updated By</strong></td><td>${escapeHtml_(updatedBy)}</td></tr>
      <tr><td><strong>Priority</strong></td><td>${escapeHtml_(ticket.priority || CONFIG.defaultPriority)}</td></tr>
      <tr><td><strong>Status</strong></td><td>${escapeHtml_(ticket.status || CONFIG.defaultStatus)}</td></tr>
      <tr><td><strong>Assignee</strong></td><td>${escapeHtml_(ticket.assignee || 'Not assigned')}</td></tr>
      <tr><td><strong>Due Date</strong></td><td>${escapeHtml_(dueDate)}</td></tr>
    </table>
    <p><strong>Changes</strong></p>
    <ul>${changes
      .map(
        (change) =>
          `<li><strong>${escapeHtml_(change.label)}</strong>: ${escapeHtml_(change.before || 'blank')} to ${escapeHtml_(
            change.after || 'blank'
          )}</li>`
      )
      .join('')}</ul>
    <p><a href="${escapeHtml_(sheetUrl)}">Open the tracker</a></p>
  `;

  return { subject, body, htmlBody };
}

function maybeHydrateMemberEmails_(ticketsSheet, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assigneesSheet = ss.getSheetByName(CONFIG.sheets.members);
  if (!assigneesSheet) {
    return;
  }

  const row = getTicketRow_(ticketsSheet, rowNumber);
  const createdBy = String(row[COL.CREATED_BY - 1] || '').trim();
  const createdByEmail = String(row[COL.CREATED_BY_EMAIL - 1] || '').trim();
  const assignee = String(row[COL.ASSIGNEE - 1] || '').trim();
  const assigneeEmail = String(row[COL.ASSIGNEE_EMAIL - 1] || '').trim();

  if (createdBy) {
    const creatorMember = findMemberByName_(assigneesSheet, createdBy);
    if (
      creatorMember &&
      creatorMember.email &&
      normalizeLookupValue_(createdByEmail) !== normalizeLookupValue_(creatorMember.email)
    ) {
      ticketsSheet.getRange(rowNumber, COL.CREATED_BY_EMAIL).setValue(creatorMember.email);
    }
  }

  if (assignee) {
    const assigneeMember = findMemberByName_(assigneesSheet, assignee);
    if (
      assigneeMember &&
      assigneeMember.email &&
      normalizeLookupValue_(assigneeEmail) !== normalizeLookupValue_(assigneeMember.email)
    ) {
      ticketsSheet.getRange(rowNumber, COL.ASSIGNEE_EMAIL).setValue(assigneeMember.email);
    }
  }
}

function getTicketRow_(sheet, rowNumber) {
  return sheet.getRange(rowNumber, 1, 1, TASK_HEADERS.length).getValues()[0];
}

function mapTicketRow_(row) {
  return {
    taskId: row[COL.TASK_ID - 1],
    createdAt: row[COL.CREATED_AT - 1],
    updatedAt: row[COL.UPDATED_AT - 1],
    createdBy: String(row[COL.CREATED_BY - 1] || '').trim(),
    createdByEmail: String(row[COL.CREATED_BY_EMAIL - 1] || '').trim(),
    title: String(row[COL.TITLE - 1] || '').trim(),
    description: String(row[COL.DESCRIPTION - 1] || '').trim(),
    priority: String(row[COL.PRIORITY - 1] || '').trim(),
    status: String(row[COL.STATUS - 1] || '').trim(),
    assignee: String(row[COL.ASSIGNEE - 1] || '').trim(),
    assigneeEmail: String(row[COL.ASSIGNEE_EMAIL - 1] || '').trim(),
    dueDate: row[COL.DUE_DATE - 1],
    assignmentNotificationStatus: String(row[COL.ASSIGNMENT_NOTIFICATION_STATUS - 1] || '').trim(),
    assignmentNotifiedEmail: String(row[COL.ASSIGNMENT_NOTIFIED_EMAIL - 1] || '').trim(),
    assignmentNotifiedAt: row[COL.ASSIGNMENT_NOTIFIED_AT - 1],
    updateNotificationStatus: String(row[COL.UPDATE_NOTIFICATION_STATUS - 1] || '').trim(),
    updateNotifiedAt: row[COL.UPDATE_NOTIFIED_AT - 1],
    notes: String(row[COL.NOTES - 1] || '').trim(),
  };
}

function isBlankTicketRow_(row) {
  const importantValues = [
    row[COL.CREATED_BY - 1],
    row[COL.CREATED_BY_EMAIL - 1],
    row[COL.TITLE - 1],
    row[COL.DESCRIPTION - 1],
    row[COL.ASSIGNEE - 1],
    row[COL.ASSIGNEE_EMAIL - 1],
    row[COL.DUE_DATE - 1],
    row[COL.NOTES - 1],
  ];

  return importantValues.every((value) => String(value || '').trim() === '');
}

function isClosedStatus_(status) {
  return ['completed', 'cancelled'].includes(String(status || '').toLowerCase());
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function emptyTicketRow_() {
  return Array(TASK_HEADERS.length).fill('');
}

function showDialog_(templateName, title, data) {
  const template = HtmlService.createTemplateFromFile(templateName);
  template.data = JSON.stringify(data);

  const html = template
    .evaluate()
    .setWidth(760)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function defaultTaskDialogTicket_(currentUser) {
  const creator = currentUser || { name: '', email: '' };

  return {
    taskId: '',
    createdBy: creator.name || '',
    createdByEmail: creator.email || '',
    title: '',
    description: '',
    priority: getSetting_('Default Priority', CONFIG.defaultPriority),
    status: getSetting_('Default Status', CONFIG.defaultStatus),
    assignee: '',
    assigneeEmail: '',
    dueDate: '',
    notes: '',
  };
}

function serializeTicketForDialog_(ticket) {
  return {
    taskId: String(ticket.taskId || ticket.ticketId || '').trim(),
    createdBy: String(ticket.createdBy || '').trim(),
    createdByEmail: String(ticket.createdByEmail || '').trim(),
    title: String(ticket.title || '').trim(),
    description: String(ticket.description || '').trim(),
    priority: String(ticket.priority || '').trim() || getSetting_('Default Priority', CONFIG.defaultPriority),
    status: String(ticket.status || '').trim() || getSetting_('Default Status', CONFIG.defaultStatus),
    assignee: String(ticket.assignee || '').trim(),
    assigneeEmail: String(ticket.assigneeEmail || '').trim(),
    dueDate: formatDateForInput_(ticket.dueDate),
    notes: String(ticket.notes || '').trim(),
  };
}

function normalizeTaskForm_(form, assigneesSheet, currentUser) {
  const creator = currentUser || { name: '', email: '' };
  const formCreatedBy = String(form.createdBy || '').trim();
  const formCreatedByEmail = String(form.createdByEmail || '').trim();
  const useCurrentUser = !formCreatedBy && !formCreatedByEmail;
  const createdBy = hydrateMemberFromForm_(
    assigneesSheet,
    useCurrentUser ? creator.name : formCreatedBy,
    useCurrentUser ? creator.email : formCreatedByEmail,
    {
      nameLabel: 'Created by',
      emailLabel: 'Created by email',
      requireName: true,
      requireEmail: true,
    }
  );
  const assignee = hydrateMemberFromForm_(assigneesSheet, form.assignee, form.assigneeEmail, {
    nameLabel: 'Assignee',
    emailLabel: 'Assignee email',
    requireName: false,
    requireEmail: true,
  });
  const dueDateText = String(form.dueDate || '').trim();

  return {
    createdBy: createdBy.name,
    createdByEmail: createdBy.email,
    title: requireString_(form.title, 'Task title'),
    description: String(form.description || '').trim(),
    priority: requireChoice_(form.priority, CONFIG.priorities, 'Priority'),
    status: requireChoice_(form.status, CONFIG.statuses, 'Status'),
    assignee: assignee.name,
    assigneeEmail: assignee.email,
    dueDate: dueDateText ? requireDate_(dueDateText, 'Due date') : '',
    notes: String(form.notes || '').trim(),
  };
}

function hydrateMemberFromForm_(assigneesSheet, nameValue, emailValue, options) {
  let name = String(nameValue || '').trim();
  let email = String(emailValue || '').trim();
  const memberByName = name ? findMemberByName_(assigneesSheet, name) : null;

  if (memberByName) {
    name = memberByName.name;
    if (!email || !isValidEmail_(email)) {
      email = memberByName.email;
    }
  }

  const memberByEmail = email ? findMemberByEmail_(assigneesSheet, email) : null;
  if (memberByEmail && !name) {
    name = memberByEmail.name;
  }

  if (options.requireName) {
    name = requireString_(name, options.nameLabel);
  }

  if (options.requireEmail) {
    email = requireEmail_(email, options.emailLabel);
  }

  return { name, email };
}

function normalizeMode_(mode, allowedModes) {
  const normalizedMode = String(mode || '').trim().toLowerCase();
  if (!allowedModes.includes(normalizedMode)) {
    throw new Error('Invalid form mode.');
  }

  return normalizedMode;
}

function requireString_(value, label) {
  const text = String(value || '').trim();
  if (!text) {
    throw new Error(`${label} is required.`);
  }

  return text;
}

function requireEmail_(value, label) {
  const email = requireString_(value, label);
  if (!isValidEmail_(email)) {
    throw new Error(`${label} must be a valid email address.`);
  }

  return email;
}

function requireChoice_(value, options, label) {
  const text = requireString_(value, label);
  const match = options.find((option) => normalizeLookupValue_(option) === normalizeLookupValue_(text));
  if (!match) {
    throw new Error(`${label} must be one of: ${options.join(', ')}.`);
  }

  return match;
}

function requireDate_(value, label) {
  const date = parseDateInput_(value);
  if (!date) {
    throw new Error(`${label} must use YYYY-MM-DD format.`);
  }

  return date;
}

function getSettingsForDialog_(settingsSheet) {
  const settings = getSettingsMap_(settingsSheet);
  const settingNames = new Set(DEFAULT_SETTINGS.map((row) => row[0]));
  Object.keys(settings).forEach((settingName) => settingNames.add(settingName));

  return Array.from(settingNames).map((settingName) => ({
    name: settingName,
    value: String(getSetting_(settingName, '') || ''),
  }));
}

function parseDateInput_(value) {
  const input = String(value || '').trim();
  if (!input) {
    return null;
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(input)) {
    return null;
  }

  const parts = input.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (Number.isNaN(date.getTime())) {
    return null;
  }

  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    return null;
  }

  return date;
}

function findMemberByName_(assigneesSheet, name) {
  const normalizedName = normalizeLookupValue_(name);
  return getMembers_(assigneesSheet).find((member) => normalizeLookupValue_(member.name) === normalizedName) || null;
}

function findMemberByEmail_(assigneesSheet, email) {
  const normalizedEmail = normalizeLookupValue_(email);
  return getMembers_(assigneesSheet).find((member) => normalizeLookupValue_(member.email) === normalizedEmail) || null;
}

function findAnyMemberByEmail_(assigneesSheet, email) {
  const normalizedEmail = normalizeLookupValue_(email);
  return getAllMembers_(assigneesSheet).find((member) => normalizeLookupValue_(member.email) === normalizedEmail) || null;
}

function getMembers_(assigneesSheet) {
  return getAllMembers_(assigneesSheet)
    .filter((member) => member.name && normalizeLookupValue_(member.active) !== 'no');
}

function getAllMembers_(assigneesSheet) {
  const lastRow = assigneesSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return assigneesSheet
    .getRange(2, 1, lastRow - 1, MEMBER_HEADERS.length)
    .getValues()
    .map((row, index) => ({
      rowNumber: index + 2,
      name: String(row[0] || '').trim(),
      email: String(row[1] || '').trim(),
      active: String(row[2] || 'Yes').trim(),
    }));
}

function normalizeLookupValue_(value) {
  return String(value || '').trim().toLowerCase();
}

function getStoredTicketSnapshotForRow_(sheet, rowNumber) {
  const ticket = mapTicketRow_(getTicketRow_(sheet, rowNumber));
  if (!ticket.taskId) {
    return null;
  }

  return getStoredTicketSnapshot_(ticket.taskId);
}

function getStoredTicketSnapshot_(taskId) {
  const props = PropertiesService.getDocumentProperties();
  const rawValue = props.getProperty(ticketSnapshotKey_(taskId)) || props.getProperty(`TICKET_SNAPSHOT_${taskId}`);
  if (!rawValue) {
    return null;
  }

  try {
    return JSON.parse(rawValue);
  } catch (error) {
    console.error(error);
    return null;
  }
}

function saveTicketSnapshot_(ticket) {
  if (!ticket || !ticket.taskId) {
    return;
  }

  PropertiesService.getDocumentProperties().setProperty(ticketSnapshotKey_(ticket.taskId), JSON.stringify(snapshotTicket_(ticket)));
}

function syncTicketSnapshots_(ticketsSheet) {
  const lastRow = ticketsSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  for (let row = 2; row <= lastRow; row += 1) {
    const ticket = mapTicketRow_(getTicketRow_(ticketsSheet, row));
    if (ticket.taskId && !isBlankTicketRow_(getTicketRow_(ticketsSheet, row))) {
      saveTicketSnapshot_(ticket);
    }
  }
}

function ticketSnapshotKey_(taskId) {
  return `TASK_SNAPSHOT_${taskId}`;
}

function snapshotTicket_(ticket) {
  return {
    taskId: String(ticket.taskId || ticket.ticketId || '').trim(),
    createdBy: String(ticket.createdBy || '').trim(),
    createdByEmail: String(ticket.createdByEmail || '').trim(),
    title: String(ticket.title || '').trim(),
    description: String(ticket.description || '').trim(),
    priority: String(ticket.priority || '').trim(),
    status: String(ticket.status || '').trim(),
    assignee: String(ticket.assignee || '').trim(),
    assigneeEmail: String(ticket.assigneeEmail || '').trim(),
    dueDate: formatDateForSnapshot_(ticket.dueDate),
    notes: String(ticket.notes || '').trim(),
  };
}

function getTicketChanges_(previousTicket, currentTicket) {
  return UPDATE_NOTIFICATION_FIELDS.reduce((changes, field) => {
    const key = field[0];
    const label = field[1];
    const before = String(previousTicket[key] || '').trim();
    const after = String(currentTicket[key] || '').trim();

    if (before !== after) {
      changes.push({ label, before, after });
    }

    return changes;
  }, []);
}

function formatDateForSnapshot_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function getEditActorEmail_(e) {
  try {
    if (e && e.user && typeof e.user.getEmail === 'function') {
      return e.user.getEmail() || getActiveUserEmail_();
    }
  } catch (error) {
    console.log('Editor email unavailable.');
  }

  return getActiveUserEmail_();
}

function getActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim();
  } catch (error) {
    return '';
  }
}

function getCurrentUserProfile_(assigneesSheet, emailOverride) {
  const email = String(emailOverride || getActiveUserEmail_() || '').trim();
  if (!email) {
    return { name: '', email: '' };
  }

  const member = assigneesSheet ? findAnyMemberByEmail_(assigneesSheet, email) : null;
  return {
    name: member && member.name ? member.name : inferNameFromEmail_(email),
    email,
  };
}

function inferNameFromEmail_(email) {
  const localPart = String(email || '').split('@')[0];
  const name = localPart
    .split(/[._-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(' ');

  return name || email;
}

function isSettingEnabled_(key, fallback) {
  const value = String(getSetting_(key, fallback ? 'Yes' : 'No') || '').trim().toLowerCase();
  return ['yes', 'true', 'enabled', 'on', '1'].includes(value);
}

function nextTaskId_() {
  const props = PropertiesService.getDocumentProperties();
  const currentValue = Number(props.getProperty('NEXT_TASK_NUMBER') || props.getProperty('NEXT_TICKET_NUMBER') || '1');
  const nextValue = Number.isFinite(currentValue) && currentValue > 0 ? currentValue : 1;
  props.setProperty('NEXT_TASK_NUMBER', String(nextValue + 1));
  props.setProperty('NEXT_TICKET_NUMBER', String(nextValue + 1));
  return `${CONFIG.taskPrefix}-${String(nextValue).padStart(4, '0')}`;
}

function syncTicketCounter_(ticketsSheet) {
  const lastRow = ticketsSheet.getLastRow();
  if (lastRow <= 1) {
    PropertiesService.getDocumentProperties().setProperty('NEXT_TASK_NUMBER', '1');
    PropertiesService.getDocumentProperties().setProperty('NEXT_TICKET_NUMBER', '1');
    return;
  }

  const ticketIds = ticketsSheet.getRange(2, COL.TASK_ID, lastRow - 1, 1).getValues().flat();
  const highestNumber = ticketIds.reduce((highest, ticketId) => {
    const match = String(ticketId || '').match(new RegExp(`^${CONFIG.taskPrefix}-(\\d+)$`));
    return match ? Math.max(highest, Number(match[1])) : highest;
  }, 0);

  PropertiesService.getDocumentProperties().setProperty('NEXT_TASK_NUMBER', String(highestNumber + 1));
  PropertiesService.getDocumentProperties().setProperty('NEXT_TICKET_NUMBER', String(highestNumber + 1));
}

function setHeaderRow_(sheet, headers, aliases) {
  const headerAliases = aliases || {};
  const lastRow = sheet.getLastRow();
  const existingColumnCount = Math.max(sheet.getLastColumn(), headers.length);
  const currentHeaders = sheet
    .getRange(1, 1, 1, existingColumnCount)
    .getValues()[0]
    .map((header) => String(header || '').trim());
  const headersAlreadyCurrent = headers.every((header, index) => currentHeaders[index] === header);

  if (headersAlreadyCurrent) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const recognizedHeaders = currentHeaders.filter((header) => headers.includes(header) || headerAliases[header]);

  if (recognizedHeaders.length > 0 && lastRow > 1) {
    const sourceValues = sheet.getRange(2, 1, lastRow - 1, existingColumnCount).getValues();
    const sourceIndexByHeader = {};

    currentHeaders.forEach((header, index) => {
      const targetHeader = headerAliases[header] || header;
      if (headers.includes(targetHeader) && sourceIndexByHeader[targetHeader] === undefined) {
        sourceIndexByHeader[targetHeader] = index;
      }
    });

    const migratedValues = sourceValues.map((sourceRow) =>
      headers.map((header) =>
        sourceIndexByHeader[header] === undefined ? '' : sourceRow[sourceIndexByHeader[header]]
      )
    );

    sheet.getRange(1, 1, lastRow, existingColumnCount).clearContent();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, migratedValues.length, headers.length).setValues(migratedValues);
  } else {
    sheet.getRange(1, 1, 1, existingColumnCount).clearContent();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
}

function seedSettings_(settingsSheet) {
  const existing = getSettingsMap_(settingsSheet);
  const rowsToAdd = DEFAULT_SETTINGS.filter(([setting]) => !existing[setting]);

  if (rowsToAdd.length > 0) {
    settingsSheet
      .getRange(settingsSheet.getLastRow() + 1, 1, rowsToAdd.length, SETTINGS_HEADERS.length)
      .setValues(rowsToAdd);
  }
}

function applyValidations_(ticketsSheet, assigneesSheet) {
  const maxRows = ticketsSheet.getMaxRows();
  if (maxRows <= 1) {
    return;
  }

  ticketsSheet.getRange(2, 1, maxRows - 1, TASK_HEADERS.length).clearDataValidations();

  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.priorities, true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.statuses, true)
    .setAllowInvalid(false)
    .build();

  const assigneeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(assigneesSheet.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();

  const dueDateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();

  ticketsSheet.getRange(2, COL.PRIORITY, maxRows - 1, 1).setDataValidation(priorityRule);
  ticketsSheet.getRange(2, COL.STATUS, maxRows - 1, 1).setDataValidation(statusRule);
  ticketsSheet.getRange(2, COL.CREATED_BY, maxRows - 1, 1).setDataValidation(assigneeRule);
  ticketsSheet.getRange(2, COL.ASSIGNEE, maxRows - 1, 1).setDataValidation(assigneeRule);
  ticketsSheet.getRange(2, COL.DUE_DATE, maxRows - 1, 1).setDataValidation(dueDateRule);

  const activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.activeOptions, true)
    .setAllowInvalid(false)
    .build();

  assigneesSheet.getRange(2, 3, Math.max(assigneesSheet.getMaxRows() - 1, 1), 1).setDataValidation(activeRule);
}

function formatTicketsSheet_(sheet) {
  sheet.getRange(1, 1, 1, TASK_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(2, COL.CREATED_AT, Math.max(sheet.getMaxRows() - 1, 1), 2).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, COL.DUE_DATE, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet
    .getRange(2, COL.ASSIGNMENT_NOTIFIED_AT, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');
  sheet
    .getRange(2, COL.UPDATE_NOTIFIED_AT, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');

  const widths = [110, 145, 145, 160, 220, 220, 320, 110, 130, 160, 220, 120, 300, 210, 220, 145, 190, 145];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  ensureFilter_(sheet);
}

function formatAssigneesSheet_(sheet) {
  sheet.getRange(1, 1, 1, MEMBER_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidths(1, 1, 180);
  sheet.setColumnWidths(2, 1, 240);
  sheet.setColumnWidths(3, 1, 90);
  ensureFilter_(sheet);
}

function formatSettingsSheet_(sheet) {
  sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidths(1, 1, 200);
  sheet.setColumnWidths(2, 1, 260);
}

function ensureFilter_(sheet) {
  if (!sheet.getFilter()) {
    sheet.getDataRange().createFilter();
  }
}

function getOrCreateSheet_(ss, sheetName, legacySheetNames) {
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    return existingSheet;
  }

  const legacyNames = legacySheetNames || [];
  for (let index = 0; index < legacyNames.length; index += 1) {
    const legacySheet = ss.getSheetByName(legacyNames[index]);
    if (legacySheet) {
      legacySheet.setName(sheetName);
      return legacySheet;
    }
  }

  return ss.insertSheet(sheetName);
}

function getSetting_(key, fallback) {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.settings);
  if (!settingsSheet) {
    return fallback;
  }

  const settings = getSettingsMap_(settingsSheet);
  return settings[key] || fallback;
}

function getSettingsMap_(settingsSheet) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }

  return settingsSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .reduce((settings, row) => {
      const key = String(row[0] || '').trim();
      if (key) {
        settings[key] = row[1];
      }
      return settings;
    }, {});
}

function setSettingValue_(settingsSheet, key, value) {
  const rowNumber = findSettingRow_(settingsSheet, key);
  if (rowNumber) {
    settingsSheet.getRange(rowNumber, 2).setValue(value);
    return;
  }

  settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, 1, SETTINGS_HEADERS.length).setValues([[key, value]]);
}

function findSettingRow_(settingsSheet, key) {
  const normalizedKey = normalizeLookupValue_(key);
  const lastRow = settingsSheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }

  const values = settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const foundIndex = values.findIndex((value) => normalizeLookupValue_(value) === normalizedKey);
  return foundIndex === -1 ? null : foundIndex + 2;
}

function formatDateForEmail_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || 'Not set');
}

function formatDateForInput_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function alertUser_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    console.log(message);
  }
}
