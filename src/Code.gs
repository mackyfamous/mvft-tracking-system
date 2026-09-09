const CONFIG = {
  trackerName: 'MVFT Tracking System',
  notificationPrefix: '[MVFT]',
  sheets: {
    tickets: 'Tickets',
    assignees: 'Assignees',
    settings: 'Settings',
  },
  priorities: ['Low', 'Medium', 'High', 'Urgent'],
  statuses: ['Open', 'In Progress', 'Waiting', 'Completed', 'Cancelled'],
  activeOptions: ['Yes', 'No'],
  defaultPriority: 'Medium',
  defaultStatus: 'Open',
  ticketPrefix: 'MVFT',
  clearToken: 'CLEAR',
};

const TICKET_HEADERS = [
  'Ticket ID',
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
  'Assignment Notification Status',
  'Assignment Notified Email',
  'Assignment Notified At',
  'Update Notification Status',
  'Update Notified At',
  'Notes',
];

const ASSIGNEE_HEADERS = ['Name', 'Email', 'Active'];
const SETTINGS_HEADERS = ['Setting', 'Value'];

const COL = {
  TICKET_ID: 1,
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
  ASSIGNMENT_NOTIFICATION_STATUS: 13,
  ASSIGNMENT_NOTIFIED_EMAIL: 14,
  ASSIGNMENT_NOTIFIED_AT: 15,
  UPDATE_NOTIFICATION_STATUS: 16,
  UPDATE_NOTIFIED_AT: 17,
  NOTES: 18,
};

const TICKET_HEADER_ALIASES = {
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
  const ticketsSheet = getOrCreateSheet_(ss, CONFIG.sheets.tickets);
  const assigneesSheet = getOrCreateSheet_(ss, CONFIG.sheets.assignees);
  const settingsSheet = getOrCreateSheet_(ss, CONFIG.sheets.settings);

  setHeaderRow_(ticketsSheet, TICKET_HEADERS, TICKET_HEADER_ALIASES);
  setHeaderRow_(assigneesSheet, ASSIGNEE_HEADERS);
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
    .filter((trigger) => trigger.getHandlerFunction() === 'handleTicketEdit')
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('handleTicketEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  alertUser_('Email notification trigger installed.');
}

function handleTicketEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.sheets.tickets) {
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
    for (let row = Math.max(firstRow, 2); row <= lastRow; row += 1) {
      const previousTicket = getStoredTicketSnapshotForRow_(sheet, row);
      normalizeTicketRow_(sheet, row);
      maybeHydrateMemberEmails_(sheet, row);
      maybeSendAssignmentEmail_(sheet, row);
      const currentTicket = mapTicketRow_(getTicketRow_(sheet, row));
      maybeSendCreatorUpdateEmail_(sheet, row, previousTicket, currentTicket, getEditActorEmail_(e));
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
    alertUser_('No tickets found.');
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
  const { ticketsSheet, assigneesSheet } = ensureTrackerReady_();
  const title = 'Add Task';

  const createdBy = promptMember_(assigneesSheet, title, 'Created by', {
    requireName: true,
    requireEmail: true,
  });
  if (!createdBy) {
    return;
  }

  const taskTitle = promptRequired_(title, 'Task title');
  if (taskTitle === null) {
    return;
  }

  const description = promptOptional_(title, 'Description');
  if (description === null) {
    return;
  }

  const priority = promptChoice_(title, 'Priority', CONFIG.priorities, getSetting_('Default Priority', CONFIG.defaultPriority));
  if (priority === null) {
    return;
  }

  const status = promptChoice_(title, 'Status', CONFIG.statuses, getSetting_('Default Status', CONFIG.defaultStatus));
  if (status === null) {
    return;
  }

  const assignee = promptMember_(assigneesSheet, title, 'Assignee', {
    requireName: false,
    requireEmail: true,
  });
  if (!assignee) {
    return;
  }

  const dueDate = promptDate_(title, 'Due date in YYYY-MM-DD format. Leave blank if there is no due date.', {
    allowBlank: true,
  });
  if (dueDate === null) {
    return;
  }

  const notes = promptOptional_(title, 'Notes');
  if (notes === null) {
    return;
  }

  const now = new Date();
  const rowNumber = ticketsSheet.getLastRow() + 1;
  const ticketId = nextTicketId_();
  const row = emptyTicketRow_();

  row[COL.TICKET_ID - 1] = ticketId;
  row[COL.CREATED_AT - 1] = now;
  row[COL.UPDATED_AT - 1] = now;
  row[COL.CREATED_BY - 1] = createdBy.name;
  row[COL.CREATED_BY_EMAIL - 1] = createdBy.email;
  row[COL.TITLE - 1] = taskTitle;
  row[COL.DESCRIPTION - 1] = description;
  row[COL.PRIORITY - 1] = priority;
  row[COL.STATUS - 1] = status;
  row[COL.ASSIGNEE - 1] = assignee.name;
  row[COL.ASSIGNEE_EMAIL - 1] = assignee.email;
  row[COL.DUE_DATE - 1] = dueDate || '';
  row[COL.NOTES - 1] = notes;

  ticketsSheet.getRange(rowNumber, 1, 1, TICKET_HEADERS.length).setValues([row]);
  maybeSendAssignmentEmail_(ticketsSheet, rowNumber);
  saveTicketSnapshot_(mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber)));
  ticketsSheet.setActiveRange(ticketsSheet.getRange(rowNumber, COL.TITLE));

  alertUser_(`Task ${ticketId} was created.`);
}

function updateSelectedTask() {
  const { ticketsSheet, assigneesSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();

  if (!activeRange || activeRange.getSheet().getName() !== CONFIG.sheets.tickets || activeRange.getRow() <= 1) {
    alertUser_('Select a task row in the Tickets tab before using Update selected task.');
    return;
  }

  const rowNumber = activeRange.getRow();
  const previousTicket = mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber));
  if (isBlankTicketRow_(getTicketRow_(ticketsSheet, rowNumber))) {
    alertUser_('The selected ticket row is blank.');
    return;
  }

  const title = 'Update Selected Task';
  const row = getTicketRow_(ticketsSheet, rowNumber);

  const taskTitle = promptUpdateText_(title, 'Title', previousTicket.title);
  if (taskTitle === null) {
    return;
  }

  const description = promptUpdateText_(title, 'Description', previousTicket.description);
  if (description === null) {
    return;
  }

  const priority = promptUpdateChoice_(title, 'Priority', CONFIG.priorities, previousTicket.priority);
  if (priority === null) {
    return;
  }

  const status = promptUpdateChoice_(title, 'Status', CONFIG.statuses, previousTicket.status);
  if (status === null) {
    return;
  }

  const assignee = promptUpdateText_(title, 'Assignee', previousTicket.assignee);
  if (assignee === null) {
    return;
  }

  let assigneeEmail = promptUpdateEmail_(title, 'Assignee email', previousTicket.assigneeEmail);
  if (assigneeEmail === null) {
    return;
  }

  const dueDate = promptUpdateDate_(title, 'Due date', previousTicket.dueDate);
  if (dueDate === null) {
    return;
  }

  const notes = promptUpdateText_(title, 'Notes', previousTicket.notes);
  if (notes === null) {
    return;
  }

  row[COL.TITLE - 1] = taskTitle;
  row[COL.DESCRIPTION - 1] = description;
  row[COL.PRIORITY - 1] = priority;
  row[COL.STATUS - 1] = status;
  row[COL.ASSIGNEE - 1] = assignee;

  if (
    normalizeLookupValue_(assignee) !== normalizeLookupValue_(previousTicket.assignee) &&
    normalizeLookupValue_(assigneeEmail) === normalizeLookupValue_(previousTicket.assigneeEmail)
  ) {
    const assigneeMember = findMemberByName_(assigneesSheet, assignee);
    if (assigneeMember && assigneeMember.email) {
      assigneeEmail = assigneeMember.email;
    }
  }

  row[COL.ASSIGNEE_EMAIL - 1] = assigneeEmail;
  row[COL.DUE_DATE - 1] = dueDate || '';
  row[COL.NOTES - 1] = notes;

  ticketsSheet.getRange(rowNumber, 1, 1, TICKET_HEADERS.length).setValues([row]);
  normalizeTicketRow_(ticketsSheet, rowNumber);
  maybeHydrateMemberEmails_(ticketsSheet, rowNumber);
  maybeSendAssignmentEmail_(ticketsSheet, rowNumber);

  const currentTicket = mapTicketRow_(getTicketRow_(ticketsSheet, rowNumber));
  maybeSendCreatorUpdateEmail_(ticketsSheet, rowNumber, snapshotTicket_(previousTicket), currentTicket, getActiveUserEmail_());
  saveTicketSnapshot_(currentTicket);

  alertUser_(`Task ${currentTicket.ticketId || rowNumber} was updated.`);
}

function addMember() {
  const { assigneesSheet, ticketsSheet } = ensureTrackerReady_();
  const title = 'Add Member';

  const name = promptRequired_(title, 'Member name');
  if (name === null) {
    return;
  }

  const email = promptEmail_(title, 'Member email', { required: true });
  if (email === null) {
    return;
  }

  if (findMemberByEmail_(assigneesSheet, email)) {
    alertUser_('That email already exists in the Assignees tab. Use Update selected member instead.');
    return;
  }

  const active = promptChoice_(title, 'Active', CONFIG.activeOptions, 'Yes');
  if (active === null) {
    return;
  }

  const rowNumber = assigneesSheet.getLastRow() + 1;
  assigneesSheet.getRange(rowNumber, 1, 1, ASSIGNEE_HEADERS.length).setValues([[name, email, active]]);
  applyValidations_(ticketsSheet, assigneesSheet);
  assigneesSheet.setActiveRange(assigneesSheet.getRange(rowNumber, 1));

  alertUser_(`Member ${name} was added.`);
}

function updateSelectedMember() {
  const { assigneesSheet, ticketsSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();

  if (!activeRange || activeRange.getSheet().getName() !== CONFIG.sheets.assignees || activeRange.getRow() <= 1) {
    alertUser_('Select a member row in the Assignees tab before using Update selected member.');
    return;
  }

  const rowNumber = activeRange.getRow();
  const current = assigneesSheet.getRange(rowNumber, 1, 1, ASSIGNEE_HEADERS.length).getValues()[0];
  const title = 'Update Selected Member';

  const name = promptUpdateText_(title, 'Name', current[0]);
  if (name === null) {
    return;
  }

  const email = promptUpdateEmail_(title, 'Email', current[1]);
  if (email === null) {
    return;
  }

  const active = promptUpdateChoice_(title, 'Active', CONFIG.activeOptions, current[2] || 'Yes');
  if (active === null) {
    return;
  }

  assigneesSheet.getRange(rowNumber, 1, 1, ASSIGNEE_HEADERS.length).setValues([[name, email, active]]);
  applyValidations_(ticketsSheet, assigneesSheet);

  alertUser_(`Member ${name} was updated.`);
}

function updateSetting() {
  const { settingsSheet } = ensureTrackerReady_();
  const activeRange = SpreadsheetApp.getActiveRange();
  const title = 'Update Setting';
  let settingName = '';
  let currentValue = '';

  if (activeRange && activeRange.getSheet().getName() === CONFIG.sheets.settings && activeRange.getRow() > 1) {
    const row = settingsSheet.getRange(activeRange.getRow(), 1, 1, SETTINGS_HEADERS.length).getValues()[0];
    settingName = String(row[0] || '').trim();
    currentValue = row[1];
  }

  if (!settingName) {
    settingName = promptRequired_(title, `Setting name. Common options: ${DEFAULT_SETTINGS.map((row) => row[0]).join(', ')}`);
    if (settingName === null) {
      return;
    }
    currentValue = getSetting_(settingName, '');
  }

  const newValue = promptUpdateText_(title, `Value for ${settingName}`, currentValue);
  if (newValue === null) {
    return;
  }

  setSettingValue_(settingsSheet, settingName, newValue);
  alertUser_(`Setting "${settingName}" was updated.`);
}

function normalizeTicketRow_(sheet, rowNumber) {
  const row = getTicketRow_(sheet, rowNumber);
  if (isBlankTicketRow_(row)) {
    return;
  }

  const now = new Date();

  if (!row[COL.TICKET_ID - 1]) {
    sheet.getRange(rowNumber, COL.TICKET_ID).setValue(nextTicketId_());
  }

  if (!row[COL.CREATED_AT - 1]) {
    sheet.getRange(rowNumber, COL.CREATED_AT).setValue(now);
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
  const subject = `${prefix} New assignment: ${ticket.ticketId || ticket.title}`;
  const dueDate = formatDateForEmail_(ticket.dueDate);
  const assigneeName = ticket.assignee || 'there';

  const body = [
    `Hi ${assigneeName},`,
    '',
    `You have been assigned a new ${trackerName} ticket.`,
    '',
    `Ticket ID: ${ticket.ticketId || 'Pending ID'}`,
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
    <p>You have been assigned a new <strong>${escapeHtml_(trackerName)}</strong> ticket.</p>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#d9d9d9;">
      <tr><td><strong>Ticket ID</strong></td><td>${escapeHtml_(ticket.ticketId || 'Pending ID')}</td></tr>
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
  const subject = `${prefix} Ticket updated: ${ticket.ticketId || ticket.title}`;
  const dueDate = formatDateForEmail_(ticket.dueDate);
  const creatorName = ticket.createdBy || 'there';
  const updatedBy = editorEmail || 'A team member';
  const changeLines = changes.map((change) => `${change.label}: ${change.before || 'blank'} -> ${change.after || 'blank'}`);

  const body = [
    `Hi ${creatorName},`,
    '',
    `A ${trackerName} ticket you created was updated.`,
    '',
    `Ticket ID: ${ticket.ticketId || 'Pending ID'}`,
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
    <p>A <strong>${escapeHtml_(trackerName)}</strong> ticket you created was updated.</p>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#d9d9d9;">
      <tr><td><strong>Ticket ID</strong></td><td>${escapeHtml_(ticket.ticketId || 'Pending ID')}</td></tr>
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
  const assigneesSheet = ss.getSheetByName(CONFIG.sheets.assignees);
  if (!assigneesSheet) {
    return;
  }

  const row = getTicketRow_(ticketsSheet, rowNumber);
  const createdBy = String(row[COL.CREATED_BY - 1] || '').trim();
  const createdByEmail = String(row[COL.CREATED_BY_EMAIL - 1] || '').trim();
  const assignee = String(row[COL.ASSIGNEE - 1] || '').trim();
  const assigneeEmail = String(row[COL.ASSIGNEE_EMAIL - 1] || '').trim();

  if (createdBy && !createdByEmail) {
    const creatorMember = findMemberByName_(assigneesSheet, createdBy);
    if (creatorMember && creatorMember.email) {
      ticketsSheet.getRange(rowNumber, COL.CREATED_BY_EMAIL).setValue(creatorMember.email);
    }
  }

  if (assignee && !assigneeEmail) {
    const assigneeMember = findMemberByName_(assigneesSheet, assignee);
    if (assigneeMember && assigneeMember.email) {
      ticketsSheet.getRange(rowNumber, COL.ASSIGNEE_EMAIL).setValue(assigneeMember.email);
    }
  }
}

function getTicketRow_(sheet, rowNumber) {
  return sheet.getRange(rowNumber, 1, 1, TICKET_HEADERS.length).getValues()[0];
}

function mapTicketRow_(row) {
  return {
    ticketId: row[COL.TICKET_ID - 1],
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
  return Array(TICKET_HEADERS.length).fill('');
}

function promptRequired_(title, message) {
  const ui = SpreadsheetApp.getUi();

  while (true) {
    const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const value = String(response.getResponseText() || '').trim();
    if (value) {
      return value;
    }

    ui.alert('This field is required.');
  }
}

function promptOptional_(title, message) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(title, `${message}. Leave blank if not needed.`, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  return String(response.getResponseText() || '').trim();
}

function promptChoice_(title, fieldName, options, defaultValue) {
  const ui = SpreadsheetApp.getUi();
  const optionsText = options.join(', ');

  while (true) {
    const response = ui.prompt(
      title,
      `${fieldName}. Options: ${optionsText}. Leave blank for ${defaultValue}.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const value = String(response.getResponseText() || '').trim() || defaultValue;
    const match = options.find((option) => option.toLowerCase() === value.toLowerCase());
    if (match) {
      return match;
    }

    ui.alert(`Please enter one of: ${optionsText}.`);
  }
}

function promptEmail_(title, fieldName, options) {
  const ui = SpreadsheetApp.getUi();
  const required = Boolean(options && options.required);

  while (true) {
    const response = ui.prompt(
      title,
      `${fieldName}${required ? '' : '. Leave blank if not needed.'}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const value = String(response.getResponseText() || '').trim();
    if (!value && !required) {
      return '';
    }

    if (isValidEmail_(value)) {
      return value;
    }

    ui.alert('Enter a valid email address.');
  }
}

function promptDate_(title, message, options) {
  const ui = SpreadsheetApp.getUi();
  const allowBlank = Boolean(options && options.allowBlank);

  while (true) {
    const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const value = String(response.getResponseText() || '').trim();
    if (!value && allowBlank) {
      return '';
    }

    const date = parseDateInput_(value);
    if (date) {
      return date;
    }

    ui.alert('Enter a valid date in YYYY-MM-DD format.');
  }
}

function promptUpdateText_(title, fieldName, currentValue) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    title,
    `${fieldName}. Current: ${formatPromptValue_(currentValue)}. Leave blank to keep it, or type ${CONFIG.clearToken} to clear it.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  return resolveUpdateValue_(response.getResponseText(), currentValue);
}

function promptUpdateEmail_(title, fieldName, currentValue) {
  const ui = SpreadsheetApp.getUi();

  while (true) {
    const response = ui.prompt(
      title,
      `${fieldName}. Current: ${formatPromptValue_(currentValue)}. Leave blank to keep it, or type ${CONFIG.clearToken} to clear it.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const value = resolveUpdateValue_(response.getResponseText(), currentValue);
    if (!value || isValidEmail_(value)) {
      return value;
    }

    ui.alert('Enter a valid email address.');
  }
}

function promptUpdateChoice_(title, fieldName, options, currentValue) {
  const ui = SpreadsheetApp.getUi();
  const optionsText = options.join(', ');

  while (true) {
    const response = ui.prompt(
      title,
      `${fieldName}. Current: ${formatPromptValue_(currentValue)}. Options: ${optionsText}. Leave blank to keep it.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const rawValue = String(response.getResponseText() || '').trim();
    if (!rawValue) {
      return currentValue || '';
    }

    const match = options.find((option) => option.toLowerCase() === rawValue.toLowerCase());
    if (match) {
      return match;
    }

    ui.alert(`Please enter one of: ${optionsText}.`);
  }
}

function promptUpdateDate_(title, fieldName, currentValue) {
  const ui = SpreadsheetApp.getUi();

  while (true) {
    const response = ui.prompt(
      title,
      `${fieldName}. Current: ${formatPromptValue_(formatDateForEmail_(currentValue))}. Leave blank to keep it, or type ${CONFIG.clearToken} to clear it.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const rawValue = String(response.getResponseText() || '').trim();
    if (!rawValue) {
      return currentValue || '';
    }

    if (rawValue.toUpperCase() === CONFIG.clearToken) {
      return '';
    }

    const date = parseDateInput_(rawValue);
    if (date) {
      return date;
    }

    ui.alert('Enter a valid date in YYYY-MM-DD format.');
  }
}

function promptMember_(assigneesSheet, title, roleLabel, options) {
  const requireName = Boolean(options && options.requireName);
  const requireEmail = Boolean(options && options.requireEmail);
  let name = requireName
    ? promptRequired_(title, `${roleLabel} name`)
    : promptOptional_(title, `${roleLabel} name`);

  if (name === null) {
    return null;
  }

  let email = '';
  const memberByName = name ? findMemberByName_(assigneesSheet, name) : null;
  if (memberByName) {
    name = memberByName.name;
    email = memberByName.email;
  }

  if (email && requireEmail && !isValidEmail_(email)) {
    email = '';
  }

  if (!email && requireEmail) {
    email = promptEmail_(title, `${roleLabel} email`, { required: true });
    if (email === null) {
      return null;
    }

    const memberByEmail = findMemberByEmail_(assigneesSheet, email);
    if (memberByEmail && !name) {
      name = memberByEmail.name;
    }
  }

  return { name, email };
}

function resolveUpdateValue_(value, currentValue) {
  const nextValue = String(value || '').trim();
  if (!nextValue) {
    return currentValue || '';
  }

  if (nextValue.toUpperCase() === CONFIG.clearToken) {
    return '';
  }

  return nextValue;
}

function formatPromptValue_(value) {
  return String(value || 'blank');
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

function getMembers_(assigneesSheet) {
  const lastRow = assigneesSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return assigneesSheet
    .getRange(2, 1, lastRow - 1, ASSIGNEE_HEADERS.length)
    .getValues()
    .map((row, index) => ({
      rowNumber: index + 2,
      name: String(row[0] || '').trim(),
      email: String(row[1] || '').trim(),
      active: String(row[2] || 'Yes').trim(),
    }))
    .filter((member) => member.name && normalizeLookupValue_(member.active) !== 'no');
}

function normalizeLookupValue_(value) {
  return String(value || '').trim().toLowerCase();
}

function getStoredTicketSnapshotForRow_(sheet, rowNumber) {
  const ticket = mapTicketRow_(getTicketRow_(sheet, rowNumber));
  if (!ticket.ticketId) {
    return null;
  }

  return getStoredTicketSnapshot_(ticket.ticketId);
}

function getStoredTicketSnapshot_(ticketId) {
  const rawValue = PropertiesService.getDocumentProperties().getProperty(ticketSnapshotKey_(ticketId));
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
  if (!ticket || !ticket.ticketId) {
    return;
  }

  PropertiesService.getDocumentProperties().setProperty(ticketSnapshotKey_(ticket.ticketId), JSON.stringify(snapshotTicket_(ticket)));
}

function syncTicketSnapshots_(ticketsSheet) {
  const lastRow = ticketsSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  for (let row = 2; row <= lastRow; row += 1) {
    const ticket = mapTicketRow_(getTicketRow_(ticketsSheet, row));
    if (ticket.ticketId && !isBlankTicketRow_(getTicketRow_(ticketsSheet, row))) {
      saveTicketSnapshot_(ticket);
    }
  }
}

function ticketSnapshotKey_(ticketId) {
  return `TICKET_SNAPSHOT_${ticketId}`;
}

function snapshotTicket_(ticket) {
  return {
    ticketId: String(ticket.ticketId || '').trim(),
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
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    return '';
  }
}

function isSettingEnabled_(key, fallback) {
  const value = String(getSetting_(key, fallback ? 'Yes' : 'No') || '').trim().toLowerCase();
  return ['yes', 'true', 'enabled', 'on', '1'].includes(value);
}

function nextTicketId_() {
  const props = PropertiesService.getDocumentProperties();
  const currentValue = Number(props.getProperty('NEXT_TICKET_NUMBER') || '1');
  const nextValue = Number.isFinite(currentValue) && currentValue > 0 ? currentValue : 1;
  props.setProperty('NEXT_TICKET_NUMBER', String(nextValue + 1));
  return `${CONFIG.ticketPrefix}-${String(nextValue).padStart(4, '0')}`;
}

function syncTicketCounter_(ticketsSheet) {
  const lastRow = ticketsSheet.getLastRow();
  if (lastRow <= 1) {
    PropertiesService.getDocumentProperties().setProperty('NEXT_TICKET_NUMBER', '1');
    return;
  }

  const ticketIds = ticketsSheet.getRange(2, COL.TICKET_ID, lastRow - 1, 1).getValues().flat();
  const highestNumber = ticketIds.reduce((highest, ticketId) => {
    const match = String(ticketId || '').match(new RegExp(`^${CONFIG.ticketPrefix}-(\\d+)$`));
    return match ? Math.max(highest, Number(match[1])) : highest;
  }, 0);

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
  sheet.getRange(1, 1, 1, TICKET_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(2, COL.CREATED_AT, Math.max(sheet.getMaxRows() - 1, 1), 2).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, COL.DUE_DATE, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet
    .getRange(2, COL.ASSIGNMENT_NOTIFIED_AT, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');
  sheet
    .getRange(2, COL.UPDATE_NOTIFIED_AT, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');

  const widths = [110, 145, 145, 160, 220, 220, 320, 110, 130, 160, 220, 120, 210, 220, 145, 190, 145, 300];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  ensureFilter_(sheet);
}

function formatAssigneesSheet_(sheet) {
  sheet.getRange(1, 1, 1, ASSIGNEE_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
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

function getOrCreateSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
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
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MMM d, yyyy');
  }

  return String(value || 'Not set');
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
