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
};

const TICKET_HEADERS = [
  'Ticket ID',
  'Created At',
  'Updated At',
  'Title',
  'Description',
  'Priority',
  'Status',
  'Assignee',
  'Assignee Email',
  'Due Date',
  'Notification Status',
  'Notified Email',
  'Notified At',
  'Notes',
];

const ASSIGNEE_HEADERS = ['Name', 'Email', 'Active'];
const SETTINGS_HEADERS = ['Setting', 'Value'];

const COL = {
  TICKET_ID: 1,
  CREATED_AT: 2,
  UPDATED_AT: 3,
  TITLE: 4,
  DESCRIPTION: 5,
  PRIORITY: 6,
  STATUS: 7,
  ASSIGNEE: 8,
  ASSIGNEE_EMAIL: 9,
  DUE_DATE: 10,
  NOTIFICATION_STATUS: 11,
  NOTIFIED_EMAIL: 12,
  NOTIFIED_AT: 13,
  NOTES: 14,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MVFT Tracker')
    .addItem('Set up tracker', 'setupTracker')
    .addItem('Install email trigger', 'installTriggers')
    .addSeparator()
    .addItem('Send pending notifications', 'sendPendingAssignmentNotifications')
    .addToUi();
}

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketsSheet = getOrCreateSheet_(ss, CONFIG.sheets.tickets);
  const assigneesSheet = getOrCreateSheet_(ss, CONFIG.sheets.assignees);
  const settingsSheet = getOrCreateSheet_(ss, CONFIG.sheets.settings);

  setHeaderRow_(ticketsSheet, TICKET_HEADERS);
  setHeaderRow_(assigneesSheet, ASSIGNEE_HEADERS);
  setHeaderRow_(settingsSheet, SETTINGS_HEADERS);

  seedSettings_(settingsSheet);
  formatTicketsSheet_(ticketsSheet);
  formatAssigneesSheet_(assigneesSheet);
  formatSettingsSheet_(settingsSheet);
  applyValidations_(ticketsSheet, assigneesSheet);
  syncTicketCounter_(ticketsSheet);

  alertUser_('MVFT Tracking System setup is complete.');
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
      normalizeTicketRow_(sheet, row);
      maybeSendAssignmentEmail_(sheet, row);
    }
  } finally {
    lock.releaseLock();
  }
}

function sendPendingAssignmentNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.tickets);
  if (!sheet) {
    alertUser_('Tickets sheet was not found. Run setupTracker first.');
    return;
  }

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
      if (maybeSendAssignmentEmail_(sheet, row)) {
        sentCount += 1;
      }
    }
  } finally {
    lock.releaseLock();
  }

  alertUser_(`Pending notification run complete. Emails sent: ${sentCount}.`);
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
  const row = getTicketRow_(sheet, rowNumber);
  const ticket = mapTicketRow_(row);

  if (!ticket.title) {
    return false;
  }

  if (isClosedStatus_(ticket.status)) {
    return false;
  }

  if (!ticket.assigneeEmail) {
    sheet.getRange(rowNumber, COL.NOTIFICATION_STATUS).setValue('Missing Email');
    return false;
  }

  if (!isValidEmail_(ticket.assigneeEmail)) {
    sheet.getRange(rowNumber, COL.NOTIFICATION_STATUS).setValue('Invalid Email');
    return false;
  }

  if (
    String(ticket.notificationStatus).toLowerCase() === 'sent' &&
    String(ticket.notifiedEmail).toLowerCase() === String(ticket.assigneeEmail).toLowerCase()
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

    sheet.getRange(rowNumber, COL.NOTIFICATION_STATUS).setValue('Sent');
    sheet.getRange(rowNumber, COL.NOTIFIED_EMAIL).setValue(ticket.assigneeEmail);
    sheet.getRange(rowNumber, COL.NOTIFIED_AT).setValue(new Date());
    return true;
  } catch (error) {
    sheet.getRange(rowNumber, COL.NOTIFICATION_STATUS).setValue(`Error: ${error.message}`);
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

function getTicketRow_(sheet, rowNumber) {
  return sheet.getRange(rowNumber, 1, 1, TICKET_HEADERS.length).getValues()[0];
}

function mapTicketRow_(row) {
  return {
    ticketId: row[COL.TICKET_ID - 1],
    createdAt: row[COL.CREATED_AT - 1],
    updatedAt: row[COL.UPDATED_AT - 1],
    title: String(row[COL.TITLE - 1] || '').trim(),
    description: String(row[COL.DESCRIPTION - 1] || '').trim(),
    priority: String(row[COL.PRIORITY - 1] || '').trim(),
    status: String(row[COL.STATUS - 1] || '').trim(),
    assignee: String(row[COL.ASSIGNEE - 1] || '').trim(),
    assigneeEmail: String(row[COL.ASSIGNEE_EMAIL - 1] || '').trim(),
    dueDate: row[COL.DUE_DATE - 1],
    notificationStatus: String(row[COL.NOTIFICATION_STATUS - 1] || '').trim(),
    notifiedEmail: String(row[COL.NOTIFIED_EMAIL - 1] || '').trim(),
    notifiedAt: row[COL.NOTIFIED_AT - 1],
    notes: String(row[COL.NOTES - 1] || '').trim(),
  };
}

function isBlankTicketRow_(row) {
  const importantValues = [
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

function setHeaderRow_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function seedSettings_(settingsSheet) {
  const defaults = [
    ['Tracker Name', CONFIG.trackerName],
    ['Notification Prefix', CONFIG.notificationPrefix],
    ['Default Priority', CONFIG.defaultPriority],
    ['Default Status', CONFIG.defaultStatus],
  ];

  const existing = getSettingsMap_(settingsSheet);
  const rowsToAdd = defaults.filter(([setting]) => !existing[setting]);

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
  sheet.getRange(2, COL.NOTIFIED_AT, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');

  const widths = [110, 145, 145, 220, 320, 110, 130, 160, 220, 120, 160, 220, 145, 300];
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

