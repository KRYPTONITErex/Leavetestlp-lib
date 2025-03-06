

const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || 'default-webhook-url';
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || 'default-bot-token';
const SLACK_CHANNEL = PropertiesService.getScriptProperties().getProperty('SLACK_CHANNEL') || '#default-channel';

const USER_LIST_URL = PropertiesService.getScriptProperties().getProperty('USER_LIST_URL');

const LOOK_UP_BY_EMAIL_URL = PropertiesService.getScriptProperties().getProperty('LOOK_UP_BY_EMAIL_URL');
const POST_MSG_URL = PropertiesService.getScriptProperties().getProperty('POST_MSG_URL');

function onFormSubmit(e) {
  if (!e) {
    logToSheet('No event object received');
    return;
  }

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColumnIndex = 12;

  if (!sheet.getRange(row, statusColumnIndex).getValue()) {
    sheet.getRange(row, statusColumnIndex).setValue("Pending");
  }

  const message = buildLeaveRequestMessage(data);

  sendToSlackApp(message);
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRange = e.range;
  const statusColumn = 12;  // Column number where status is located

  // Check if the edited column is the status column
  if (editedRange.getColumn() === statusColumn) {
    const row = editedRange.getRow();
    const status = e.value;  // "Approved" or "Rejected"
    const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    const name = data[2];
    const leaveType = data[3];
    const leaveFrom = formatDate(data[5]);
    const leaveTo = formatDate(data[6]);
    const reason = data[8];
    
    // Get the approver's email (the user who edited the sheet)
    const approverEmail = Session.getActiveUser().getEmail();
    Logger.log('Approver email: ' + approverEmail);
    
    const requesterEmail = data[1];  // Assuming requester email is in the second column
    sendApprovalEmail(requesterEmail, name, status, leaveType, leaveFrom, leaveTo, reason, approverEmail, data);

    const rejectReason = data[12];
    const message = buildApprovalMessage(status, name, leaveFrom, leaveTo, reason, rejectReason, formatTimestamp(new Date()));
    sendToSlackApp(message);


    sendLeaveResponseSlackMessage({
      userName: name,
      userId: getSlackUserIdByEmail(requesterEmail),
      leaveType,
      status,
      fromDate: leaveFrom,
      toDate: leaveTo,
      managerEmail: approverEmail,
      rejectedReason: rejectReason ?? '-'
    });
  }
}
