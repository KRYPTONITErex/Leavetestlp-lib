// Main.gs

const PROPS = PropertiesService.getScriptProperties();
const SLACK = {
  WEBHOOK_URL: PROPS.getProperty('SLACK_WEBHOOK_URL') || 'default-webhook-url',
  BOT_TOKEN: PROPS.getProperty('SLACK_BOT_TOKEN') || 'default-bot-token',
  CHANNEL: PROPS.getProperty('SLACK_CHANNEL') || '#default-channel',
  USER_LIST_URL: PROPS.getProperty('USER_LIST_URL'),
  LOOKUP_BY_EMAIL_URL: PROPS.getProperty('LOOK_UP_BY_EMAIL_URL'),
  POST_MSG_URL: PROPS.getProperty('POST_MSG_URL')
};

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
  
  sendToSlackApp(buildLeaveRequestMessage(data));
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRange = e.range;
  const statusColumn = 12;
  
  if (editedRange.getColumn() === statusColumn) {
    const row = editedRange.getRow();
    const status = e.value;
    const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const name = data[2];
    const leaveType = data[3];
    const leaveFrom = formatDate(data[5]);
    const leaveTo = formatDate(data[6]);
    const reason = data[8];
    const rejectReason = data[12];
    
    // Handle approval process
    const approverEmail = Session.getActiveUser().getEmail();
    const requesterEmail = data[1];
    
    // Send email notification
    sendApprovalEmail(requesterEmail, name, status, leaveType, leaveFrom, leaveTo, reason, approverEmail, data);
    
    // Send Slack notifications
    sendToSlackApp(buildApprovalMessage(status, name, leaveFrom, leaveTo, reason, rejectReason, formatTimestamp(new Date())));
    
    const requesterUserId = getSlackUserIdByEmail(requesterEmail);
    logToSheet(`Requester Slack User ID: ${requesterUserId || 'Not found'}`);
    
    if (requesterUserId) {
      sendLeaveResponseSlackMessage({
        userName: name,
        userId: requesterUserId,
        leaveType,
        status,
        fromDate: leaveFrom,
        toDate: leaveTo,
        managerEmail: approverEmail,
        rejectedReason: rejectReason ?? '-'
      });
    } else {
      logToSheet(`Cannot send Slack message - no user ID found for ${requesterEmail}`);
    }
  }
}