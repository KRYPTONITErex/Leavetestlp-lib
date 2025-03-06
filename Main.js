

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



function checkSlackConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs") || 
                SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
  
  sheet.appendRow(["Timestamp", "Config Check"]);
  
  for (const key in props) {
    if (key.includes('SLACK') || key.includes('URL')) {
      // Only log the first few characters of tokens for security
      const value = props[key].startsWith('xoxb-') ? 
                    props[key].substring(0, 10) + '...' : 
                    props[key];
      
      sheet.appendRow([new Date(), `${key}: ${value}`]);
    }
  }
}


function fixScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentProps = scriptProperties.getProperties();
  
  // Log current values
  logToSheet("Current script properties:");
  for (const [key, value] of Object.entries(currentProps)) {
    if (key.includes('URL') || key.includes('TOKEN')) {
      // Mask tokens for security
      const maskedValue = value.startsWith('xoxb-') ? 
                         value.substring(0, 10) + '...' : 
                         value;
      logToSheet(`${key}: ${maskedValue}`);
    }
  }
  
  // Fix LOOK_UP_BY_EMAIL_URL if it has quotes
  if (currentProps.LOOK_UP_BY_EMAIL_URL && 
      currentProps.LOOK_UP_BY_EMAIL_URL.includes('"')) {
    const fixedValue = currentProps.LOOK_UP_BY_EMAIL_URL.replace(/"/g, '');
    scriptProperties.setProperty('LOOK_UP_BY_EMAIL_URL', fixedValue);
    logToSheet(`Fixed LOOK_UP_BY_EMAIL_URL: ${fixedValue}`);
  }
  
  // Fix other URL properties if needed
  if (currentProps.USER_LIST_URL && 
      currentProps.USER_LIST_URL.includes('"')) {
    const fixedValue = currentProps.USER_LIST_URL.replace(/"/g, '');
    scriptProperties.setProperty('USER_LIST_URL', fixedValue);
    logToSheet(`Fixed USER_LIST_URL: ${fixedValue}`);
  }
  
  if (currentProps.POST_MSG_URL && 
      currentProps.POST_MSG_URL.includes('"')) {
    const fixedValue = currentProps.POST_MSG_URL.replace(/"/g, '');
    scriptProperties.setProperty('POST_MSG_URL', fixedValue);
    logToSheet(`Fixed POST_MSG_URL: ${fixedValue}`);
  }
  
  // Verify SLACK_BOT_TOKEN format
  if (currentProps.SLACK_BOT_TOKEN && 
      !currentProps.SLACK_BOT_TOKEN.startsWith('xoxb-')) {
    logToSheet(`Warning: SLACK_BOT_TOKEN may have an incorrect format. Tokens usually start with 'xoxb-'`);
  }
  
  logToSheet("Script properties check/fix completed");
}
