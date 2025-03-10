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


function updateLeaveBalanceSheet(name, totalLeaveDays) {
  try {
    // Get the leave balance sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const balanceSheet = ss.getSheetByName("Leave Balance"); // Adjust sheet name if different
    
    if (!balanceSheet) {
      logToSheet("Error: Leave Balance sheet not found");
      return;
    }
    
    // Find the row for this user
    const dataRange = balanceSheet.getDataRange();
    const values = dataRange.getValues();
    
    let userRow = -1;
    for (let i = 1; i < values.length; i++) { // Start from 1 to skip header
      if (values[i][1] === name) { // Assuming name is in column B (index 1)
        userRow = i + 1; // Add 1 because array is 0-indexed but sheets are 1-indexed
        break;
      }
    }
    
    if (userRow === -1) {
      logToSheet(`Error: User ${name} not found in Leave Balance sheet`);
      return;
    }
    
    // Get current taken leave value
    const takenLeaveCell = balanceSheet.getRange(userRow, 3); // Assuming Taken Leave is column C
    const currentTakenLeave = Number(takenLeaveCell.getValue()) || 0; // Convert to number, default to 0 if NaN
    
    // Get total leave allocation
    const totalLeaveCell = balanceSheet.getRange(userRow, 4); // Assuming Total Leave is column D
    const totalLeave = Number(totalLeaveCell.getValue()) || 0; // Convert to number, default to 0 if NaN
    
    // Ensure totalLeaveDays is a number
    const leaveDaysNum = Number(totalLeaveDays) || 0;
    
    // Calculate new values
    const newTakenLeave = currentTakenLeave + leaveDaysNum;
    
    // Calculate remaining leave
    const remainingLeave = totalLeave - newTakenLeave;
    
    // Update cells
    takenLeaveCell.setValue(newTakenLeave);
    balanceSheet.getRange(userRow, 5).setValue(remainingLeave); // Assuming Remaining Leave is column E
    
    logToSheet(`Successfully updated leave balance for ${name}: Added ${leaveDaysNum} days, New total taken: ${newTakenLeave}, Remaining: ${remainingLeave}`);
  } catch (error) {
    logToSheet(`Error updating leave balance: ${error.toString()}`);
  }
}

// Modify the existing onEdit function to include leave balance updates
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
    const totalLeaveDays = data[7]; // Get the total leave days from column 8
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
    
    // Update leave balance sheet if status is Approved
    if (status === 'Approved' && totalLeaveDays > 0) {
      updateLeaveBalanceSheet(name, totalLeaveDays);
    }
  }
}