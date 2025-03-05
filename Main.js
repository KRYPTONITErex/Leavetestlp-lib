

function setProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SLACK_WEBHOOK_URL', 'your-webhook-url');
  scriptProperties.setProperty('SLACK_BOT_TOKEN', 'your-bot-token');
  scriptProperties.setProperty('SLACK_CHANNEL', '#your-channel');
  scriptProperties.setProperty('USER_LIST_URL', 'your-user-list-url');
  scriptProperties.setProperty('POST_MSG_URL', 'your-post-msg-url');
  scriptProperties.setProperty('LOOK_UP_BY_EMAIL_URL', 'your-lookup-by-email-url');
}

const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || 'default-webhook-url';
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || 'default-bot-token';
const SLACK_CHANNEL = PropertiesService.getScriptProperties().getProperty('SLACK_CHANNEL') || '#default-channel';

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


    sendLeaveResponseSlackMessage(data);
  }
}

function sendApprovalEmail(requesterEmail, name, status, leaveType, leaveFrom, leaveTo, reason, approverEmail, data) {
  const subject = `Leave Request Status: ${status === 'Approved' ? '✅ Approved' : '❌ Rejected'}`;

  const rejectReason = data[12];

  const body = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <p>Dear <strong>${name}</strong>,</p>

      <p>We would like to inform you about the status of your leave request.</p>

      <p><strong>📅 Leave Request Status:</strong> <span style="color: ${status === 'Approved' ? '#4CAF50' : '#F44336'}; font-weight: bold;">${status === 'Approved' ? '✅ Approved' : '❌ Rejected'}</span></p>
      <p><strong>📝 Leave Type:</strong> ${leaveType}</p>
      <p><strong>⏳ Leave Duration:</strong> ${leaveFrom} to ${leaveTo}</p>
      <p><strong>💬 Reason for Leave:</strong> ${reason}</p>
      
      ${status === 'Rejected' && rejectReason ? `<p><strong>❌ Rejected Reason:</strong> ${rejectReason}</p>` : ''}

      <p><strong>👤 Approver:</strong> ${approverEmail}</p>

      <p>If you have any further questions or require assistance, please don’t hesitate to contact the HR department.</p>

      <p>Best regards,</p>
      <p><strong>Leave Management Team</strong><br>P I I T S Co.,Ltd</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: requesterEmail,
    subject: subject,
    body: body.replace(/<\/?[^>]+(>|$)/g, ""), // Plain text version
    htmlBody: body
  });
}