

// function onFormSubmit(e) {
//   if (!e) {
//     logToSheet('No event object received');
//     return;
//   }

//   const sheet = e.range.getSheet();
//   const row = e.range.getRow();
//   const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
//   const statusColumnIndex = 12;

//   if (!sheet.getRange(row, statusColumnIndex).getValue()) {
//     sheet.getRange(row, statusColumnIndex).setValue("Pending");
//   }

//   const message = buildLeaveRequestMessage(data);

//   sendToSlackApp(message);
// }

// function onEdit(e) {
//   const sheet = e.source.getActiveSheet();
//   const editedRange = e.range;
//   const statusColumn = 12;  // Column number where status is located

//   // Check if the edited column is the status column
//   if (editedRange.getColumn() === statusColumn) {
//     const row = editedRange.getRow();
//     const status = e.value;  // "Approved" or "Rejected"
//     const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

//     const name = data[2];
//     const leaveType = data[3];
//     const leaveFrom = formatDate(data[5]);
//     const leaveTo = formatDate(data[6]);
//     const reason = data[8];
    
//     // Get the approver's email (the user who edited the sheet)
//     const approverEmail = Session.getActiveUser().getEmail();
//     Logger.log('Approver email: ' + approverEmail);
    
//     const requesterEmail = data[1];  // Assuming requester email is in the second column
//     sendApprovalEmail(requesterEmail, name, status, leaveType, leaveFrom, leaveTo, reason, approverEmail, data);

//     const rejectReason = data[12];
//     const message = buildApprovalMessage(status, name, leaveFrom, leaveTo, reason, rejectReason, formatTimestamp(new Date()));
//     sendToSlackApp(message);
//   }
// }

// function sendApprovalEmail(requesterEmail, name, status, leaveType, leaveFrom, leaveTo, reason, approverEmail, data) {
//   const subject = `Leave Request Status: ${status === 'Approved' ? '✅ Approved' : '❌ Rejected'}`;

//   const rejectReason = data[12];

//   const body = `
//     <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
//       <p>Dear <strong>${name}</strong>,</p>

//       <p>We would like to inform you about the status of your leave request.</p>

//       <p><strong>📅 Leave Request Status:</strong> <span style="color: ${status === 'Approved' ? '#4CAF50' : '#F44336'}; font-weight: bold;">${status === 'Approved' ? '✅ Approved' : '❌ Rejected'}</span></p>
//       <p><strong>📝 Leave Type:</strong> ${leaveType}</p>
//       <p><strong>⏳ Leave Duration:</strong> ${leaveFrom} to ${leaveTo}</p>
//       <p><strong>💬 Reason for Leave:</strong> ${reason}</p>
      
//       ${status === 'Rejected' && rejectReason ? `<p><strong>❌ Rejected Reason:</strong> ${rejectReason}</p>` : ''}

//       <p><strong>👤 Approver:</strong> ${approverEmail}</p>

//       <p>If you have any further questions or require assistance, please don’t hesitate to contact the HR department.</p>

//       <p>Best regards,</p>
//       <p><strong>Leave Management Team</strong><br>P I I T S Co.,Ltd</p>
//     </div>
//   `;
  
//   MailApp.sendEmail({
//     to: requesterEmail,
//     subject: subject,
//     body: body.replace(/<\/?[^>]+(>|$)/g, ""), // Plain text version
//     htmlBody: body
//   });
// }

// function buildLeaveRequestMessage(values) {
//   const timestamp = formatTimestamp(values[0]);
//   const email = values[1];
//   const name = values[2];
//   const leaveType = values[3];
//   const informedTo = values[4];
//   const leaveFrom = formatDate(values[5]);
//   const leaveTo = formatDate(values[6]);
//   const totalLeave = values[7];
//   const reason = values[8];
//   const attachedFile = values[10] || 'No file attached';


//   const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

//   const message = {
//     blocks: [
//       {
//         type: 'header',
//         text: {
//           type: 'plain_text',
//           text: '📝 New Leave Request Submitted',
//           emoji: true
//         }
//       },
//       { type: 'divider' },
//       {
//         type: 'section',
//         fields: [
//           { type: 'mrkdwn', text: `*👤 Name:*\n${name}` },
//           { type: 'mrkdwn', text: `*📩 Email:*\n${email}` }
//         ]
//       },
//       {
//         type: 'section',
//         fields: [
//           { type: 'mrkdwn', text: `*📌 Leave Type:*\n${leaveType}` },
//           { type: 'mrkdwn', text: `*👨🏻‍💻 Informed To:*\n${informedTo}` }
//         ]
//       },
//       {
//         type: 'section',
//         fields: [
//           { type: 'mrkdwn', text: `*📅 From Date:*\n${leaveFrom}` },
//           { type: 'mrkdwn', text: `*📅 To Date:*\n${leaveTo}  (_${totalLeave} days_)` }
//         ]
//       },
//       {
//         type: 'section',
//         text: { type: 'mrkdwn', text: `*📝 Reason:*\n>${reason}` }
//       },
//       {
//         type: 'section',
//         text: { type: 'mrkdwn', text: `*📎 Attachment:* ${attachedFile}` }
//       },
//       { type: 'divider' },
//       {
//         type: 'context',
//         elements: [{ type: 'mrkdwn', text: `*🕒 Submitted On:* ${timestamp}` }]
//       },
//       {
//         type: 'actions',
//         elements: [
//           {
//             type: 'button',
//             text: { type: 'plain_text', text: '🔍 Review Leave Request' },
//             url: spreadsheetUrl,
//             style: 'primary'
//           }
//         ]
//       }
//     ]
//   };

//   return message;
// }

// function buildApprovalMessage(approvalStatus, name, leaveFrom, leaveTo, reason, rejectReason, timestamp) {
//   const isApproved = approvalStatus === 'Approved';
//   const statusText = isApproved ? '*✅ Approved*' : '*❌ Rejected*';

//   const message = {
//     blocks: [
//       {
//         type: 'header',
//         text: { type: 'plain_text', text: `📢 Leave Request ${approvalStatus}`, emoji: true }
//       },
//       { type: 'divider' },
//       {
//         type: 'section',
//         fields: [
//           { type: 'mrkdwn', text: `*👤 Name:*\n${name}` },
//           { type: 'mrkdwn', text: `*📌 Status:* ${statusText}` }
//         ]
//       },
//       {
//         type: 'section',
//         fields: [
//           { type: 'mrkdwn', text: `*📅 From:* ${leaveFrom}` },
//           { type: 'mrkdwn', text: `*📅 To:* ${leaveTo}` },
//         ]
//       },
//       {
//         type: 'section',
//         fields: [
//         { type: 'mrkdwn', text: `*📝 Reason:*\n>${reason}` },
//         ...(approvalStatus === 'Rejected' && rejectReason
//         ? [{ type: 'mrkdwn', text: `*💁🏻 Reject because:*\n${rejectReason}` }]
//         : [])
//         ]
//       },
//       { type: 'divider' },
//       {
//         type: 'context',
//         elements: [{ type: 'mrkdwn', text: `*🕒 Submitted On:* ${timestamp}` }]
//       },
//     ]
//   };

//   return message;
// }


// function getSlackUserId(informedTo) {
//   try{
//     const options = {
//       method: "get",
//       headers: {
//         Authorization: `Bearer ${SLACK_BOT_TOKEN}`, // Fixed variable name from BOT_TOKEN to SLACK_BOT_TOKEN
//       },
//     };
//     const response = UrlFetchApp.fetch(`${USER_LIST_URL}`, options);
//     const users = JSON.parse(response).members;
//     const user = users.filter((u) => !u.deleted).find(u => u.real_name === informedTo);
//     if (user) {
//       logToSheet(`getSlackUserId: User found - ${user.real_name}, ID: ${user.id}`); // Added logging
//       return user.id;
//     } else {
//       logToSheet(`getSlackUserId: User not found for ${informedTo}`); // Added logging
//       return null;
//     }
//   } catch(err) {
//     logToSheet(`getSlackUserId Error: ${err}`); // Improved error logging
//   }
// }

// function getSlackUserIdByEmail(email) {
//   try {
//     const url = `${LOOK_UP_BY_EMAIL_URL}?email=${encodeURIComponent(email)}`;
//     const options = {
//       "method": "get",
//       "headers": {
//         "Authorization": "Bearer " + BOT_TOKEN
//       }
//     };
//     const response = UrlFetchApp.fetch(url, options);
//     const result = JSON.parse(response.getContentText());
//     if (result.ok) {
//       var userId = result.user.id;
//       return userId;
//     } else {
//       throw new Error('User ID not found!')
//     }
//   } catch(err) {
//     logToSheet(err);
//   }
// }



// function sendToSlackApp(message) {
//   const options = {
//     method: 'post',
//     contentType: 'application/json',
//     payload: JSON.stringify({
//       channel: SLACK_CHANNEL,
//       text: message, // Plaintext fallback
//       blocks: message.blocks // Optional interactive blocks
//     }),
//     headers: {
//       Authorization: `Bearer ${SLACK_BOT_TOKEN}`
//     }
//   };

//   try {
//     UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
//     logToSheet('Slack notification sent successfully.');
//   } catch (err) {
//     logToSheet('Failed to send Slack notification: ' + err);
//   }
// }


// function formatDate(date) {
//   const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd');
//   return formattedDate;
// }

// function formatTimestamp(timestamp) {
//   const formattedTimestamp = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');
//   return formattedTimestamp;
// }

// function logToSheet(message) {
//   let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
//   if (!sheet) {
//     sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
//     sheet.appendRow(["Timestamp", "Log Message"]);
//   }
//   sheet.appendRow([new Date(), message]); // Log the message with timestamp
// }


