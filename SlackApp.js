function buildLeaveRequestMessage(values) {
  const timestamp = formatTimestamp(values[0]);
  const email = values[1];
  const name = values[2];
  const leaveType = values[3];
  const informedTo = values[4];
  const leaveFrom = formatDate(values[5]);
  const leaveTo = formatDate(values[6]);
  const totalLeave = values[7];
  const reason = values[8];
  const attachedFile = values[10] || 'No file attached';


  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const message = {
    blocks: [
      {
        type: 'header',
        text: {
          type: 'plain_text',
          text: '📝 New Leave Request Submitted',
          emoji: true
        }
      },
      { type: 'divider' },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*👤 Name:*\n${name}` },
          { type: 'mrkdwn', text: `*📩 Email:*\n${email}` }
        ]
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*📌 Leave Type:*\n${leaveType}` },
          { type: 'mrkdwn', text: `*👨🏻‍💻 Informed To:*\n${informedTo}` }
        ]
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*📅 From Date:*\n${leaveFrom}` },
          { type: 'mrkdwn', text: `*📅 To Date:*\n${leaveTo}  (_${totalLeave} days_)` }
        ]
      },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: `*📝 Reason:*\n>${reason}` }
      },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: `*📎 Attachment:* ${attachedFile}` }
      },
      { type: 'divider' },
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: `*🕒 Submitted On:* ${timestamp}` }]
      },
      {
        type: 'actions',
        elements: [
          {
            type: 'button',
            text: { type: 'plain_text', text: '🔍 Review Leave Request' },
            url: spreadsheetUrl,
            style: 'primary'
          }
        ]
      }
    ]
  };

  return message;
}

function buildApprovalMessage(approvalStatus, name, leaveFrom, leaveTo, reason, rejectReason, timestamp) {
  const isApproved = approvalStatus === 'Approved';
  const statusText = isApproved ? '*✅ Approved*' : '*❌ Rejected*';

  const message = {
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: `📢 Leave Request ${approvalStatus}`, emoji: true }
      },
      { type: 'divider' },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*👤 Name:*\n${name}` },
          { type: 'mrkdwn', text: `*📌 Status:* ${statusText}` }
        ]
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*📅 From:* ${leaveFrom}` },
          { type: 'mrkdwn', text: `*📅 To:* ${leaveTo}` },
        ]
      },
      {
        type: 'section',
        fields: [
        { type: 'mrkdwn', text: `*📝 Reason:*\n>${reason}` },
        ...(approvalStatus === 'Rejected' && rejectReason
        ? [{ type: 'mrkdwn', text: `*💁🏻 Reject because:*\n${rejectReason}` }]
        : [])
        ]
      },
      { type: 'divider' },
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: `*🕒 Submitted On:* ${timestamp}` }]
      },
    ]
  };

  return message;
}


function getSlackUserId(informedTo) {
  try{
    const options = {
      method: "get",
      headers: {
        Authorization: `Bearer ${SLACK_BOT_TOKEN}`, // Fixed variable name from BOT_TOKEN to SLACK_BOT_TOKEN
      },
    };
    const response = UrlFetchApp.fetch(`${USER_LIST_URL}`, options);
    const users = JSON.parse(response).members;
    const user = users.filter((u) => !u.deleted).find(u => u.real_name === informedTo);
    if (user) {
      logToSheet(`getSlackUserId: User found - ${user.real_name}, ID: ${user.id}`); // Added logging
      return user.id;
    } else {
      logToSheet(`getSlackUserId: User not found for ${informedTo}`); // Added logging
      return null;
    }
  } catch(err) {
    logToSheet(`getSlackUserId Error: ${err}`); // Improved error logging
  }
}

function getSlackUserIdByEmail(email) {
  try {
    const url = `${LOOK_UP_BY_EMAIL_URL}?email=${encodeURIComponent(email)}`;
    const options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + SLACK_BOT_TOKEN
      }
    };
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.ok) {
      var userId = result.user.id;
      return userId;
    } else {
      throw new Error('User ID not found!')
    }
  } catch(err) {
    logToSheet(err);
  }
}



function sendToSlackApp(message) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      channel: SLACK_CHANNEL,
      text: message, // Plaintext fallback
      blocks: message.blocks // Optional interactive blocks
    }),
    headers: {
      Authorization: `Bearer ${SLACK_BOT_TOKEN}`
    }
  };

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    logToSheet('Slack notification sent successfully.');
  } catch (err) {
    logToSheet('Failed to send Slack notification: ' + err);
  }
}


// function sendLeaveResponseSlackMessage({userName, userId, leaveType, status, fromDate, toDate, managerEmail, rejectedReason}) {
//   const isApproved = status === 'Approved';
//   const statusText = isApproved ? '*✅ Approved*' : '*❌ Rejected*';
  
//   const message = {
//     channel: userId,
//     blocks: [
//       {
//         type: "header",
//         text: {
//           type: "plain_text",
//           text: `👋 Hello ${userName}, Your leave request has been ${status}.`,
//           emoji: true
//         }
//       },
//       { type: "divider" },
//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*📌 Leave Type :*\n         ${leaveType}`
//           },
//           {
//             type: "mrkdwn",
//             text: `*📊 Status :*\n         ${statusText}`
//           }
//         ]
//       },

//       ...(status === 'Rejected' && rejectedReason ? [{
//         type: 'section',
//         fields: [
//           {
//             type: 'mrkdwn',
//             text: `*💁🏻 Reject because:*\n${rejectedReason}`
//           }
//         ]
//       }] : []),
//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*📅 From Date :*\n         ${formatDate(fromDate)}`
//           },
//           {
//             type: "mrkdwn",
//             text: `*📅 To Date :*\n         ${formatDate(toDate)}`
//           }
//         ]
//       },
//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*👨🏻‍💼 Approver :*\n         ${managerEmail}`
//           }
//         ]
//       }
//     ],
//   };

//   const payload = JSON.stringify(message);

//   const options = {
//     method: "POST",
//     contentType  :   "application/json",
//     headers: {
//       Authorization: `Bearer ${BOT_TOKEN}`,
//     },
//     payload,
//   };

//   try {
//     UrlFetchApp.fetch(POST_MSG_URL, options);
//     logToSheet("Slack notification sent successfully.");
//   } catch(err) {
//     logToSheet(err);
//   }
// }


//just test ment for clasp push

// function sendLeaveResponseSlackMessage({
//   userName,
//   userId,
//   leaveType,
//   status,
//   fromDate,
//   toDate,
//   managerEmail,
//   rejectedReason = '-'  // Default value if rejectReason is not provided
// }) {
//   const isApproved = status === 'Approved';
//   const statusText = isApproved ? '*✅ Approved*' : '*❌ Rejected*';
  
//   const message = {
//     channel: userId,  // Direct message to the user (via their Slack userId)
//     blocks: [
//       {
//         type: "header",
//         text: {
//           type: "plain_text",
//           text: `👋 Hello ${userName}, Your leave request has been ${status}.`,
//           emoji: true
//         }
//       },
//       { type: "divider" },
//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*📌 Leave Type :*\n         ${leaveType}`
//           },
//           {
//             type: "mrkdwn",
//             text: `*📊 Status :*\n         ${statusText}`
//           }
//         ]
//       },

//       ...(status === 'Rejected' && rejectedReason !== '-' ? [{
//         type: 'section',
//         fields: [
//           {
//             type: 'mrkdwn',
//             text: `*💁🏻 Reject because:*\n${rejectedReason}`
//           }
//         ]
//       }] : []),

//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*📅 From Date :*\n         ${formatDate(fromDate)}`
//           },
//           {
//             type: "mrkdwn",
//             text: `*📅 To Date :*\n         ${formatDate(toDate)}`
//           }
//         ]
//       },
//       {
//         type: "section",
//         fields: [
//           {
//             type: "mrkdwn",
//             text: `*👨🏻‍💼 Approver :*\n         ${managerEmail}`
//           }
//         ]
//       }
//     ],
//   };

//   const payload = JSON.stringify(message);

//   const options = {
//     method: "POST",
//     contentType  :   "application/json",
//     headers: {
//       Authorization: `Bearer ${SLACK_BOT_TOKEN}`,
//     },
//     payload,
//   };

//   try {
//     UrlFetchApp.fetch(POST_MSG_URL, options);
//     logToSheet("Slack notification sent successfully.");
//   } catch(err) {
//     logToSheet(`Error sending Slack message: ${err}`);
//   }
// }


