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
  try {
    const options = {
      method: "get",
      headers: {
        Authorization: `Bearer ${SLACK_BOT_TOKEN}`,
      },
    };
    const response = UrlFetchApp.fetch(`${USER_LIST_URL}`, options);
    const users = JSON.parse(response).members;
    const user = users.filter((u) => !u.deleted).find(u => u.real_name === informedTo);
    if (user) {
      Logger.log(`User found - ${user.real_name}, ID: ${user.id}`);
      return user.id;
    } else {
      Logger.log(`User not found for ${informedTo}`);
      return null;
    }
  } catch (err) {
    Logger.log(`getSlackUserId Error: ${err}`);
    return null; 
  }
}

function getSlackUserIdByEmail(email) {
  if (!email) {
    logToSheet("Error: No email provided to getSlackUserIdByEmail");
    return null;
  }
  
  try {
    const url = `${LOOK_UP_BY_EMAIL_URL}?email=${encodeURIComponent(email)}`;
    logToSheet(`Looking up Slack user ID for email: ${email}`);
    
    const options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + SLACK_BOT_TOKEN
      }
    };
    
    // Log the full URL being called (but mask the token in logs)
    logToSheet(`API Call URL: ${url}`);
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      logToSheet(`User found: ${result.user.id} for email: ${email}`);
      return result.user.id;
    } else {
      logToSheet(`Error finding user for email ${email}: ${result.error}`);
      return null;
    }
  } catch (err) {
    logToSheet(`Exception in getSlackUserIdByEmail: ${err.toString()}`);
    
    // Additional troubleshooting
    logToSheet(`LOOK_UP_BY_EMAIL_URL value: ${LOOK_UP_BY_EMAIL_URL}`);
    return null;
  }
}

function sendToSlackApp(message) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      channel: SLACK_CHANNEL,
      text: message, 
      blocks: message.blocks
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


function sendLeaveResponseSlackMessage({
  userName,
  userId,
  leaveType,
  status,
  fromDate,
  toDate,
  managerEmail,
  rejectedReason = '-'
}) {
  if (!userId) {
    logToSheet("Error: Cannot send Slack message - userID is null or empty");
    return;
  }

  const isApproved = status === 'Approved';
  const statusText = isApproved ? '*✅ Approved*' : '*❌ Rejected*';
  
  const message = {
    channel: userId,
    blocks: [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: `👋 Hello ${userName}, Your leave request has been ${status}.`,
          emoji: true
        }
      },
      { type: "divider" },
      {
        type: "section",
        fields: [
          {
            type: "mrkdwn",
            text: `*📌 Leave Type :*\n         ${leaveType}`
          },
          {
            type: "mrkdwn",
            text: `*📊 Status :*\n         ${statusText}`
          }
        ]
      },

      ...(status === 'Rejected' && rejectedReason !== '-' ? [{
        type: 'section',
        fields: [
          {
            type: 'mrkdwn',
            text: `*💁🏻 Reject because:*\n${rejectedReason}`
          }
        ]
      }] : []),

      {
        type: "section",
        fields: [
          {
            type: "mrkdwn",
            text: `*📅 From Date :*\n         ${fromDate}`
          },
          {
            type: "mrkdwn",
            text: `*📅 To Date :*\n         ${toDate}`
          }
        ]
      },
      {
        type: "section",
        fields: [
          {
            type: "mrkdwn",
            text: `*👨🏻‍💼 Approver :*\n         ${managerEmail}`
          }
        ]
      }
    ],
  };

  logToSheet(`Preparing to send Slack message to user ID: ${userId}`);
  const payload = JSON.stringify(message);
  logToSheet(`Message payload: ${payload.substring(0, 200)}...`);

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${SLACK_BOT_TOKEN}`,
    },
    payload,
  };

  try {
    const response = UrlFetchApp.fetch(POST_MSG_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    logToSheet(`Slack API response code: ${responseCode}`);
    logToSheet(`Slack API response: ${responseText}`);
    
    if (responseCode === 200) {
      logToSheet(`Slack notification sent successfully to user ID ${userId}`);
    } else {
      logToSheet(`Failed to send Slack notification. Status code: ${responseCode}`);
    }
  } catch(err) {
    logToSheet(`Error sending Slack message: ${err.toString()}`);
  }
}