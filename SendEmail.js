

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