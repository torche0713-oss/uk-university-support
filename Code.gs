/**
 * Google Apps Script - Web App
 * Deploy this as a web app to receive form submissions and write to Google Sheets.
 * 
 * HOW TO SETUP:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire code
 * 4. Replace SHEET_ID with your Google Sheet's ID (from the URL)
 * 5. Replace NOTIFY_EMAIL with your email address(es)
 * 6. Deploy > New deployment > Web app
 * 7. Set "Execute as" to "Me" and "Who has access" to "Anyone"
 * 8. Copy the web app URL and paste it into book.html form action
 * 9. After deploying, run setupHourlyReminder() once manually
 *    to activate the reminder system
 */

const SHEET_ID = '1MExCIMoC6_hITUmadTaImzyZUSFS_j3R2G2rPqP_zTk';
const SHEET_NAME = 'Sheet1';
const NOTIFY_EMAIL = 'torche0713@gmail.com,nierose4602@gmail.com,bailoub88888@gmail.com';

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = e.parameter;
    
    // Append headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Full Name', 'Email', 'Phone', 'Postcode', 'Age',
        'Course Interested In', 'Residency Years', 'Preferred Call Date',
        'Preferred Call Time', 'Too Busy', 'Confused About Documents', 'Reminder Sent'
      ]);
    }
    
    // Append form data
    sheet.appendRow([
      new Date(),
      data.fullName || '',
      data.email || '',
      data.phone || '',
      data.postcode || '',
      data.age || '',
      data.course || '',
      data.residencyYears || '',
      data.callDate || '',
      data.callTime || '',
      data.tooBusy === 'on' ? 'Yes' : 'No',
      data.confused === 'on' ? 'Yes' : 'No',
      'No'
    ]);
    
    // Send confirmation email to the student
    var studentEmail = data.email || '';
    if (studentEmail) {
      try {
        MailApp.sendEmail({
          to: studentEmail,
          subject: 'Your free call is booked - UK University Support',
          htmlBody:
            '<div style="font-family:sans-serif;max-width:560px;margin:0 auto;">' +
            '<div style="background:linear-gradient(135deg,#0F172A,#2563EB);color:#fff;padding:32px;border-radius:16px 16px 0 0;text-align:center;">' +
            '<h1 style="margin:0;font-size:1.5rem;">You\'re all set!</h1>' +
            '</div>' +
            '<div style="background:#fff;border:1px solid #e2e8f0;padding:32px;border-radius:0 0 16px 16px;">' +
            '<p>Hi <strong>' + (data.fullName || 'there') + '</strong>,</p>' +
            '<p>Your free 15-minute planning call has been booked. Here\'s a summary:</p>' +
            '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
            '<tr><td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;width:120px;">Date</td><td style="padding:8px;border:1px solid #e2e8f0;">' + (data.callDate || 'TBC') + '</td></tr>' +
            '<tr><td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;">Time</td><td style="padding:8px;border:1px solid #e2e8f0;">' + (data.callTime || 'TBC') + '</td></tr>' +
            '<tr><td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;">Course Interest</td><td style="padding:8px;border:1px solid #e2e8f0;">' + (data.course || 'TBC') + '</td></tr>' +
            '</table>' +
            '<p>A personal admission consultant will call you at the scheduled time. You\'ll also receive a reminder email 1 hour before the call.</p>' +
            '<p style="color:#64748b;font-size:0.85rem;">Need to reschedule? Reply to this email.</p>' +
            '</div>' +
            '</div>'
        });
      } catch (emailErr) {
        // Confirmation email failed silently - don't block the submission
      }
    }
    
    // Send notification email to team
    var recipients = NOTIFY_EMAIL.split(',').map(function(e) { return e.trim(); });
    for (var i = 0; i < recipients.length; i++) {
      MailApp.sendEmail({
        to: recipients[i],
        subject: 'New lead: ' + (data.fullName || 'Unknown'),
        htmlBody: '<h2>New Student Lead</h2>' +
        '<table style="border-collapse:collapse;width:100%;">' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Name</td><td style="padding:8px;border:1px solid #ddd;">' + (data.fullName || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Email</td><td style="padding:8px;border:1px solid #ddd;">' + (data.email || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Phone</td><td style="padding:8px;border:1px solid #ddd;">' + (data.phone || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Postcode</td><td style="padding:8px;border:1px solid #ddd;">' + (data.postcode || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Course</td><td style="padding:8px;border:1px solid #ddd;">' + (data.course || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Call Date</td><td style="padding:8px;border:1px solid #ddd;">' + (data.callDate || '') + '</td></tr>' +
        '<tr><td style="padding:8px;border:1px solid #ddd;font-weight:600;">Call Time</td><td style="padding:8px;border:1px solid #ddd;">' + (data.callTime || '') + '</td></tr>' +
        '</table>'
      });
    }
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="en-GB">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Thank You</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #0F172A, #2563EB);
            color: #fff;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            text-align: center;
          }
          .card {
            background: rgba(255,255,255,0.1);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 20px;
            padding: 48px 40px;
            max-width: 480px;
          }
          h1 { font-size: 1.8rem; margin-bottom: 12px; }
          p { color: rgba(255,255,255,0.8); line-height: 1.6; }
          .btn {
            display: inline-block;
            margin-top: 24px;
            padding: 14px 32px;
            background: #10B981;
            color: #fff;
            border-radius: 12px;
            text-decoration: none;
            font-weight: 700;
          }
        </style>
      </head>
      <body>
        <div class="card">
          <div style="font-size:3rem;margin-bottom:16px;">✅</div>
          <h1>You're all set!</h1>
          <p>We've received your details. A confirmation email has been sent to your inbox. A personal admission consultant will call you at your scheduled time to help with your next steps.</p>
          <a href="${ScriptApp.getService().getUrl().replace('/exec', '/exec')}" class="btn">Back to Home</a>
        </div>
      </body>
      </html>
    `).setTitle('Thank You');
    
  } catch (err) {
    return HtmlService.createHtmlOutput(`
      <h2>Something went wrong</h2>
      <p>Please try again or contact support.</p>
      <p style="color:#999;font-size:0.85rem;">${err.message}</p>
    `).setTitle('Error');
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html><body>
      <h2>Form endpoint is active</h2>
      <p>This is a Google Apps Script webhook for receiving form submissions.</p>
    </body></html>
  `);
}

/**
 * checkReminders()
 * Runs every hour. Finds calls within the next hour that haven't
 * had a reminder sent, and emails the student.
 */
function checkReminders() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return; // No data rows
  
  var headers = data[0];
  var emailCol = headers.indexOf('Email');
  var nameCol = headers.indexOf('Full Name');
  var dateCol = headers.indexOf('Preferred Call Date');
  var timeCol = headers.indexOf('Preferred Call Time');
  var reminderCol = headers.indexOf('Reminder Sent');
  
  if (emailCol === -1 || dateCol === -1 || timeCol === -1 || reminderCol === -1) return;
  
  var now = new Date();
  var oneHourFromNow = new Date(now.getTime() + 60 * 60 * 1000);
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][reminderCol] === 'Yes') continue; // Already reminded
    
    var callDate = data[i][dateCol];
    var callTime = data[i][timeCol];
    var studentEmail = data[i][emailCol];
    var studentName = data[i][nameCol] || 'there';
    
    if (!callDate || !callTime || !studentEmail) continue;
    
    // Parse call date and time
    var callDateTime = new Date(callDate + 'T' + callTime);
    
    // Check if call is within the next hour
    if (callDateTime > now && callDateTime <= oneHourFromNow) {
      // Send reminder email
      try {
        MailApp.sendEmail({
          to: studentEmail,
          subject: 'Reminder: Your free call is in 1 hour',
          htmlBody:
            '<div style="font-family:sans-serif;max-width:560px;margin:0 auto;">' +
            '<div style="background:linear-gradient(135deg,#0F172A,#2563EB);color:#fff;padding:32px;border-radius:16px 16px 0 0;text-align:center;">' +
            '<h1 style="margin:0;font-size:1.5rem;">Call Reminder</h1>' +
            '</div>' +
            '<div style="background:#fff;border:1px solid #e2e8f0;padding:32px;border-radius:0 0 16px 16px;">' +
            '<p>Hi <strong>' + studentName + '</strong>,</p>' +
            '<p>Your free 15-minute planning call is in <strong>1 hour</strong>.</p>' +
            '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
            '<tr><td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;width:120px;">Date</td><td style="padding:8px;border:1px solid #e2e8f0;">' + callDate + '</td></tr>' +
            '<tr><td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;">Time</td><td style="padding:8px;border:1px solid #e2e8f0;">' + callTime + '</td></tr>' +
            '</table>' +
            '<p>A personal admission consultant will be calling you. Please be available at the scheduled time.</p>' +
            '<p style="color:#64748b;font-size:0.85rem;">Need to reschedule? Reply to this email.</p>' +
            '</div>' +
            '</div>'
        });
      } catch (e) {
        continue;
      }
      
      // Mark reminder as sent
      sheet.getRange(i + 1, reminderCol + 1).setValue('Yes');
    }
  }
}

/**
 * setupHourlyReminder()
 * Run this ONCE manually after deployment to activate hourly reminders.
 * Go to Run > Run function > setupHourlyReminder, then authorize.
 */
function setupHourlyReminder() {
  // Remove existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new hourly trigger
  ScriptApp.newTrigger('checkReminders')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('Hourly reminder trigger created successfully.');
}
