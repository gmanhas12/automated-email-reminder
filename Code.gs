function sendReminders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet= spreadsheet.getSheetByName("Sheet1");
  var allRows = sheet.getDataRange().getValues();
  var today =new Date();

  // Loop through each row (skip the header row)
  for (var i = 1; i < allRows.length; i++) {  
    var personName =allRows[i][0];
    var recipientEmail = allRows[i][1];
    var dueDate =new Date(allRows[i][2]);
    var taskStatus= allRows[i][3];

    //remaining days 
    var timeDifference = dueDate -today;
    var daysLeft = Math.ceil(timeDifference/(1000 * 60 * 60 * 24));

    // skip if task is already completed or not in range
    if (taskStatus.toLowerCase() !== "pending") {
      continue;
    }
    if (daysLeft > 3 || daysLeft < 0) {
      continue;
    }
    var emailSubject;
    var emailMessage;

    if (daysLeft === 0) {
      emailSubject = "⚠️ Important: Your task is due TODAY!";
      emailMessage= `
        <p>Hey ${personName},</p>
        <p>This is an automated reminder that your deadline is <strong>today</strong>. If you haven’t completed your task yet, please do so as soon as possible.</p>
        <p>Best,</p>
        <p><strong>Reminder Bot</strong></p>
      `;
    } else {
      emailSubject = `Reminder: ${daysLeft} Day(s) Left to Complete Your Task`;
      emailMessage = `
        <p>Hi ${personName},</p>
        <p>This is an automated reminder that your task is due in <strong>${daysLeft} days</strong> (${dueDate.toDateString()}).</p>
        <p>Best,</p>
        <p><strong>Reminder Bot</strong></p>
      `;
    }

    // send reminder
    MailApp.sendEmail({
      to: recipientEmail,
      subject: emailSubject,
      htmlBody: emailMessage
    });

    // update sheet
    sheet.getRange(i+1, 4).setValue("Reminder Sent");
  }
}
