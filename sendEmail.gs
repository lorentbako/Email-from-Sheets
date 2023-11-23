function sendEmail() {
  // Get the email address and message from the sheet
  const sheet = SpreadsheetApp.getActive()
  const messageSheet = sheet.getSheetByName('Email')
  const emailAddress = messageSheet.getRange("J2").getValue();
  const emailSubject = messageSheet.getRange("I2").getValue();

  let ccAddress = messageSheet.getRange("K2:K").getValues();
  //filter empty rows
  ccAddress = ccAddress.filter(row => row[0]).flat()

  //join all ccAddresses
  ccAddress = ccAddress.join(",")

  //get the range of the table from an input
  let inputRange = Browser.inputBox("Please enter the range of the message (e.g. A1:F9)");
  let message = messageSheet.getRange(inputRange).getValues();
  // remove empty rows
  message = message.filter(row => row[0])

  if (message.length > 0) {
    // Create an HTML table to display the message
    var html = '<table border="1px" cellpadding="4" style="border-collapse: collapse; margin: 25px 0; text-align: left; overflow: scroll;">';
    for (var i = 0; i < message.length; i++) {
      html += '<tr>';
      for (var j = 0; j < message[i].length; j++) {
        if (j == 1 && i != 0) {
          var date = new Date(message[i][j]);
          html += '<td>' + '  ' + Utilities.formatDate(date, "GMT", "dd/MM/yyyy") + '  ' + '</td>';
        } else {
          html += '<td>' + message[i][j] + '</td>';
        }
      }
      html += '</tr>';
    }
    html += '</table>';

    // Send the email
    GmailApp.sendEmail(emailAddress, emailSubject, "", {
      cc: ccAddress,
      htmlBody: html
    });

    // Show a message that the email was sent
    sheet.toast(`E-Mail was sent successfully to ${emailAddress} and ${ccAddress}`);
  }

  if (message.length === 0) {
    Browser.msgBox("Please enter a range that has non-empty fields!")
  }
}
