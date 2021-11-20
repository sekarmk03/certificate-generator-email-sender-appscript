function sendEmails() {
  
  let email = 0;
  let name = 1;

  const emailTemp = HtmlService.createTemplateFromFile("email");
  const spFile = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails");
  const data = spFile.getRange("A2:B" + spFile.getLastRow()).getValues();
  const mailSubject = "Announcement: Google DSC UPI Member Acceptance";

  data.forEach(function(row){
    emailTemp.name = row[name];
    let htmlMessage = emailTemp.evaluate().getContent();
    GmailApp.sendEmail(
      row[email],
      mailSubject,
      "Your email doesn't support HTML.",
      {name: "Google Developer Student Clubs UPI", htmlBody: htmlMessage}
    );
  });

}