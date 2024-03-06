const APP_TITLE = "env(APP_TITLE)";
const slideTemplateId = "env(SLIDE_TEMPLATE_ID)";
const tempFolder = getFolderByName('Temp Certificates');
const archiveFolder = getFolderByName('Archive Certificates');

function generateCertificates() {
  const template = DriveApp.getFileById(slideTemplateId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const dataNameIndex = headers.indexOf('Cert Name');
  const dataUseridIndex = headers.indexOf('User ID');
  const dataCertidIndex = headers.indexOf('File ID');
  const dataStatusIndex = headers.indexOf('Status');
  const dataCertURLIndex = headers.indexOf('File URL');

  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    const dataName = rowData[dataNameIndex];
    const dataUserid = rowData[dataUseridIndex];
    const dataStatus = rowData[dataStatusIndex];

    if (dataStatus != 'CREATED' && dataStatus != 'SENT' && dataStatus != 'ARCHIVED') {
      ss.toast(`Generate certificate for ${dataName}`, APP_TITLE, 3);

      const cpSlideId = template.makeCopy(tempFolder).setName(dataName + '-' + dataUserid).getId();
      const slide = SlidesApp.openById(cpSlideId).getSlides()[0];
      slide.replaceAllText('<<NAME>>', dataName);

      const file = DriveApp.getFileById(cpSlideId);
      file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      const baseURL = 'https://docs.google.com/presentation/d/';

      sheet.getRange(i + 1, dataCertidIndex + 1).setValue(cpSlideId);
      sheet.getRange(i + 1, dataCertURLIndex + 1).setValue(baseURL + cpSlideId);
      sheet.getRange(i + 1, dataStatusIndex + 1).setValue('CREATED');
      SpreadsheetApp.flush();
    }
  }
}

function sendCertificates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const dataNameIndex = headers.indexOf('Name');
  const dataEmailIndex = headers.indexOf('Email');
  const dataCertidIndex = headers.indexOf('File ID');
  const dataStatusIndex = headers.indexOf('Status');

  const emailTemp = HtmlService.createTemplateFromFile("email");
  const mailSubject = 'DevFest Bandung 2023 | E-Certificate and Event Wrap Up';
  const senderName = 'GDG Bandung';

  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    const dataName = rowData[dataNameIndex];
    const dataEmail = rowData[dataEmailIndex];
    const dataCertid = rowData[dataCertidIndex];
    const dataStatus = rowData[dataStatusIndex];
    const attachment = DriveApp.getFileById(dataCertid);

    if (dataStatus == 'CREATED') {
      ss.toast(`Sending email for ${dataName}`, APP_TITLE, 3);
      emailTemp.NAME = dataName;

      let htmlMailBody = emailTemp.evaluate().getContent();
      GmailApp.sendEmail(
        dataEmail,
        mailSubject,
        "Your email doesn't support HTML.",
        {
          name: senderName,
          htmlBody: htmlMailBody,
          attachments: [attachment.getAs(MimeType.PDF)]
        }
      )

      sheet.getRange(i + 1, dataStatusIndex + 1).setValue('SENT');
      SpreadsheetApp.flush();
    }
  }
}

function slideToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const dataNameIndex = headers.indexOf('Cert Name');
  const dataUseridIndex = headers.indexOf('User ID');
  const dataCertidIndex = headers.indexOf('File ID');
  const dataStatusIndex = headers.indexOf('Status');
  const dataPDFIndex = headers.indexOf('PDF');

  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    const dataName = rowData[dataNameIndex];
    const dataFileID = rowData[dataCertidIndex];
    const dataUserid = rowData[dataUseridIndex];
    const dataStatus = rowData[dataStatusIndex];
    
    if (dataStatus == 'CREATED' || dataStatus == 'SENT') {
      ss.toast(`Generate PDF archive for ${dataName}`, APP_TITLE, 3);

      const blob = DriveApp.getFileById(dataFileID).getBlob();
      const file = archiveFolder.createFile(blob);
      file.setName(dataName + '-' + dataUserid + '.pdf');

      file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      const PDFURL = 'https://drive.google.com/file/d/' + file.getId();

      sheet.getRange(i + 1, dataPDFIndex + 1).setValue(PDFURL);
      sheet.getRange(i + 1, dataStatusIndex + 1).setValue('ARCHIVED');
      SpreadsheetApp.flush();
    }
  }
}