function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificate')
    .addItem('Generate Certificates', 'generateCertificates')
    .addSeparator()
    .addItem('Send Certificates', 'sendCertificates')
    .addSeparator()
    .addItem('PDF Archive', 'slideToPDF')
    .addToUi();
}