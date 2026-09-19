function doPost(e) {
  let data = {};

  try {
    data = JSON.parse(e.postData.contents || '{}');
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Invalid JSON payload' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Registrations') || spreadsheet.insertSheet('Registrations');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Batch', 'Attendance', 'Note']);
  }

  sheet.appendRow([
    new Date(),
    data.name || '',
    data.phone || '',
    data.batch || '',
    data.status || '',
    data.note || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Registration saved successfully' }))
    .setMimeType(ContentService.MimeType.JSON);
}