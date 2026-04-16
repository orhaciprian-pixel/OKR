const SHEET_NAME = 'OKR_Data';

function doGet(e) {
  const quarter = (e && e.parameter && e.parameter.quarter) ? e.parameter.quarter : 'Q1-2026';
  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();
  const row = rows.find(r => r[0] === quarter);
  const payload = (row && row[1]) ? row[1] : JSON.stringify({ objectives: [] });

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const quarter = body.quarter;
  const payload = JSON.stringify(body.data);

  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === quarter) {
      sheet.getRange(i + 1, 2).setValue(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'updated' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow([quarter, payload]);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'created' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  let ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {
    ss = null;
  }
  if (!ss) {
    ss = SpreadsheetApp.create('OKR Personal');
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Quarter', 'Data']);
  }
  return sheet;
}
