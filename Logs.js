function writeLog(type, message) {

  const ss = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.LOG_SHEET_NAME
  );

  sheet.appendRow([
    new Date(),
    type,
    message
  ]);
}

function debugLog(message) {

  Logger.log(message);

  writeLog(
    "DEBUG",
    message
  );
}