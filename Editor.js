function insertToEditorSheet(rowData) {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.EDITOR_SHEET_NAME
    );

  const data =
    sheet.getDataRange()
      .getValues();

  const exists =
    data.some(row =>
      row[0] === rowData.bookingId
    );

  if (exists) {
    return;
  }

  sheet.appendRow([
    rowData.bookingId,
    rowData.client,
    rowData.university,
    rowData.package,
    rowData.deadline,
    "",
    "Waiting Edit",
    new Date()
  ]);

  writeLog(
    "EDITOR_CREATED",
    rowData.bookingId
  );
}

function updateEditorStatus(
  bookingId,
  status
) {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.EDITOR_SHEET_NAME
    );

  const data =
    sheet.getDataRange()
      .getValues();

  for (let i = 1; i < data.length; i++) {

    if (
      data[i][0] === bookingId
    ) {

      sheet.getRange(
        i + 1,
        CONFIG.EDITOR_COL.STATUS_EDITED
      ).setValue(status);

      sheet.getRange(
        i + 1,
        CONFIG.EDITOR_COL.LAST_UPDATE
      ).setValue(new Date());

      return;
    }
  }
}