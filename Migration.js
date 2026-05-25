function importExisting2026Bookings() {

  const formSS = SpreadsheetApp.openById(
    CONFIG.FORM_SPREADSHEET_ID
  );

  const formSheet = formSS.getSheetByName(
    CONFIG.FORM_SHEET_NAME
  );

  const data = formSheet
    .getDataRange()
    .getDisplayValues();

  // REMOVE HEADER
  data.shift();

  let imported = 0;

  data.forEach((row, index) => {
    const formRowId = index + 2;
    const tanggal = row[8];

    // ONLY 2026
    if (!isBooking2026(tanggal)) {
      return;
    }

    try {

      processBooking(row, formRowId);

      imported++;

    } catch (err) {

      writeLog(
        "IMPORT_ERROR",
        err.toString()
      );
    }
  });

  Logger.log(
    `${imported} bookings imported`
  );
}