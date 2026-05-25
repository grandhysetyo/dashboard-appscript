function getPhotographers() {

  const ss = SpreadsheetApp.openById(
    CONFIG.PHOTOGRAPHER_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.PHOTOGRAPHER_SHEET_NAME
  );

  const data = sheet.getDataRange().getValues();

  data.shift();

  return data.map(row => ({

    nama: row[2],
    kota: row[3],
    whatsapp: row[4],
    email: row[5],
    instagram: row[6],
    bank: row[7],
    rates: row[8],
    hex: row[9]
  }));
}

function getPhotographerByName(name) {

  const photographers = getPhotographers();

  return photographers.find(p =>
    p.nama === name
  );
}

function setupPhotographerDropdown() {

  const ss = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );

  const photographers = getPhotographers();

  const names = photographers.map(
    p => p.nama
  );

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(names)
    .build();

  sheet
    .getRange("O2:O1000")
    .setDataValidation(rule);
}

function refreshPhotographerDropdown() {

  const dashboardSS = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const dashboardSheet = dashboardSS.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );

  const photographers = getPhotographers();

  const names = photographers
    .map(p => p.nama)
    .filter(name => name);

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(false)
    .build();

  const lastRow = Math.max(
    dashboardSheet.getLastRow(),
    1000
  );

  dashboardSheet
    .getRange(
      2,
      CONFIG.COL.PHOTOGRAPHER,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "REFRESH_DROPDOWN",
    `${names.length} photographers loaded`
  );
}