function refreshPhotographerDropdown() {

  const sheet =
    getDashboardSheet();

  const lastRow =
    getDashboardLastRow(sheet);

  const photographers =
    getPhotographers()
      .map(p => p.nama)
      .filter(Boolean);

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        photographers,
        true
      )
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      CONFIG.COL.PHOTOGRAPHER,
      lastRow,
      1
    )
    .setDataValidation(rule);

    sheet
    .getRange(
      2,
      CONFIG.COL.PHOTOGRAPHER_VIDEOGRAPHER,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "DROPDOWN",
    "Photographer refreshed"
  );
}

function refreshActionCalendarDropdown() {

  const sheet =
    getDashboardSheet();

  const lastRow =
    getDashboardLastRow(sheet);

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList([
        "Update G-Cal",
        "Added"
      ])
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      CONFIG.COL.ACTION_CALENDAR,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "DROPDOWN",
    "Action Calendar refreshed"
  );
}

function refreshPaymentDropdown() {

  const sheet =
    getDashboardSheet();

  const lastRow =
    getDashboardLastRow(sheet);

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList([
        "DP",
        "FP"
      ])
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      CONFIG.COL.PAYMENT_STATUS,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "DROPDOWN",
    "Payment refreshed"
  );
}

function refreshProjectDropdown() {

  const sheet =
    getDashboardSheet();

  const lastRow =
    getDashboardLastRow(sheet);

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList([
        "Not Assign",
        "Assign",
        "Assign - Rescheduled",
        "Uploded",
        "Choosen",
        "Revisi",
        "Done"
      ])
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      CONFIG.COL.PROJECT_STATUS,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "DROPDOWN",
    "Project refreshed"
  );
}

function refreshEditorStatusDropdown() {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.EDITOR_SHEET_NAME
    );

  const lastRow = Math.max(
    sheet.getLastRow(),
    1000
  );

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList([
        "Waiting Edit",
        "Finish Edited",
        "Revisi",
        "Done"
      ], true)
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      CONFIG.EDITOR_COL.STATUS_EDITED,
      lastRow,
      1
    )
    .setDataValidation(rule);

  writeLog(
    "DROPDOWN",
    "Editor Status refreshed"
  );
}

function refreshAllDropdowns() {

  refreshPhotographerDropdown();

  refreshActionCalendarDropdown();

  refreshPaymentDropdown();

  refreshProjectDropdown();

  SpreadsheetApp.getActive()
    .toast(
      "All dropdowns refreshed",
      "RAYS SYSTEM",
      5
    );
}