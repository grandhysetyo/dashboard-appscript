function createTriggers() {

  const formSS = SpreadsheetApp.openById(
    CONFIG.FORM_SPREADSHEET_ID
  );

  // FORM SUBMIT
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(formSS)
    .onFormSubmit()
    .create();

  // DASHBOARD EDIT
  ScriptApp.newTrigger("onDashboardEdit")
    .forSpreadsheet(
      SpreadsheetApp.openById(
        CONFIG.DASHBOARD_SPREADSHEET_ID
      )
    )
    .onEdit()
    .create();

    setupAllDropdowns();
}

function testColor() {

  Logger.log(
    getEventColor(
      "Personal Packages Diamond - 2 Hours"
    )
  );
}