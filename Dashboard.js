function insertBookingToDashboard(data) {

  const ss = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );
  console.log(
    JSON.stringify(data)
  );
  sheet.appendRow([
    data.timestamp,
    data.formRowId,
    data.calendarId1,
    data.calendarId2,
    data.bookingId,
    data.nama,
    data.universitas,
    data.fakultas,
    formatDate(data.tanggal),
    formatTime24(data.jam1),
    formatTime24(data.jam2),
    data.lokasi,
    data.paket,
    data.instagram,
    data.whatsapp,
    data.notes,
    data.summary,
    "", // photographer
    "", // photographer/videographer

    "", // action calendar

    CONFIG.PAYMENT_DEFAULT,
    CONFIG.PROJECT_DEFAULT,

    "", // deadline
    "", // lain lain

    new Date(),

    ""
  ]);

  const lastRow =
    sheet.getLastRow();

  // FORCE JAM 1 FORMAT
  sheet.getRange(
    lastRow,
    CONFIG.COL.TIME_1
  )
  .setNumberFormat("@")
  .setValue(
    "'" + formatTime24(data.jam1)
  );

  // FORCE JAM 2 FORMAT
  sheet.getRange(
    lastRow,
    CONFIG.COL.TIME_2
  )
  .setNumberFormat("@")
  .setValue(
    data.jam2
      ? "'" + formatTime24(data.jam2)
      : ""
  );
}

function findDashboardRowByCalendarId(
  calendarId
) {

  const ss = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (data[i][2] === calendarId) {
      return i + 1;
    }
  }

  return null;
}

function findDashboardRowByFormRowId(
  dashboardData,
  formRowId
) {

  for (
    let i = 1;
    i < dashboardData.length;
    i++
  ) {

    if (
      Number(dashboardData[i][1]) ===
      Number(formRowId)
    ) {

      return {
        row: i + 1,
        data: dashboardData[i]
      };
    }
  }

  return null;
}

function updateDashboardBooking(
  sheet,
  row,
  formData
) {

  const nama = formData[2];
  const universitas = formData[3];
  const fakultas = formData[4];
  const whatsapp = formData[5];
  const instagram = formData[6];
  const paket = formData[7];
  const tanggal = formData[8];
  const jam1 = formData[9];
  const jam2 = formData[10];
  const lokasi = formData[13];
  const notes = formData[15];

  const bookingId =
    sheet.getRange(
      row,
      CONFIG.COL.BOOKING_ID
    ).getValue();

  const summary =
    generateSummary({
      bookingId,
      nama,
      universitas,
      tanggal,
      jam1,
      jam2,
      lokasi,
      paket,
      instagram,
      whatsapp,
      notes
    });

  // UPDATE DASHBOARD
  sheet.getRange(
    row,
    CONFIG.COL.DATE
  ).setValue(
    formatDate(tanggal)
  );

  const formattedTime1 =
  "'" + formatTime24(jam1);

  sheet.getRange(
    row,
    CONFIG.COL.TIME_1
  )
  .setNumberFormat("@")
  .setValue(formattedTime1);

  const formattedTime2 =
  "'" + formatTime24(jam2);

  sheet.getRange(
    row,
    CONFIG.COL.TIME_2
  )
  .setNumberFormat("@")
  .setValue(formattedTime2);

  sheet.getRange(
    row,
    CONFIG.COL.LOCATION
  ).setValue(lokasi);

  sheet.getRange(
    row,
    CONFIG.COL.NOTES
  ).setValue(notes);

  sheet.getRange(
    row,
    CONFIG.COL.SUMMARY
  ).setValue(summary);

  // STATUS
  sheet.getRange(
    row,
    CONFIG.COL.PROJECT_STATUS
  ).setValue(
    "Assign - Rescheduled"
  );

  // ACTION
  sheet.getRange(
    row,
    CONFIG.COL.ACTION_CALENDAR
  ).setValue(
    "Update G-Cal"
  );

  // SYSTEM MESSAGE
  setSystemMessage(
    sheet,
    row,
    "⚠ Booking updated from Google Form",
    "#FFF2CC"
  );

  writeLog(
    "SYNC_UPDATE",
    bookingId
  );
}

function isFormRowIdExists(
  formRowId
) {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.DASHBOARD_SHEET_NAME
    );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const values =
    sheet
      .getRange(
        2,
        CONFIG.COL.FORM_ROW_ID,
        lastRow - 1,
        1
      )
      .getValues();

  return values.some(
    row =>
      Number(row[0]) ===
      Number(formRowId)
  );
}

function testInsertBooking() {

  insertBookingToDashboard({

    timestamp:
      new Date(),

    formRowId:
      9999,

    calendarId1:
      "TEST_EVENT_1",

    calendarId2:
      "TEST_EVENT_2",

    bookingId:
      "TEST-BOOKING",

    nama:
      "TEST CLIENT",

    universitas:
      "TEST UNIVERSITY",

    fakultas:
      "TEST FACULTY",

    tanggal:
      new Date(),

    jam1:
      "08:00",

    jam2:
      "15:00",

    lokasi:
      "TEST LOCATION",

    paket:
      "Personal Packages Silver - 2 Hours",

    instagram:
      "testig",

    whatsapp:
      "08123456789",

    notes:
      "TEST NOTES",

    summary:
      "TEST SUMMARY"
  });

  Logger.log(
    "TEST INSERT SUCCESS"
  );
}
