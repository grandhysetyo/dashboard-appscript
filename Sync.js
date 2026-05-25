function syncFutureBookings() {

  const formSS =
    SpreadsheetApp.openById(
      CONFIG.FORM_SPREADSHEET_ID
    );

  const formSheet =
    formSS.getSheetByName(
      CONFIG.FORM_SHEET_NAME
    );

  const formData =
    formSheet
      .getDataRange()
      .getDisplayValues();

  const dashboardSS =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const dashboardSheet =
    dashboardSS.getSheetByName(
      CONFIG.DASHBOARD_SHEET_NAME
    );

  const dashboardData =
    dashboardSheet
      .getDataRange()
      .getDisplayValues();

  // REMOVE HEADER
  formData.shift();

  const today =
    new Date();

  today.setHours(
    0, 0, 0, 0
  );

  formData.forEach((row, index) => {

    const formRowId =
      index + 2;

    // =========================
    // FORM DATE
    // =========================

    const tanggal =
      row[8];

    if (!tanggal) {
      return;
    }

    // =========================
    // PARSE DATE
    // MM/dd/yyyy
    // =========================

    const parsedDate =
      combineDateTime(
        tanggal,
        "00:00"
      );

    // INVALID DATE
    if (
      isNaN(
        parsedDate.getTime()
      )
    ) {

      writeLog(
        "INVALID_DATE",
        tanggal
      );

      return;
    }

    parsedDate.setHours(
      0, 0, 0, 0
    );

    // =========================
    // SKIP PAST BOOKINGS
    // =========================

    if (
      parsedDate < today
    ) {

      return;
    }

    // =========================
    // FIND DASHBOARD ROW
    // =========================

    const dashboardRow =
      findDashboardRowByFormRowId(
        dashboardData,
        formRowId
      );

    // =========================
    // INSERT NEW BOOKING
    // =========================

    if (!dashboardRow) {

      writeLog(
        "NEW_BOOKING_FOUND",
        formRowId
      );

      const bookingData = {

        timestamp:
          row[0],

        formRowId,

        calendarId1:
          "",

        calendarId2:
          "",

        bookingId:
          generateBookingId(),

        nama:
          row[2],

        universitas:
          row[3],

        fakultas:
          row[4],

        tanggal:
          row[8],

        jam1:
          row[9],

        jam2:
          row[10],

        lokasi:
          row[13],

        paket:
          row[7],

        instagram:
          row[6],

        whatsapp:
          row[5],

        notes:
          row[15],

        summary:
          generateSummary({

            bookingId:
              generateBookingId(),

            nama:
              row[2],

            universitas:
              row[3],

            tanggal:
              row[8],

            jam1:
              row[9],

            jam2:
              row[10],

            lokasi:
              row[13],

            paket:
              row[7],

            instagram:
              row[6],

            whatsapp:
              row[5],

            notes:
              row[15]
          })
      };

      insertBookingToDashboard(
        bookingData
      );

      return;
    }

    // =========================
    // EXISTING DATA
    // =========================

    const jam1 =
      row[9];

    const jam2 =
      row[10];

    const lokasi =
      row[13];

    const notes =
      row[15];

    const dashboardTanggal =
      dashboardRow.data[
        CONFIG.COL.DATE - 1
      ];

    const dashboardJam1 =
      dashboardRow.data[
        CONFIG.COL.TIME_1 - 1
      ];

    const dashboardJam2 =
      dashboardRow.data[
        CONFIG.COL.TIME_2 - 1
      ];

    const dashboardLokasi =
      dashboardRow.data[
        CONFIG.COL.LOCATION - 1
      ];

    const dashboardNotes =
      dashboardRow.data[
        CONFIG.COL.NOTES - 1
      ];

    // =========================
    // NORMALIZE
    // =========================

    const formDate =
      formatDate(tanggal);

    const dashDate =
      formatDate(dashboardTanggal);

    const formJam1 =
      formatTime24(jam1);

    const dashJam1 =
      formatTime24(dashboardJam1);

    const formJam2 =
      formatTime24(jam2);

    const dashJam2 =
      formatTime24(dashboardJam2);

    // =========================
    // DETECT CHANGE
    // =========================

    const changed =

      formDate !== dashDate

      ||

      formJam1 !== dashJam1

      ||

      formJam2 !== dashJam2

      ||

      (lokasi || "").trim()
      !==
      (dashboardLokasi || "").trim()

      ||

      (notes || "").trim()
      !==
      (dashboardNotes || "").trim();

    if (!changed) {

      return;
    }

    // =========================
    // UPDATE DASHBOARD
    // =========================

    writeLog(
      "BOOKING_UPDATED",
      formRowId
    );

    updateDashboardBooking(

      dashboardSheet,

      dashboardRow.row,

      row
    );
  });

  SpreadsheetApp.getActive()
    .toast(

      "Future booking sync completed",

      "RAYS SYSTEM",

      5
    );
}

function refreshDashboardDatesOnly() {

  const formSS =
    SpreadsheetApp.openById(
      CONFIG.FORM_SPREADSHEET_ID
    );

  const formSheet =
    formSS.getSheetByName(
      CONFIG.FORM_SHEET_NAME
    );

  const formData =
    formSheet
      .getDataRange()
      .getDisplayValues();

  const dashboardSS =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const dashboardSheet =
    dashboardSS.getSheetByName(
      CONFIG.DASHBOARD_SHEET_NAME
    );

  const dashboardData =
    dashboardSheet
      .getDataRange()
      .getValues();

  // REMOVE HEADER
  formData.shift();

  const today =
    new Date();

  formData.forEach((formRow, index) => {

    const formRowId =
      index + 2;

    // =========================
    // FORM DATE
    // =========================

    const formTanggal =
      formRow[8];

    if (!formTanggal) {
      return;
    }

    // =========================
    // PARSE FORM DATE
    // FORM = MM/dd/yyyy
    // =========================

    const parts =
      formTanggal
        .toString()
        .split("/");

    const month =
      Number(parts[0]) - 1;

    const day =
      Number(parts[1]);

    const year =
      Number(parts[2]);

    const parsedDate =
      new Date(
        year,
        month,
        day
      );

    // INVALID DATE
    if (
      isNaN(
        parsedDate.getTime()
      )
    ) {

      writeLog(
        "INVALID_FORM_DATE",
        formTanggal
      );

      return;
    }

    // ONLY TODAY & FUTURE
    if (
      parsedDate < today
    ) {

      return;
    }

    // =========================
    // FIND DASHBOARD ROW
    // =========================

    const dashboardRow =
      findDashboardRowByFormRowId(
        dashboardData,
        formRowId
      );

    if (!dashboardRow) {

      writeLog(
        "FORM_ROW_NOT_FOUND",
        formRowId
      );

      return;
    }

    // =========================
    // NORMALIZED DATE
    // =========================

    const normalized =
      Utilities.formatDate(

        parsedDate,

        Session.getScriptTimeZone(),

        "MM/dd/yyyy"
      );

    dashboardSheet.getRange(

      dashboardRow.row,

      CONFIG.COL.DATE

    ).setValue(
      normalized
    );

    writeLog(

      "DATE_REFRESHED",

      `Row ${dashboardRow.row}: ${normalized}`
    );
  });

  SpreadsheetApp.getActive()
    .toast(

      "Dashboard dates refreshed",

      "RAYS SYSTEM",

      5
    );
}

function refreshFutureBookingsFromForm() {

  const formSS =
    SpreadsheetApp.openById(
      CONFIG.FORM_SPREADSHEET_ID
    );

  const formSheet =
    formSS.getSheetByName(
      CONFIG.FORM_SHEET_NAME
    );

  const formData =
    formSheet
      .getDataRange()
      .getDisplayValues();

  const dashboardSS =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  const dashboardSheet =
    dashboardSS.getSheetByName(
      CONFIG.DASHBOARD_SHEET_NAME
    );

  const dashboardData =
    dashboardSheet
      .getDataRange()
      .getValues();

  // REMOVE HEADER
  formData.shift();

  const today =
    new Date();

  today.setHours(
    0, 0, 0, 0
  );

  formData.forEach((row, index) => {

    const formRowId =
      index + 2;

    // =========================
    // PARSE DATE
    // FORM = MM/dd/yyyy
    // =========================

    const formTanggal =
      row[8];

    if (!formTanggal) {
      return;
    }

    const parts =
      formTanggal
        .toString()
        .split("/");

    const month =
      Number(parts[0]) - 1;

    const day =
      Number(parts[1]);

    const year =
      Number(parts[2]);

    const parsedDate =
      new Date(
        year,
        month,
        day
      );

    // INVALID DATE
    if (
      isNaN(
        parsedDate.getTime()
      )
    ) {

      writeLog(
        "INVALID_DATE",
        formTanggal
      );

      return;
    }

    parsedDate.setHours(
      0, 0, 0, 0
    );

    // SKIP PAST BOOKINGS
    if (
      parsedDate < today
    ) {

      return;
    }

    // =========================
    // FIND DASHBOARD ROW
    // =========================

    const dashboardRow =
      findDashboardRowByFormRowId(
        dashboardData,
        formRowId
      );

    if (!dashboardRow) {

      writeLog(
        "ROW_NOT_FOUND",
        formRowId
      );

      return;
    }

    const dashboardRowNumber =
      dashboardRow.row;

    // =========================
    // FORM VALUES
    // =========================

    const nama =
      row[2];

    const universitas =
      row[3];

    const fakultas =
      row[4];

    const whatsapp =
      row[5];

    const instagram =
      row[6];

    const paket =
      row[7];

    const jam1 =
      row[9];

    const jam2 =
      row[10];

    const lokasi =
      row[13];

    const notes =
      row[15];

    // =========================
    // BOOKING ID
    // =========================

    const bookingId =
      dashboardSheet
        .getRange(

          dashboardRowNumber,

          CONFIG.COL.BOOKING_ID

        )
        .getValue();

    // =========================
    // SUMMARY
    // =========================

    const summary =
      generateSummary({

        bookingId,

        nama,

        universitas,

        tanggal:
          formTanggal,

        jam1,

        jam2,

        lokasi,

        paket,

        instagram,

        whatsapp,

        notes
      });

    // =========================
    // UPDATE DASHBOARD
    // =========================

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.CLIENT_NAME

    ).setValue(
      nama
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.UNIVERSITY

    ).setValue(
      universitas
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.FACULTY

    ).setValue(
      fakultas
    );

   dashboardSheet.getRange(
      dashboardRowNumber,

      CONFIG.COL.DATE

    )
    .setNumberFormat("@")
    .setValue(
      "'" + formatDate(formTanggal)
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.TIME_1

    )
    .setNumberFormat("@")
    .setValue(
      "'" + formatTime24(jam1)
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.TIME_2

    )
    .setNumberFormat("@")
    .setValue(
      "'" + formatTime24(jam2)
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.LOCATION

    ).setValue(
      lokasi
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.PACKAGE

    ).setValue(
      paket
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.INSTAGRAM

    ).setValue(
      instagram
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.WHATSAPP

    ).setValue(
      whatsapp
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.NOTES

    ).setValue(
      notes
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.SUMMARY

    ).setValue(
      summary
    );

    dashboardSheet.getRange(

      dashboardRowNumber,

      CONFIG.COL.LAST_UPDATE

    ).setValue(
      new Date()
    );

    writeLog(
      "BOOKING_REFRESHED",
      bookingId
    );
  });

  SpreadsheetApp.getActive()
    .toast(

      "Future bookings refreshed",

      "RAYS SYSTEM",

      5
    );
}