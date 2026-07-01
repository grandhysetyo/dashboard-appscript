function syncFutureBookings() {

  const formSS =
    SpreadsheetApp.openById(
      CONFIG.FORM_SPREADSHEET_ID
    );

  const formSheet =
    formSS.getSheetByName(
      CONFIG.FORM_SHEET_NAME
    );

  const masterData =
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
  masterData.shift();

  const today =
    new Date();

  today.setHours(
    0, 0, 0, 0
  );

  masterData.forEach((row, index) => {

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

    const bookingId = generateBookingId();

    if (!dashboardRow) {

      writeLog(
        "NEW_BOOKING_FOUND",
        formRowId
      );

      const sessionDate =
        combineDateTime(
          row[8],
          "00:00"
        );
      
      const folder =
        createBookingFolder(
          bookingId,
          row[2], // NAMA
          row[3], // UNIVERSITAS
          sessionDate, // TANGGAL
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
          bookingId,
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
              bookingId,
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
              row[15],
            driveFolder:
              folder.url
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