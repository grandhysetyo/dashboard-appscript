function onFormSubmit(e) {

  const formRowId =
    e.range.getRow();

  const data =
    e.values;

  // CHECK EXISTING
  const exists =
    isFormRowIdExists(
      formRowId
    );

  // IF EXISTS
  // MEANS EDIT RESPONSE
  if (exists) {

    writeLog(
      "SKIPPED_SUBMIT",
      `Form Row ID ${formRowId} already exists`
    );

    return;
  }

  processBooking(
    data,
    formRowId
  );
}

function onDashboardEdit(e) {

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    const sheet =
      e.range.getSheet();

    if (
      sheet.getName() !==
      CONFIG.DASHBOARD_SHEET_NAME
    ) return;

    const row =
      e.range.getRow();

    const col =
      e.range.getColumn();

    if (row < 2) return;

    const data = sheet
      .getRange(row, 1, 1, 24)
      .getValues()[0];

    const rowData = {
      formRowId:
        data[
          CONFIG.COL.FORM_ROW_ID - 1
        ],
      calendarId1:
        data[
          CONFIG.COL.CALENDAR_ID_1 - 1
        ],
      calendarId2:
        data[
          CONFIG.COL.CALENDAR_ID_2 - 1
        ],
      bookingId:
        data[
          CONFIG.COL.BOOKING_ID - 1
        ],
      client:
        data[
          CONFIG.COL.CLIENT_NAME - 1
        ],
      university:
        data[
          CONFIG.COL.UNIVERSITY - 1
        ],
      faculty:
        data[
          CONFIG.COL.FACULTY - 1
        ],
      tanggal:
        data[
          CONFIG.COL.DATE - 1
        ],
      jam1:
        data[
          CONFIG.COL.TIME_1 - 1
        ],
      jam2:
        data[
          CONFIG.COL.TIME_2 - 1
        ],
      lokasi:
        data[
          CONFIG.COL.LOCATION - 1
        ],
      package:
        data[
          CONFIG.COL.PACKAGE - 1
        ],
      instagram:
        data[
          CONFIG.COL.INSTAGRAM - 1
        ],
      whatsapp:
        data[
          CONFIG.COL.WHATSAPP - 1
        ],
      notes:
        data[
          CONFIG.COL.NOTES - 1
        ],
      summary:
        data[
          CONFIG.COL.SUMMARY - 1
        ],
      photographer:
        data[
          CONFIG.COL.PHOTOGRAPHER - 1
        ],
      actionCalendar:
        data[
          CONFIG.COL.ACTION_CALENDAR - 1
        ],
      paymentStatus:
        data[
          CONFIG.COL.PAYMENT_STATUS - 1
        ],
      projectStatus:
        data[
          CONFIG.COL.PROJECT_STATUS - 1
        ]
    };

    // =========================
    // PHOTOGRAPHER COLOR
    // =========================

    if (
      col === CONFIG.COL.PHOTOGRAPHER
    ) {

      const cell =
        sheet.getRange(
          row,
          CONFIG.COL.PHOTOGRAPHER
        );

      if (!e.value) {

        cell
          .setBackground("#FFFFFF")
          .setFontColor("#000000");

        return;
      }

      // =========================
      // REGENERATE SUMMARY
      // =========================

      const newSummary =
        generateSummary({
          bookingId:
            rowData.bookingId,
          nama:
            rowData.client,
          universitas:
            rowData.university,
          tanggal:
            rowData.tanggal,
          jam1:
            rowData.jam1,
          jam2:
            rowData.jam2,
          lokasi:
            rowData.lokasi,
          paket:
            rowData.package,
          instagram:
            rowData.instagram,
          whatsapp:
            rowData.whatsapp,
          notes:
            rowData.notes
        });

      // UPDATE SUMMARY COLUMN
      sheet.getRange(
        row,
        CONFIG.COL.SUMMARY
      ).setValue(newSummary);

      // UPDATE OBJECT
      rowData.summary =
        newSummary;

      const photographer =
        getPhotographerByName(
          e.value
        );

      if (photographer) {

        const hex =
          photographer.hex || "#FFFFFF";

        cell.setBackground(hex);

        cell.setFontColor(
          getContrastColor(hex)
        );
      }
    }

    // =========================
    // UPDATE G-CAL
    // =========================

    if (
      col === CONFIG.COL.ACTION_CALENDAR &&
      e.value === "Update G-Cal"
    ) {

      try {

        // VALIDATE PHOTOGRAPHER
        if (!rowData.photographer) {

          setSystemMessage(
            sheet,
            row,
            "❌ Photographer empty",
            "#F4CCCC"
          );

          return;
        }

        const photographer =
          getPhotographerByName(
            rowData.photographer
          );

        rowData.photographerEmail =
          photographer?.email || "";

        // UPDATE EVENT
        const success =
          updateCalendarEvent(
            rowData,
            sheet,
            row
          );

        SpreadsheetApp.flush();

        // UPDATE STATUS
        if (success) {
          // STATUS PROJECT
          sheet.getRange(
            row,
            CONFIG.COL.PROJECT_STATUS
          ).setValue("Assign");

          // ACTION CALENDAR
          sheet.getRange(
            row,
            CONFIG.COL.ACTION_CALENDAR
          ).setValue("Added");

          // LAST UPDATE
          sheet.getRange(
            row,
            CONFIG.COL.LAST_UPDATE
          ).setValue(new Date());

          // SUCCESS MESSAGE
          setSystemMessage(
            sheet,
            row,
            "✅ Google Calendar Updated",
            "#D9EAD3"
          );

          SpreadsheetApp.getActive()
            .toast(
              `Booking ${rowData.bookingId} updated`,
              "RAYS SYSTEM",
              5
            );
        }        

        writeLog(
          "UPDATE_GCAL",
          rowData.bookingId
        );

      } catch (err) {

        setSystemMessage(
          sheet,
          row,
          `❌ ${err}`,
          "#F4CCCC"
        );

        writeLog(
          "UPDATE_GCAL_ERROR",
          err.toString()
        );
      }
    }

    // =========================
    // CHOOSEN
    // =========================

    if (
      col === CONFIG.COL.PROJECT_STATUS &&
      e.value === "Choosen"
    ) {

      sheet.getRange(
        row,
        CONFIG.COL.DEADLINE
      ).setValue(
        calculateDeadline()
      );

      sheet.getRange(
        row,
        CONFIG.COL.LAST_UPDATE
      ).setValue(new Date());

      setSystemMessage(
        sheet,
        row,
        "✅ Deadline Generated",
        "#D9EAD3"
      );

      writeLog(
        "CHOOSEN",
        rowData.bookingId
      );
    }

    // =========================
    // DONE
    // =========================

    if (
      col === CONFIG.COL.PROJECT_STATUS &&
      e.value === "Done"
    ) {

      completeCalendarEvent([
        rowData.calendarId1,
        rowData.calendarId2
      ]);

      sheet.getRange(
        row,
        CONFIG.COL.LAST_UPDATE
      ).setValue(new Date());

      setSystemMessage(
        sheet,
        row,
        "✅ Project Done",
        "#D9EAD3"
      );

      writeLog(
        "DONE",
        rowData.bookingId
      );
    }

  } catch (err) {

    writeLog(
      "ERROR",
      err.toString()
    );

  } finally {

    lock.releaseLock();
  }
}

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("RAYS SYSTEM")

    // =========================
    // DROPDOWNS
    // =========================

    .addItem(
      "Refresh Photographer Dropdown",
      "refreshPhotographerDropdown"
    )

    .addItem(
      "Refresh Action Dropdown",
      "refreshActionCalendarDropdown"
    )

    .addItem(
      "Refresh Payment Dropdown",
      "refreshPaymentDropdown"
    )

    .addItem(
      "Refresh Project Dropdown",
      "refreshProjectDropdown"
    )

    .addSeparator()

    .addItem(
      "Refresh All Dropdowns",
      "refreshAllDropdowns"
    )

    .addSeparator()

    // =========================
    // IMPORT
    // =========================

    .addItem(
      "Sync Bookings",
      "syncFutureBookings"
    )

    .addSeparator()

    // =========================
    // MAINTENANCE
    // =========================

    .addItem(
      "Delete All Rays Events",
      "deleteAllRaysEvents"
    )

    .addItem(
      "Clear Dashboard Data",
      "clearDashboardData"
    )

    .addToUi();
}

function createSyncTrigger() {

  ScriptApp.newTrigger(
    "syncUpdatedBookings"
  )
  .timeBased()
  .everyMinutes(15)
  .create();
}