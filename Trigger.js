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
      .getRange(row, 1, 1, 26)
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
      videographerphotographer:
        data[
          CONFIG.COL.PHOTOGRAPHER_VIDEOGRAPHER - 1
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
    // PHOTOGRAPHER_VIDEOGRAPHER COLOR
    // =========================

    if (
      col === CONFIG.COL.PHOTOGRAPHER_VIDEOGRAPHER
    ) {

      const cell =
        sheet.getRange(
          row,
          CONFIG.COL.PHOTOGRAPHER_VIDEOGRAPHER
        );

      if (!e.value) {

        cell
          .setBackground("#FFFFFF")
          .setFontColor("#000000");

        return;
      }

      const photographer_videogrpaher =
        getPhotographerByName(
          e.value
        );

      if (photographer_videogrpaher) {

        const hex =
        photographer_videogrpaher.hex || "#FFFFFF";

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
        const photographer =
          getPhotographerByName(
            rowData.photographer
          );

        const videographer =
          getPhotographerByName(
            rowData.videographerphotographer
          );
        
        rowData.photographerEmail =
          photographer?.email || "";
        
        rowData.videographerEmail =
          videographer?.email || "";

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

      insertToEditorSheet({
        bookingId:
          rowData.bookingId,
        client:
          rowData.client,
        university:
          rowData.university,
        package:
          rowData.package,
        deadline:
          calculateDeadline()
      });

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
    // Revisi
    // =========================

    if (
      col === CONFIG.COL.PROJECT_STATUS &&
      e.value === "Revisi"
    ){

      updateEditorStatus(
        rowData.bookingId,
        "Revisi"
      );
      setSystemMessage(
        sheet,
        row,
        "✅ Project Revisi",
        "#D9EAD3"
      );

      writeLog(
        "REVISI",
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

      updateEditorStatus(
        rowData.bookingId,
        "Done"
      );

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
      "Refresh Status Project Dropdown",
      "refreshProjectDropdown"
    )

    .addItem(
      "Refresh Editor Dropdown",
      "refreshEditorStatusDropdown"
    )

    .addSeparator()

    // =========================
    // Sync
    // =========================

    .addItem(
      "Sync Bookings",
      "syncFutureBookings"
    )

    // =========================
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