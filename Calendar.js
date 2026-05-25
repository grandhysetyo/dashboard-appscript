function safeGetEvent(
  calendar,
  eventId
) {

  try {

    // =========================
    // VALIDATE ID
    // =========================

    if (
      !eventId ||
      eventId.toString().trim() === ""
    ) {

      return null;
    }

    const cleanId =
      eventId
        .toString()
        .trim();

    // =========================
    // GET EVENT
    // =========================

    let event =
      calendar.getEventById(
        cleanId
      );

    // =========================
    // TRY @google.com
    // =========================

    if (
      !event &&
      !cleanId.includes("@google.com")
    ) {

      event =
        calendar.getEventById(
          cleanId + "@google.com"
        );
    }

    // =========================
    // EVENT NOT FOUND
    // =========================

    if (!event) {

      return null;
    }

    // =========================
    // VALIDATE REAL EVENT
    // =========================
    // IMPORTANT:
    // Google Apps Script can
    // return ghost/stub object
    // even if event deleted
    // =========================

    try {

      const title =
        event.getTitle();

      const startTime =
        event.getStartTime();

      // INVALID TITLE
      if (
        !title ||
        title.toString().trim() === ""
      ) {

        return null;
      }

      // INVALID START TIME
      if (
        !startTime ||
        isNaN(startTime.getTime())
      ) {

        return null;
      }

      return event;

    } catch (validationError) {

      writeLog(
        "INVALID_EVENT_STUB",
        validationError.toString()
      );

      return null;
    }

  } catch (err) {

    writeLog(
      "SAFE_GET_EVENT_ERROR",
      err.stack || err.toString()
    );

    return null;
  }
}

function recreateMissingCalendarEvent(
  rowData
) {
  
    Logger.log('START',rowData.jam1)
    const bookingData = {
      clientName:
        rowData.client,

      universitas:
        rowData.university,

      fakultas:
        rowData.faculty,

      whatsapp:
        rowData.whatsapp,

      instagram:
        rowData.instagram,

      paket:
        rowData.package,

      tanggal:
        rowData.tanggal,

      jam1:
        rowData.jam,

      lokasi:
        rowData.lokasi,

      notes:
        rowData.notes,

      summary:
        rowData.summary,

      photographer:
        rowData.photographer
    };

    writeLog(
      "RECREATE_EVENT_START",
      JSON.stringify(bookingData)
    );

    const eventIds =
      createCalendarEvents(
        bookingData
      );

    console.log("RECREATE_CREATED_IDS",JSON.stringify(events));
    // =========================
    // VALIDATION
    // =========================

    if (
      !eventIds ||
      !Array.isArray(eventIds) ||
      eventIds.length === 0
    ) {

      writeLog(
        "RECREATE_FAILED",
        rowData.bookingId
      );

      throw new Error(
        "Failed to recreate calendar event"
      );
    }

    // =========================
    // VERIFY EVENT EXISTS
    // =========================

    const calendar =
      CalendarApp.getCalendarById(
        CONFIG.CALENDAR_ID
      );

    eventIds.forEach(id => {

      const event =
        safeGetEvent(
          calendar,
          id
        );

      if (!event) {

        throw new Error(
          `Recreated event verification failed: ${id}`
        );
      }
    });

    // =========================
    // UPDATE DASHBOARD
    // =========================

    const sheet =
      getDashboardSheet();

    const data =
      sheet
        .getDataRange()
        .getValues();

    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const bookingId =
        data[i][
          CONFIG.COL.BOOKING_ID - 1
        ];

      if (
        bookingId ===
        rowData.bookingId
      ) {

        sheet.getRange(
          i + 1,
          CONFIG.COL.CALENDAR_ID
        )
        .setValue(
          eventIds.join(",")
        );

        break;
      }
    }

    writeLog(
      "RECREATE_SUCCESS",
      eventIds.join(",")
    );
    console.log(
      "RECREATE_RETURN",
      JSON.stringify(events)
    );
    return eventIds;
}

function createCalendarEvents(data) {

  const calendar = 
    CalendarApp.getCalendarById(
      CONFIG.CALENDAR_ID
    );

  const events = [];

  const duration =
    getPackageDuration(data.paket);

  const firstName =
    getFirstName(data.clientName);

  if (!firstName) {
    throw new Error(
      "Client name missing while creating calendar event"
    );
  }

  // =========================
  // 2 HOURS + 2 SESSION
  // =========================

  if (
    duration === 2 &&
    data.jam1 &&
    data.jam2
  ) {

    const event1 =
      createSingleEvent({
        calendar,
        title:
          `RAYS | Not Assign - ${firstName} - ${data.universitas}`,
        tanggal: data.tanggal,
        jam: data.jam1,
        duration: 1,
        summary: data.summary
      });

    const event2 =
      createSingleEvent({
        calendar,
        title:
          `RAYS | Not Assign - ${firstName} - ${data.universitas}`,
        tanggal: data.tanggal,
        jam: data.jam2,
        duration: 1,
        summary: data.summary
      });

    events.push(event1);
    events.push(event2);

    if (!event1 || !event2) {
      throw new Error(
        "Event creation failed"
      );
    }

    return events;
  }

  // =========================
  // NORMAL EVENT
  // =========================
  const event =
    createSingleEvent({
      calendar,
      title:
        `RAYS | Not Assign - ${firstName} - ${data.universitas}`,
      tanggal: data.tanggal,
      jam: data.jam1,
      duration,
      summary: data.summary
    });

  if (!event) {
    throw new Error(
      "Event creation failed"
    );
  }

  events.push(event);
  return events;
}

function createSingleEvent(data) {
    const startTime =
      combineDateTime(
        data.tanggal,
        data.jam
      );

    const endTime =
      new Date(
        startTime.getTime() +
        data.duration * 60 * 60 * 1000
      );

    const event =
      data.calendar.createEvent(
        data.title,
        startTime,
        endTime,
        {
          description: data.summary
        }
      );

    event.setColor(
      CalendarApp.EventColor.RED
    );

    if (!event) {
      throw new Error(
        "Calendar createEvent returned null"
      );
    }
  
    const eventId =
      event.getId();

    writeLog(
      "CREATE_SINGLE_EVENT_SUCCESS",
      eventId
    );
    return eventId;
}

function updateCalendarEvent(
  rowData,
  sheet,
  row
) {

  try {

    const calendar =
      CalendarApp.getCalendarById(
        CONFIG.CALENDAR_ID
      );

    // =========================
    // VALIDATE REQUIRED DATA
    // =========================

    if (!rowData.tanggal) {

      throw new Error(
        "Tanggal kosong"
      );
    }

    if (!rowData.jam1) {

      throw new Error(
        "Jam 1 kosong"
      );
    }

    if (!rowData.package) {

      throw new Error(
        "Package kosong"
      );
    }

    // =========================
    // DETERMINE EVENT COUNT
    // =========================

    const duration =
      getPackageDuration(
        rowData.package
      );

    const expectedEventCount =
      duration === 2
        ? 2
        : 1;

    // =========================
    // LOAD EXISTING EVENTS
    // =========================

    const existingEventIds = [

      rowData.calendarId1,

      rowData.calendarId2

    ].filter(Boolean);

    const existingEvents = [];

    let hasMissingEvent = false;

    existingEventIds.forEach(id => {

      const event =
        safeGetEvent(
          calendar,
          id
        );

      if (event) {

        existingEvents.push(event);

      } else {

        hasMissingEvent = true;
      }
    });

    // =========================
    // CHECK EVENT COUNT
    // =========================

    if (
      existingEvents.length !==
      expectedEventCount
    ) {

      hasMissingEvent = true;
    }

    // =========================
    // DETECT MAJOR CHANGE
    // =========================

    const currentDate =
      formatDate(
        rowData.tanggal
      );

    const currentJam1 =
      formatTime24(
        rowData.jam1
      );

    const currentJam2 =
      formatTime24(
        rowData.jam2
      );

    const storedDate =
      sheet.getRange(
        row,
        CONFIG.COL.DATE
      ).getDisplayValue();

    const storedJam1 =
      sheet.getRange(
        row,
        CONFIG.COL.TIME_1
      ).getDisplayValue();

    const storedJam2 =
      sheet.getRange(
        row,
        CONFIG.COL.TIME_2
      ).getDisplayValue();

    const storedPackage =
      sheet.getRange(
        row,
        CONFIG.COL.PACKAGE
      ).getDisplayValue();

    const needRecreate =

      hasMissingEvent

      ||

      currentDate !== storedDate

      ||

      currentJam1 !== storedJam1

      ||

      currentJam2 !== storedJam2

      ||

      rowData.package !==
      storedPackage;

    // =========================
    // PATH A
    // FULL REBUILD
    // =========================

    if (needRecreate) {

      writeLog(
        "REBUILD_EVENT",
        rowData.bookingId
      );

      // DELETE OLD EVENTS
      existingEvents.forEach(event => {

        try {

          event.deleteEvent();

        } catch (err) {

          writeLog(
            "DELETE_SKIP",
            err.toString()
          );
        }
      });

      // CREATE NEW EVENTS
      const newEventIds =
        createCalendarEvents({

          clientName:
            rowData.client,

          universitas:
            rowData.university,

          tanggal:
            rowData.tanggal,

          jam1:
            rowData.jam1,

          jam2:
            rowData.jam2,

          paket:
            rowData.package,

          summary:
            rowData.summary
        });

      if (
        !newEventIds ||
        newEventIds.length === 0
      ) {

        throw new Error(
          "Failed creating calendar events"
        );
      }

      // RELOAD NEW EVENTS
      const rebuiltEvents = [];

      newEventIds.forEach(id => {

        const event =
          safeGetEvent(
            calendar,
            id
          );

        if (!event) {

          throw new Error(
            `Created event not found: ${id}`
          );
        }

        rebuiltEvents.push(event);
      });

      // UPDATE EVENT CONTENT
      rebuiltEvents.forEach((event, index) => {

        updateSingleCalendarEvent({

          event,

          rowData,

          index
        });
      });

      // SAVE IDS
      sheet.getRange(
        row,
        CONFIG.COL.CALENDAR_ID_1
      ).setValue(
        newEventIds[0] || ""
      );

      sheet.getRange(
        row,
        CONFIG.COL.CALENDAR_ID_2
      ).setValue(
        newEventIds[1] || ""
      );

    }

    // =========================
    // PATH B
    // NORMAL UPDATE
    // =========================

    else {

      writeLog(
        "NORMAL_UPDATE",
        rowData.bookingId
      );

      existingEvents.forEach((event, index) => {

        updateSingleCalendarEvent({

          event,

          rowData,

          index
        });
      });
    }

    // =========================
    // UPDATE STATUS
    // =========================

    sheet.getRange(
      row,
      CONFIG.COL.ACTION_CALENDAR
    ).setValue(
      "Added"
    );

    sheet.getRange(
      row,
      CONFIG.COL.PROJECT_STATUS
    ).setValue(
      "Assign"
    );

    sheet.getRange(
      row,
      CONFIG.COL.LAST_UPDATE
    ).setValue(
      new Date()
    );

    setSystemMessage(

      sheet,

      row,

      "✅ Google Calendar updated",

      "#D9EAD3"
    );

    // =========================
    // SUCCESS LOG
    // =========================

    writeLog(
      "UPDATE_GCAL_SUCCESS",
      rowData.bookingId
    );

    return true;

  } catch (err) {

    writeLog(
      "UPDATE_GCAL_ERROR",
      err.stack || err.toString()
    );

    setSystemMessage(

      sheet,

      row,

      `❌ ${err.toString()}`,

      "#F4CCCC"
    );

    throw err;
  }
}

function updateSingleCalendarEvent({
  event,
  rowData,
  index
}) {

  const photographer =
    getFirstName(
      rowData.photographer
    );

  const client =
    getFirstName(
      rowData.client
    );

  const title =
    `RAYS | ${photographer} - ${client} - ${rowData.university}`;

  event.setTitle(title);

  event.setDescription(
    rowData.summary
  );

  // =========================
  // TIME
  // =========================

  const duration =
    getPackageDuration(
      rowData.package
    );

  const eventTime =
    index === 0
      ? rowData.jam1
      : rowData.jam2;

  const startTime =
    combineDateTime(
      rowData.tanggal,
      eventTime
    );

  const endTime =
    new Date(
      startTime.getTime()
      +
      duration * 60 * 60 * 1000
    );

  event.setTime(
    startTime,
    endTime
  );

  // =========================
  // COLOR
  // =========================

  const color =
    getEventColor(
      rowData.package
    );

  event.setColor(
    Number(color)
  );

  // =========================
  // GUEST
  // =========================

  if (
    rowData.photographerEmail
  ) {

    try {

      event.addGuest(
        rowData.photographerEmail
      );

    } catch (err) {

      writeLog(
        "ADD_GUEST_SKIP",
        err.toString()
      );
    }
  }
}

function completeCalendarEvent(eventId) {

  const calendar = CalendarApp.getCalendarById(
    CONFIG.CALENDAR_ID
  );

  eventId
  .toString()
  .split(",")
  .forEach(id => {

    const event =
      safeGetEvent(
        calendar,
        id
      );

    if (!event) return;

    const title =
      event.getTitle();

    if (
      !title.startsWith("✅")
    ) {

      event.setTitle(
        "✅ " + title
      );
    }
  });

  if (!event) return;

  const title = event.getTitle();

  if (!title.startsWith("✅")) {

    event.setTitle(
      "✅ " + title
    );
  }
}


function testCreateEvent() {
  const events =
    createCalendarEvents({
      clientName:
        "GR CLIENT",
      universitas:
        "GR UNIVERSITY",
      tanggal:
        new Date(),
      jam1:
        "08:00",
      jam2:
        "",
      paket:
        "Personal Packages Silver - 1 Hours",

      summary:
        "TEST SUMMARY"
    });

  console.log(events);
}

function testRecreateFlow() {

  const rowData = {

    bookingId:
      "TEST-RECREATE",

    calendarId1:
      "gp44dg53a9f5g045761vlldd0c@google.com",

    calendarId2:
      "",

    client:
      "TEST CLIENT",

    university:
      "TEST UNIVERSITY",

    tanggal:
      new Date(),

    jam1:
      "08:00",

    jam2:
      "",

    lokasi:
      "TEST LOCATION",

    package:
      "Personal Packages Silver - 1 Hours",

    photographer:
      "Grandhys",

    photographerEmail:
      "grandhysetyo@gmail.com",

    summary:
      "TEST SUMMARY"
  };

  updateCalendarEvent(
    rowData
  );
}

function testRecreateMissing() {

  try {

    const rowData = {

      bookingId:
        "TEST-RECREATE",

      calendarId1:
        "gp44dg53a9f5g045761vlldd0c@google.com",

      calendarId2:
        "",

      client:
        "TEST CLIENT",

      university:
        "TEST UNIVERSITY",

      tanggal:
        new Date(),

      jam1:
        "08:00",

      jam2:
        "",

      lokasi:
        "TEST LOCATION",

      package:
        "Personal Packages Silver - 1 Hours",

      photographer:
        "Grandhys",

      photographerEmail:
        "grandhysetyo@gmail.com",

      summary:
        "TEST SUMMARY"
    };

    const result =
      recreateMissingCalendarEvent(
        rowData
      );

    writeLog.log(result);

  } catch (err) {

    writeLog.log(
      err.stack || err.toString()
    );

    throw err;
  }
}
