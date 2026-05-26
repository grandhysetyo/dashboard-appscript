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

    const hasAssignedTeam =rowData.photographer || rowData.videographerphotographer;

    sheet.getRange(
      row,
      CONFIG.COL.PROJECT_STATUS
    ).setValue(
      hasAssignedTeam
        ? "Assign"
        : "Not Assign"
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

  const videographer =
    getFirstName(
      rowData.videographerphotographer
    );

  const client =
    getFirstName(
      rowData.client
    );

  const isAssigned =
    photographer || videographer;
  
  let assignLabel = "Not Assign";
  
  if (photographer && videographer) {
  
    assignLabel =
      `${photographer}/${videographer}`;
  
  } else if (photographer) {
  
    assignLabel = photographer;
  
  } else if (videographer) {
  
    assignLabel = videographer;
  }

  const title =
    `RAYS | ${assignLabel} - ${client} - ${rowData.university}`;

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

  let color;

  if (!isAssigned) {

    // RED
    color =
      CalendarApp.EventColor.RED;

  } else {

    color =
      getEventColor(
        rowData.package
      );
  }

  event.setColor(
    Number(color)
  );

  // =========================
  // GUEST SYNC
  // =========================

  const guests = [

    rowData.photographerEmail,

    rowData.videographerEmail

  ].filter(Boolean);

  // REMOVE DUPLICATE
  const uniqueGuests =
    [...new Set(guests)];

  try {

    // =========================
    // REMOVE OLD GUESTS
    // =========================

    const existingGuests =
      event.getGuestList();

    existingGuests.forEach(guest => {

      try {

        event.removeGuest(
          guest.getEmail()
        );

      } catch (err) {

        writeLog(
          "REMOVE_GUEST_SKIP",
          err.toString()
        );
      }
    });

    // =========================
    // ADD NEW GUESTS
    // =========================

    uniqueGuests.forEach(email => {

      try {

        event.addGuest(email);

      } catch (err) {

        writeLog(
          "ADD_GUEST_SKIP",
          err.toString()
        );
      }
    });

  } catch (err) {

    writeLog(
      "SYNC_GUEST_ERROR",
      err.toString()
    );
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

