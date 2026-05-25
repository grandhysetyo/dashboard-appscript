function deleteAllRaysEvents() {

  const calendar = CalendarApp.getCalendarById(
    CONFIG.CALENDAR_ID
  );

  const start = new Date("2026-01-01");
  const end = new Date("2027-01-01");

  const events = calendar.getEvents(
    start,
    end
  );

  let deleted = 0;

  events.forEach(event => {

    const title = event.getTitle();

    // ONLY DELETE RAYS EVENTS
    if (title.includes("RAYS")) {

      event.deleteEvent();

      deleted++;
    }
  });

  Logger.log(
    `${deleted} events deleted`
  );

  writeLog(
    "DELETE_EVENTS",
    `${deleted} calendar events deleted`
  );
}