function getOrCreateFolder(
  parentFolder,
  folderName
) {

  const folders =
    parentFolder.getFoldersByName(
      folderName
    );

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(
    folderName
  );
}

function findBookingFolder(
  bookingId
) {

  const root =
    DriveApp.getFolderById(
      CONFIG.DRIVE_ROOT_FOLDER_ID
    );

  const years =
    root.getFolders();

  while (years.hasNext()) {

    const yearFolder =
      years.next();

    const dates =
      yearFolder.getFolders();

    while (dates.hasNext()) {

      const dateFolder =
        dates.next();

      const clients =
        dateFolder.getFolders();

      while (clients.hasNext()) {

        const clientFolder =
          clients.next();

        if (
          clientFolder
            .getName()
            .startsWith(
              bookingId + " - "
            )
        ) {

          return clientFolder;
        }
      }
    }
  }

  return null;
}

function createBookingFolder(
  bookingId,
  clientName,
  university,
  sessionDate
) {

  const rootFolder =
    DriveApp.getFolderById(
      CONFIG.DRIVE_ROOT_FOLDER_ID
    );

  const year =
    Utilities.formatDate(
      sessionDate,
      Session.getScriptTimeZone(),
      "yyyy"
    );

  const monthDay =
    Utilities.formatDate(
      sessionDate,
      Session.getScriptTimeZone(),
      "MM-dd"
    );

  // =====================
  // YEAR
  // =====================

  const yearFolder =
    getOrCreateFolder(
      rootFolder,
      year
    );

  // =====================
  // DATE
  // =====================

  const dateFolder =
    getOrCreateFolder(
      yearFolder,
      monthDay
    );

  // =====================
  // CLIENT
  // =====================

  const clientFolderName =
    `${bookingId} - ${university} - ${clientName}`;

  const clientFolder =
    getOrCreateFolder(
      dateFolder,
      clientFolderName
    );

  return {

    url:
      clientFolder.getUrl(),

    name:
      clientFolder.getName()
  };
}

function moveBookingFolder(
  bookingId,
  sessionDate
) {

  const bookingFolder =
    findBookingFolder(
      bookingId
    );

  if (!bookingFolder) {

    writeLog(
      "FOLDER_NOT_FOUND",
      bookingId
    );

    return;
  }

  const root =
    DriveApp.getFolderById(
      CONFIG.DRIVE_ROOT_FOLDER_ID
    );

  const year =
    Utilities.formatDate(
      sessionDate,
      Session.getScriptTimeZone(),
      "yyyy"
    );

  const monthDay =
    Utilities.formatDate(
      sessionDate,
      Session.getScriptTimeZone(),
      "MM-dd"
    );

  const yearFolder =
    getOrCreateFolder(
      root,
      year
    );

  const dateFolder =
    getOrCreateFolder(
      yearFolder,
      monthDay
    );

  const parents =
    bookingFolder.getParents();

  while (
    parents.hasNext()
  ) {

    const oldParent =
      parents.next();

    dateFolder.addFolder(
      bookingFolder
    );

    oldParent.removeFolder(
      bookingFolder
    );
  }

  writeLog(
    "MOVE_FOLDER_SUCCESS",
    bookingId
  );
}

function syncBookingFolderDate(
  rowData,
  sheet,
  row
) {

  const bookingFolder =
    findBookingFolder(
      rowData.bookingId
    );

  // =====================
  // FOLDER NOT FOUND
  // CREATE RETROACTIVE
  // =====================

  if (!bookingFolder) {

    writeLog(
      "FOLDER_NOT_FOUND_CREATE",
      rowData.bookingId
    );

    const folder =
      createBookingFolder(
        rowData.bookingId,
        rowData.client,
        rowData.university,
        rowData.tanggal
      );

    sheet.getRange(
      row,
      CONFIG.COL.DRIVE_FOLDER
    ).setValue(
      folder.url
    );

    writeLog(
      "FOLDER_CREATED_RETROACTIVE",
      rowData.bookingId
    );

    return;
  }

  // =====================
  // TARGET DATE FOLDER
  // =====================

  const root =
    DriveApp.getFolderById(
      CONFIG.DRIVE_ROOT_FOLDER_ID
    );

  const year =
    Utilities.formatDate(
      rowData.tanggal,
      Session.getScriptTimeZone(),
      "yyyy"
    );

  const monthDay =
    Utilities.formatDate(
      rowData.tanggal,
      Session.getScriptTimeZone(),
      "MM-dd"
    );

  const yearFolder =
    getOrCreateFolder(
      root,
      year
    );

  const targetDateFolder =
    getOrCreateFolder(
      yearFolder,
      monthDay
    );

  // =====================
  // CURRENT PARENT
  // =====================

  const parents =
    bookingFolder.getParents();

  if (!parents.hasNext()) {

    writeLog(
      "FOLDER_PARENT_NOT_FOUND",
      rowData.bookingId
    );

    return;
  }

  const currentParent =
    parents.next();

  // =====================
  // ALREADY CORRECT
  // =====================

  if (
    currentParent.getId() ===
    targetDateFolder.getId()
  ) {

    writeLog(
      "FOLDER_ALREADY_SYNC",
      rowData.bookingId
    );

    return;
  }

  // =====================
  // MOVE FOLDER
  // =====================

  targetDateFolder.addFolder(
    bookingFolder
  );

  currentParent.removeFolder(
    bookingFolder
  );

  writeLog(
    "MOVE_FOLDER_SUCCESS",
    `${rowData.bookingId} -> ${monthDay}`
  );

}

function testCreateFolder() {

  const bookingId =
    "RAYS-TEST-001";

  const folder =
    createBookingFolder(
      bookingId,
      "Budi Test",
      "UBAYA",
      new Date("2026-05-23")
    );

  Logger.log(folder.url);
}

function testRescheduleFolder() {

  syncBookingFolderDate(
    "RAYS-TEST-001",
    new Date("2026-05-30")
  );

}