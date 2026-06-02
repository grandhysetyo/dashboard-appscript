function processBooking(
  data,
  formRowId = null
){

  const timestamp = data[0];
  const email = data[1];
  const nama = data[2];
  const universitas = data[3];
  const fakultas = data[4];
  const whatsapp = data[5];
  const instagram = data[6];
  const paket = data[7];
  const tanggal = data[8];
  const jam1 = data[9];
  const jam2 = data[10];
  const buktiTransfer = data[11];
  const agreement = data[12];
  const lokasi = data[13];
  const nominal = data[14];
  const notes = data[15];

  // ONLY 2026
  if (!isBooking2026(tanggal)) {

    writeLog(
      "SKIPPED",
      `Booking ignored because not 2026: ${nama}`
    );

    return;
  }

  const bookingId =
    generateBookingId();

  const folder =
    createBookingFolder(
      bookingId,
      nama,
      universitas,
      tanggal
    );

  // =========================
  // SUMMARY
  // =========================

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
      notes,
      driveFolder: folder.url
    });


  // =========================
  // CREATE EVENTS
  // =========================
  const events =
    createCalendarEvents({
      clientName: nama,
      universitas,
      tanggal,
      jam1,
      jam2,
      paket,
      summary
    });

  // =========================
  // SAVE DASHBOARD
  // =========================
  insertBookingToDashboard({
    timestamp,
    formRowId,
    calendarId1: events[0],
    calendarId2: events[1],
    bookingId,
    nama,
    universitas,
    fakultas,
    tanggal,
    jam1,
    jam2,
    lokasi,
    paket,
    instagram,
    whatsapp,
    notes,
    summary, 
    driveFolder: folder.url
  });

  writeLog(
    "CREATE",
    `${bookingId} created`
  );
}