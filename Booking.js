function generateBookingId() {

  const ss = SpreadsheetApp.openById(
    CONFIG.DASHBOARD_SPREADSHEET_ID
  );

  const sheet = ss.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return buildBookingId(1);
  }

  const values = sheet
    .getRange(2, CONFIG.COL.BOOKING_ID, lastRow - 1, 1)
    .getValues();

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyMMdd"
  );

  let count = 0;

  values.forEach(row => {

    const id = row[0];

    if (id && id.includes(today)) {
      count++;
    }
  });

  return buildBookingId(count + 1);
}

function buildBookingId(number) {

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyMMdd"
  );

  return `RAYS-${today}-${String(number).padStart(3, "0")}`;
}

function generateSummary(data) {

  let jamText =
    formatTime24(data.jam1);

  if (
    data.jam2 &&
    data.jam2.toString().trim() !== ""
  ) {

    jamText +=
      ` & ${formatTime24(data.jam2)}`;
  }

  return `
━━━━━━━━━━━━━━━━━━━━━━━━━━
🎓 RAYS MOMENTS BOOKING
━━━━━━━━━━━━━━━━━━━━━━━━━━

👤 CLIENT
Nama       : ${data.nama}
Instagram  : ${data.instagram}
Whatsapp   : ${'https://wa.me/62'+data.whatsapp}

🏫 SESSION
Universitas : ${data.universitas}
Lokasi      : ${data.lokasi}

📅 SCHEDULE
Tanggal : ${formatDate(data.tanggal)}
Jam     : ${jamText}

📦 PACKAGE
${data.paket}

━━━━━━━━━━━━━━━━━━━━━━━━━━
📝 NOTES
━━━━━━━━━━━━━━━━━━━━━━━━━━

${data.notes || "-"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
📂 GOOGLE DRIVE
━━━━━━━━━━━━━━━━━━━━━━━━━━

${data.driveFolder || "-"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
Booking ID:
${data.bookingId || "-"}
━━━━━━━━━━━━━━━━━━━━━━━━━━
  `.trim();
}