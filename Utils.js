function formatDate(
  date
) {

  let parsedDate;

  // =========================
  // DATE OBJECT
  // =========================

  if (
    date instanceof Date
  ) {

    parsedDate = date;
  }

  // =========================
  // STRING MM/dd/yyyy
  // =========================

  else {

    const parts =
      date
        .toString()
        .trim()
        .split("/");

    const month =
      Number(parts[0]) - 1;

    const day =
      Number(parts[1]);

    const year =
      Number(parts[2]);

    parsedDate =
      new Date(
        year,
        month,
        day
      );
  }

  // =========================
  // RETURN FIXED STRING
  // M/d/yyyy
  // =========================

  return Utilities.formatDate(

    parsedDate,

    Session.getScriptTimeZone(),

    "M/d/yyyy"
  );
}

function combineDateTime(dateValue,timeValue) {
  let day;
  let month;
  let year;

  // =========================
  // DATE OBJECT
  // =========================

  if (
    dateValue instanceof Date
  ) {

    day =
      dateValue.getDate();

    month =
      dateValue.getMonth();

    year =
      dateValue.getFullYear();
  }

  // =========================
  // STRING MM/dd/yyyy
  // =========================

  else {

    const parts =
      dateValue
        .toString()
        .split("/");

    month =
      Number(parts[0]) - 1;

    day =
      Number(parts[1]);

    year =
      Number(parts[2]);
  }

  // =========================
  // TIME
  // =========================

  const formattedTime =
    formatTime24(
      timeValue
    );

  const timeParts =
    formattedTime.split(":");

  const hours =
    Number(timeParts[0]);

  const minutes =
    Number(timeParts[1]);

  // =========================
  // FINAL DATE
  // =========================

  return new Date(
    year,
    month,
    day,
    hours,
    minutes,

    0
  );
}

function getFirstName(name) {

  if (!name) return "";

  return name.toString().split(" ")[0];
}

function getPackageDuration(pkg) {

  if (pkg.includes("1.5")) return 1.5;
  if (pkg.includes("2")) return 2;

  return 1;
}


function getEventColor(pkg) {

  // 1.5 HOURS
  if (
    pkg &&
    pkg.toString().includes("1.5")
  ) {

    // BLUEBERRY
    return "9";
  }

  // 2 HOURS
  if (
    pkg &&
    pkg.toString().includes("2")
  ) {

    // ORANGE
    return "6";
  }

  // DEFAULT 1 HOUR
  // SAGE
  return "2";
}

function calculateDeadline() {

  const now = new Date();

  now.setDate(now.getDate() + 3);

  return now;
}

function isBooking2026(dateValue) {

  const date = new Date(dateValue);

  return date.getFullYear() === 2026;
}

function formatTime24(timeValue) {

  if (!timeValue) {
    return "";
  }

  // =========================
  // DATE OBJECT
  // =========================

  if (timeValue instanceof Date) {

    return Utilities.formatDate(
      timeValue,
      Session.getScriptTimeZone(),
      "HH:mm"
    );
  }

  // =========================
  // STRING
  // =========================

  let str =
    timeValue.toString().trim();

  // =========================
  // HANDLE AM/PM
  // =========================

  const isPM =
    str.toUpperCase().includes("PM");

  const isAM =
    str.toUpperCase().includes("AM");

  // REMOVE AM PM
  str = str
    .replace(/AM/i, "")
    .replace(/PM/i, "")
    .trim();



  if (!str) {
    return "";
  }
  
  const parts =
    str.split(":");

  let hours =
    Number(parts[0]);

  const minutes =
    parts[1]
      ? Number(parts[1])
      : 0;

  // PM CONVERT
  if (
    isPM &&
    hours < 12
  ) {
    hours += 12;
  }

  // 12 AM
  if (
    isAM &&
    hours === 12
  ) {
    hours = 0;
  }

  return (
    String(hours)
      .padStart(2, "0")
    +
    ":"
    +
    String(minutes)
      .padStart(2, "0")
  );
}

function getContrastColor(hex) {

  if (!hex) return "#000000";

  hex = hex.replace("#", "");

  const r =
    parseInt(hex.substr(0, 2), 16);

  const g =
    parseInt(hex.substr(2, 2), 16);

  const b =
    parseInt(hex.substr(4, 2), 16);

  const brightness =
    ((r * 299) +
    (g * 587) +
    (b * 114)) / 1000;

  return brightness > 125
    ? "#000000"
    : "#FFFFFF";
}

function setSystemMessage(
  sheet,
  row,
  message,
  color = "#FFFFFF"
) {

  const cell =
    sheet.getRange(
      row,
      CONFIG.COL.SYSTEM_MESSAGE
    );

  cell.setValue(message);
  cell.setBackground(color);
}

function getDashboardSheet() {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.DASHBOARD_SPREADSHEET_ID
    );

  return ss.getSheetByName(
    CONFIG.DASHBOARD_SHEET_NAME
  );
}

function getDashboardLastRow(sheet) {

  return Math.max(
    sheet.getLastRow(),
    1000
  );
}


function sendAssignmentEmail(
  email,
  rowData
) {

  if (!email) return;

  const subject =
    `[RAYS] ${rowData.client} - ${formatDate(rowData.tanggal)}`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: rowData.summary
  });

  writeLog(
    "EMAIL_SENT",
    `${rowData.bookingId} -> ${email}`
  );
}
