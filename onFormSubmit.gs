function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // 請按你實際欄位改：
  const ATTEND_COL = 3; 
  const EMAIL_COL  = 5; 
  const STATUS_COL = 8; 

  const attendRaw = String(sheet.getRange(row, ATTEND_COL).getValue() || "");
  const attend = attendRaw.replace(/\s+/g, ""); // 去空白
  const email = String(sheet.getRange(row, EMAIL_COL).getValue() || "").trim();

  // 沒 email 就唔 send（因為係 optional）
  if (!email) return;

  Logger.log("Email=" + email);

  // Swift like logic: attend.contains("會出席") || attend.contains("未確定")
  const isAttending = attend.includes('會出席') || attend.includes('未確定');
  if (!isAttending) return;

  Logger.log("attend=" + attend);

  const statusCell = sheet.getRange(row, STATUS_COL);
  const status = String(statusCell.getValue() || "").trim();
  if (status === 'SENT') return;

  // Event details
  const calendar = CalendarApp.getDefaultCalendar();
  const eventTitle = "Tak & Natalie's Big Day 💍";

  const startTime = new Date('2026-04-25T14:00:00-04:00');
  const endTime   = new Date('2026-04-25T15:00:00-04:00');

  const location = "Markham Civic Centre, 101 Town Centre Blvd, Markham ON L3R 9W3, Canada";
  const description =
    "Tak & Natalie 的婚禮儀式 🤍\n\n" +
    "日期: 4 月 25 日\n" +
    "時間: 2:00 - 3:00 PM(Eastern Time, Toronto)\n" +
    "地點: Markham Civic Centre\n" +
    "場地: Wedding Chapel\n\n" +
    "好期待到時見到你～";

  Logger.log("Attend=" + sheet.getRange(row, ATTEND_COL).getValue());
  Logger.log("Email=" + sheet.getRange(row, EMAIL_COL).getValue());
  Logger.log("Status(before)=" + sheet.getRange(row, STATUS_COL).getValue());

  // Create event + send invite
  calendar.createEvent(eventTitle, startTime, endTime, {
    guests: email,
    sendInvites: true,
    location: location,
    description: description
  });

  statusCell.setValue("SENT");
}
