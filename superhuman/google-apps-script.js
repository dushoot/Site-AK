const SHEET_ID = "1YwwWnk96e6P84mYwUhxPJiWjy-DQQbZXwpfyrY9hVKY";
const SHEET_NAME = "Лист1";

const HEADERS = ["email", "created_at", "page", "user_agent", "referrer"];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: "superhuman-waitlist" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(event) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];

  ensureHeaders(sheet);

  const params = event.parameter || {};
  const row = HEADERS.map((key) => params[key] || "");

  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeaders(sheet) {
  const currentHeaders = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getValues()[0];

  const hasHeaders = HEADERS.every((header, index) => currentHeaders[index] === header);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}
