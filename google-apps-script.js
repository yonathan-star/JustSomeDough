const SHEET_NAME = "Orders";

function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "No form data received. Submit from the website or run testOrder()." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getOrdersSheet();
  const data = e.parameter;
  const pickupDate = parseDateInput(data.preferredDate);

  const row = [
    new Date(),
    pickupDate || data.preferredDate || "",
    data.name || "",
    data.phone || "",
    data.email || "",
    data.fulfillment || "",
    data.items || "",
    Number(data.total || 0),
    data.notes || "",
  ];

  sheet.appendRow(row);
  sortOrders(sheet);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrdersSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const headers = [
    "Submitted At",
    "Pickup Date",
    "Name",
    "Phone",
    "Email",
    "Fulfillment",
    "Items",
    "Total",
    "Notes",
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange("A:A").setNumberFormat("m/d/yyyy h:mm am/pm");
  sheet.getRange("B:B").setNumberFormat("dddd, mmmm d, yyyy");
  sheet.getRange("H:H").setNumberFormat("$0");

  return sheet;
}

function sortOrders(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort([
      { column: 2, ascending: true },
      { column: 1, ascending: true },
    ]);
}

function parseDateInput(value) {
  if (!value) return null;
  const parts = String(value).split("-");
  if (parts.length !== 3) return null;

  const year = Number(parts[0]);
  const month = Number(parts[1]) - 1;
  const day = Number(parts[2]);

  if (!year || month < 0 || !day) return null;
  return new Date(year, month, day, 12, 0, 0);
}

function setupOrdersSheet() {
  getOrdersSheet();
}

function testOrder() {
  return doPost({
    parameter: {
      name: "Test Customer",
      phone: "555-555-5555",
      email: "test@example.com",
      fulfillment: "Sunday pickup",
      preferredDate: "2026-05-31",
      items: "1 x Classic Country",
      total: "12",
      notes: "Test order from Apps Script.",
    },
  });
}
