const SHEET_NAME = "Orders";

function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "No form data received. Submit from the website or run testOrder()." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getOrdersSheet();
  const data = e.parameter;

  const row = [
    new Date(),
    data.orderWeek || "",
    data.preferredWeek || "",
    data.name || "",
    data.phone || "",
    data.email || "",
    data.fulfillment || "",
    data.preferredDate || "",
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
    "Order Week",
    "Preferred Week",
    "Name",
    "Phone",
    "Email",
    "Fulfillment",
    "Preferred Date",
    "Items",
    "Total",
    "Notes",
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

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

function setupOrdersSheet() {
  getOrdersSheet();
}

function testOrder() {
  return doPost({
    parameter: {
      orderWeek: "TEST-WEEK",
      preferredWeek: "TEST-PREFERRED-WEEK",
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
