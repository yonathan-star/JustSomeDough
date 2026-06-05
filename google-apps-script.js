const SHEET_NAME = "Orders";
const AVAILABILITY_SHEET_NAME = "Available Weeks";
const MAX_ORDERS_PER_WEEK = 60;
const ROLLING_AVAILABLE_WEEKS = 8;
const BAKERY_ADDRESS = "2960 Birch Terrace, Davie, FL 33330";
const DELIVERY_ADDRESS_CONTEXT = "Broward County, FL";
const BASE_DELIVERY_FEE = 5;
const INCLUDED_DELIVERY_MILES = 21;
const DELIVERY_FEE_PER_EXTRA_MILE = 1;
const MAX_DELIVERY_MILES = 25;

function doPost(e) {
  return handleOrderRequest(e);
}

function doGet(e) {
  return handleOrderRequest(e);
}

function handleOrderRequest(e) {
  if (!e || !e.parameter) {
    return createResponse({ ok: false, error: "No form data received. Submit from the website or run testOrder()." });
  }

  const data = e.parameter;
  if (data.action === "availability") {
    return createResponse({ ok: true, dates: getAvailablePickupDates() }, data.callback);
  }

  if (data.action === "deliveryFee") {
    return createResponse(getDeliveryQuote(data.deliveryAddress), data.callback);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getOrdersSheet();
    const pickupDate = parseDateInput(data.preferredDate);

    if (!pickupDate) {
      return createResponse({ ok: false, error: "Choose a valid pickup date." }, data.callback);
    }

    if (!isPickupDateAvailable(pickupDate)) {
      return createResponse(
        {
          ok: false,
          error: "This pickup week is not available. Please choose another Sunday.",
          code: "WEEK_UNAVAILABLE",
        },
        data.callback,
      );
    }

    let deliveryQuote = null;
    if (data.fulfillment === "Delivery") {
      deliveryQuote = getDeliveryQuote(data.deliveryAddress);
      if (!deliveryQuote.ok) return createResponse(deliveryQuote, data.callback);
    }

    const currentOrders = countOrdersForPickupDate(sheet, pickupDate);
    if (currentOrders >= MAX_ORDERS_PER_WEEK) {
      return createResponse(
        {
          ok: false,
          error: "This pickup week is full. Please choose another Sunday.",
          code: "WEEK_FULL",
          limit: MAX_ORDERS_PER_WEEK,
          currentOrders,
        },
        data.callback,
      );
    }

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
      data.deliveryAddress || "",
      deliveryQuote ? deliveryQuote.fee : 0,
      deliveryQuote ? deliveryQuote.miles : "",
      data.paymentMethod || "",
    ];

    sheet.appendRow(row);
    sortOrders(sheet);

    return createResponse(
      {
        ok: true,
        limit: MAX_ORDERS_PER_WEEK,
        currentOrders: currentOrders + 1,
      },
      data.callback,
    );
  } finally {
    lock.releaseLock();
  }
}

function createResponse(payload, callback) {
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function countOrdersForPickupDate(sheet, pickupDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const targetDateKey = getDateKey(pickupDate);
  return values.filter(([value]) => getDateKey(value) === targetDateKey).length;
}

function getDeliveryQuote(deliveryAddress) {
  const destinationOptions = getDeliveryDestinationOptions(deliveryAddress);
  if (!destinationOptions.length) {
    return { ok: false, error: "Enter a delivery address.", code: "NO_DELIVERY_ADDRESS" };
  }

  try {
    const quote = getFirstDeliveryRoute(destinationOptions);
    const leg = quote && quote.leg;
    const meters = leg && leg.distance && leg.distance.value;

    if (!meters) {
      return { ok: false, error: "Could not calculate delivery distance. Please check the address.", code: "NO_ROUTE" };
    }

    const miles = meters / 1609.344;
    if (miles > MAX_DELIVERY_MILES) {
      return {
        ok: false,
        error: `Delivery is only available within ${MAX_DELIVERY_MILES} miles.`,
        code: "OUT_OF_DELIVERY_RANGE",
        miles: roundMiles(miles),
      };
    }

    const extraMiles = Math.max(0, Math.ceil(miles - INCLUDED_DELIVERY_MILES));
    const fee = BASE_DELIVERY_FEE + extraMiles * DELIVERY_FEE_PER_EXTRA_MILE;

    return {
      ok: true,
      fee,
      miles: roundMiles(miles),
      origin: BAKERY_ADDRESS,
      destination: leg.end_address || deliveryAddress,
      requestedDestination: deliveryAddress,
      resolvedFrom: quote.destination,
    };
  } catch (error) {
    return { ok: false, error: "Could not calculate delivery distance. Please check the address.", code: "DELIVERY_QUOTE_FAILED" };
  }
}

function getFirstDeliveryRoute(destinations) {
  for (let index = 0; index < destinations.length; index += 1) {
    try {
      const destination = destinations[index];
      const directions = Maps.newDirectionFinder()
        .setOrigin(BAKERY_ADDRESS)
        .setDestination(destination)
        .setMode(Maps.DirectionFinder.Mode.DRIVING)
        .getDirections();

      const route = directions.routes && directions.routes[0];
      const leg = route && route.legs && route.legs[0];
      if (leg && leg.distance && leg.distance.value) {
        return { destination, leg };
      }
    } catch (error) {
      // Try the next locally-qualified version of the address.
    }
  }

  return null;
}

function getDeliveryDestinationOptions(deliveryAddress) {
  const address = String(deliveryAddress || "").trim().replace(/\s+/g, " ");
  if (!address) return [];

  let options = [address];
  if (!hasLocalAddressContext(address)) {
    options = [
      `${address}, ${DELIVERY_ADDRESS_CONTEXT}`,
      `${address}, FL`,
      `${address}, Florida, USA`,
      address,
    ];
  }

  return dedupeStrings(options);
}

function hasLocalAddressContext(address) {
  return /\b\d{5}(?:-\d{4})?\b/.test(address) || /\bfl(?:orida)?\b/i.test(address) || address.split(",").length >= 3;
}

function dedupeStrings(values) {
  const seen = {};
  return values.filter((value) => {
    const key = value.toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function roundMiles(value) {
  return Math.round(value * 10) / 10;
}

function getAvailabilitySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(AVAILABILITY_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(AVAILABILITY_SHEET_NAME);
  }

  const headers = ["Pickup Date", "Available"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    seedAvailableWeeks(sheet);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange("A:A").setNumberFormat("dddd, mmmm d, yyyy");
  return sheet;
}

function seedAvailableWeeks(sheet) {
  refreshAvailableWeeks(sheet);
}

function refreshAvailableWeeks(sheet = getAvailabilitySheet()) {
  const existingAvailability = getExistingAvailability(sheet);
  const rows = getRollingSundayDates().map((date) => {
    const dateKey = getDateKey(date);
    return [date, existingAvailability.has(dateKey) ? existingAvailability.get(dateKey) : true];
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function getExistingAvailability(sheet) {
  const availability = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return availability;

  sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .forEach(([date, available]) => {
      const dateKey = getDateKey(date);
      if (dateKey) availability.set(dateKey, isAvailableValue(available));
    });

  return availability;
}

function getRollingSundayDates() {
  const rows = [];
  const sunday = getNextSunday(new Date());

  for (let index = 0; index < ROLLING_AVAILABLE_WEEKS; index += 1) {
    rows.push(new Date(sunday));
    sunday.setDate(sunday.getDate() + 7);
  }

  return rows;
}

function getAvailablePickupDates() {
  const sheet = getAvailabilitySheet();
  refreshAvailableWeeks(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const todayKey = getDateKey(new Date());
  return sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .map(([date, available]) => ({ date, available }))
    .filter(({ date, available }) => isAvailableValue(available) && getDateKey(date) >= todayKey)
    .map(({ date }) => ({
      value: getDateKey(date),
      label: Utilities.formatDate(normalizeDate(date), Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy"),
    }))
    .sort((a, b) => a.value.localeCompare(b.value));
}

function isPickupDateAvailable(pickupDate) {
  const pickupDateKey = getDateKey(pickupDate);
  return getAvailablePickupDates().some(({ value }) => value === pickupDateKey);
}

function isAvailableValue(value) {
  if (value === true) return true;
  return ["true", "yes", "y", "available", "open", "1"].includes(String(value).trim().toLowerCase());
}

function getDateKey(value) {
  const date = normalizeDate(value);
  if (!date) return "";

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function normalizeDate(value) {
  if (value instanceof Date) return value;

  if (typeof value === "number") {
    return new Date(Math.round((value - 25569) * 86400000));
  }

  const parsedInput = parseDateInput(value);
  if (parsedInput) return parsedInput;

  const parsedDate = new Date(value);
  if (!Number.isNaN(parsedDate.getTime())) return parsedDate;

  return null;
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
    "Address",
    "Delivery Fee",
    "Delivery Miles",
    "Payment Method",
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
  sheet.getRange("K:K").setNumberFormat("$0");

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

function getNextSunday(date) {
  const nextSunday = new Date(date);
  nextSunday.setHours(12, 0, 0, 0);
  const daysUntilSunday = (7 - nextSunday.getDay()) % 7 || 7;
  nextSunday.setDate(nextSunday.getDate() + daysUntilSunday);
  return nextSunday;
}

function setupOrdersSheet() {
  getOrdersSheet();
  getAvailabilitySheet();
}

function testOrder() {
  return doPost({
    parameter: {
      name: "Test Customer",
      phone: "555-555-5555",
      email: "test@example.com",
      fulfillment: "Sunday pickup",
      preferredDate: "2026-05-31",
      items: "1 x Regular Sourdough",
      total: "12",
    },
  });
}
