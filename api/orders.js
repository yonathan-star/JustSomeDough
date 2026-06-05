const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
import { quoteDelivery as quoteDeliveryFromModule } from "../scripts/delivery-quote-core.mjs";

const ORDER_LIMIT_PER_WEEK = 60;
const ROLLING_AVAILABLE_WEEKS = 8;
const BAKER_SUMMARY_START_ROW = 22;
const BAKER_SUMMARY_COLUMNS = 15;
const BAKERY_ADDRESS = "2960 Birch Terrace, Davie, FL 33330";
const DELIVERY_ADDRESS_CONTEXT = "Broward County, FL";
const BASE_DELIVERY_FEE = 5;
const INCLUDED_DELIVERY_MILES = 21;
const DELIVERY_FEE_PER_EXTRA_MILE = 1;
const MAX_DELIVERY_MILES = 25;
const NOMINATIM_SEARCH_URL = "https://nominatim.openstreetmap.org/search";
const OSRM_ROUTE_URL = "https://router.project-osrm.org/route/v1/driving";

const geocodeCache = new Map();
let bakeryLocationPromise = null;

const ORDER_COLUMNS = [
  "SubmittedAt",
  "PickupDate",
  "Name",
  "Phone",
  "Email",
  "Fulfillment",
  "Items",
  "Total",
  "Notes",
  "Address",
  "DeliveryFee",
  "DeliveryMiles",
  "PaymentMethod",
];

function getOrdersTableName() {
  return (process.env.MS_ORDERS_TABLE_NAME || "Orders").trim();
}

function getAvailabilityTableName() {
  return (process.env.MS_AVAILABILITY_TABLE_NAME || "AvailableWeeks").trim();
}

function getBakerSheetName() {
  return (process.env.MS_BAKER_SHEET_NAME || "Orders").trim();
}

function sendJson(res, statusCode, payload, callback) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  res.setHeader("Surrogate-Control", "no-store");

  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    res.setHeader("Content-Type", "application/javascript; charset=utf-8");
    res.status(statusCode).send(`${callback}(${JSON.stringify(payload)});`);
    return;
  }

  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.status(statusCode).json(payload);
}

function getQuery(req) {
  if (req.method === "POST" && req.body && typeof req.body === "object") {
    return { ...req.query, ...req.body };
  }

  return req.query || {};
}

function requireEnv(name) {
  let value = process.env[name];
  if (!value) throw new Error(`Missing Vercel environment variable: ${name}`);
  value = value.trim().replace(/^['"]|['"]$/g, "");

  if (name === "MS_REFRESH_TOKEN" && value.startsWith("MS_REFRESH_TOKEN=")) {
    value = value.slice("MS_REFRESH_TOKEN=".length).trim();
  }

  return value;
}

async function getAccessToken() {
  const body = new URLSearchParams({
    client_id: requireEnv("MS_CLIENT_ID"),
    client_secret: requireEnv("MS_CLIENT_SECRET"),
    refresh_token: requireEnv("MS_REFRESH_TOKEN"),
    grant_type: "refresh_token",
    scope: "offline_access Files.ReadWrite User.Read Mail.Send",
  });

  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || "Could not refresh Microsoft access token.");
  }

  return payload.access_token;
}

function encodeWorkbookPath(path) {
  return path
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

function encodeSharingUrl(url) {
  return `u!${Buffer.from(url, "utf8")
    .toString("base64")
    .replace(/=+$/, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-")}`;
}

function getWorkbookBasePath() {
  if (process.env.MS_WORKBOOK_DRIVE_ID && process.env.MS_WORKBOOK_ITEM_ID) {
    return `/drives/${encodeURIComponent(requireEnv("MS_WORKBOOK_DRIVE_ID"))}/items/${encodeURIComponent(requireEnv("MS_WORKBOOK_ITEM_ID"))}`;
  }

  if (process.env.MS_WORKBOOK_SHARE_URL) {
    return `/shares/${encodeSharingUrl(process.env.MS_WORKBOOK_SHARE_URL)}/driveItem`;
  }

  if (process.env.MS_WORKBOOK_ITEM_ID) {
    return `/me/drive/items/${encodeURIComponent(process.env.MS_WORKBOOK_ITEM_ID)}`;
  }

  const workbookPath = requireEnv("MS_WORKBOOK_PATH");
  return `/me/drive/root:/${encodeWorkbookPath(workbookPath)}:`;
}

async function graphRequest(path, options = {}) {
  const accessToken = await getAccessToken();
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  const text = await response.text();
  const payload = text ? JSON.parse(text) : null;

  if (!response.ok) {
    const message = payload?.error?.message || payload?.error_description || "Microsoft Graph request failed.";
    throw new Error(message);
  }

  return payload;
}

async function getTableRows(tableName) {
  const rows = [];
  let nextUrl = `${getWorkbookBasePath()}/workbook/tables/${encodeURIComponent(tableName)}/rows`;

  while (nextUrl) {
    const path = nextUrl.startsWith(GRAPH_BASE_URL) ? nextUrl.slice(GRAPH_BASE_URL.length) : nextUrl;
    const payload = await graphRequest(path);
    rows.push(...(payload.value || []));
    nextUrl = payload["@odata.nextLink"] || "";
  }

  return rows;
}

async function deleteTableRows(tableName, rowCount) {
  for (let index = rowCount - 1; index >= 0; index -= 1) {
    await graphRequest(
      `${getWorkbookBasePath()}/workbook/tables/${encodeURIComponent(tableName)}/rows/$/itemAt(index=${index})`,
      { method: "DELETE" },
    );
  }
}

async function addTableRows(tableName, values) {
  if (!values.length) return null;

  return graphRequest(`${getWorkbookBasePath()}/workbook/tables/${encodeURIComponent(tableName)}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values }),
  });
}

async function getWorksheetIdByName(sheetName) {
  const payload = await graphRequest(`${getWorkbookBasePath()}/workbook/worksheets`);
  const worksheets = payload.value || [];
  const target = sheetName.trim().toLowerCase();
  const worksheet =
    worksheets.find((item) => String(item.name || "").trim().toLowerCase() === target) ||
    worksheets.find((item) => String(item.name || "").toLowerCase() === target);

  if (!worksheet) {
    throw new Error(`Could not find worksheet: ${sheetName}`);
  }

  return worksheet.id;
}

async function updateWorksheetValues(sheetName, address, values) {
  const worksheetId = await getWorksheetIdByName(sheetName);
  return graphRequest(
    `${getWorkbookBasePath()}/workbook/worksheets/${encodeURIComponent(worksheetId)}/range(address='${address}')`,
    {
      method: "PATCH",
      body: JSON.stringify({ values }),
    },
  );
}

async function getWorksheetValues(sheetName, address) {
  const worksheetId = await getWorksheetIdByName(sheetName);
  const payload = await graphRequest(
    `${getWorkbookBasePath()}/workbook/worksheets/${encodeURIComponent(worksheetId)}/range(address='${address}')`,
  );

  return payload.values || [];
}

async function refreshAvailabilityTable(dates) {
  const tableName = getAvailabilityTableName();
  const existingRows = await getTableRows(tableName);
  const availabilityByDate = new Map();

  existingRows
    .filter((row) => !isBlankTableRow(row))
    .forEach((row) => {
      const values = row.values?.[0] || [];
      const dateKey = normalizeDateKey(values[0]);
      if (!dateKey) return;
      availabilityByDate.set(dateKey, isAvailabilityEnabled(values[1]));
    });

  const refreshedDates = dates.map((date) => ({
    ...date,
    available: availabilityByDate.has(date.value) ? availabilityByDate.get(date.value) : true,
  }));
  const existingValues = existingRows
    .filter((row) => !isBlankTableRow(row))
    .map((row) => row.values?.[0] || []);
  const alreadyMatches =
    existingValues.length === refreshedDates.length &&
    refreshedDates.every((date, index) => {
      const row = existingValues[index] || [];
      return normalizeDateKey(row[0]) === date.value && isAvailabilityEnabled(row[1]) === date.available;
    });

  if (alreadyMatches) return refreshedDates;

  if (existingRows.length) {
    await deleteTableRows(tableName, existingRows.length);
  }

  await addTableRows(
    tableName,
    refreshedDates.map((date) => [date.value, date.available]),
  );

  return refreshedDates;
}

function isBlankTableRow(row) {
  const values = row?.values?.[0] || [];
  return values.every((value) => String(value ?? "").trim() === "");
}

function getCell(rowValues, index) {
  return rowValues?.[index] ?? "";
}

function isAvailabilityEnabled(value) {
  if (value === true) return true;
  if (value === false) return false;

  const normalized = String(value ?? "").trim().toLowerCase();
  return ["true", "yes", "1", "open", "available"].includes(normalized);
}

function normalizeDeliveryAddress(address) {
  return String(address || "").trim().replace(/\s+/g, " ").replace(/\s+([,.])/g, "$1");
}

function hasLocalAddressContext(address) {
  return /\b\d{5}(?:-\d{4})?\b/.test(address) || /\bfl(?:orida)?\b/i.test(address) || address.split(",").length >= 3;
}

function dedupeStrings(values) {
  const seen = {};
  return values.filter((value) => {
    const key = String(value || "").toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getDeliveryDestinationOptions(deliveryAddress) {
  const address = normalizeDeliveryAddress(deliveryAddress);
  if (!address) return [];

  let options = [address];
  if (!hasLocalAddressContext(address)) {
    options = [
      `${address}, Davie, FL`,
      `${address}, Davie, Broward County, FL`,
      `${address}, ${DELIVERY_ADDRESS_CONTEXT}`,
      `${address}, FL`,
      `${address}, Florida, USA`,
      address,
    ];
  }

  return dedupeStrings(options);
}

async function fetchJson(url, options = {}, timeoutMs = 12000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
      headers: {
        Accept: "application/json",
        "User-Agent": "JustSomeDough/1.0",
        ...(options.headers || {}),
      },
    });
    const text = await response.text();
    const payload = text ? JSON.parse(text) : null;

    if (!response.ok) {
      throw new Error(payload?.error || payload?.message || `Request failed: ${response.status}`);
    }

    return payload;
  } finally {
    clearTimeout(timeout);
  }
}

async function geocodeAddress(address) {
  const normalized = normalizeDeliveryAddress(address);
  if (!normalized) return null;

  if (geocodeCache.has(normalized)) return geocodeCache.get(normalized);

  const params = new URLSearchParams({
    format: "jsonv2",
    limit: "1",
    countrycodes: "us",
    addressdetails: "1",
    q: normalized,
  });

  const payload = await fetchJson(`${NOMINATIM_SEARCH_URL}?${params.toString()}`);
  const result = Array.isArray(payload) ? payload[0] : null;
  if (!result || result.lat == null || result.lon == null) {
    geocodeCache.set(normalized, null);
    return null;
  }

  const location = {
    lat: Number(result.lat),
    lon: Number(result.lon),
    displayName: result.display_name || normalized,
  };
  geocodeCache.set(normalized, location);
  return location;
}

async function getGeocodedLocationFromVariants(address) {
  const candidates = getDeliveryDestinationOptions(address);
  for (const candidate of candidates) {
    const location = await geocodeAddress(candidate);
    if (location) {
      return { ...location, requestedAddress: normalizeDeliveryAddress(address), resolvedFrom: candidate };
    }
  }

  return null;
}

async function getBakeryLocation() {
  if (!bakeryLocationPromise) {
    bakeryLocationPromise = getGeocodedLocationFromVariants(BAKERY_ADDRESS);
  }

  return bakeryLocationPromise;
}

async function getRouteMeters(origin, destination) {
  const params = new URLSearchParams({
    overview: "false",
    alternatives: "false",
    steps: "false",
  });

  const url = `${OSRM_ROUTE_URL}/${origin.lon},${origin.lat};${destination.lon},${destination.lat}?${params.toString()}`;
  const payload = await fetchJson(url);
  const route = payload?.routes?.[0];
  return route?.distance || 0;
}

function parseOrderItems(items) {
  const parsed = {
    plain: 0,
    chocolateChip: 0,
    garlicRosemary: 0,
    blueberryLemon: 0,
    olive: 0,
    custom: 0,
    customText: "",
    totalQuantity: 0,
  };

  String(items || "")
    .split(";")
    .map((item) => item.trim())
    .filter(Boolean)
    .forEach((item) => {
      const match = item.match(/^(\d+(?:\.\d+)?)\s*x\s*(.+)$/i);
      const quantity = match ? Number(match[1]) : 1;
      const name = (match ? match[2] : item).trim();
      const normalizedName = name.toLowerCase();

      parsed.totalQuantity += quantity;

      if (normalizedName.includes("regular") || normalizedName.includes("plain")) {
        parsed.plain += quantity;
      } else if (normalizedName.includes("chocolate") || normalizedName === "cc") {
        parsed.chocolateChip += quantity;
      } else if (normalizedName.includes("garlic") || normalizedName.includes("rosemary")) {
        parsed.garlicRosemary += quantity;
      } else if (normalizedName.includes("blueberry") || normalizedName.includes("lemon")) {
        parsed.blueberryLemon += quantity;
      } else if (normalizedName.includes("olive")) {
        parsed.olive += quantity;
      } else {
        parsed.custom += quantity;
        parsed.customText = parsed.customText ? `${parsed.customText}; ${item}` : item;
      }
    });

  return parsed;
}

function asDisplayValue(value) {
  if (value === 0 || value === "0") return "";
  return value || "";
}

function getSummaryOrderKey(dateKey, name, phone) {
  return [dateKey, String(name || "").trim().toLowerCase(), String(phone || "").trim()].join("|");
}

function extractManualSummaryValues(summaryRows) {
  const manualByOrder = new Map();
  let currentDate = "";

  summaryRows.forEach((row) => {
    const dateKey = normalizeDateKey(row?.[0]);
    if (dateKey) {
      currentDate = dateKey;
      return;
    }

    if (!currentDate) return;

    const name = String(row?.[0] || "").trim();
    if (!name || ["qty", "total $"].includes(name.toLowerCase())) return;

    const phone = row?.[13] || "";
    const override = {
      made: row?.[8] ?? "",
      paid: row?.[9] ?? "",
      payment: row?.[10] ?? "",
      notes: row?.[11] ?? "",
      address: row?.[12] ?? "",
      phone: row?.[13] ?? "",
    };

    if (Object.values(override).every((value) => String(value ?? "").trim() === "")) return;

    manualByOrder.set(getSummaryOrderKey(currentDate, name, phone), override);
  });

  return manualByOrder;
}

function buildBakerSummaryRows(dates, orderRows, overflowByDate = new Map(), manualByOrder = new Map()) {
  const rows = [];
  const ordersByDate = new Map();

  orderRows
    .filter((row) => !isBlankTableRow(row))
    .forEach((row) => {
      const values = row.values?.[0] || [];
      const dateKey = normalizeDateKey(getCell(values, 1));
      if (!dateKey) return;

      if (!ordersByDate.has(dateKey)) ordersByDate.set(dateKey, []);
      ordersByDate.get(dateKey).push(values);
    });

  dates.forEach((date) => {
    const weekOrders = (ordersByDate.get(date.value) || []).sort((a, b) =>
      String(getCell(a, 0)).localeCompare(String(getCell(b, 0))),
    );
    const totals = {
      plain: 0,
      chocolateChip: 0,
      garlicRosemary: 0,
      blueberryLemon: 0,
      olive: 0,
      custom: 0,
      quantity: 0,
      money: 0,
    };
    const blank = Array(BAKER_SUMMARY_COLUMNS).fill("");

    rows.push([
      date.value,
      ...Array(BAKER_SUMMARY_COLUMNS - 4).fill(""),
      "Overflow loaves",
      asDisplayValue(overflowByDate.get(date.value) || 0),
      "",
    ]);
    rows.push([
      "",
      "Plain",
      "CC",
      "GR",
      "BL",
      "olive",
      "Custom",
      "QTY",
      "Made",
      "Paid",
      "Payment",
      "Notes",
      "Address",
      "Phone",
      "Total $",
    ]);

    weekOrders.forEach((order) => {
      const items = parseOrderItems(getCell(order, 6));
      const total = Number(getCell(order, 7) || 0);

      totals.plain += items.plain;
      totals.chocolateChip += items.chocolateChip;
      totals.garlicRosemary += items.garlicRosemary;
      totals.blueberryLemon += items.blueberryLemon;
      totals.olive += items.olive;
      totals.custom += items.custom;
      totals.quantity += items.totalQuantity;
      totals.money += total;

      rows.push([
        getCell(order, 2),
        asDisplayValue(items.plain),
        asDisplayValue(items.chocolateChip),
        asDisplayValue(items.garlicRosemary),
        asDisplayValue(items.blueberryLemon),
        asDisplayValue(items.olive),
        items.customText,
        asDisplayValue(items.totalQuantity),
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.made || "",
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.paid || "",
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.payment || getCell(order, 12),
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.notes || getCell(order, 8),
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.address || getCell(order, 9),
        manualByOrder.get(getSummaryOrderKey(date.value, getCell(order, 2), getCell(order, 3)))?.phone || getCell(order, 3),
        asDisplayValue(total),
      ]);
    });

    rows.push([
      "QTY",
      asDisplayValue(totals.plain),
      asDisplayValue(totals.chocolateChip),
      asDisplayValue(totals.garlicRosemary),
      asDisplayValue(totals.blueberryLemon),
      asDisplayValue(totals.olive),
      asDisplayValue(totals.custom),
      asDisplayValue(totals.quantity),
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
    rows.push([
      "total $",
      totals.plain ? totals.plain * 12 : "",
      totals.chocolateChip ? totals.chocolateChip * 15 : "",
      totals.garlicRosemary ? totals.garlicRosemary * 15 : "",
      totals.blueberryLemon ? totals.blueberryLemon * 15 : "",
      totals.olive ? totals.olive * 15 : "",
      "",
      "t sales",
      asDisplayValue(totals.money),
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
    rows.push(blank);
  });

  return rows;
}

function extractOverflowCounts(summaryRows) {
  const overflowByDate = new Map();

  summaryRows.forEach((row) => {
    const dateKey = normalizeDateKey(row?.[0]);
    if (!dateKey) return;

    const overflowCount = Number(row?.[13] || 0);
    if (Number.isFinite(overflowCount) && overflowCount > 0) {
      overflowByDate.set(dateKey, overflowCount);
    }
  });

  return overflowByDate;
}

async function getOverflowCounts() {
  const rows = await getWorksheetValues(getBakerSheetName(), `A${BAKER_SUMMARY_START_ROW}:O250`);
  return extractOverflowCounts(rows);
}

async function refreshBakerSummary(dates, orderRows, overflowByDate) {
  const existingRows = await getWorksheetValues(getBakerSheetName(), `A${BAKER_SUMMARY_START_ROW}:O250`);
  const counts = overflowByDate || extractOverflowCounts(existingRows);
  const manualByOrder = extractManualSummaryValues(existingRows);
  const values = buildBakerSummaryRows(dates, orderRows, counts, manualByOrder);

  // Pad with blank rows up to row 250 so stale data from previous larger summaries is cleared
  const blank = Array(BAKER_SUMMARY_COLUMNS).fill("");
  const totalRows = 250 - BAKER_SUMMARY_START_ROW + 1;
  const paddedValues = values.slice();
  while (paddedValues.length < totalRows) paddedValues.push(blank);

  await updateWorksheetValues(getBakerSheetName(), `A${BAKER_SUMMARY_START_ROW}:O250`, paddedValues);
}

function getLoafCountFromOrderValues(values) {
  return parseOrderItems(getCell(values, 6)).totalQuantity;
}

function getRequestedLoafCount(data) {
  const count = parseOrderItems(data.items).totalQuantity;
  return count > 0 ? count : 1;
}

function getLoavesByDate(rows) {
  const loavesByDate = new Map();

  rows
    .filter((row) => !isBlankTableRow(row))
    .forEach((row) => {
      const values = row.values?.[0] || [];
      const dateKey = normalizeDateKey(getCell(values, 1));
      if (!dateKey) return;
      loavesByDate.set(dateKey, (loavesByDate.get(dateKey) || 0) + getLoafCountFromOrderValues(values));
    });

  return loavesByDate;
}

function getOpenWeeks(availability, loavesByDate, requestedLoaves = 1) {
  return availability
    .filter((date) => date.available)
    .map((date) => ({
      ...date,
      currentLoaves: loavesByDate.get(date.value) || 0,
      remainingLoaves: Math.max(ORDER_LIMIT_PER_WEEK - (loavesByDate.get(date.value) || 0), 0),
    }))
    .filter((date) => date.remainingLoaves >= requestedLoaves);
}

async function removeBlankTableRows(tableName, rows) {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (isBlankTableRow(rows[index])) {
      await graphRequest(
        `${getWorkbookBasePath()}/workbook/tables/${encodeURIComponent(tableName)}/rows/$/itemAt(index=${index})`,
        { method: "DELETE" },
      );
    }
  }
}

async function addOrderRow(data) {
  const values = [
    ORDER_COLUMNS.map((column) => {
      switch (column) {
        case "SubmittedAt":
          return data.submittedAt || new Date().toISOString();
        case "PickupDate":
          return data.preferredDate || "";
        case "Name":
          return data.name || "";
        case "Phone":
          return data.phone || "";
        case "Email":
          return data.email || "";
        case "Fulfillment":
          return data.fulfillment || "";
        case "Items":
          return data.items || "";
        case "Total":
          return Number(data.total || 0);
        case "Notes":
          return data.notes || "";
        case "Address":
          return data.deliveryAddress || "";
        case "DeliveryFee":
          return Number(data.deliveryFee || 0);
        case "DeliveryMiles":
          return data.deliveryMiles ? Number(data.deliveryMiles) : "";
        case "PaymentMethod":
          return data.paymentMethod || "";
        default:
          return "";
      }
    }),
  ];

  return graphRequest(`${getWorkbookBasePath()}/workbook/tables/${encodeURIComponent(getOrdersTableName())}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values }),
  });
}

function normalizeDateKey(value) {
  if (!value) return "";

  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400000));
    return date.toISOString().slice(0, 10);
  }

  const text = String(value).trim();
  const isoMatch = text.match(/\d{4}-\d{2}-\d{2}/);
  if (isoMatch) return isoMatch[0];

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);

  return "";
}

function getNextSunday(date = new Date()) {
  const nextSunday = new Date(date);
  nextSunday.setUTCHours(12, 0, 0, 0);
  const daysUntilSunday = (7 - nextSunday.getUTCDay()) % 7 || 7;
  nextSunday.setUTCDate(nextSunday.getUTCDate() + daysUntilSunday);
  return nextSunday;
}

function formatDateLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
    timeZone: "UTC",
  }).format(date);
}

function getRollingAvailability() {
  const dates = [];
  const sunday = getNextSunday();

  for (let index = 0; index < ROLLING_AVAILABLE_WEEKS; index += 1) {
    const date = new Date(sunday);
    date.setUTCDate(sunday.getUTCDate() + index * 7);
    dates.push({
      value: date.toISOString().slice(0, 10),
      label: formatDateLabel(date),
    });
  }

  return dates;
}

function getPickupDateFromRow(row) {
  return normalizeDateKey(row?.values?.[0]?.[1]);
}

async function calculateDeliveryQuote(deliveryAddress) {
  const destinationOptions = getDeliveryDestinationOptions(deliveryAddress);
  if (!destinationOptions.length) {
    return { ok: false, error: "Enter a delivery address.", code: "NO_DELIVERY_ADDRESS" };
  }

  const origin = await getBakeryLocation();
  if (!origin) {
    return { ok: false, error: "Could not locate the bakery address.", code: "NO_BAKERY_ROUTE" };
  }

  for (const requestedDestination of destinationOptions) {
    const destination = await geocodeAddress(requestedDestination);
    if (!destination) {
      continue;
    }

    try {
      const meters = await getRouteMeters(origin, destination);
      if (!meters) {
        continue;
      }

      const miles = meters / 1609.344;
      if (miles > MAX_DELIVERY_MILES) {
        return {
          ok: false,
          error: `Delivery is only available within ${MAX_DELIVERY_MILES} miles.`,
          code: "OUT_OF_DELIVERY_RANGE",
          miles: roundMiles(miles),
          requestedDestination,
          resolvedFrom: destination.resolvedFrom || requestedDestination,
        };
      }

      const extraMiles = Math.max(0, Math.ceil(miles - INCLUDED_DELIVERY_MILES));
      const fee = BASE_DELIVERY_FEE + extraMiles * DELIVERY_FEE_PER_EXTRA_MILE;
      return {
        ok: true,
        fee,
        miles: roundMiles(miles),
        origin: BAKERY_ADDRESS,
        destination: String(destination.displayName || requestedDestination),
        requestedDestination: String(requestedDestination),
        resolvedFrom: String(destination.resolvedFrom || requestedDestination),
        includedMiles: INCLUDED_DELIVERY_MILES,
        feePerExtraMile: DELIVERY_FEE_PER_EXTRA_MILE,
        maxMiles: MAX_DELIVERY_MILES,
      };
    } catch (error) {
      continue;
    }
  }

  return {
    ok: false,
    error: "Could not calculate delivery distance. Please check the address.",
    code: "DELIVERY_QUOTE_FAILED",
  };
}

async function getDeliveryQuote(deliveryAddress) {
  return calculateDeliveryQuote(deliveryAddress);
}

async function sendOrderSmsNotification(data, pickupDate) {
  const notifyAddress = (process.env.SMS_NOTIFY_ADDRESS || "").trim();
  if (!notifyAddress) return;

  const name = data.name || "Unknown";
  const items = data.items || "—";
  const fulfillment = data.fulfillment || "";
  const total = data.total ? `$${data.total}` : "";
  const phone = data.phone || "";

  const body = [name, items, fulfillment, pickupDate, total, phone].filter(Boolean).join(" | ");

  await graphRequest("/me/sendMail", {
    method: "POST",
    body: JSON.stringify({
      message: {
        subject: "New Dough Order",
        body: { contentType: "Text", content: body.slice(0, 160) },
        toRecipients: [{ emailAddress: { address: notifyAddress } }],
      },
      saveToSentItems: false,
    }),
  });
}

async function handleSubmitOrder(data, deliveryQuote = null) {
  const pickupDate = normalizeDateKey(data.preferredDate);
  if (!pickupDate) {
    return { ok: false, error: "Choose a valid pickup date." };
  }

  const availability = await refreshAvailabilityTable(getRollingAvailability());
  const availableDates = new Set(availability.filter((date) => date.available).map((date) => date.value));
  if (!availableDates.has(pickupDate)) {
    return {
      ok: false,
      error: "This pickup week is not available. Please choose another Sunday.",
      code: "WEEK_UNAVAILABLE",
    };
  }

  if (data.fulfillment === "Delivery") {
    const quotedDelivery = deliveryQuote || (await quoteDeliveryFromModule(data.deliveryAddress));
    if (!quotedDelivery.ok) return quotedDelivery;
  }

  const ordersTableName = getOrdersTableName();
  const rows = await getTableRows(ordersTableName);
  await removeBlankTableRows(ordersTableName, rows);

  const loavesByDate = getLoavesByDate(rows);
  const requestedLoaves = getRequestedLoafCount(data);
  const currentLoaves = loavesByDate.get(pickupDate) || 0;
  if (currentLoaves + requestedLoaves > ORDER_LIMIT_PER_WEEK) {
    const overflowByDate = await getOverflowCounts();
    overflowByDate.set(pickupDate, (overflowByDate.get(pickupDate) || 0) + requestedLoaves);
    await refreshBakerSummary(availability, rows, overflowByDate);

    return {
      ok: false,
      error: "This pickup week is full. Please choose another available Sunday.",
      code: "WEEK_FULL",
      limit: ORDER_LIMIT_PER_WEEK,
      currentLoaves,
      requestedLoaves,
      availableWeeks: getOpenWeeks(availability, loavesByDate, requestedLoaves),
    };
  }

  await addOrderRow({ ...data, preferredDate: pickupDate });

  try {
    await sendOrderSmsNotification(data, pickupDate);
  } catch (_) {
    // Non-fatal: don't fail the order if the SMS notification fails
  }

  const refreshedRows = await getTableRows(ordersTableName);
  await refreshBakerSummary(availability, refreshedRows);

  return {
    ok: true,
    limit: ORDER_LIMIT_PER_WEEK,
    currentLoaves: currentLoaves + requestedLoaves,
  };
}

export const __test__ = {
  buildBakerSummaryRows,
  extractManualSummaryValues,
  extractOverflowCounts,
  getLoavesByDate,
  getOpenWeeks,
  getRequestedLoafCount,
  getDeliveryDestinationOptions,
  hasLocalAddressContext,
  calculateDeliveryQuote,
  getDeliveryQuote,
  geocodeAddress,
  getGeocodedLocationFromVariants,
  getBakeryLocation,
  getRouteMeters,
  normalizeDeliveryAddress,
  parseOrderItems,
};

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    sendJson(res, 200, { ok: true });
    return;
  }

  if (!["GET", "POST"].includes(req.method)) {
    sendJson(res, 405, { ok: false, error: "Method not allowed." });
    return;
  }

  const data = getQuery(req);

  try {
    if (data.action === "availability") {
      const rollingDates = getRollingAvailability();
      const dates = await refreshAvailabilityTable(rollingDates);
      // Refresh baker summary so manually-added Excel orders appear in week totals
      try {
        const orderRows = await getTableRows(getOrdersTableName());
        await refreshBakerSummary(dates, orderRows);
      } catch (_) {
        // Non-fatal: don't block availability response if summary refresh fails
      }
      sendJson(res, 200, { ok: true, dates }, data.callback);
      return;
    }

    if (data.action === "deliveryFee") {
      const deliveryQuote = await quoteDeliveryFromModule(data.deliveryAddress);
      sendJson(res, 200, deliveryQuote, data.callback);
      return;
    }

    let deliveryQuote = null;
    if (data.fulfillment === "Delivery") {
      deliveryQuote = await quoteDeliveryFromModule(data.deliveryAddress);
    }

    const response = await handleSubmitOrder(data, deliveryQuote);
    sendJson(res, 200, response, data.callback);
  } catch (error) {
    sendJson(
      res,
      500,
      {
        ok: false,
        error: "Order system is not configured yet. Check the Vercel environment variables and Excel workbook.",
        detail: error.message,
      },
      data.callback,
    );
  }
}
