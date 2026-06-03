const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

const ORDER_LIMIT_PER_WEEK = 40;
const ROLLING_AVAILABLE_WEEKS = 8;
const BAKER_SUMMARY_START_ROW = 22;
const BAKER_SUMMARY_COLUMNS = 15;
const BASE_DELIVERY_FEE = 5;
const INCLUDED_DELIVERY_MILES = 21;
const DELIVERY_FEE_PER_EXTRA_MILE = 1;
const MAX_DELIVERY_MILES = 25;

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
    scope: "offline_access Files.ReadWrite User.Read",
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

async function refreshAvailabilityTable(dates) {
  const tableName = getAvailabilityTableName();
  const existingRows = await getTableRows(tableName);

  if (existingRows.length) {
    await deleteTableRows(tableName, existingRows.length);
  }

  await addTableRows(
    tableName,
    dates.map((date) => [date.value, true]),
  );
}

function isBlankTableRow(row) {
  const values = row?.values?.[0] || [];
  return values.every((value) => String(value ?? "").trim() === "");
}

function getCell(rowValues, index) {
  return rowValues?.[index] ?? "";
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

function buildBakerSummaryRows(dates, orderRows) {
  const rows = [];
  const ordersByDate = new Map();

  orderRows
    .filter((row) => !isBlankTableRow(row))
    .forEach((row) => {
      const values = row.values?.[0] || [];
      const dateKey = getDateKey(getCell(values, 1));
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

    rows.push([date.value, ...Array(BAKER_SUMMARY_COLUMNS - 1).fill("")]);
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
        "",
        "",
        getCell(order, 12),
        getCell(order, 8),
        getCell(order, 9),
        getCell(order, 3),
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

async function refreshBakerSummary(dates, orderRows) {
  const values = buildBakerSummaryRows(dates, orderRows);
  const clearRows = Math.max(350, values.length + 20);
  const clearValues = Array.from({ length: clearRows }, () => Array(BAKER_SUMMARY_COLUMNS).fill(""));
  const clearEndRow = BAKER_SUMMARY_START_ROW + clearRows - 1;
  const writeEndRow = BAKER_SUMMARY_START_ROW + values.length - 1;

  await updateWorksheetValues(getBakerSheetName(), `A${BAKER_SUMMARY_START_ROW}:O${clearEndRow}`, clearValues);

  if (values.length) {
    await updateWorksheetValues(getBakerSheetName(), `A${BAKER_SUMMARY_START_ROW}:O${writeEndRow}`, values);
  }
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

function getDeliveryQuote(deliveryAddress) {
  if (!String(deliveryAddress || "").trim()) {
    return { ok: false, error: "Enter a delivery address.", code: "NO_DELIVERY_ADDRESS" };
  }

  return {
    ok: true,
    fee: BASE_DELIVERY_FEE,
    miles: 0,
    includedMiles: INCLUDED_DELIVERY_MILES,
    feePerExtraMile: DELIVERY_FEE_PER_EXTRA_MILE,
    maxMiles: MAX_DELIVERY_MILES,
  };
}

async function handleSubmitOrder(data) {
  const pickupDate = normalizeDateKey(data.preferredDate);
  if (!pickupDate) {
    return { ok: false, error: "Choose a valid pickup date." };
  }

  const availableDates = new Set(getRollingAvailability().map((date) => date.value));
  if (!availableDates.has(pickupDate)) {
    return {
      ok: false,
      error: "This pickup week is not available. Please choose another Sunday.",
      code: "WEEK_UNAVAILABLE",
    };
  }

  if (data.fulfillment === "Delivery") {
    const deliveryQuote = getDeliveryQuote(data.deliveryAddress);
    if (!deliveryQuote.ok) return deliveryQuote;
  }

  const ordersTableName = getOrdersTableName();
  const rows = await getTableRows(ordersTableName);
  await removeBlankTableRows(ordersTableName, rows);

  const currentOrders = rows
    .filter((row) => !isBlankTableRow(row))
    .filter((row) => getPickupDateFromRow(row) === pickupDate).length;
  if (currentOrders >= ORDER_LIMIT_PER_WEEK) {
    return {
      ok: false,
      error: "This pickup week is full. Please choose another Sunday.",
      code: "WEEK_FULL",
      limit: ORDER_LIMIT_PER_WEEK,
      currentOrders,
    };
  }

  await addOrderRow({ ...data, preferredDate: pickupDate });
  const refreshedRows = await getTableRows(ordersTableName);
  await refreshBakerSummary(getRollingAvailability(), refreshedRows);

  return {
    ok: true,
    limit: ORDER_LIMIT_PER_WEEK,
    currentOrders: currentOrders + 1,
  };
}

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
      const dates = getRollingAvailability();
      await refreshAvailabilityTable(dates);
      await refreshBakerSummary(dates, await getTableRows(getOrdersTableName()));
      sendJson(res, 200, { ok: true, dates }, data.callback);
      return;
    }

    if (data.action === "deliveryFee") {
      sendJson(res, 200, getDeliveryQuote(data.deliveryAddress), data.callback);
      return;
    }

    const response = await handleSubmitOrder(data);
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
