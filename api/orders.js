const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

const ORDER_LIMIT_PER_WEEK = 40;
const ROLLING_AVAILABLE_WEEKS = 8;
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

  const rows = await getTableRows(getOrdersTableName());
  const currentOrders = rows.filter((row) => getPickupDateFromRow(row) === pickupDate).length;
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
      sendJson(res, 200, { ok: true, dates: getRollingAvailability() }, data.callback);
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
