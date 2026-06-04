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
      `${address}, Broward County, FL`,
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

export async function quoteDelivery(deliveryAddress) {
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
          miles: Math.round(miles * 10) / 10,
          requestedDestination,
          resolvedFrom: destination.resolvedFrom || requestedDestination,
        };
      }

      const extraMiles = Math.max(0, Math.ceil(miles - INCLUDED_DELIVERY_MILES));
      const fee = BASE_DELIVERY_FEE + extraMiles * DELIVERY_FEE_PER_EXTRA_MILE;
      return {
        ok: true,
        fee,
        miles: Math.round(miles * 10) / 10,
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
