const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbylmClssVR4MnXjlAQd1iztXXSVWLr3Vd73XaX9lPE-449d1JQkkpb6sIpge4j9LFUGfw/exec";

const products = [
  {
    id: "regular-sourdough",
    name: "Regular Sourdough",
    price: 12,
    image: "regular-sourdough.jpg",
    description: "A naturally leavened loaf with a crisp crust and open, chewy crumb.",
  },
  {
    id: "chocolate-chip",
    name: "Chocolate Chip",
    price: 15,
    image: "Snapchat-856069574.jpg",
    description: "A sweet sourdough loaf folded with chocolate chips for breakfast or dessert.",
  },
  {
    id: "garlic-rosemary",
    name: "Garlic Rosemary",
    price: 15,
    image: "garlic-rosemary.jpg",
    description: "Savory sourdough made for dinner boards, soups, and sandwiches.",
  },
  {
    id: "blueberry-lemon",
    name: "Blueberry Lemon",
    price: 15,
    image: "blueberry-lemon-new.jpg",
    description: "Bright lemon and blueberry folded into a tender sourdough loaf.",
  },
  {
    id: "olive",
    name: "Olive",
    price: 15,
    image: "olive-sourdough.jpg",
    description: "A savory olive sourdough loaf with a briny, rich finish.",
  },
];

const cart = new Map();
const productGrid = document.querySelector("#productGrid");
const cartItems = document.querySelector("#cartItems");
const cartTotal = document.querySelector("#cartTotal");
const orderForm = document.querySelector("#orderForm");
const formStatus = document.querySelector("#formStatus");
const deliveryFeeLine = document.querySelector("#deliveryFeeLine");
const deliveryFeeAmount = document.querySelector("#deliveryFeeAmount");
const deliveryAddressField = document.querySelector("#deliveryAddressField");
const deliveryFeeStatus = document.querySelector("#deliveryFeeStatus");
const pickupDateSelect = document.querySelector("#preferredDate");
const orderItemsInput = document.querySelector("#orderItems");
const orderTotalInput = document.querySelector("#orderTotal");
let currentDeliveryQuote = null;
let deliveryQuoteRequestId = 0;

const currency = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0,
});

function renderProducts() {
  productGrid.innerHTML = products
    .map(
      (product) => `
        <article class="product-card">
          ${
            product.image
              ? `<img src="${product.image}" alt="${product.name} sourdough loaf preview" />`
              : `<div class="product-image-placeholder" aria-hidden="true"></div>`
          }
          <div class="product-body">
            <div class="product-meta">
              <h3>${product.name}</h3>
              <span>${currency.format(product.price)}</span>
            </div>
            <p>${product.description}</p>
            <div class="qty-row" aria-label="${product.name} quantity">
              <button type="button" data-action="decrease" data-id="${product.id}" aria-label="Decrease ${product.name} quantity">-</button>
              <output id="qty-${product.id}">0</output>
              <button type="button" data-action="increase" data-id="${product.id}" aria-label="Increase ${product.name} quantity">+</button>
            </div>
          </div>
        </article>
      `,
    )
    .join("");
}

function getCartLines() {
  return products
    .map((product) => ({ ...product, quantity: cart.get(product.id) || 0 }))
    .filter((product) => product.quantity > 0);
}

function getFulfillmentFee() {
  return orderForm.elements.fulfillment.value === "Delivery" && currentDeliveryQuote
    ? currentDeliveryQuote.fee
    : 0;
}

function updateDeliveryAddressField() {
  const isDelivery = orderForm.elements.fulfillment.value === "Delivery";
  const input = orderForm.elements.deliveryAddress;
  deliveryAddressField.hidden = !isDelivery;
  input.required = isDelivery;
  if (!isDelivery) {
    input.value = "";
    currentDeliveryQuote = null;
    deliveryFeeStatus.textContent = "Delivery fee will be calculated from 2960 Birch Terr.";
  }
}

function updateCart() {
  const lines = getCartLines();
  const subtotal = lines.reduce((sum, line) => sum + line.price * line.quantity, 0);
  const deliveryFee = getFulfillmentFee();
  const total = subtotal + deliveryFee;

  products.forEach((product) => {
    const output = document.querySelector(`#qty-${product.id}`);
    if (output) output.value = cart.get(product.id) || 0;
  });

  if (!lines.length) {
    cartItems.className = "cart-items empty";
    cartItems.textContent = "No loaves selected yet.";
  } else {
    cartItems.className = "cart-items";
    cartItems.innerHTML = lines
      .map(
        (line) => `
          <div class="cart-line">
            <span>${line.quantity} x ${line.name}</span>
            <strong>${currency.format(line.price * line.quantity)}</strong>
          </div>
        `,
      )
      .join("");
  }

  deliveryFeeLine.hidden = deliveryFee === 0;
  deliveryFeeAmount.textContent = currency.format(deliveryFee);
  cartTotal.textContent = currency.format(total);
  orderItemsInput.value = lines.map((line) => `${line.quantity} x ${line.name}`).join("; ");
  orderTotalInput.value = String(total);
}

function getNextSunday(date = new Date()) {
  const nextSunday = new Date(date);
  nextSunday.setHours(12, 0, 0, 0);
  const daysUntilSunday = (7 - nextSunday.getDay()) % 7 || 7;
  nextSunday.setDate(nextSunday.getDate() + daysUntilSunday);
  return nextSunday;
}

function isSunday(dateString) {
  if (!dateString) return false;
  return new Date(`${dateString}T12:00:00`).getDay() === 0;
}

function buildPayload() {
  const formData = new FormData(orderForm);
  const deliveryAddress = formData.get("deliveryAddress");
  const notes = formData.get("notes");
  const deliveryFee = getFulfillmentFee();

  formData.set("submittedAt", new Date().toISOString());
  formData.set("items", orderItemsInput.value);
  formData.set("total", orderTotalInput.value);
  formData.set("deliveryFee", String(deliveryFee));
  if (deliveryAddress) {
    const deliveryDetails = currentDeliveryQuote
      ? `Delivery address: ${deliveryAddress} | Delivery fee: ${currency.format(deliveryFee)} | Distance: ${currentDeliveryQuote.miles} mi`
      : `Delivery address: ${deliveryAddress}`;
    formData.set("notes", `${deliveryDetails}${notes ? ` | ${notes}` : ""}`);
  }
  return new URLSearchParams(formData);
}

function requestScript(payload) {
  return new Promise((resolve, reject) => {
    const callbackName = `handleOrderResponse_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Order submission timed out."));
    }, 30000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (response) => {
      cleanup();
      resolve(response);
    };

    payload.set("callback", callbackName);
    payload.set("_", Date.now().toString());
    script.async = true;
    script.src = `${GOOGLE_SCRIPT_URL}?${payload.toString()}`;
    script.onerror = () => {
      cleanup();
      reject(new Error("Order submission failed."));
    };
    document.body.append(script);
  });
}

function submitOrder(payload) {
  return requestScript(payload);
}

function loadAvailability() {
  const payload = new URLSearchParams({ action: "availability" });
  return requestScript(payload);
}

function loadDeliveryFee(address) {
  const payload = new URLSearchParams({
    action: "deliveryFee",
    deliveryAddress: address,
  });
  return requestScript(payload);
}

async function submitOrderLegacy(payload) {
  await fetch(GOOGLE_SCRIPT_URL, {
    method: "POST",
    body: payload,
    mode: "no-cors",
  });

  return { ok: true, legacy: true };
}

function renderAvailability(dates) {
  pickupDateSelect.innerHTML = "";

  if (!dates.length) {
    pickupDateSelect.innerHTML = `<option value="">No available pickup weeks</option>`;
    pickupDateSelect.disabled = true;
    formStatus.textContent = "No pickup weeks are available right now.";
    return;
  }

  pickupDateSelect.disabled = false;
  pickupDateSelect.innerHTML = dates
    .map((date) => `<option value="${date.value}">${date.label}</option>`)
    .join("");
}

async function setAvailableDates() {
  pickupDateSelect.innerHTML = `<option value="">Loading available weeks...</option>`;
  pickupDateSelect.disabled = true;

  try {
    const response = await loadAvailability();
    if (!response.ok) throw new Error(response.error || "Could not load available weeks.");
    renderAvailability(response.dates || []);
  } catch (error) {
    pickupDateSelect.innerHTML = `<option value="">Available weeks could not load</option>`;
    formStatus.textContent = "Available pickup weeks could not load. Please refresh or contact the baker.";
  }
}

async function updateDeliveryQuote() {
  const address = orderForm.elements.deliveryAddress.value.trim();
  const requestId = (deliveryQuoteRequestId += 1);

  currentDeliveryQuote = null;
  updateCart();

  if (orderForm.elements.fulfillment.value !== "Delivery") return true;

  if (!address) {
    deliveryFeeStatus.textContent = "Enter a delivery address to calculate the fee.";
    return false;
  }

  deliveryFeeStatus.textContent = "Calculating delivery fee...";

  try {
    const quote = await loadDeliveryFee(address);
    if (requestId !== deliveryQuoteRequestId) return false;

    if (!quote.ok) {
      deliveryFeeStatus.textContent = quote.error || "Could not calculate delivery fee.";
      return false;
    }

    currentDeliveryQuote = quote;
    deliveryFeeStatus.textContent = `Delivery: ${currency.format(quote.fee)} (${quote.miles} mi from 2960 Birch Terr)`;
    updateCart();
    return true;
  } catch (error) {
    if (requestId === deliveryQuoteRequestId) {
      deliveryFeeStatus.textContent = "Could not calculate delivery fee. Please check the address.";
    }
    return false;
  }
}

productGrid.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) return;

  const product = products.find((item) => item.id === button.dataset.id);
  if (!product) return;

  const current = cart.get(product.id) || 0;
  const next = button.dataset.action === "increase" ? current + 1 : Math.max(0, current - 1);
  if (next === 0) cart.delete(product.id);
  else cart.set(product.id, next);
  updateCart();
});

orderForm.elements.fulfillment.addEventListener("change", () => {
  updateDeliveryAddressField();
  updateCart();
  if (orderForm.elements.fulfillment.value === "Delivery") updateDeliveryQuote();
});

orderForm.elements.deliveryAddress.addEventListener("change", () => {
  updateDeliveryQuote();
});

orderForm.elements.deliveryAddress.addEventListener("blur", () => {
  updateDeliveryQuote();
});

orderForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!getCartLines().length) {
    formStatus.textContent = "Add at least one loaf before submitting.";
    return;
  }

  if (!orderForm.elements.preferredDate.value) {
    formStatus.textContent = "Choose an available pickup week.";
    return;
  }

  if (orderForm.elements.fulfillment.value === "Delivery" && !(await updateDeliveryQuote())) {
    formStatus.textContent = "Enter a valid delivery address before submitting.";
    return;
  }

  if (!GOOGLE_SCRIPT_URL) {
    formStatus.textContent =
      "Order is ready, but the Google Apps Script URL has not been added to app.js yet.";
    return;
  }

  formStatus.textContent = "Submitting order...";

  try {
    const payload = buildPayload();
    let response;

    try {
      response = await submitOrder(payload);
    } catch (error) {
      response = await submitOrderLegacy(payload);
    }

    if (!response.ok) {
      formStatus.textContent = response.error || "Could not submit the order. Please try another pickup date.";
      return;
    }

    orderForm.reset();
    cart.clear();
    await setAvailableDates();
    updateCart();
    formStatus.textContent = "Order submitted. The baker will confirm shortly.";
  } catch (error) {
    formStatus.textContent = "Could not submit the order. Please try again or text the baker.";
  }
});

renderProducts();
setAvailableDates();
updateDeliveryAddressField();
updateCart();
