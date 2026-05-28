const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwUe5PVjXh39iDZjz2sHNbC6OJ9le5vaB7_88yMszezRv0dL6xdoXgdN7jJJGqKKF1GXQ/exec";

const products = [
  {
    id: "regular-sourdough",
    name: "Regular Sourdough",
    price: 12,
    image: "Snapchat-1793413218.jpg",
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
    id: "general",
    name: "General",
    price: 15,
    image: "Snapchat-2127436635.jpg",
    description: "A fresh sourdough loaf baked with the week's classic dough.",
  },
  {
    id: "garlic-rosemary",
    name: "Garlic Rosemary",
    price: 15,
    image: "Snapchat-1804780126.jpg",
    description: "Savory sourdough made for dinner boards, soups, and sandwiches.",
  },
  {
    id: "blueberry-lemon",
    name: "Blueberry Lemon",
    price: 15,
    image: "Snapchat-1276294950.jpg",
    description: "Bright lemon and blueberry folded into a tender sourdough loaf.",
  },
  {
    id: "olive",
    name: "Olive",
    price: 15,
    image: "Snapchat-830501156.jpg",
    description: "A savory olive sourdough loaf with a briny, rich finish.",
  },
];

const cart = new Map();
const productGrid = document.querySelector("#productGrid");
const cartItems = document.querySelector("#cartItems");
const cartTotal = document.querySelector("#cartTotal");
const orderForm = document.querySelector("#orderForm");
const formStatus = document.querySelector("#formStatus");
const orderItemsInput = document.querySelector("#orderItems");
const orderTotalInput = document.querySelector("#orderTotal");

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
          <img src="${product.image}" alt="${product.name} sourdough loaf preview" />
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

function updateCart() {
  const lines = getCartLines();
  const total = lines.reduce((sum, line) => sum + line.price * line.quantity, 0);

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

function setDefaultDate() {
  const dateInput = orderForm.elements.preferredDate;
  dateInput.min = new Date().toISOString().slice(0, 10);
  dateInput.value = getNextSunday().toISOString().slice(0, 10);
}

function buildPayload() {
  const formData = new FormData(orderForm);
  formData.set("submittedAt", new Date().toISOString());
  formData.set("fulfillment", "Sunday pickup");
  formData.set("items", orderItemsInput.value);
  formData.set("total", orderTotalInput.value);
  return formData;
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

orderForm.elements.preferredDate.addEventListener("change", (event) => {
  if (!isSunday(event.target.value)) {
    formStatus.textContent = "Pickup is only available on Sundays. Please choose a Sunday.";
    event.target.value = getNextSunday(new Date(`${event.target.value}T12:00:00`)).toISOString().slice(0, 10);
  } else {
    formStatus.textContent = "";
  }

});

orderForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!getCartLines().length) {
    formStatus.textContent = "Add at least one loaf before submitting.";
    return;
  }

  if (!isSunday(orderForm.elements.preferredDate.value)) {
    formStatus.textContent = "Pickup is only available on Sundays. Please choose a Sunday.";
    return;
  }

  if (!GOOGLE_SCRIPT_URL) {
    formStatus.textContent =
      "Order is ready, but the Google Apps Script URL has not been added to app.js yet.";
    return;
  }

  formStatus.textContent = "Submitting order...";

  try {
    await fetch(GOOGLE_SCRIPT_URL, {
      method: "POST",
      body: buildPayload(),
      mode: "no-cors",
    });

    orderForm.reset();
    cart.clear();
    setDefaultDate();
    updateCart();
    formStatus.textContent = "Order submitted. The baker will confirm shortly.";
  } catch (error) {
    formStatus.textContent = "Could not submit the order. Please try again or text the baker.";
  }
});

renderProducts();
setDefaultDate();
updateCart();
