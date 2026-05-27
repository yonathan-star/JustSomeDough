# Just Some Dough website

This is a static preorder site for Just Some Dough. Open `index.html` in a browser to preview it.

## Connect orders to Google Sheets

1. Create a Google Sheet for orders.
2. In the Sheet, go to Extensions > Apps Script.
3. Paste the contents of `google-apps-script.js`.
4. Click Deploy > New deployment.
5. Select Web app.
6. Set "Execute as" to yourself.
7. Set "Who has access" to Anyone.
8. Deploy and copy the Web app URL.
9. Open `app.js` and paste the URL into:

```js
const GOOGLE_SCRIPT_URL = "PASTE_WEB_APP_URL_HERE";
```

Orders will be appended to the `Orders` tab and sorted by `Order Week`, then submission time. The website also sends a `Preferred Week` based on the preferred pickup date.

## Updating products

Edit the `products` array in `app.js` to change names, prices, descriptions, and photos. When individual product photos arrive, place them in this folder and update each product's `image` value.
