# Just Some Dough website

This is a static preorder site with a Vercel API that writes orders into an Excel workbook through Microsoft Graph.

## Excel workbook

Put the workbook in the OneDrive account that will own the orders.

Create a real Excel table named `Orders` with these columns:

```text
SubmittedAt
PickupDate
Name
Phone
Email
Fulfillment
Items
Total
Notes
Address
DeliveryFee
DeliveryMiles
PaymentMethod
```

The old `AvailableWeeks` table is no longer required. The API automatically shows the next 8 Sundays.

## Microsoft app setup

1. Go to `https://entra.microsoft.com`.
2. Open `Applications > App registrations`.
3. Create a new registration.
4. Name it `JustSomeDough Orders`.
5. For supported account types, choose:

```text
Accounts in any organizational directory and personal Microsoft accounts
```

6. Add this redirect URI as a Web redirect URI:

```text
http://localhost:53682/callback
```

7. Copy the Application client ID.
8. Go to `Certificates & secrets`.
9. Create a new client secret and copy the secret value immediately.

## Get the refresh token

In PowerShell, run:

```powershell
$env:MS_CLIENT_ID='PASTE_CLIENT_ID_HERE'
$env:MS_CLIENT_SECRET='PASTE_CLIENT_SECRET_VALUE_HERE'
node scripts/microsoft-auth.js
```

Open the URL printed in the terminal, sign in to the Microsoft account that owns the Excel file, and approve access.
The terminal will print `MS_REFRESH_TOKEN`.

## Vercel environment variables

In Vercel, add these environment variables:

```text
MS_CLIENT_ID
MS_CLIENT_SECRET
MS_REFRESH_TOKEN
MS_WORKBOOK_SHARE_URL
```

Set `MS_WORKBOOK_SHARE_URL` to the shared Excel URL, for example:

```text
https://posnack-my.sharepoint.com/:x:/r/personal/.../Doc.aspx?...&file=SourDough.xlsx
```

If the workbook is in your own OneDrive, you can use `MS_WORKBOOK_PATH` instead:

```text
Documents/JustSomeDough.xlsx
```

If you know the OneDrive file item ID, you can use `MS_WORKBOOK_ITEM_ID` instead.

## Deploy

Deploy this folder to Vercel. The website calls:

```text
/api/orders
```

Test these URLs after deploy:

```text
https://YOUR-VERCEL-DOMAIN.vercel.app/api/orders?action=availability
https://YOUR-VERCEL-DOMAIN.vercel.app/api/orders?action=deliveryFee&deliveryAddress=1534%20Johnson%20St
```

Then submit a test order from the website and confirm it appears in the Excel `Orders` table.

## Updating products

Edit the `products` array in `app.js` to change names, prices, descriptions, and photos.
