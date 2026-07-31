# Squarespace order formatter — web tool

A tiny browser tool: drop the Squarespace **orders CSV**, get back a clean,
grouped Excel packing list. **Everything runs in the browser** — the order
file (with customer names/addresses) is never uploaded anywhere.

## Files

- `index.html` — the page (UI)
- `format-core.js` — the logic (CSV parse → fill-down → group → trim → format)
- `exceljs.min.js` — the library that writes the `.xlsx`, vendored so there's
  no external/CDN dependency at runtime
- `staticwebapp.config.json` — Azure Static Web Apps routing config

## Use it

Open the page (double-click `index.html`, or visit the hosted URL) → drop the
CSV → a formatted `..._formatted.xlsx` downloads. Works offline.

## What it does

1. Fills the **Shipping Name** onto every line of a multi-item order (falls
   back to Billing Name if an order has none).
2. **Sorts by Shipping Name** so each customer — repeat buyers included — is
   grouped together.
3. Keeps only **Order ID, Lineitem quantity, Lineitem name, Lineitem variant,
   Shipping Name**.
4. Centers text, styles + freezes the header, auto-sizes columns, shades each
   customer block. Order numbers keep their leading zero (stored as text).

## Deploy to Azure Static Web Apps (free)

No server code — this is a pure static site, so the **Free** plan is enough.
In the [Azure Portal](https://portal.azure.com):

1. **Create a resource → Static Web App.**
2. Sign in; pick your **Subscription** and a **Resource Group** (create one,
   e.g. `orders-tool`).
3. **Name** it (e.g. `squarespace-orders`); **Plan type: Free.**
4. **Deployment source: GitHub** → authorize → choose:
   - **Organization:** `illfindu0742`
   - **Repository:** `battlecode`
   - **Branch:** the branch this lives on
5. **Build details:** Build Presets → **Custom**
   - **App location:** `/web`
   - **Api location:** *(leave blank)*
   - **Output location:** *(leave blank)*
6. **Review + create.**

Azure adds a GitHub Actions workflow to the repo and deploys automatically;
you get a URL like `https://<name>.azurestaticapps.net`. Every later push
re-deploys on its own. Bookmark that URL — that's the tool.
