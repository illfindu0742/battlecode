# battlecode — Squarespace Order Formatter

Tools that turn a **Squarespace orders CSV export** (Orders → Export) into a
clean, grouped **Excel packing list**. Built for a small jewelry shop; the
person running it day-to-day is non-technical, so ease of use is the priority.

## The transformation (single source of truth — all three tools do exactly this)

1. **Fill Shipping Name down** onto every line of a multi-item order. Squarespace
   repeats the Order ID on each line item but only fills the customer/address
   fields on the first row. Falls back to **Billing Name** if an order has no
   shipping name at all.
2. **Sort by Shipping Name** so each customer's rows sit together — including a
   repeat buyer's separate orders.
3. **Keep only 5 columns, in this order:** `Order ID`, `Lineitem quantity`,
   `Lineitem name`, `Lineitem variant`, `Shipping Name`.
4. **Format:** all cells centered; header bold white on slate `#44546A`, frozen;
   columns auto-sized (capped); light per-customer banding `#F2F2F2`; autofilter;
   **Order ID stored as text** so leading zeros survive (e.g. `01437`).

(There is intentionally **no** "group identical line items" step — the user
explicitly dropped that idea.)

## Three implementations (pick by who's using it)

1. **Browser web app — `web/` (primary, simplest, cross-platform).**
   Static, 100% client-side: drop the CSV → formatted `.xlsx` downloads. Order
   data never leaves the browser; no server. Works offline / any browser.
   - `web/index.html` — drag-and-drop UI (theme-aware light/dark)
   - `web/format-core.js` — the logic as pure UMD functions (`parseCSV`,
     `transform`, `buildWorkbook(ExcelJS, result)`); runs in browser and Node
   - `web/exceljs.min.js` — **vendored** ExcelJS (writes the styled xlsx
     in-browser; vendored on purpose so there's no CDN dependency — CSP/offline safe)
   - `web/staticwebapp.config.json` — Azure Static Web Apps routing
   - `web/README.md` — usage + Azure deploy steps
2. **In-Excel button (VBA) — `FormatSquarespaceOrders.bas`.**
   A "Format Orders" button on Excel's Quick Access Toolbar; works on the
   currently-open sheet and writes a new formatted workbook (original untouched).
   Setup guide: `MAC-SETUP.md`. Printable setup guide artifact also exists.
3. **CLI — `format_orders.py`** (`requirements.txt` → `openpyxl`). Same result
   from a terminal: `python format_orders.py orders.csv [out.xlsx]`.

## Key files
- `EXAMPLE-orders.csv` — synthetic, **PII-free** sample with the real 47-column
  Squarespace header; exercises fill-down (order 00101), repeat-customer grouping
  (Grace Hopper 00102+00104), and billing-name fallback (00105). Use this for tests.
- `README.md` — top-level overview (leads with the web app).

## Running / testing
- **Web/core logic in Node:** `require('./web/format-core.js')` + ExcelJS; run
  against `EXAMPLE-orders.csv`. ExcelJS for Node isn't a repo dep — `npm i exceljs`
  in a temp dir if needed (the vendored `web/exceljs.min.js` UMD also `require`s).
- **CLI:** `pip install -r requirements.txt && python format_orders.py EXAMPLE-orders.csv`
- No CI configured in this repo. No test framework; verification is running the
  tools against `EXAMPLE-orders.csv` and checking the output workbook.

## Conventions & gotchas (important)
- **Customer PII:** real `.csv`/`.xlsx` are gitignored — never commit a real
  order export. Only `EXAMPLE-orders.csv` (synthetic) is committed (force-added).
- **Order-number leading zeros:** Squarespace shows `01437`; Excel strips the
  zero when it *opens* a CSV. VBA + core pad numeric order IDs to
  `ORDER_ID_MIN_WIDTH = 5` to restore them; xlsx cells use text format `@`.
- **Mac VBA constraints (kept for cross-platform):** no `Scripting.Dictionary`
  (absent on Mac — uses Collections), no file-system access (operates on the
  open sheet, builds a new workbook in memory).
- **Excel-for-Mac menu note:** "Tools → Macro" lives in the top menu bar, or use
  **View tab → Macros**; VBA editor = **⌥+F11** (Mac) / **Alt+F11** (Windows).
- Squarespace order export = 47 columns; only 6 are read (the 5 output columns +
  Billing Name for fallback). Columns are matched by header name, case-insensitive.

## Deployment status / open next step
- The web app is **built and committed but not yet deployed.** Hosting options:
  - **Azure Static Web Apps** (Free tier): Portal → Static Web App → GitHub deploy,
    **App location `/web`**, Api/Output blank. Steps in `web/README.md`.
  - **Netlify**: a Netlify MCP connector *is* available in sessions and can deploy
    end-to-end without user login (if Azure isn't a hard requirement).
  - **Double-click `web/index.html`** locally (zero hosting).
- **No Azure MCP connector is attached to sessions** (verified via ToolSearch +
  MCP server listing). The Microsoft connector present is **Microsoft 365 / Graph**
  (mail/calendar/SharePoint) — NOT Azure resource management, so it can't deploy.
  To deploy via a connector, the user must authorize an Azure connector on claude.ai.

## Git / branch
- Active work branch: **`claude/excel-plugin-feasibility-cu84ec`**.
- **PR #1** (open) → base `claude/volunteer-cabin-bookings-r2qOx` (this repo's
  default branch; there is no `main`).
- History note: an earlier `volunteercabinrentals.com` occupancy scraper that
  used to be the only code here was removed — it's unrelated to this project.
