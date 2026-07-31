# Squarespace order formatter

Turns a **Squarespace orders export** into a clean, grouped packing list —
names filled in, sorted so each customer's rows sit together, trimmed to the
five columns that matter, centered and formatted.

## The Excel button (recommended, for Mac)

A one-click **Format Orders** button that lives right in Excel. No Python, no
internet, no files to move around.

- **`FormatSquarespaceOrders.bas`** — the macro
- **`MAC-SETUP.md`** — one-time setup (≈5 min, done once)
- **`EXAMPLE-orders.csv`** — fake-data sample for testing the button

**Everyday use:** download the orders CSV from Squarespace, open it in Excel,
click **Format Orders** — a clean new workbook appears. Follow
[`MAC-SETUP.md`](MAC-SETUP.md) to install the button.

## The command-line script (optional / advanced)

For anyone who'd rather run it from a terminal (or on a server), the same logic
is available as a Python script:

```bash
pip install -r requirements.txt
python format_orders.py orders.csv          # -> orders_formatted.xlsx
```

Both paths do exactly the same thing:

1. **Fill the Shipping Name down** into the blank continuation rows of a
   multi-item order (falls back to Billing Name if an order has no shipping
   name at all).
2. **Sort by Shipping Name** so each customer — including a repeat buyer's
   separate orders — is grouped together.
3. **Keep only:** Order ID, Lineitem quantity, Lineitem name, Lineitem variant,
   Shipping Name.
4. **Format:** centered text, styled + frozen header, auto-sized columns, light
   per-customer shading. Order numbers keep their leading zero (e.g. `01437`).

## A note on privacy

Real order exports contain customer names, addresses, and phone numbers.
`.gitignore` is set so real `.csv`/`.xlsx` files are **never** committed. The
only data file in the repo is the fake `EXAMPLE-orders.csv`.
