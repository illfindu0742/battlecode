# Squarespace order formatter

Turns a **Squarespace orders export (CSV)** into a clean, grouped packing
list as an Excel file — no manual clean-up.

## What it does

Given the CSV you download from Squarespace (**Orders → Export**), it:

1. **Fills the Shipping Name down** into the blank rows of multi-item orders.
   Squarespace repeats the Order ID on every line item but only puts the
   customer name on the first row — this fills in the rest.
2. **Sorts by Shipping Name** so every row for the same customer is together,
   including a repeat customer's separate orders.
3. **Keeps only 5 columns:** Order ID, Lineitem quantity, Lineitem name,
   Lineitem variant, Shipping Name.
4. **Formats it:** all text centered, a styled and frozen header row,
   auto-sized columns, and light shading that alternates for each customer so
   the blocks are easy to see.

The output is a new `.xlsx` next to your CSV — your original file is never
changed.

## The easy way: double-click (no typing)

Keep all these files together in one folder. Then just:

- **Mac:** double-click **`Format Orders (Mac).command`**
- **Windows:** double-click **`Format Orders (Windows).bat`**

A window pops up asking you to pick your Squarespace CSV. Choose it, and a
formatted `..._formatted.xlsx` appears next to it and opens in Excel. That's it.

You can also **drag your CSV straight onto the launcher icon** to skip the
picker.

> The very first time, it quietly installs one small helper (`openpyxl`).
> If a launcher says Python isn't installed, get it from
> [python.org](https://www.python.org/downloads/) — on Windows, tick
> **"Add python.exe to PATH"** during install — then double-click again.
>
> On Mac, the first double-click may warn that it's from an unidentified
> developer: **right-click → Open → Open** once, and it's trusted after that.

## The manual way (terminal)

First time, from this folder: `pip install -r requirements.txt`
(use `pip3` / `python3` on some Macs). Then:

```bash
python format_orders.py orders.csv                 # -> orders_formatted.xlsx
python format_orders.py orders.csv packing.xlsx    # choose the output name
```

## A note on privacy

Order exports contain customer names, addresses, and phone numbers. The
`.gitignore` here is set up so CSV and XLSX files are **never** committed to
git. Keep it that way.
