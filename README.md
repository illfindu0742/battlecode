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

## First-time setup (once per computer)

You need Python 3 installed (already on most Macs; on Windows, get it from
[python.org](https://www.python.org/downloads/) and check "Add to PATH").

Then, in a terminal, from this folder:

```bash
pip install -r requirements.txt
```

## Every time you want to format an export

```bash
python format_orders.py orders.csv
```

That writes `orders_formatted.xlsx` in the same folder. To choose the output
name yourself:

```bash
python format_orders.py orders.csv packing-list.xlsx
```

> On some Macs the commands are `pip3` and `python3` instead of `pip` and
> `python`.

## A note on privacy

Order exports contain customer names, addresses, and phone numbers. The
`.gitignore` here is set up so CSV and XLSX files are **never** committed to
git. Keep it that way.
