# One-time setup: the "Format Orders" button in Excel for Mac

Do this **once** on your wife's Mac. It takes about 5 minutes. After that she
never sees any of this again — she just opens a CSV and clicks a button.

You'll need the file **`FormatSquarespaceOrders.bas`** (the macro) on the Mac.

---

## Step 1 — Allow macros to run

1. Open **Excel**.
2. Menu bar: **Excel → Settings** (older versions: **Preferences**).
3. Click **Security & Privacy**.
4. Under *Macro Security*, choose **"Disable all macros with notification"**
   (this is usually already the default). Close Settings.

This lets your own macro run while still warning about macros from strangers.

---

## Step 2 — Create the "Personal Macro Workbook"

This is an invisible workbook Excel loads every time it starts — the perfect
home for a button you want on *every* file.

1. Menu bar: **Tools → Macro → Record Macro…**
   *(No Tools menu? Use the **View** tab → **Macros ▾ → Record Macro…**.)*
2. In the dialog, set **Store macro in: → Personal Macro Workbook**. Click **OK**.
3. Immediately stop: **Tools → Macro → Stop Recording**
   *(or **View → Macros ▾ → Stop Recording**).*

That's it — you just created the Personal Macro Workbook. (The throwaway
recording doesn't matter; we replace it in the next step.)

---

## Step 3 — Add the macro

1. Open the code editor: press **⌥ Option + F11**
   *(or **Tools → Macro → Visual Basic Editor**).*
2. On the left is the **Project** panel (if you don't see it, **View → Project
   Explorer**). Find and click **VBAProject (PERSONAL.XLSB)**.
3. **Import the macro:**
   - **File → Import File…**, choose **`FormatSquarespaceOrders.bas`**, click **Open**.
   - *If there's no Import option:* instead do **Insert → Module**, then open
     `FormatSquarespaceOrders.bas` in TextEdit, copy **everything except the
     very first line** (the line starting with `Attribute`), and paste it into
     the module.
4. Press **⌘S** to save. If asked, keep the format as-is (it saves the Personal
   Macro Workbook).
5. Close the code editor (**⌘Q** closes just the editor, back to Excel).

---

## Step 4 — Pin the "Format Orders" button

1. Menu bar: **Excel → Settings → Ribbon & Toolbar**.
2. Click the **Quick Access Toolbar** tab.
3. Under **Choose commands from:** pick **Macros**.
4. In the list, click **FormatSquarespaceOrders** (it may show as
   `PERSONAL.XLSB!FormatSquarespaceOrders`).
5. Click **▶ (Add)** to move it to the right-hand list. Click **Save**.

A little button now sits at the **top-left of every Excel window**. Hovering
over it shows "FormatSquarespaceOrders." That's the button she clicks.

---

## Step 5 — Make CSV files open in Excel (recommended)

By default a Mac may open `.csv` files in **Numbers** instead of Excel. Fix it
once:

1. In Finder, **right-click any `.csv`** file → **Get Info**.
2. Expand **Open with:** and choose **Microsoft Excel**.
3. Click **Change All…** and confirm.

Now every downloaded orders file opens straight into Excel.

---

## Step 6 — Test it

1. Open the included **`EXAMPLE-orders.csv`** in Excel (double-click).
2. Click the **Format Orders** button on the toolbar.
3. A new workbook should appear — names filled in, grouped by customer, five
   tidy centered columns. A "Done!" message tells you how many orders it
   processed.

If anything looks off or you get an error, note the exact message and send it
over — it's an easy fix.

---

## What your wife does from now on (the whole routine)

1. **Squarespace → Orders → Export** — downloads the orders CSV.
2. **Double-click** the downloaded file — it opens in Excel.
3. Click **Format Orders**.
4. **⌘S** to save the clean sheet, or **⌘P** to print the packing list.

No setup, no terminal, no typing — just open and click.
