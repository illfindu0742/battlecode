/*
 * Squarespace order formatter — shared core logic.
 * Pure functions, no DOM. Runs in the browser (loaded as a plain <script>,
 * exposes window.OrderFormatter) and in Node (module.exports) for testing.
 *
 * buildWorkbook() takes ExcelJS as an argument so this file doesn't care where
 * the library comes from (global in the browser, require() in Node).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.OrderFormatter = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var ORDER_ID_MIN_WIDTH = 5;

  // Output columns, in the exact order requested.
  var COLUMNS = [
    { header: "Order ID", key: "orderId" },
    { header: "Lineitem quantity", key: "qty" },
    { header: "Lineitem name", key: "name" },
    { header: "Lineitem variant", key: "variant" },
    { header: "Shipping Name", key: "shipping" },
  ];

  // --- CSV parsing (handles quotes, embedded commas/newlines, "" escapes) ---
  function parseCSV(text) {
    var rows = [];
    var row = [];
    var field = "";
    var inQuotes = false;
    var i = 0;
    // strip a BOM if present
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    while (i < text.length) {
      var c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ",") { row.push(field); field = ""; i++; continue; }
      if (c === "\r") { i++; continue; }
      if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++;
    }
    // flush last field/row if the file didn't end with a newline
    if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
    return rows;
  }

  function clean(v) {
    if (v === undefined || v === null) return "";
    return String(v).trim();
  }

  function findCol(headers, name) {
    for (var c = 0; c < headers.length; c++) {
      if (clean(headers[c]).toLowerCase() === name.toLowerCase()) return c;
    }
    return -1;
  }

  function isAllDigits(s) {
    return s.length > 0 && /^[0-9]+$/.test(s);
  }

  function padOrder(s) {
    if (s.length === 0) return "";
    if (isAllDigits(s) && s.length < ORDER_ID_MIN_WIDTH) {
      return ("0".repeat(ORDER_ID_MIN_WIDTH - s.length)) + s;
    }
    return s;
  }

  // --- transform: matrix (incl. header row) -> sorted, filled, trimmed rows ---
  function transform(matrix) {
    if (!matrix || matrix.length < 2) {
      throw new Error("This file doesn't look like an orders export (no rows).");
    }
    var headers = matrix[0];
    var cOrder = findCol(headers, "Order ID");
    var cQty = findCol(headers, "Lineitem quantity");
    var cName = findCol(headers, "Lineitem name");
    var cVar = findCol(headers, "Lineitem variant");
    var cShip = findCol(headers, "Shipping Name");
    var cBill = findCol(headers, "Billing Name");

    var missing = [];
    if (cOrder < 0) missing.push("Order ID");
    if (cQty < 0) missing.push("Lineitem quantity");
    if (cName < 0) missing.push("Lineitem name");
    if (cVar < 0) missing.push("Lineitem variant");
    if (cShip < 0) missing.push("Shipping Name");
    if (missing.length) {
      throw new Error(
        "This doesn't look like a Squarespace orders export. Missing column(s): " +
          missing.join(", ") + "."
      );
    }

    // pass 1: first non-empty shipping name per order id
    var shipByOrder = {};
    var r, oid, s;
    for (r = 1; r < matrix.length; r++) {
      oid = clean(matrix[r][cOrder]);
      if (!oid) continue;
      s = clean(matrix[r][cShip]);
      if (s && !(oid in shipByOrder)) shipByOrder[oid] = s;
    }
    // fallback to billing name for orders with no shipping name at all
    if (cBill >= 0) {
      for (r = 1; r < matrix.length; r++) {
        oid = clean(matrix[r][cOrder]);
        if (!oid || oid in shipByOrder) continue;
        var b = clean(matrix[r][cBill]);
        if (b) shipByOrder[oid] = b;
      }
    }

    // build output rows (skip fully-blank rows)
    var out = [];
    for (r = 1; r < matrix.length; r++) {
      oid = clean(matrix[r][cOrder]);
      var name = clean(matrix[r][cName]);
      if (!oid && !name) continue;
      s = clean(matrix[r][cShip]);
      if (!s) s = shipByOrder[oid] || "";
      var qtyRaw = clean(matrix[r][cQty]);
      var qty = /^[0-9]+$/.test(qtyRaw) ? parseInt(qtyRaw, 10) : qtyRaw;
      out.push({
        orderId: padOrder(oid),
        qty: qty,
        name: name,
        variant: clean(matrix[r][cVar]),
        shipping: s,
        _i: out.length, // stable tiebreaker
      });
    }
    if (!out.length) throw new Error("No order rows found in this file.");

    // sort: shipping name, then order id, then original order
    out.sort(function (a, b) {
      var sa = a.shipping.toLowerCase(), sb = b.shipping.toLowerCase();
      if (sa < sb) return -1; if (sa > sb) return 1;
      if (a.orderId < b.orderId) return -1; if (a.orderId > b.orderId) return 1;
      return a._i - b._i;
    });

    var custSet = {};
    out.forEach(function (row) { custSet[row.shipping.toLowerCase()] = true; });
    return { rows: out, lineItems: out.length, customers: Object.keys(custSet).length };
  }

  // --- build a styled ExcelJS workbook from the transformed rows ---
  function buildWorkbook(ExcelJS, result) {
    var rows = result.rows;
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet("Orders", { views: [{ state: "frozen", ySplit: 1 }] });

    // column widths from content (capped)
    var caps = { orderId: 14, qty: 18, name: 62, variant: 40, shipping: 34 };
    ws.columns = COLUMNS.map(function (col) {
      var longest = col.header.length;
      rows.forEach(function (row) {
        var len = String(row[col.key] == null ? "" : row[col.key]).length;
        if (len > longest) longest = len;
      });
      var width = Math.min(caps[col.key], Math.max(10, longest + 2));
      return { header: col.header, key: col.key, width: width };
    });

    var center = { horizontal: "center", vertical: "center", wrapText: false };
    var thin = { style: "thin", color: { argb: "FFD9D9D9" } };
    var border = { top: thin, left: thin, bottom: thin, right: thin };

    // header styling
    ws.getRow(1).eachCell(function (cell) {
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF44546A" } };
      cell.alignment = center;
      cell.border = border;
    });
    ws.getRow(1).height = 20;

    // data rows
    var band = false, prevKey = null;
    rows.forEach(function (row) {
      var excelRow = ws.addRow({
        orderId: row.orderId,
        qty: row.qty,
        name: row.name,
        variant: row.variant,
        shipping: row.shipping,
      });
      var key = row.shipping.toLowerCase();
      if (key !== prevKey) { band = !band; prevKey = key; }
      excelRow.eachCell(function (cell, colNumber) {
        cell.font = { name: "Arial" };
        cell.alignment = center;
        cell.border = border;
        if (colNumber === 1) cell.numFmt = "@"; // Order ID as text (keep leading zeros)
        if (band) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
        }
      });
    });

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length + 1, column: COLUMNS.length } };
    return wb;
  }

  return {
    ORDER_ID_MIN_WIDTH: ORDER_ID_MIN_WIDTH,
    COLUMNS: COLUMNS,
    parseCSV: parseCSV,
    transform: transform,
    buildWorkbook: buildWorkbook,
  };
});
