#!/bin/bash
# Double-click this to turn a Squarespace orders CSV into a formatted Excel file.
# Keep it in the same folder as format_orders.py.
cd "$(dirname "$0")"

# 1. Find Python.
PY=""
for c in python3 python; do
  command -v "$c" >/dev/null 2>&1 && { PY="$c"; break; }
done
if [ -z "$PY" ]; then
  osascript -e 'display dialog "Python is not installed yet.\n\nInstall it from python.org, then double-click this again." buttons {"OK"} default button "OK" with icon caution'
  exit 1
fi

# 2. Make sure the one dependency is present.
"$PY" -c "import openpyxl" >/dev/null 2>&1 || "$PY" -m pip install --quiet --user openpyxl

# 3. Choose the CSV -- either dropped onto this icon, or via a picker.
CSV="$1"
if [ -z "$CSV" ]; then
  CSV=$(osascript -e 'POSIX path of (choose file with prompt "Choose your Squarespace orders export (.csv):" of type {"csv","public.comma-separated-values-text"})' 2>/dev/null)
fi
[ -z "$CSV" ] && exit 0

# 4. Run it and open the result.
OUT="${CSV%.*}_formatted.xlsx"
if "$PY" format_orders.py "$CSV" "$OUT"; then
  open "$OUT"
else
  osascript -e 'display dialog "Something went wrong. Make sure you picked a Squarespace orders CSV." buttons {"OK"} with icon caution'
fi
