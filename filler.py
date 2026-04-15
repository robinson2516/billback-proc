"""Fills the 1 REBILL BLANK.xlsm template with extracted data."""
import io
from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).parent
TEMPLATE = BASE_DIR / "1 REBILL_BLANK_updated.xlsm"

LINE_ITEM_START_ROW = 15   # line items begin at row 15

# Vendor names from Keys sheet (col A, rows 2–15)
_VENDOR_LIST = [
    "Peterbilt of Utah",
    "Vernal Petebilt",
    "St. George Peterbilt",
    "Idaho Falls Peterbilt",
    "Boise Peterbilt",
    "Caldwell Peterbilt",
    "Jerome Peterbilt",
    "Magic Valley Petebilt",
    "Grand Junction Peterbilt",
    "Elko Peterbilt",
    "Utah Valley Peterbilt",
    "Ogden Peterbilt",
    "Salina Peterbilt",
    "Ontario Peterbilt",
]


def _match_vendor(name: str) -> str:
    """Return the exact Keys-sheet vendor string closest to `name`, or `name` if no match."""
    if not name:
        return name
    normalized = name.lower()
    # Exact match first
    for v in _VENDOR_LIST:
        if v.lower() == normalized:
            return v
    # Any word in the extracted name matches a word in a vendor entry
    words = set(normalized.split())
    for v in _VENDOR_LIST:
        if any(w in v.lower().split() for w in words):
            return v
    return name


def fill_rebill_sheet(data: dict) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE, keep_vba=True)
    ws = wb["Rebill Sheet"]

    # ── Header fields ────────────────────────────────────────────
    ws.cell(row=1, column=4).value  = _match_vendor(data.get("vendor") or "")
    ws.cell(row=1, column=17).value = data.get("customer") or ""
    ws.cell(row=2, column=4).value  = data.get("invoice_number") or ""
    ws.cell(row=2, column=17).value = data.get("taxable") or "Yes"
    ws.cell(row=3, column=4).value  = data.get("date") or ""
    ws.cell(row=4, column=4).value  = data.get("lease_or_rental") or "Lease"
    ws.cell(row=4, column=15).value = data.get("invoice_wording") or ""

    # ── Misc / shop supplies → K9 ────────────────────────────────
    try:
        misc = float(data.get("misc") or 0)
    except (ValueError, TypeError):
        misc = 0.0
    ws.cell(row=9, column=11).value = misc if misc else None

    # ── Line items ────────────────────────────────────────────────
    # Each item goes into ONE column based on type:
    #   Col B (2)  — Rebill
    #   Col D (4)  — Internal PMs
    #   Col E (5)  — Internal Repairs
    line_items = data.get("line_items", [])
    total = 0.0
    for i, item in enumerate(line_items):
        row = LINE_ITEM_START_ROW + i
        try:
            cost = float(item.get("cost") or 0)
        except (ValueError, TypeError):
            cost = 0.0
        if cost == 0:
            continue

        total += cost
        item_type = (item.get("type") or "repair").lower()
        if item_type == "pm":
            ws.cell(row=row, column=4).value = cost
        elif item_type == "rebill":
            ws.cell(row=row, column=2).value = cost
        else:
            ws.cell(row=row, column=5).value = cost

    # ── Invoice total (line items + misc) → L12 ──────────────────
    inv_total = total + misc
    ws.cell(row=12, column=12).value = round(inv_total, 2) if inv_total else None

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
