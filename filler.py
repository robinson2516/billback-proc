"""Fills the 1 REBILL BLANK.xlsm template with extracted data."""
import io
from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).parent
TEMPLATE = BASE_DIR / "1 REBILL BLANK.xlsm"

# Row where individual line items begin in the template
LINE_ITEM_START_ROW = 18


def fill_rebill_sheet(data: dict) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE, keep_vba=True)
    ws = wb["Rebill Sheet"]

    # ── Header fields ────────────────────────────────────────────
    ws.cell(row=1, column=4).value  = data.get("vendor") or ""
    ws.cell(row=1, column=17).value = data.get("customer") or ""
    ws.cell(row=2, column=4).value  = data.get("invoice_number") or ""
    ws.cell(row=2, column=17).value = data.get("taxable") or "Yes"
    ws.cell(row=3, column=4).value  = data.get("date") or ""
    ws.cell(row=4, column=4).value  = data.get("lease_or_rental") or "Lease"
    ws.cell(row=4, column=15).value = data.get("invoice_wording") or ""

    # ── Line items ───────────────────────────────────────────────
    # Each item occupies one row. Costs go into:
    #   Col B (2) — Rebill       (only if item["rebill"] is True)
    #   Col D (4) — Internal PMs  (if type == "pm")
    #   Col E (5) — Internal Repairs (if type == "repair")
    line_items = data.get("line_items", [])
    for i, item in enumerate(line_items):
        row = LINE_ITEM_START_ROW + i

        try:
            cost = float(item.get("cost") or 0)
        except (ValueError, TypeError):
            cost = 0.0
        if cost == 0:
            continue

        item_type = (item.get("type") or "repair").lower()
        if item_type == "pm":
            ws.cell(row=row, column=4).value = cost   # Internal PMs
        else:
            ws.cell(row=row, column=5).value = cost   # Internal Repairs

        if item.get("rebill"):
            ws.cell(row=row, column=2).value = cost   # Rebill

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
