"""Fills the 1 REBILL BLANK.xlsm template with extracted data."""
import io
from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).parent
TEMPLATE = BASE_DIR / "1 REBILL BLANK.xlsm"

LINE_ITEM_START_ROW = 15   # line items begin at row 15


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

    # ── Misc / shop supplies → Row 9, Col K (11) ─────────────────
    try:
        misc = float(data.get("misc") or 0)
    except (ValueError, TypeError):
        misc = 0.0
    if misc:
        ws.cell(row=9, column=11).value = misc

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

    # ── Invoice total → Row 12, Col L (12) ───────────────────────
    if total:
        ws.cell(row=12, column=12).value = total

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
