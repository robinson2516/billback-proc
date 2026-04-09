"""Loads unit spreadsheet for Lease/Rental lookup."""
from pathlib import Path
from datetime import date
import openpyxl

BASE_DIR = Path(__file__).parent

_unit_data = None
_unit_loaded_date = None


def _load_units() -> dict:
    global _unit_data, _unit_loaded_date
    today = date.today()
    if _unit_data is not None and _unit_loaded_date == today:
        return _unit_data

    path = BASE_DIR / "Unit numbers and descriptions .xlsx"
    if not path.exists():
        _unit_data = {}
        return _unit_data

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        unit = row[0]
        if not unit:
            continue
        unit_str = str(unit).strip()
        name = str(row[11]).strip() if len(row) > 11 and row[11] else ""  # Col L: Lease/Rental
        if name.lower() in ("lease", "rental"):
            data[unit_str] = name.capitalize()

    _unit_data = data
    _unit_loaded_date = today
    return _unit_data


def lookup_unit(unit_number: str) -> str:
    """Returns 'Lease', 'Rental', or 'Lease' (default) for a given unit number."""
    if not unit_number:
        return "Lease"
    data = _load_units()
    return data.get(str(unit_number).strip(), "Lease")
