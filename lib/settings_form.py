"""
Downloadable/uploadable Excel "settings form" -- an alternative way to set
an organization's colors, fiscal-year display order, and axis minimums,
for people who'd rather fill in an Excel sheet than click through the web
page's color pickers. Both paths feed the same OrgProfile fields and lead
to the same result.

Layout (fixed cell addresses, shared between build_settings_form and
parse_settings_form so they can never drift apart):

  Colors                          value cell
    Banner - primary                C5
    Banner - secondary              C6
    Accent strip                    C7
    Bar chart - Actual               C8
    Bar chart - Budgeted            C9
    Line series 1                   C10
    Line series 2                   C11
    Line series 3                   C12
    Line series 4                   C13

  Fiscal year display order (1-4, must be the 4 numbers 1-4 with no repeats)
    Line series 1 -> position       C17
    Line series 2 -> position       C18
    Line series 3 -> position       C19
    Line series 4 -> position       C20

  Axis minimums ($)
    Contributions bar chart          C23
    Income/Expense line charts       C24
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from lib.org_profile import OrgProfile

SHEET_NAME = "Settings"

COLOR_CELLS = {
    "banner_primary": ("B5", "C5", "Banner - primary"),
    "banner_secondary": ("B6", "C6", "Banner - secondary"),
    "accent": ("B7", "C7", "Accent strip"),
    "bar_actual_color": ("B8", "C8", "Bar chart - Actual"),
    "bar_budget_color": ("B9", "C9", "Bar chart - Budgeted"),
}
LINE_COLOR_CELLS = ["C10", "C11", "C12", "C13"]  # line_colors[0..3]
DISPLAY_ORDER_CELLS = ["C17", "C18", "C19", "C20"]  # display_order slot for series 1..4
BAR_AXIS_MIN_CELL = "C23"
LINE_AXIS_MIN_CELL = "C24"


class SettingsFormError(ValueError):
    """Raised when an uploaded settings form has missing or invalid values."""


def build_settings_form(profile: OrgProfile) -> bytes:
    """Return .xlsx file bytes: a filled-in-with-current-values settings form."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 28

    bold = Font(bold=True)
    title_font = Font(bold=True, size=14)

    ws["B2"] = "Organization Dashboard Styler -- Settings Form"
    ws["B2"].font = title_font
    ws["B3"] = ("Fill in the values below, save this file, then upload it in the app "
                "under \"Upload a filled-in settings form.\"")
    ws["B3"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("B3:D3")
    ws.row_dimensions[3].height = 30

    ws["B4"] = "COLORS (hex, e.g. #14495A)"
    ws["B4"].font = bold

    for field, (label_cell, value_cell, label) in COLOR_CELLS.items():
        ws[label_cell] = label
        value = getattr(profile, field)
        ws[value_cell] = value
        ws[value_cell].fill = PatternFill(start_color="FF" + value.lstrip("#").upper(),
                                           end_color="FF" + value.lstrip("#").upper(),
                                           fill_type="solid")

    for i, cell in enumerate(LINE_COLOR_CELLS):
        ws[f"B{10 + i}"] = f"Line series {i + 1}"
        value = profile.line_colors[i] if i < len(profile.line_colors) else "#888888"
        ws[cell] = value
        ws[cell].fill = PatternFill(start_color="FF" + value.lstrip("#").upper(),
                                     end_color="FF" + value.lstrip("#").upper(),
                                     fill_type="solid")

    ws["B15"] = "FISCAL YEAR DISPLAY ORDER (1-4, no repeats)"
    ws["B15"].font = bold
    ws["B16"] = "1 = listed first in the legend AND drawn on top of the other lines."
    ws.merge_cells("B16:D16")

    display_order = profile.display_order or [0, 1, 2, 3]
    # Convert internal 0-indexed slot-order back into "series N -> position" form
    # for the form: position[series_index] = slot + 1
    positions_by_series = [0, 0, 0, 0]
    for slot, series_idx in enumerate(display_order):
        if 0 <= series_idx < 4:
            positions_by_series[series_idx] = slot + 1

    for i, cell in enumerate(DISPLAY_ORDER_CELLS):
        ws[f"B{17 + i}"] = f"Line series {i + 1} -> display position"
        ws[cell] = positions_by_series[i] or (i + 1)

    dv = DataValidation(type="whole", operator="between", formula1=1, formula2=4,
                         showErrorMessage=True, errorTitle="Invalid position",
                         error="Enter a whole number from 1 to 4.")
    ws.add_data_validation(dv)
    for cell in DISPLAY_ORDER_CELLS:
        dv.add(ws[cell])

    ws["B22"] = "AXIS MINIMUMS ($)"
    ws["B22"].font = bold
    ws["B23"] = "Contributions bar chart"
    ws[BAR_AXIS_MIN_CELL] = profile.bar_axis_min
    ws["B24"] = "Income/Expense line charts"
    ws[LINE_AXIS_MIN_CELL] = profile.line_axis_min

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def parse_settings_form(file_like) -> dict:
    """
    Read a filled-in settings form (path or file-like/bytes buffer) and return
    a dict of OrgProfile field updates: colors, display_order, axis minimums.

    Raises SettingsFormError with a plain-English message if a value is
    missing or the display-order column isn't a clean 1-4 permutation.
    """
    wb = openpyxl.load_workbook(file_like, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise SettingsFormError(
            f"This doesn't look like a settings form -- no '{SHEET_NAME}' tab found."
        )
    ws = wb[SHEET_NAME]

    updates: dict = {}

    for field, (_, value_cell, label) in COLOR_CELLS.items():
        value = ws[value_cell].value
        if not value or not str(value).strip():
            raise SettingsFormError(f"'{label}' is empty -- please fill in a color.")
        updates[field] = _normalize_hex(str(value), label)

    line_colors = []
    for i, cell in enumerate(LINE_COLOR_CELLS):
        value = ws[cell].value
        if not value or not str(value).strip():
            raise SettingsFormError(f"'Line series {i + 1}' color is empty.")
        line_colors.append(_normalize_hex(str(value), f"Line series {i + 1}"))
    updates["line_colors"] = line_colors

    positions_by_series = []
    for i, cell in enumerate(DISPLAY_ORDER_CELLS):
        value = ws[cell].value
        try:
            pos = int(value)
        except (TypeError, ValueError):
            raise SettingsFormError(
                f"'Line series {i + 1} -> display position' must be a whole number from 1 to 4."
            )
        positions_by_series.append(pos)

    if sorted(positions_by_series) != [1, 2, 3, 4]:
        raise SettingsFormError(
            "The 4 display-position values must be exactly 1, 2, 3, and 4, each used once. "
            f"Found: {positions_by_series}."
        )

    # Convert "position per series" back into display_order[slot] = series_index
    display_order = [0, 0, 0, 0]
    for series_idx, pos in enumerate(positions_by_series):
        display_order[pos - 1] = series_idx
    updates["display_order"] = display_order

    for field, cell in (("bar_axis_min", BAR_AXIS_MIN_CELL), ("line_axis_min", LINE_AXIS_MIN_CELL)):
        value = ws[cell].value
        try:
            updates[field] = float(value)
        except (TypeError, ValueError):
            raise SettingsFormError(f"Axis minimum in cell {cell} must be a number.")

    return updates


def _normalize_hex(value: str, label: str) -> str:
    value = value.strip()
    hex_part = value.lstrip("#").upper()
    if len(hex_part) != 6 or any(c not in "0123456789ABCDEF" for c in hex_part):
        raise SettingsFormError(
            f"'{label}' isn't a valid hex color ({value!r}). Use a format like #14495A."
        )
    return "#" + hex_part
