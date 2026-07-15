"""
minitab_utils.py

Reusable functions for writing new lot data into the FF1 5K Lysing
Minitab project (.mpx), confirmed against the real worksheet structure:

  "Tensile Data (Boxplot)" worksheet:
    - Each lot is its own COLUMN.
    - Column NAME is the fill date as a string (e.g. '9/5/2025'), with
      a '_1' suffix for a second same-day lot (matches the original
      written instructions' duplicate-lot rule exactly).
    - Always exactly 38 rows. Shorter lots are padded with Minitab's
      own native missing-value marker ('*') — confirmed visible in
      real data, not something we need custom logic to emulate.

  "Tensile Data (Control Chart)" worksheet:
    - C1 = Lot No. (Text), C2 = WO No., C3 = Fill Date (Date/Time),
      C4...C41 = Bag 1...Bag 38.
    - Each lot is its own ROW. Since every row already has the same
      fixed 38 bag columns, there's no padding problem here — that's
      specific to the Boxplot sheet.

NOTE ON DATE FORMATTING: existing Boxplot column names are
inconsistent (some 2-digit year like '6/10/26', some 4-digit like
'6/24/2026'). New columns written by this module use a single
consistent format (see BOXPLOT_DATE_FORMAT below) rather than adding
a third variant.
"""

import os
import math
from datetime import datetime, date
import win32com.client
import config


MINITAB_MISSING_VALUE = float("nan")  # Confirmed correct: Minitab's
                                       # real internal missing-value
                                       # sentinel reads back as
                                       # 1.23456e+30 via COM — that's
                                       # expected, not a bug. Displays
                                       # as '*' in the Minitab GUI.


# ---------------------------------------------------------------------------
# Project / worksheet access
# ---------------------------------------------------------------------------

def open_minitab_project(project_path=None, visible=True):
    """Open the Minitab project via COM. Returns (mtb, project)."""
    project_path = project_path or config.MINITAB_PROJECT_PATH
    if not os.path.exists(project_path):
        raise FileNotFoundError(f"Minitab project not found: {project_path}")

    mtb = win32com.client.Dispatch("Mtb.Application.1")
    mtb.UserInterface.Visible = visible
    mtb.Open(project_path)

    project = mtb.ActiveProject
    return mtb, project


def get_worksheet(project, sheet_name):
    """Get a specific worksheet by name (not just ActiveWorksheet,
    since this project has multiple worksheets)."""
    for i in range(1, project.Worksheets.Count + 1):
        worksheet = project.Worksheets.Item(i)
        if worksheet.Name == sheet_name:
            return worksheet
    raise ValueError(f"Worksheet {sheet_name!r} not found in project.")


# ---------------------------------------------------------------------------
# Boxplot worksheet
# ---------------------------------------------------------------------------

def format_boxplot_column_name(fill_date, is_duplicate=False):
    """Format a fill date as a Boxplot column name, matching the
    existing convention (with '_1' suffix for same-day duplicates,
    per the original written instructions).

    NOTE: deliberately not using strftime's leading-zero-strip flags
    (%-m/%-d are Linux-only, %#m/%#d are Windows-only) — building the
    string manually instead so this works regardless of platform.
    """
    formatted = f"{fill_date.month}/{fill_date.day}/{fill_date.year}"
    if is_duplicate:
        formatted += "_1"
    return formatted


def write_boxplot_column(worksheet, column_name, seal_values, max_rows=38):
    """Write one lot's data as a new column on the Boxplot worksheet.
    Pads to max_rows with Minitab's native missing-value marker if the
    lot has fewer real readings than that."""
    if len(seal_values) > max_rows:
        raise ValueError(
            f"{len(seal_values)} readings exceeds max_rows={max_rows} — "
            f"the worksheet's fixed row count may need to increase."
        )

    padded_values = list(seal_values) + [MINITAB_MISSING_VALUE] * (
        max_rows - len(seal_values)
    )

    # Columns.Add() is the real method for creating a new column
    # (confirmed via exploration — Item() alone can't create one,
    # despite accepting an index argument that suggested it might).
    new_column = worksheet.Columns.Add()
    new_column.Name = column_name
    new_column.SetData(padded_values)

    return new_column.Number


# ---------------------------------------------------------------------------
# Control Chart worksheet
# ---------------------------------------------------------------------------

def find_next_open_control_chart_row(worksheet):
    """Find the first empty row, using the Lot No. column (C1) as the
    reference — always populated for real rows."""
    lot_column = worksheet.Columns.Item(1)
    return lot_column.RowCount + 1


def write_control_chart_row(worksheet, lot_number, wo_number, fill_date,
                             seal_values, max_bags=38):
    """Write one lot's data as a new row on the Control Chart
    worksheet: C1=Lot No., C2=WO No., C3=Fill Date, C4+=Bag 1...Bag N.
    Unused bag columns for this row are left as missing (Minitab
    handles ragged row lengths within a column naturally — no need to
    pre-pad every column, unlike the Boxplot sheet)."""
    if len(seal_values) > max_bags:
        raise ValueError(
            f"{len(seal_values)} readings exceeds max_bags={max_bags}."
        )

    row = find_next_open_control_chart_row(worksheet)

    # COM date marshalling typically expects datetime.datetime, not
    # datetime.date — a raw date object likely won't write correctly.
    # This conversion is confirmed as necessary in the old project's
    # convert_excel_values_for_minitab().
    if isinstance(fill_date, date) and not isinstance(fill_date, datetime):
        fill_date = datetime(fill_date.year, fill_date.month, fill_date.day)

    columns = worksheet.Columns
    columns.Item(1).SetData([str(lot_number)], row)
    # WO No. column is Numeric (confirmed: no "-T" text-type suffix
    # in the real worksheet), so this needs an actual number, not a
    # string — a mismatched type raises a COM error otherwise.
    columns.Item(2).SetData([int(wo_number)], row)
    columns.Item(3).SetData([fill_date], row)

    for i, value in enumerate(seal_values):
        bag_col_index = 4 + i  # C4 = Bag 1
        columns.Item(bag_col_index).SetData([value], row)

    # Explicitly pad remaining, unused Bag columns for this row with
    # the missing-value marker — leaving them untouched resulted in
    # genuinely empty cells rather than a proper '*', which is
    # inconsistent with how the Boxplot sheet represents missing data.
    for i in range(len(seal_values), max_bags):
        bag_col_index = 4 + i
        columns.Item(bag_col_index).SetData([MINITAB_MISSING_VALUE], row)

    return row


# ---------------------------------------------------------------------------
# Orchestration helper
# ---------------------------------------------------------------------------

def write_new_lot_to_minitab(project, lot_number, wo_number, fill_date,
                              seal_values, is_duplicate_date=False):
    """Writes one lot's data to both Minitab worksheets."""
    boxplot_sheet = get_worksheet(project, config.DEST_BOXPLOT_SHEET)
    control_chart_sheet = get_worksheet(project, config.DEST_CONTROL_CHART_SHEET)

    column_name = format_boxplot_column_name(fill_date, is_duplicate_date)
    boxplot_col = write_boxplot_column(boxplot_sheet, column_name, seal_values)

    control_chart_row = write_control_chart_row(
        control_chart_sheet, lot_number, wo_number, fill_date, seal_values
    )

    return {
        "boxplot_column": boxplot_col,
        "boxplot_column_name": column_name,
        "control_chart_row": control_chart_row,
    }