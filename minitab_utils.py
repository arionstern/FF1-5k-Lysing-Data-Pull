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
    """Format a fill date as a Boxplot column name (base name only,
    no dedup check — see determine_unique_boxplot_column_name for the
    real duplicate-date logic used by write_new_lot_to_minitab)."""
    formatted = f"{fill_date.month}/{fill_date.day}/{fill_date.year}"
    if is_duplicate:
        formatted += "_1"
    return formatted


def get_existing_boxplot_column_names(worksheet):
    """Read every existing column name on the Boxplot worksheet."""
    return [
        worksheet.Columns.Item(i).Name
        for i in range(1, worksheet.Columns.Count + 1)
    ]


def get_boxplot_column_names_from_year(worksheet, start_year=None):
    """Filter existing Boxplot column names to everything from
    start_year onward (NOT a single rolling year — cumulative, so this
    doesn't reset/exclude 2026 data once 2027 starts). Parsed
    dynamically from each column's date-based name.
    Defaults to config.CHART_DATA_START_YEAR.

    Returns names sorted by ACTUAL PARSED DATE, not worksheet position
    — these normally match (new columns append rightmost), but break
    if a column ever lands out of position (e.g. stray empty columns
    shifting a write, or backfilling an older lot after newer ones
    already exist). Confirmed real bug: a column at a later position
    but earlier date plotted in the wrong order on the chart.
    """
    from datetime import datetime
    start_year = start_year or config.CHART_DATA_START_YEAR

    all_names = get_existing_boxplot_column_names(worksheet)
    dated_names = []
    for name in all_names:
        # Names look like "6/10/2026" or "6/10/2026_1" — strip any
        # "_N" suffix before parsing the date
        base = name.split("_")[0]
        parts = base.split("/")
        if len(parts) != 3:
            continue
        try:
            month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
            if year < 100:  # 2-digit year like "26" -> 2026, confirmed
                             # real existing columns use this format
                year += 2000
            if year >= start_year:
                dated_names.append((datetime(year, month, day), name))
        except ValueError:
            continue

    dated_names.sort(key=lambda x: x[0])
    return [name for _, name in dated_names]


def determine_unique_boxplot_column_name(worksheet, fill_date):
    """Determine the correct column name for a new lot, automatically
    detecting same-day duplicates by checking existing column names —
    matches the original written instructions' '_1' suffix rule, but
    derived from real worksheet state instead of a caller-provided flag.

    First lot on a date: '7/1/2026'. Second: '7/1/2026_1'. Third (if
    it ever happens): '7/1/2026_2', etc. — not capped at one duplicate,
    since the original instructions only ever showed 2/day as an
    example, not a hard limit.
    """
    base_name = f"{fill_date.month}/{fill_date.day}/{fill_date.year}"
    existing_names = get_existing_boxplot_column_names(worksheet)

    if base_name not in existing_names:
        return base_name

    suffix = 1
    while f"{base_name}_{suffix}" in existing_names:
        suffix += 1
    return f"{base_name}_{suffix}"


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
# Boxplot chart generation
# ---------------------------------------------------------------------------

def regenerate_boxplot_chart(project, boxplot_sheet=None):
    """Regenerate the combined Boxplot chart (Multiple Y's, Simple)
    covering all data from config.CHART_DATA_START_YEAR onward.

    Confirmed working syntax, including title AND axis labels:
    range selection + Overlay/IQRBox/Outlier + Title + AxLabel 1/2.
    AxLabel is a SESSION SUBCOMMAND (found via Minitab's own
    "Help Boxplot." documentation) — it must be nested inside the
    Boxplot command's own subcommand block, not issued as a separate
    command afterward (that fails with "Unknown Minitab command").
    K=1 targets the X-axis, K=2 targets the Y-axis.

    Unlike Xbar (which auto-updates on the Control Chart worksheet as
    new rows are added), Boxplot does NOT auto-update since each new
    lot is a brand-new column — this needs to be re-run every time new
    data is written.
    """
    boxplot_sheet = boxplot_sheet or get_worksheet(project, config.DEST_BOXPLOT_SHEET)

    column_names = get_boxplot_column_names_from_year(boxplot_sheet)
    if not column_names:
        raise ValueError(
            f"No columns found from {config.CHART_DATA_START_YEAR} "
            f"onward — nothing to chart."
        )

    # Explicit column list, NOT a 'first'-'last' range — a range spans
    # every column POSITION in between, including any stray empty
    # columns (e.g. leftover from testing), which breaks the chart
    # with "No data in column CN" even though our real data is fine.
    # column_names already excludes unnamed/empty columns (they fail
    # the date-parsing check in get_boxplot_column_names_from_year).
    # NOTE: this specific combination (explicit column LIST + Overlay)
    # hasn't been directly tested yet — we've confirmed range+Overlay
    # works, and separately confirmed a plain list WITHOUT Overlay
    # produces separate tiled charts (bad). Worth verifying this
    # combination produces one combined chart, not tiled ones.
    column_list = " ".join(f"'{name}'" for name in column_names)
    chart_config = config.MINITAB_BOXPLOT_CHART

    command_text = (
        f"Boxplot {column_list};\n"
        f"  Overlay;\n"
        f"  IQRBox;\n"
        f"  Outlier;\n"
        f"  Title \"{chart_config['title']}\";\n"
        f"  AxLabel 1 \"{chart_config['x_axis']}\";\n"
        f"  AxLabel 2 \"{chart_config['y_axis']}\"."
    )

    commands_before = project.Commands.Count
    project.ExecuteCommand(command_text)
    commands_after = project.Commands.Count

    if commands_after <= commands_before:
        raise RuntimeError(
            "Boxplot command did not create a new chart — check "
            "Minitab's Session window for the real error (ExecuteCommand "
            "does not raise on Minitab-side syntax errors)."
        )

    return project.Commands.Item(commands_after)


# ---------------------------------------------------------------------------
# Orchestration helper
# ---------------------------------------------------------------------------

def write_new_lot_to_minitab(project, lot_number, wo_number, fill_date,
                              seal_values):
    """Writes one lot's data to both Minitab worksheets. Same-day
    duplicate detection is automatic — no flag needed from the caller."""
    boxplot_sheet = get_worksheet(project, config.DEST_BOXPLOT_SHEET)
    control_chart_sheet = get_worksheet(project, config.DEST_CONTROL_CHART_SHEET)

    column_name = determine_unique_boxplot_column_name(boxplot_sheet, fill_date)
    boxplot_col = write_boxplot_column(boxplot_sheet, column_name, seal_values)

    control_chart_row = write_control_chart_row(
        control_chart_sheet, lot_number, wo_number, fill_date, seal_values
    )

    return {
        "boxplot_column": boxplot_col,
        "boxplot_column_name": column_name,
        "control_chart_row": control_chart_row,
    }