"""
excel_utils.py

Reusable functions for reading the source GEM log sheet (embedded in
SAP's Office Integration) and writing to the three destination sheets
in FF1 5k Lysing Tensile Data.xlsx: WO Data, Tensile Data (Boxplot),
and Tensile Data (Control Chart).

All cell/column references here are confirmed against real data — see
tests/sap/test_locate_fill_line_and_data.py for how the source layout
was found, and the chat history for how the destination layout was
confirmed via screenshots.

"Next open row/column" lookups use Excel's End(xlUp)/End(xlToLeft)
navigation (same as pressing Ctrl+Up or Ctrl+Left from the bottom/right
edge of the sheet) rather than looping cell-by-cell — much faster once
these sheets have hundreds of rows/columns of history, borrowed from
the pattern used in the last project's excel_utils.py.
"""

import win32com.client
import config


# Excel COM constants (win32com doesn't expose these by name the way
# VBA does, so they're spelled out here)
XL_UP = -4162
XL_TO_LEFT = -4159


# ---------------------------------------------------------------------------
# Source sheet (SOPGEM-45-2, embedded in SAP)
# ---------------------------------------------------------------------------

def attach_to_source_excel():
    """Attach to the Excel process SAP's Office Integration opened.
    Returns the Excel Application COM object.

    NOTE: ActiveWorkbook is unreliable here (returns None) since COM
    attachment doesn't give Excel a "focused" window — callers should
    grab the workbook by index instead, e.g. excel_app.Workbooks(1).
    """
    return win32com.client.GetObject(Class="Excel.Application")


def read_fill_line(sheet):
    """Read the header-level fill line (step 4.5.1), e.g. 'FF1'."""
    return sheet.Range(config.SOURCE_FILL_LINE_CELL).Value


def validate_fill_line(fill_line):
    """Returns True if fill_line is a valid value per config."""
    return fill_line in config.VALID_FILL_LINES


def read_lot_metadata(sheet):
    """Read product name, part number, WO number from the source
    sheet header. NOTE: part_number was missing entirely until this
    fix — Run_Lysing_Pull.py's WO Data append would have silently
    written a blank Part Number for every genuinely new lot."""
    return {
        "product_name": sheet.Range(config.SOURCE_PRODUCT_NAME_CELL).Value,
        "part_number": sheet.Range(config.SOURCE_PART_NUMBER_CELL).Value,
        "wo_number": sheet.Range(config.SOURCE_WO_NUMBER_CELL).Value,
    }


def read_fill_date(workbook):
    """Read the real fill date from the dedicated 'GEM Fill Logs
    Header Page' sheet (F16) — a DIFFERENT sheet than SOURCE_SHEET_NAME
    (which only has fill_line/product_name/etc in its own embedded
    header). Deliberately raises a clear error if this cell is blank
    or not a real date, rather than silently falling back to a date
    derived from the lot code — a missing/invalid value here likely
    means the header page wasn't filled out correctly, which is worth
    surfacing rather than masking with a guessed value."""
    from datetime import datetime, date as date_type

    try:
        header_sheet = workbook.Sheets(config.GEM_FILL_LOGS_HEADER_SHEET_NAME)
    except Exception as e:
        raise ValueError(
            f"Could not find sheet {config.GEM_FILL_LOGS_HEADER_SHEET_NAME!r} "
            f"in the source workbook: {e}"
        )

    raw_value = header_sheet.Range(config.FILL_DATE_CELL).Value

    if raw_value is None or raw_value == "":
        raise ValueError(
            f"Fill date cell {config.FILL_DATE_CELL!r} on "
            f"{config.GEM_FILL_LOGS_HEADER_SHEET_NAME!r} is blank — "
            f"the header page likely wasn't filled out correctly for "
            f"this lot. Not defaulting to a lot-code-derived date, "
            f"since that could mask a real data problem."
        )

    if isinstance(raw_value, datetime):
        return raw_value
    if isinstance(raw_value, date_type):
        return datetime(raw_value.year, raw_value.month, raw_value.day)

    raise ValueError(
        f"Fill date cell {config.FILL_DATE_CELL!r} on "
        f"{config.GEM_FILL_LOGS_HEADER_SHEET_NAME!r} contains "
        f"{raw_value!r}, which isn't a recognizable date."
    )


def read_seal_strength_values(sheet):
    """Read the real (non-placeholder) Seal Strength values from the
    Tensile Tester Use Log table, stopping at the first true blank row
    (the table is a fixed-size template padded with 'N/A' text for
    unused rows — those get filtered out, not treated as data).

    NOTE: this one still loops row-by-row rather than using End(xlUp),
    since the table is a fixed-size template with 'N/A' placeholders
    scattered inside real data (not a clean contiguous block) — a
    single End() jump can't distinguish "real data ends here" from
    "template placeholder starts here." Loop is capped at 500 rows as
    a safety net, well above the largest template seen so far.

    Returns a list of float values, in row order (i.e. bag 1 first).
    """
    seal_col = config.SOURCE_TABLE_COLUMNS["seal_strength"]
    start_row = config.SOURCE_TABLE_DATA_START_ROW

    values = []
    row = start_row
    max_rows_to_check = 500  # safety cap
    while row < start_row + max_rows_to_check:
        cell_value = sheet.Range(f"{seal_col}{row}").Value
        if cell_value is None or cell_value == "":
            break  # true end of the template
        if isinstance(cell_value, (int, float)):
            values.append(float(cell_value))
        # non-numeric (e.g. 'N/A') rows are silently skipped, not
        # treated as the end of data — real rows can be followed by
        # more placeholder rows within the same fixed-size template
        row += 1

    return values


# ---------------------------------------------------------------------------
# Destination — WO Data sheet
# ---------------------------------------------------------------------------

def find_next_open_wo_data_row(sheet):
    """Find the first empty row below the existing data, using
    End(xlUp) from the bottom of the sheet instead of looping down
    from row 2 — checks the Fill Lot # column, since that's always
    populated for real rows."""
    max_row = sheet.Rows.Count
    last_used_row = sheet.Range(
        f"{config.COL_LOT_NUMBER}{max_row}"
    ).End(XL_UP).Row

    # If the sheet is genuinely empty below the header, End(xlUp) from
    # the bottom lands back on row 1 (the header) — handle that case
    # explicitly rather than returning row 2 by accident of arithmetic.
    if last_used_row == 1 and sheet.Range(f"{config.COL_LOT_NUMBER}1").Value in (None, ""):
        return 2

    return last_used_row + 1


def get_last_known_lot(sheet):
    """Get the most recent lot number in WO Data column E, using the
    same End(xlUp) trick as find_next_open_wo_data_row — the row just
    above the next-open row is the last known lot. Used by
    sap_utils.find_new_lots() as the stopping point for its scan."""
    next_open_row = find_next_open_wo_data_row(sheet)
    last_row = next_open_row - 1
    if last_row < 2:
        return None  # sheet is genuinely empty
    return str(sheet.Range(f"{config.COL_LOT_NUMBER}{last_row}").Value)


def append_wo_data_row(sheet, fill_date, filler, part_number, product_name,
                        lot_number, wo_number):
    """Append a new row to WO Data with the given values.
    Column mapping confirmed against real headers:
    A=Fill Date, B=Filler, C=Part Number, D=Product Name,
    E=Fill Lot #, F=Fill WO #.
    """
    row = find_next_open_wo_data_row(sheet)
    sheet.Range(f"A{row}").Value = fill_date
    sheet.Range(f"B{row}").Value = filler
    sheet.Range(f"C{row}").Value = part_number
    sheet.Range(f"D{row}").Value = product_name
    sheet.Range(f"E{row}").Value = lot_number
    sheet.Range(f"F{row}").Value = wo_number
    return row


# ---------------------------------------------------------------------------
# Destination — Tensile Data (Boxplot) sheet
# ---------------------------------------------------------------------------

def _column_letter(col_index):
    """Convert a 1-based column index to an Excel column letter."""
    letters = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def find_next_open_boxplot_column(sheet, header_row=None):
    """Find the first empty column to the right of existing lot data,
    using End(xlToLeft) from the right edge of the sheet instead of
    looping right from column B."""
    header_row = header_row or config.DEST_BOXPLOT_HEADER_ROWS["lot"]
    max_col = sheet.Columns.Count
    last_used_col = sheet.Cells(header_row, max_col).End(XL_TO_LEFT).Column

    # If column A (row labels) is the only thing populated, End(xlToLeft)
    # lands on column A itself — handle explicitly so data starts at B.
    if last_used_col == 1:
        return 2

    return last_used_col + 1


def write_boxplot_data(sheet, lot_number, wo_number, fill_date, seal_values):
    """Write one lot's data into the next open column on the Boxplot
    sheet: lot/WO/date in the header rows, seal values going down
    starting at DEST_BOXPLOT_DATA_START_ROW."""
    col_index = find_next_open_boxplot_column(sheet)
    col_letter = _column_letter(col_index)

    rows = config.DEST_BOXPLOT_HEADER_ROWS
    sheet.Cells(rows["lot"], col_index).Value = lot_number
    sheet.Cells(rows["wo"], col_index).Value = wo_number
    sheet.Cells(rows["fill_date"], col_index).Value = fill_date

    start_row = config.DEST_BOXPLOT_DATA_START_ROW
    for i, value in enumerate(seal_values):
        sheet.Cells(start_row + i, col_index).Value = value

    return col_letter


# ---------------------------------------------------------------------------
# Destination — Tensile Data (Control Chart) sheet
# ---------------------------------------------------------------------------

def find_next_open_control_chart_row(sheet):
    """Find the first empty row below existing lot data, using
    End(xlUp) from the bottom of the sheet (column A, Lot No.)
    instead of looping down."""
    max_row = sheet.Rows.Count
    last_used_row = sheet.Range(f"A{max_row}").End(XL_UP).Row

    header_row = config.DEST_CONTROL_CHART_HEADER_ROW
    if last_used_row <= header_row:
        return header_row + 1

    return last_used_row + 1


def write_control_chart_data(sheet, lot_number, wo_number, fill_date, seal_values):
    """Write one lot's data into the next open row on the Control
    Chart sheet — this is the transpose of write_boxplot_data: seal
    values go ACROSS starting at DEST_CONTROL_CHART_DATA_START_COL,
    instead of down a column."""
    row = find_next_open_control_chart_row(sheet)

    sheet.Range(f"A{row}").Value = lot_number
    sheet.Range(f"B{row}").Value = wo_number
    sheet.Range(f"C{row}").Value = fill_date

    start_col_index = ord(config.DEST_CONTROL_CHART_DATA_START_COL) - ord("A") + 1
    for i, value in enumerate(seal_values):
        sheet.Cells(row, start_col_index + i).Value = value

    return row


# ---------------------------------------------------------------------------
# Orchestration helper
# ---------------------------------------------------------------------------

def write_new_lot_to_all_sheets(wo_data_sheet, boxplot_sheet, control_chart_sheet,
                                 fill_date, filler, part_number, product_name,
                                 lot_number, wo_number, seal_values):
    """Writes one lot's data to all three destination sheets. Does NOT
    handle the _1 duplicate-lot suffix (step Notes) — that decision
    should be made by the caller before passing lot_number in here,
    since it depends on checking whether today's date already has an
    entry, which this function doesn't have visibility into on its own.
    """
    wo_row = append_wo_data_row(
        wo_data_sheet, fill_date, filler, part_number, product_name,
        lot_number, wo_number,
    )
    boxplot_col = write_boxplot_data(
        boxplot_sheet, lot_number, wo_number, fill_date, seal_values,
    )
    control_chart_row = write_control_chart_data(
        control_chart_sheet, lot_number, wo_number, fill_date, seal_values,
    )
    return {
        "wo_data_row": wo_row,
        "boxplot_column": boxplot_col,
        "control_chart_row": control_chart_row,
    }