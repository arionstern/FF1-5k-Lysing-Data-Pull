"""
tests/sap/test_read_full_seal_strength_column.py

Exploratory script — not an automated test suite (no assertions).
Validates the fill line (step 4.5.1), then reads the FULL Seal
Strength column (not just the first ~10 rows) to find where the real
data actually ends, converting the Excel time-serial values into
readable times along the way.

Requires: SAP GUI open, sitting on the routing/documents screen with
the spreadsheet visible.
"""

import sys
import win32com.client

sys.path.append("../..")
import config


def excel_time_to_string(time_fraction):
    """Convert an Excel time serial (fraction of a day) to HH:MM AM/PM.
    Returns the raw value with a flag if it's not a numeric fraction —
    some rows may have text or unexpected formats instead."""
    if time_fraction is None:
        return None
    if not isinstance(time_fraction, (int, float)):
        return f"NON-NUMERIC: {time_fraction!r}"
    total_seconds = round(time_fraction * 24 * 60 * 60)
    hours = (total_seconds // 3600) % 24
    minutes = (total_seconds % 3600) // 60
    period = "AM" if hours < 12 else "PM"
    display_hour = hours % 12
    if display_hour == 0:
        display_hour = 12
    return f"{display_hour}:{minutes:02d} {period}"


def main():
    try:
        excel_app = win32com.client.GetObject(Class="Excel.Application")
    except Exception as e:
        print(f"FAILED to attach to Excel: {e}")
        return

    workbook = excel_app.Workbooks(1)
    sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)
    print(f"Sheet: {sheet.Name}")

    # Step 4.5.1: confirm fill line
    fill_line = sheet.Range(config.SOURCE_FILL_LINE_CELL).Value
    print(f"\nFill line ({config.SOURCE_FILL_LINE_CELL}): {fill_line!r}")
    if fill_line in config.VALID_FILL_LINES:
        print(f"OK: '{fill_line}' is a valid fill line.")
    else:
        print(f"WARNING: '{fill_line}' is NOT in {config.VALID_FILL_LINES} "
              f"— per step 4.5.1, this pull may not be valid.")

    # Read the full Seal Strength column until we hit a blank row
    seal_col = config.SOURCE_TABLE_COLUMNS["seal_strength"]
    time_col = config.SOURCE_TABLE_COLUMNS["time"]
    start_row = config.SOURCE_TABLE_DATA_START_ROW

    print(f"\nReading column {seal_col} starting at row {start_row}, "
          f"stopping at first blank...")

    row = start_row
    values = []
    max_rows_to_check = 500  # safety cap, in case there's no blank row
    while row < start_row + max_rows_to_check:
        seal_value = sheet.Range(f"{seal_col}{row}").Value
        if seal_value is None or seal_value == "":
            break
        time_value = sheet.Range(f"{time_col}{row}").Value
        values.append((row, excel_time_to_string(time_value), seal_value))
        row += 1

    print(f"\nFound {len(values)} total rows before first true blank "
          f"(rows {start_row}-{row - 1}), including template placeholders.")

    real_values = [(r, t, s) for r, t, s in values if isinstance(s, (int, float))]
    placeholder_values = [(r, t, s) for r, t, s in values if not isinstance(s, (int, float))]

    print(f"\nOf those: {len(real_values)} are REAL numeric readings, "
          f"{len(placeholder_values)} are template placeholders (e.g. 'N/A').")

    print(f"\nFirst 5 real readings:")
    for row_num, time_str, seal_val in real_values[:5]:
        print(f"  Row {row_num}: time={time_str}, seal_strength={seal_val}")
    print(f"Last 5 real readings:")
    for row_num, time_str, seal_val in real_values[-5:]:
        print(f"  Row {row_num}: time={time_str}, seal_strength={seal_val}")

    if placeholder_values:
        first_placeholder_row = placeholder_values[0][0]
        print(f"\nPlaceholders start at row {first_placeholder_row} "
              f"(i.e. real data for this lot ends at row "
              f"{first_placeholder_row - 1}).")

    print(f"\nThis lot has {len(real_values)} real bag readings — worth")
    print(f"comparing against config.MINITAB_BAG_COLUMN_RANGE "
          f"(currently bag1-bag38) to see if that range assumption holds")
    print(f"across lots, or if it genuinely varies per lot as the chart")
    print(f"title 'Variable Subgroup Size' suggests.")

    if row >= start_row + max_rows_to_check:
        print(f"\nWARNING: hit the {max_rows_to_check}-row safety cap "
              f"without finding a blank row — the real data may extend "
              f"further, or there's a gap/formatting issue worth checking.")


if __name__ == "__main__":
    main()