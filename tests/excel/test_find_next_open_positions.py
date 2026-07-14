"""
tests/excel/test_find_next_open_positions.py

Exploratory script — not an automated test suite (no assertions).
Opens the destination workbook (respects config.USE_TEST_PATHS — make
sure that's True before running this against the Arion Stern test
copy, not the real production file) and calls the three
find_next_open_* functions from excel_utils.py, printing what each one
finds. Does NOT write anything — read-only check before we trust these
functions with real writes.
"""

import sys
import win32com.client

sys.path.append("../..")
import config
import excel_utils


def main():
    if config.USE_TEST_PATHS:
        print(f"USE_TEST_PATHS is True — using: {config.LYSING_WORKBOOK_PATH}")
    else:
        print("WARNING: USE_TEST_PATHS is False — this would read from the "
              "REAL production workbook. Stopping rather than risk opening "
              "it accidentally during a test run.")
        return

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = True

    print(f"\nOpening: {config.LYSING_WORKBOOK_PATH}")
    workbook = excel_app.Workbooks.Open(config.LYSING_WORKBOOK_PATH)

    try:
        wo_data_sheet = workbook.Sheets(config.DEST_WO_SHEET_NAME)
        boxplot_sheet = workbook.Sheets(config.DEST_BOXPLOT_SHEET)
        control_chart_sheet = workbook.Sheets(config.DEST_CONTROL_CHART_SHEET)

        print("\n--- WO Data ---")
        next_wo_row = excel_utils.find_next_open_wo_data_row(wo_data_sheet)
        print(f"Next open row: {next_wo_row}")
        # Show what's actually in the row just above, as a sanity check
        prev_row = next_wo_row - 1
        prev_lot = wo_data_sheet.Range(f"{config.COL_LOT_NUMBER}{prev_row}").Value
        print(f"  (Row {prev_row}, last real row, Lot #: {prev_lot!r} — "
              f"confirm this matches the actual last lot you can see in Excel)")

        print("\n--- Tensile Data (Boxplot) ---")
        next_col_index = excel_utils.find_next_open_boxplot_column(boxplot_sheet)
        next_col_letter = excel_utils._column_letter(next_col_index)
        print(f"Next open column: {next_col_letter} (index {next_col_index})")
        prev_col_letter = excel_utils._column_letter(next_col_index - 1)
        prev_lot_boxplot = boxplot_sheet.Cells(
            config.DEST_BOXPLOT_HEADER_ROWS["lot"], next_col_index - 1
        ).Value
        print(f"  (Column {prev_col_letter}, last real column, Lot #: "
              f"{prev_lot_boxplot!r} — confirm this matches what you see in Excel)")

        print("\n--- Tensile Data (Control Chart) ---")
        next_cc_row = excel_utils.find_next_open_control_chart_row(control_chart_sheet)
        print(f"Next open row: {next_cc_row}")
        prev_cc_row = next_cc_row - 1
        prev_lot_cc = control_chart_sheet.Range(f"A{prev_cc_row}").Value
        print(f"  (Row {prev_cc_row}, last real row, Lot #: {prev_lot_cc!r} — "
              f"confirm this matches what you see in Excel)")

        print("\nDone. Nothing was written — this only reads. Compare the")
        print("'last real row/column' values above against what you can")
        print("see in the actual open Excel window.")

    finally:
        # Deliberately NOT closing the workbook — leaving it open so
        # you can visually cross-check against the printed values
        # while it's still on screen.
        pass


if __name__ == "__main__":
    main()