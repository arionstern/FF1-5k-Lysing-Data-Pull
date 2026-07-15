"""
tests/minitab/test_write_functions.py

Tests write_boxplot_column() and write_control_chart_row() against
the TEST .mpx copy, using 260610D's real data (already verified
against Excel and trusted Minitab values earlier) so there's a known-
correct result to compare against.

Requires: config.USE_TEST_PATHS = True. Refuses to run otherwise.
"""

import sys
from datetime import date

sys.path.append("../..")
import config
import minitab_utils

# Known-good real data for 260610D, already verified earlier
LOT_NUMBER = "260610D"
WO_NUMBER = "10556543"
FILL_DATE = date(2026, 6, 10)
SEAL_VALUES = [
    27.706, 27.244, 29.104, 27.265, 30.896, 28.614, 26.515, 22.54,
    21.37, 31.717, 23.552, 30.159, 29.458, 27.336, 27.089, 28.415,
    26.878, 28.125, 29.421, 27.646,
]  # 20 values, matches the Boxplot column BF we already confirmed


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping rather than "
              "risk writing to the REAL production Minitab project.")
        return

    print(f"Opening: {config.MINITAB_PROJECT_PATH}")
    mtb, project = minitab_utils.open_minitab_project()

    print("\n--- Testing write_boxplot_column ---")
    boxplot_sheet = minitab_utils.get_worksheet(project, config.DEST_BOXPLOT_SHEET)
    column_name = minitab_utils.format_boxplot_column_name(FILL_DATE)
    print(f"Column name: {column_name!r}")

    col_index = minitab_utils.write_boxplot_column(
        boxplot_sheet, column_name, SEAL_VALUES
    )
    print(f"Wrote to column index: {col_index}")

    # Read it back to confirm
    written_column = boxplot_sheet.Columns.Item(col_index)
    written_data = list(written_column.GetData())
    print(f"Read back: {written_data[:5]} ... {written_data[-5:]}")
    print(f"Column name confirmed as: {written_column.Name!r}")

    print("\n--- Testing write_control_chart_row ---")
    control_chart_sheet = minitab_utils.get_worksheet(
        project, config.DEST_CONTROL_CHART_SHEET
    )
    row = minitab_utils.write_control_chart_row(
        control_chart_sheet, LOT_NUMBER, WO_NUMBER, FILL_DATE, SEAL_VALUES
    )
    print(f"Wrote to row: {row}")

    # Read back the row across a few columns to confirm
    columns = control_chart_sheet.Columns
    lot_check = list(columns.Item(1).GetData())[row - 1]
    wo_check = list(columns.Item(2).GetData())[row - 1]
    bag1_check = list(columns.Item(4).GetData())[row - 1]
    print(f"Read back — Lot: {lot_check!r}, WO: {wo_check!r}, "
          f"Bag 1: {bag1_check!r}")

    print("\nDone. Minitab window left open (not saved) — visually")
    print("inspect both worksheets before deciding whether to save.")


if __name__ == "__main__":
    main()