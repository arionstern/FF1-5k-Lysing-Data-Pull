"""
tests/minitab/test_duplicate_date_suffix.py

Writes TWO lots with the SAME fill date and verifies the second one
actually gets a '_1' suffix -- a real test of the duplicate-date logic,
not just checking against an empty/unwritten worksheet.

Uses a throwaway test date (2099-1-1) that won't collide with any
real production data.

Requires: config.USE_TEST_PATHS = True.
"""

import sys
from datetime import date

sys.path.append("../..")
import config
import minitab_utils

TEST_DATE = date(2099, 1, 1)  # throwaway, won't collide with real data
TEST_VALUES_1 = [25.0, 26.0, 27.0]  # small fake values, just for the test
TEST_VALUES_2 = [28.0, 29.0, 30.0]


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    mtb, project = minitab_utils.open_minitab_project()
    boxplot_sheet = minitab_utils.get_worksheet(project, config.DEST_BOXPLOT_SHEET)

    print(f"Writing FIRST lot for {TEST_DATE}...")
    name_1 = minitab_utils.determine_unique_boxplot_column_name(
        boxplot_sheet, TEST_DATE
    )
    print(f"  Column name: {name_1!r}")
    minitab_utils.write_boxplot_column(boxplot_sheet, name_1, TEST_VALUES_1)

    print(f"\nWriting SECOND lot for the SAME date {TEST_DATE}...")
    name_2 = minitab_utils.determine_unique_boxplot_column_name(
        boxplot_sheet, TEST_DATE
    )
    print(f"  Column name: {name_2!r}")
    minitab_utils.write_boxplot_column(boxplot_sheet, name_2, TEST_VALUES_2)

    print("\n--- Verification ---")
    if name_1 == "1/1/2099" and name_2 == "1/1/2099_1":
        print("OK: first lot got the base name, second got '_1' suffix.")
    else:
        print(f"FAILED: expected '1/1/2099' and '1/1/2099_1', "
              f"got {name_1!r} and {name_2!r}")

    print(f"\nMinitab project left open, NOT saved -- verify visually,")
    print(f"then close WITHOUT saving to discard this test data.")


if __name__ == "__main__":
    main()