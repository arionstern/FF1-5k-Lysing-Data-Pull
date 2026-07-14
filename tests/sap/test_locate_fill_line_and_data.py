"""
tests/sap/test_locate_fill_line_and_data.py

Exploratory script — not an automated test suite (no assertions).
The first dump only covered columns A-F and missed the actual VALUES
next to each label (PRODUCT NAME:, W.O. NUMBER:, GEM FILLER:, etc.),
which live one column further right (column G). This widens the dump
to A-I and down through the Tensile Tester Use Log table, printing
each cell's real address so there's no more guessing from screenshots.

Requires: SAP GUI open, sitting on the routing/documents screen with
the spreadsheet visible, same as the last few scripts.
"""

import sys
import win32com.client

sys.path.append("../..")
import config


COLUMN_LETTERS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]


def main():
    try:
        excel_app = win32com.client.GetObject(Class="Excel.Application")
    except Exception as e:
        print(f"FAILED to attach to Excel: {e}")
        return

    workbook = excel_app.Workbooks(1)
    sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)
    print(f"Sheet: {sheet.Name}\n")

    print("Header block (rows 1-8, cols A-I) with real addresses:")
    for row in range(1, 9):
        for col_idx, col_letter in enumerate(COLUMN_LETTERS[:9], start=1):
            value = sheet.Cells(row, col_idx).Value
            if value not in (None, ""):
                print(f"  {col_letter}{row}: {value!r}")

    print("\nTensile Tester Use Log table (rows 9-20, cols A-I):")
    for row in range(9, 21):
        for col_idx, col_letter in enumerate(COLUMN_LETTERS[:9], start=1):
            value = sheet.Cells(row, col_idx).Value
            if value not in (None, ""):
                print(f"  {col_letter}{row}: {value!r}")

    print("\nUse the addresses above to confirm:")
    print(f"  - Fill line value (should be FF1 or FF2, per "
          f"config.VALID_FILL_LINES = {config.VALID_FILL_LINES})")
    print(f"  - '{config.SOURCE_COLUMN_SEAL_STRENGTH}' column letter")
    print("  - Which row the data actually starts on")


if __name__ == "__main__":
    main()