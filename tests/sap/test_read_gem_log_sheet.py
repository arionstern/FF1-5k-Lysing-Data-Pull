"""
tests/sap/test_read_gem_log_sheet.py

Exploratory script — not an automated test suite (no assertions).
Attaches to the Excel workbook SAP's Office Integration opened, checks
how many workbooks are open (flagging if more than one — affects how
excel_utils.py should pick the right one later), navigates to sheet
SOPGEM-45-2, confirms the fill line (step 4.5.1), and reads back a few
raw values from the Seal Strength column so we can see the real shape
of the data before writing parsing logic.

Requires: SAP GUI open, sitting on the routing/documents screen with
the spreadsheet visible (same state as the last two scripts).
"""

import sys
import win32com.client

sys.path.append("../..")
import config


def main():
    try:
        excel_app = win32com.client.GetObject(Class="Excel.Application")
    except Exception as e:
        print(f"FAILED to attach to Excel: {e}")
        return

    workbook_count = excel_app.Workbooks.Count
    print(f"Open workbooks: {workbook_count}")
    for wb in excel_app.Workbooks:
        print(f"  - {wb.Name}")

    if workbook_count > 1:
        print("\nNOTE: more than one workbook is open. excel_utils.py")
        print("will need a way to pick the right one (e.g. match by")
        print("filename pattern) rather than assuming Workbooks(1).")

    # NOTE: ActiveWorkbook returned None when tried here — attaching via
    # COM doesn't give Excel a "focused" window/workbook, even though
    # one is clearly open. Grabbing by index instead, since we already
    # know from the count above that exactly one is open right now.
    workbook = excel_app.Workbooks(1)
    print(f"\nWorkbook: {workbook.Name}")

    # Step 4.5: go to sheet SOPGEM-45-2
    try:
        sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)
        print(f"Found sheet: {sheet.Name}")
    except Exception as e:
        print(f"FAILED to find sheet '{config.SOURCE_SHEET_NAME}': {e}")
        print("Available sheets:")
        for s in workbook.Sheets:
            print(f"  - {s.Name}")
        return

    # Step 4.5.1: confirm fill line is FF1 or FF2
    # NOTE: don't yet know which cell holds "fill line" — printing the
    # top-left corner of the sheet to help spot it visually for now.
    print("\nTop-left corner of sheet (rows 1-10, cols A-F), to help")
    print("locate the fill-line field and the Seal Strength column:")
    for row in range(1, 11):
        row_values = []
        for col in range(1, 7):  # A-F
            cell = sheet.Cells(row, col)
            row_values.append(str(cell.Value))
        print(f"  Row {row}: {row_values}")

    print(f"\nOnce the fill-line cell and '{config.SOURCE_COLUMN_SEAL_STRENGTH}'")
    print("column are located above, update this script (or config.py)")
    print("with the real cell/column references for real parsing logic.")


if __name__ == "__main__":
    main()