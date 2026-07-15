"""
tests/minitab/test_explore_minitab_project.py

Exploratory script — not an automated test suite (no assertions).
Opens the FF1 5K Lysing Minitab project via COM and prints the real
worksheet name, column names, and a sample of data — same approach
used for SAP/Excel: confirm real structure before writing minitab_utils.py
logic, rather than guessing column names.

Requires: Minitab installed. Respects config.USE_TEST_PATHS — make
sure that's set correctly before running.
"""

import sys
import os
import win32com.client

sys.path.append("../..")
import config


def main():
    if not os.path.exists(config.MINITAB_PROJECT_PATH):
        print(f"FAILED: project not found at {config.MINITAB_PROJECT_PATH}")
        return

    print(f"Opening: {config.MINITAB_PROJECT_PATH}")
    print(f"(USE_TEST_PATHS = {config.USE_TEST_PATHS})")

    mtb = win32com.client.Dispatch("Mtb.Application.1")
    mtb.UserInterface.Visible = True
    mtb.Open(config.MINITAB_PROJECT_PATH)

    project = mtb.ActiveProject
    worksheet = project.ActiveWorksheet

    print(f"\nActive worksheet name: {worksheet.Name!r}")

    columns = worksheet.Columns
    print(f"Column count: {columns.Count}\n")

    for i in range(1, columns.Count + 1):
        column = columns.Item(i)
        try:
            data_type = column.DataType  # 0=Text, 1=Numeric, 2=Date/Time
            data_type_name = {0: "Text", 1: "Numeric", 2: "Date/Time"}.get(
                data_type, f"Unknown({data_type})"
            )
        except Exception:
            data_type_name = "?"

        try:
            row_count = column.RowCount
        except Exception:
            row_count = "?"

        print(f"  Col {i}: {column.Name!r} — type={data_type_name}, "
              f"rows={row_count}")

    # Print a small sample from the first few columns to see real data
    print("\nSample data (first 5 rows, first 10 columns):")
    sample_col_count = min(columns.Count, 10)
    for i in range(1, sample_col_count + 1):
        column = columns.Item(i)
        try:
            data = list(column.GetData())
            print(f"  {column.Name!r}: {data[:5]}")
        except Exception as e:
            print(f"  {column.Name!r}: FAILED to read ({e})")

    print("\nDone. Use this to confirm real column names/layout before "
          "writing minitab_utils.py logic for appending new lot data "
          "and generating charts.")


if __name__ == "__main__":
    main()