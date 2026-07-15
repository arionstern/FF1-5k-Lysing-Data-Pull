"""
tests/integration/test_end_to_end_260610C.py

Full pipeline integration test — the first time all three modules run
together in one continuous pass, with real data flowing from SAP
straight through to both Excel and Minitab (no hardcoded values
anywhere, unlike earlier module-level tests).

Also completes the last unconfirmed backfill lot (260610C) as a
side effect of being a genuinely useful real test case.

Requires: SAP GUI open and logged in. config.USE_TEST_PATHS = True —
refuses to run otherwise (this writes to both the TEST Excel workbook
and TEST Minitab project).
"""

import sys

sys.path.append("../..")
import config
import sap_utils
import excel_utils
import minitab_utils

LOT_NUMBER = "260610C"
WO_NUMBER = "10556542"
FILL_DATE = "6/10/2026"  # from WO Data — not re-derived from source


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping rather than "
              "risk writing to REAL production files.")
        return

    # -----------------------------------------------------------------
    # Step 1: SAP -> read real source data
    # -----------------------------------------------------------------
    print(f"=== Step 1: Reading source data from SAP (WO {WO_NUMBER}) ===")
    source_workbook, source_sheet = sap_utils.get_gem_log_sheet_for_wo(WO_NUMBER)

    try:
        fill_line = excel_utils.read_fill_line(source_sheet)
        print(f"Fill line: {fill_line}")
        if not excel_utils.validate_fill_line(fill_line):
            print(f"WARNING: '{fill_line}' not in {config.VALID_FILL_LINES}")

        metadata = excel_utils.read_lot_metadata(source_sheet)
        print(f"Metadata: {metadata}")

        seal_values = excel_utils.read_seal_strength_values(source_sheet)
        print(f"Real seal strength readings: {len(seal_values)}")
        print(f"First 5: {seal_values[:5]}")
        print(f"Last 5: {seal_values[-5:]}")
    finally:
        source_workbook.Close(SaveChanges=False)
        print("Source workbook closed.\n")

    # -----------------------------------------------------------------
    # Step 2: Excel -> write to Boxplot + Control Chart
    # (NOT WO Data — 260610C's row already exists there)
    # -----------------------------------------------------------------
    print("=== Step 2: Writing to Excel destination (TEST) ===")
    import win32com.client
    dest_excel = win32com.client.Dispatch("Excel.Application")
    dest_excel.Visible = True
    dest_workbook = dest_excel.Workbooks.Open(config.LYSING_WORKBOOK_PATH)

    boxplot_sheet_xl = dest_workbook.Sheets(config.DEST_BOXPLOT_SHEET)
    control_chart_sheet_xl = dest_workbook.Sheets(config.DEST_CONTROL_CHART_SHEET)

    boxplot_col = excel_utils.write_boxplot_data(
        boxplot_sheet_xl, LOT_NUMBER, WO_NUMBER, FILL_DATE, seal_values
    )
    print(f"Wrote to Excel Boxplot sheet, column {boxplot_col}")

    control_chart_row_xl = excel_utils.write_control_chart_data(
        control_chart_sheet_xl, LOT_NUMBER, WO_NUMBER, FILL_DATE, seal_values
    )
    print(f"Wrote to Excel Control Chart sheet, row {control_chart_row_xl}")
    print("Excel workbook left open, NOT saved.\n")

    # -----------------------------------------------------------------
    # Step 3: Minitab -> write to Boxplot + Control Chart
    # -----------------------------------------------------------------
    print("=== Step 3: Writing to Minitab destination (TEST) ===")
    from datetime import date
    fill_date_obj = date(2026, 6, 10)

    mtb, project = minitab_utils.open_minitab_project()
    boxplot_sheet_mtb = minitab_utils.get_worksheet(project, config.DEST_BOXPLOT_SHEET)
    control_chart_sheet_mtb = minitab_utils.get_worksheet(
        project, config.DEST_CONTROL_CHART_SHEET
    )

    column_name = minitab_utils.format_boxplot_column_name(fill_date_obj)
    boxplot_col_mtb = minitab_utils.write_boxplot_column(
        boxplot_sheet_mtb, column_name, seal_values
    )
    print(f"Wrote to Minitab Boxplot sheet, column {boxplot_col_mtb} "
          f"({column_name!r})")

    control_chart_row_mtb = minitab_utils.write_control_chart_row(
        control_chart_sheet_mtb, LOT_NUMBER, WO_NUMBER, fill_date_obj, seal_values
    )
    print(f"Wrote to Minitab Control Chart sheet, row {control_chart_row_mtb}")
    print("Minitab project left open, NOT saved.\n")

    # -----------------------------------------------------------------
    print("=== DONE ===")
    print(f"Lot {LOT_NUMBER} (WO {WO_NUMBER}) flowed through the full")
    print("pipeline: SAP -> Excel (Boxplot + Control Chart) -> Minitab")
    print("(Boxplot + Control Chart), all from one live source read.")
    print("\nReview both open windows before saving anything. This")
    print("also completes the last unconfirmed backfill lot, once")
    print("verified against Minitab's already-trusted values for")
    print("260610C.")


if __name__ == "__main__":
    main()