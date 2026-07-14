"""
tests/excel/test_backfill_missed_lot.py

Backfills ONE missed lot (260610C, WO 10556542) end-to-end:
  1. Navigate SAP directly to the known WO number (skips the F4
     multi-select popup entirely, since we already know the exact WO
     — simpler and more reliable than the earlier material-search flow).
  2. Read real source data using the already-built excel_utils
     read functions (fill line, metadata, seal strength values).
  3. Open the TEST destination workbook and write the data into
     Boxplot + Control Chart ONLY (NOT WO Data — that row already
     exists for this lot, so append_wo_data_row is deliberately not
     called here).

Requires: SAP GUI open and logged in. config.USE_TEST_PATHS must be
True — this refuses to run otherwise.
"""

import sys
import time
import win32com.client

sys.path.append("../..")
import config
import excel_utils


# DIAGNOSTIC: now testing 260610D (WO 10556543) — the other missing
# backfill lot, filled same day as 260610C. Using the same routing
# item key as a starting guess since same-day lots likely share
# identical routing structure. Testing whether the "completed order"
# theory holds (should fail the same way if so) or whether 260610C
# was uniquely broken for some other reason.
# Known values for this specific backfill case (from the WO Data sheet)
LOT_NUMBER = "260610D"
WO_NUMBER = "10556543"


def get_sap_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def navigate_to_wo_directly(session, wo_number):
    """Navigate straight to a known WO number, skipping the F4
    multi-select popup used when searching by material instead."""
    session.findById("wnd[0]").maximize()
    # /n forces a full reset to a fresh transaction, even if the
    # session is currently sitting deep inside a previous screen
    # (e.g. left mid-routing from an earlier failed run) — without
    # it, re-typing the transaction code alone doesn't reliably
    # return to the initial entry screen.
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_WI"
    session.findById("wnd[0]").sendVKey(0)

    order_field_id = "wnd[0]/usr/ctxtZZWOSCAN-AUFNR"
    session.findById(order_field_id).text = wo_number
    session.findById("wnd[0]").sendVKey(0)  # Enter
    print(f"Navigated directly to WO {wo_number}, "
          f"transaction: {session.Info.Transaction}")


def open_via_documents_list(session, wo_number):
    """Open the GEM log document via the Documents List tab (a fixed,
    standard SAP ALV grid — not the fragile custom routing tree),
    searching for the document number we can now DERIVE directly from
    the WO number: '0000' + WO_NUMBER + '000006' + '0030' + '-01'
    (confirmed against two real lots)."""
    expected_doc_number = f"0000{wo_number}0000060030-01"
    print(f"  Expected document number: {expected_doc_number}")

    docs_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB02"
    print("  Selecting Documents List tab...")
    session.findById(docs_tab_id).select()
    print("  OK")
    time.sleep(1)

    grid_id = (
        "wnd[0]/usr/tabsTABS_0200/tabpTAB02/"
        "ssubSUBSCREEN:SAPLZPPDB:0100/cntlZPPDB_CONT/shellcont/shell"
    )
    grid = session.findById(grid_id)

    # Try to find the row matching our derived document number by
    # reading the real DOKNR column — much more robust than a
    # hardcoded row index, since this is a standard ALV grid with
    # real column names.
    matched_row = None
    try:
        row_count = grid.RowCount
        print(f"  Grid has {row_count} rows. Searching DOKNR column...")
        for row in range(row_count):
            doc_number = grid.GetCellValue(row, "DOKNR")
            if doc_number and doc_number.strip() == expected_doc_number:
                matched_row = row
                print(f"  MATCH at row {row}: {doc_number!r}")
                break
    except Exception as e:
        print(f"  Dynamic DOKNR search failed ({e}), falling back to "
              f"row 2 (the value from the recording — may not "
              f"generalize).")
        matched_row = 2

    if matched_row is None:
        print(f"  No row matched {expected_doc_number!r} — falling back "
              f"to row 2 from the recording.")
        matched_row = 2

    print(f"  Opening row {matched_row}...")
    grid.setCurrentCell(matched_row, "DOKNR")
    grid.selectedRows = str(matched_row)
    grid.clickCurrentCell()
    print("  OK")
    time.sleep(1.5)

    # This lands on a standard SAP DMS document viewer screen — open
    # the first (only) attached file.
    file_tree_id = (
        "wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/"
        "ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/"
        "shell/shellcont[1]/shell[1]"
    )
    print("  Opening attached file...")
    file_tree = session.findById(file_tree_id)
    file_tree.selectNode("          1")
    file_tree.doubleClickNode("          1")
    print("  OK")


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping rather than "
              "risk writing to the REAL production workbook.")
        return

    session = get_sap_session()

    try:
        navigate_to_wo_directly(session, WO_NUMBER)
        open_via_documents_list(session, WO_NUMBER)
    except Exception as e:
        print(f"FAILED during SAP navigation: {e}")
        return

    print("\nWaiting for SAP to write the file to the local temp folder...")

    import os
    doc_number = f"0000{WO_NUMBER}0000060030-01"
    expected_path = os.path.join(
        os.environ["LOCALAPPDATA"], "SAP", "SAP GUI", "tmp",
        f"{doc_number}.xlsx"
    )
    print(f"  Expected path: {expected_path}")

    source_sheet = None
    max_attempts = 10
    for attempt in range(1, max_attempts + 1):
        time.sleep(1)
        if os.path.exists(expected_path):
            print(f"  Attempt {attempt}: file found on disk.")
            try:
                source_excel = win32com.client.Dispatch("Excel.Application")
                source_workbook = source_excel.Workbooks.Open(expected_path)
                source_sheet = source_workbook.Sheets(config.SOURCE_SHEET_NAME)
                print(f"  Opened directly and found sheet "
                      f"'{config.SOURCE_SHEET_NAME}'.")
                break
            except Exception as e:
                print(f"  File exists but failed to open/read it: {e}")
                break
        else:
            print(f"  Attempt {attempt}: not written yet...")

    if source_sheet is None:
        print(f"FAILED to find the source file after {max_attempts} "
              f"attempts. Expected: {expected_path}")
        return

    fill_line = excel_utils.read_fill_line(source_sheet)
    print(f"\nFill line: {fill_line}")
    if not excel_utils.validate_fill_line(fill_line):
        print(f"WARNING: '{fill_line}' not in {config.VALID_FILL_LINES}")

    metadata = excel_utils.read_lot_metadata(source_sheet)
    print(f"Metadata: {metadata}")

    seal_values = excel_utils.read_seal_strength_values(source_sheet)
    print(f"Real seal strength readings: {len(seal_values)}")
    print(f"First 5: {seal_values[:5]}")
    print(f"Last 5: {seal_values[-5:]}")

    # Step 3: write to TEST destination — Boxplot + Control Chart only
    print(f"\nOpening TEST destination workbook: {config.LYSING_WORKBOOK_PATH}")
    dest_excel = win32com.client.Dispatch("Excel.Application")
    dest_excel.Visible = True
    dest_workbook = dest_excel.Workbooks.Open(config.LYSING_WORKBOOK_PATH)

    boxplot_sheet = dest_workbook.Sheets(config.DEST_BOXPLOT_SHEET)
    control_chart_sheet = dest_workbook.Sheets(config.DEST_CONTROL_CHART_SHEET)

    # Fill date pulled from WO Data originally — hardcoding here since
    # we already know it (6/10/2026) rather than re-reading from source
    fill_date = "6/10/2026"

    boxplot_col = excel_utils.write_boxplot_data(
        boxplot_sheet, LOT_NUMBER, WO_NUMBER, fill_date, seal_values
    )
    print(f"\nWrote to Boxplot sheet, column {boxplot_col}")

    control_chart_row = excel_utils.write_control_chart_data(
        control_chart_sheet, LOT_NUMBER, WO_NUMBER, fill_date, seal_values
    )
    print(f"Wrote to Control Chart sheet, row {control_chart_row}")

    print("\nDone. Workbook left open (not saved) — review the written")
    print("values against what's already trusted in Minitab before")
    print("saving. If they match, this confirms the full read+write")
    print("path works for a real backfill case.")


if __name__ == "__main__":
    main()