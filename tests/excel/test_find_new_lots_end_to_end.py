"""
tests/integration/test_find_new_lots_end_to_end.py

Tests excel_utils.get_last_known_lot() against the real TEST workbook
(should return '260630C'), then feeds that directly into
sap_utils.find_new_lots() — the first time these two pieces run
together, instead of using a hardcoded LAST_KNOWN_LOT.

Requires: SAP GUI open and logged in. config.USE_TEST_PATHS = True.
"""

import sys
import win32com.client

sys.path.append("../..")
import config
import excel_utils
import sap_utils


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    print(f"Opening TEST workbook: {config.LYSING_WORKBOOK_PATH}")
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = True
    workbook = excel_app.Workbooks.Open(config.LYSING_WORKBOOK_PATH)
    wo_data_sheet = workbook.Sheets(config.DEST_WO_SHEET_NAME)

    last_known_lot = excel_utils.get_last_known_lot(wo_data_sheet)
    print(f"\nget_last_known_lot() returned: {last_known_lot!r}")
    if last_known_lot == "260630C":
        print("OK: matches expected value.")
    else:
        print(f"WARNING: expected '260630C', got {last_known_lot!r} — "
              f"either the sheet has changed, or something's off.")

    print(f"\nFeeding {last_known_lot!r} into sap_utils.find_new_lots()...")
    session = sap_utils.get_sap_session()
    new_lots = sap_utils.find_new_lots(session, last_known_lot)

    print(f"\nFound {len(new_lots)} new lot(s):")
    for order, data in sorted(new_lots.items()):
        print(f"  Batch={data['batch']!r} Order={order!r} "
              f"GRQty={data['gr_qty']!r}")

    ready_lots = sap_utils.filter_ready_lots(new_lots)
    print(f"\nOf those, {len(ready_lots)} are READY (GR qty != 0):")
    for order, data in sorted(ready_lots.items()):
        print(f"  Batch={data['batch']!r} Order={order!r} "
              f"GRQty={data['gr_qty']!r}")

    print("\nDone. Excel workbook left open (not saved/modified).")


if __name__ == "__main__":
    main()