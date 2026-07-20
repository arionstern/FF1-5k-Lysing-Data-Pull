"""
tests/integration/test_read_fill_date.py

Isolated test for excel_utils.read_fill_date() BEFORE trusting the
Run_Lysing_Pull.py restructuring that now calls it internally.

Uses a real, known lot: WO 10577439 / lot 260630C.
Expected fill date: datetime(2026, 6, 30) — confirmed against a real
screenshot of the "GEM Fill Logs Header Page" sheet, cell F16.

This does NOT touch the destination Excel/Minitab files at all — it
only opens the SAP-hosted source workbook, reads the date, and closes
it without saving. Safe to run regardless of USE_TEST_PATHS.
"""

import os
import sys
from datetime import datetime

# Allow running this script directly from tests/integration/ — add the
# project root (two levels up) to sys.path so sap_utils/excel_utils
# can be imported without needing to install the project as a package.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import sap_utils
import excel_utils

WO_NUMBER = "10577439"
EXPECTED_LOT = "260630C"
EXPECTED_FILL_DATE = datetime(2026, 6, 30)


def main():
    print(f"Opening source GEM log document for WO {WO_NUMBER} "
          f"(expected lot {EXPECTED_LOT})...")

    workbook, sheet = sap_utils.get_gem_log_sheet_for_wo(WO_NUMBER)

    try:
        # Sanity-check we opened the right lot before trusting the date
        metadata = excel_utils.read_lot_metadata(sheet)
        print(f"  Metadata: {metadata}")

        fill_date = excel_utils.read_fill_date(workbook)
        print(f"  Fill date read: {fill_date!r}")

        assert isinstance(fill_date, datetime), (
            f"Expected a datetime, got {type(fill_date)}"
        )
        assert fill_date == EXPECTED_FILL_DATE, (
            f"Expected {EXPECTED_FILL_DATE!r}, got {fill_date!r}"
        )

        print("PASS: read_fill_date() returned the expected date.")

    finally:
        workbook.Close(SaveChanges=False)
        print("Source workbook closed (not saved).")


if __name__ == "__main__":
    main()