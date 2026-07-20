"""
tests/minitab/test_isolate_axis_label_reset.py

Isolates WHY the Xbar chart's Y-axis label reverts to "Sample Mean" —
adds ONE throwaway test row to the Control Chart sheet (triggering
Xbar's auto-update, since it's linked to that data), and nothing else.
If the label reverts after just this, that confirms the auto-update
mechanism itself resets custom labels — a Minitab quirk, not something
caused by our other code (write_new_lot_to_minitab, chart
regeneration, etc).

IMPORTANT: before running this, manually set the Xbar Y-axis label to
"Seal Strength (lbf)" and confirm it's set, so there's something to
watch for reverting.

Requires: config.USE_TEST_PATHS = True.
"""

import sys
from datetime import date

sys.path.append("../..")
import config
import minitab_utils

# Throwaway test values, clearly fake and easy to spot/delete after
TEST_LOT_NUMBER = "TESTLOT_DELETE_ME"
TEST_WO_NUMBER = "99999999"
TEST_DATE = date(2099, 1, 1)
TEST_SEAL_VALUES = [25.0, 25.0, 25.0]


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    print("BEFORE running this: manually confirm the Xbar chart's")
    print("Y-axis label currently says 'Seal Strength (lbf)' — if it")
    print("already says 'Sample Mean', there's nothing to observe.")
    input("\nPress Enter once confirmed...")

    mtb, project = minitab_utils.open_minitab_project()
    control_chart_sheet = minitab_utils.get_worksheet(
        project, config.DEST_CONTROL_CHART_SHEET
    )

    print(f"\nAdding ONE throwaway row: {TEST_LOT_NUMBER}...")
    row = minitab_utils.write_control_chart_row(
        control_chart_sheet, TEST_LOT_NUMBER, TEST_WO_NUMBER,
        TEST_DATE, TEST_SEAL_VALUES
    )
    print(f"Added at row {row}.")

    print("\nNothing else was touched — no chart regeneration, no")
    print("other writes. Check the Xbar chart NOW:")
    print("  - Did it auto-update to show the new throwaway point?")
    print("  - Did the Y-axis label revert to 'Sample Mean'?")
    print("\nIf the label reverted, that confirms the auto-update")
    print("mechanism itself resets it — independent of anything else")
    print("in our code.")
    print(f"\nDon't forget to delete row {row} (or the whole test)")
    print("afterward and NOT save, to discard this throwaway data.")


if __name__ == "__main__":
    main()