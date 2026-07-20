"""
tests/minitab/test_duplicate_prevention.py

Direct test of lot_exists_in_control_chart() against 260630C, which
is already confirmed present in the test Minitab project. Also tests
against a lot known NOT to exist, as a control case (to catch a
false-positive that always returns True).

Requires: config.USE_TEST_PATHS = True.
"""

import sys

sys.path.append("../..")
import config
import minitab_utils


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    mtb, project = minitab_utils.open_minitab_project()
    control_chart_sheet = minitab_utils.get_worksheet(
        project, config.DEST_CONTROL_CHART_SHEET
    )

    # Test 1: a lot we KNOW is present
    exists = minitab_utils.lot_exists_in_control_chart(
        control_chart_sheet, "260630C"
    )
    print(f"lot_exists_in_control_chart('260630C') -> {exists}")
    if exists:
        print("OK: correctly detected as existing.")
    else:
        print("FAILED: expected True, got False.")

    # Test 2: a lot we KNOW is NOT present (control case)
    fake_lot = "999999Z"
    not_exists = minitab_utils.lot_exists_in_control_chart(
        control_chart_sheet, fake_lot
    )
    print(f"\nlot_exists_in_control_chart({fake_lot!r}) -> {not_exists}")
    if not not_exists:
        print("OK: correctly detected as NOT existing.")
    else:
        print(f"FAILED: expected False, got True — false positive!")


if __name__ == "__main__":
    main()