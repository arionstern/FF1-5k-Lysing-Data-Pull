"""
tests/sap/test_sap_utils_pipeline.py

Tests the full sap_utils.get_gem_log_sheet_for_wo() pipeline end to
end: navigate to a WO, open its GEM log document via Documents List,
read real data via excel_utils, and properly close the workbook when
done (important — leaving workbooks open across runs is what caused
the "locked for editing" conflicts seen earlier).

Requires: SAP GUI open and logged in.
"""

import sys

sys.path.append("../..")
import excel_utils
import sap_utils

# Known-good active lot, used to confirm the pipeline still works
# after the ReadOnly fix
WO_NUMBER = "10577439"  # 260630C


def main():
    print(f"Fetching GEM log sheet for WO {WO_NUMBER}...")
    workbook, sheet = sap_utils.get_gem_log_sheet_for_wo(WO_NUMBER)

    try:
        fill_line = excel_utils.read_fill_line(sheet)
        print(f"Fill line: {fill_line}")
        print(f"Valid: {excel_utils.validate_fill_line(fill_line)}")

        metadata = excel_utils.read_lot_metadata(sheet)
        print(f"Metadata: {metadata}")

        seal_values = excel_utils.read_seal_strength_values(sheet)
        print(f"Readings: {len(seal_values)}")
        print(f"First 5: {seal_values[:5]}")

    finally:
        # Always close, even if something above fails — this is what
        # was missing before, and why the file showed as "locked for
        # editing" on the next run.
        print("Closing workbook (no changes saved)...")
        workbook.Close(SaveChanges=False)
        print("Done.")


if __name__ == "__main__":
    main()