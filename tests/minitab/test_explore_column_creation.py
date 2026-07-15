"""
tests/minitab/test_explore_column_creation.py

Exploratory script — not an automated test suite (no assertions).
Item(index, CreateIfNecessary) wasn't a real signature — Item() only
takes one positional argument via this COM interface. This script
tries several other plausible approaches for creating a new column in
a Minitab worksheet via COM, to find the one that actually works.

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
    boxplot_sheet = minitab_utils.get_worksheet(project, config.DEST_BOXPLOT_SHEET)

    print(f"Current column count: {boxplot_sheet.Columns.Count}")

    # Approach 1: dir() on the Columns object to see what's exposed
    print("\n--- Approach 1: dir(Columns) ---")
    try:
        attrs = [a for a in dir(boxplot_sheet.Columns) if not a.startswith("_")]
        print(attrs)
    except Exception as e:
        print(f"FAILED: {e}")

    # Approach 2: dir() on a single existing Column, to see if there's
    # a way to reference/create a NEW one via similar mechanics
    print("\n--- Approach 2: dir(single Column) ---")
    try:
        col1 = boxplot_sheet.Columns.Item(1)
        attrs = [a for a in dir(col1) if not a.startswith("_")]
        print(attrs)
    except Exception as e:
        print(f"FAILED: {e}")

    # Approach 3: try .Add() directly on Columns
    print("\n--- Approach 3: Columns.Add() ---")
    try:
        new_col = boxplot_sheet.Columns.Add()
        print(f"SUCCESS: {new_col}")
    except Exception as e:
        print(f"FAILED: {e}")

    # Approach 4: try referencing an out-of-range index directly via
    # default indexer (not .Item) — some COM collections auto-create
    # on assignment even if Item() alone doesn't
    print("\n--- Approach 4: direct index assignment ---")
    try:
        next_index = boxplot_sheet.Columns.Count + 1
        boxplot_sheet.Columns(next_index).Name = "TEST_COLUMN"
        print("SUCCESS via direct call syntax")
    except Exception as e:
        print(f"FAILED: {e}")

    # Approach 5: try the worksheet-level command execution approach
    # (Minitab Command Language / session commands) — a more
    # universally documented automation path that sidesteps the
    # Columns collection entirely
    print("\n--- Approach 5: ExecuteCommand (MTB command language) ---")
    try:
        print(f"mtb has ExecuteCommand: {hasattr(mtb, 'ExecuteCommand')}")
        print(f"project has ExecuteCommand: {hasattr(project, 'ExecuteCommand')}")
        print(f"mtb attrs sample: "
              f"{[a for a in dir(mtb) if 'xecut' in a or 'ommand' in a]}")
    except Exception as e:
        print(f"FAILED: {e}")

    print(f"\nFinal column count: {boxplot_sheet.Columns.Count}")
    print("\nDone. Whichever approach printed SUCCESS (and actually")
    print("increased the column count) is the one to build on.")


if __name__ == "__main__":
    main()