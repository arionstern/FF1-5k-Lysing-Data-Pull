"""
tests/minitab/test_xbarchart_command.py

Tests the real XBARCHART session command (found via Minitab's own
documentation, not guessed) against the Control Chart worksheet's
real columns: subgroups in rows (RSUB), Bag 1-Bag 38 for the data,
Fill Date column via STAMP for real date labels on the X-axis
(inference, not confirmed), plus AxLabel and Title.

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

    # Confirm real column count for Bag N (matches earlier confirmed
    # structure: C4=Bag 1 ... up to however many bag columns exist)
    max_bag_col_index = control_chart_sheet.Columns.Count
    print(f"Control Chart sheet has {max_bag_col_index} columns total.")

    # Compute EXCLUDE range dynamically (real chart excludes rows
    # before CHART_DATA_START_YEAR, per its footnote "Results exclude
    # specified rows: 1:25") — NOT hardcoded, since that count changes
    # over time as more pre-2026 rows exist, or once 2027 starts.
    fill_date_column = control_chart_sheet.Columns.Item(3)  # C3 = Fill Date
    dates = list(fill_date_column.GetData())
    rows_before_cutoff = 0
    for d in dates:
        if d is None:
            continue
        try:
            if d.year < config.CHART_DATA_START_YEAR:
                rows_before_cutoff += 1
            else:
                break  # dates are chronological, stop at first in-range row
        except AttributeError:
            continue

    print(f"Rows before {config.CHART_DATA_START_YEAR}: {rows_before_cutoff}")

    exclude_clause = ""
    if rows_before_cutoff > 0:
        exclude_clause = f"  EXCLUDE;\n    ROWS 1:{rows_before_cutoff};\n"

    command_text = (
        "XBARCHART;\n"
        "  RSUB 'Bag 1'-'Bag 38';\n"
        "  STAMP 'Fill Date';\n"
        f"{exclude_clause}"
        "  TEST 0;\n"  # UNVERIFIED GUESS: attempting to disable all
                       # special-cause tests (the real chart shows no
                       # red violation markers) — "0" as "none" is a
                       # common Minitab convention, not confirmed here
        "  AxLabel 1 \"Fill Date\";\n"
        "  AxLabel 2 \"Seal Strength (lbf)\";\n"
        "  Title \"2026 FF1 5K Lysing (Variable Subgroup Size)\"."
    )

    print(f"\nAttempting:\n{command_text}\n")

    commands_before = project.Commands.Count
    project.ExecuteCommand(command_text)
    commands_after = project.Commands.Count

    if commands_after > commands_before:
        command = project.Commands.Item(commands_after)
        print(f"Command created: {command.Name}")
        print("Check the chart visually — does it look like the real")
        print("Xbar chart, with correct dates on the X-axis and both")
        print("axis labels set correctly?")
    else:
        print("FAILED: command count unchanged — check Minitab's "
              "Session window for the real error text.")


if __name__ == "__main__":
    main()