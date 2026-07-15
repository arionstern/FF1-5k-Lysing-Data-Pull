"""
tests/minitab/test_column_label_axis.py

Quick test — checks whether setting a column's Label property (distinct
from Name, seen earlier in dir(column) but never used) causes Minitab
to auto-populate the X-axis label when a chart is created from it. If
this doesn't work, axis labels likely need to stay a manual step —
no exposed API/command for editing them post-creation was found.

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

    column_names = minitab_utils.get_boxplot_column_names_from_year(boxplot_sheet)
    first_col, last_col = column_names[0], column_names[-1]

    # Set Label on every column in range to "Fill Date" — testing
    # whether this auto-populates the X-axis label on chart creation
    print("Setting Label = 'Fill Date' on all columns in range...")
    for i in range(1, boxplot_sheet.Columns.Count + 1):
        column = boxplot_sheet.Columns.Item(i)
        if column.Name in column_names:
            column.Label = "Fill Date"

    chart_config = config.MINITAB_BOXPLOT_CHART
    command_text = (
        f"Boxplot '{first_col}'-'{last_col}';\n"
        f"  Overlay;\n"
        f"  IQRBox;\n"
        f"  Outlier;\n"
        f"  Title \"{chart_config['title']}\"."
    )
    project.ExecuteCommand(command_text)

    print("\nChart created. Check visually: did the X-axis label")
    print("automatically show 'Fill Date' without an XLabel subcommand?")
    print("\nIf yes: column Label is the real mechanism, we can automate")
    print("this (set once, applies going forward).")
    print("If no: axis labels stay a manual 2-click step — no exposed")
    print("API/command found for this after two dead ends (command")
    print("language, Graph COM object).")


if __name__ == "__main__":
    main()