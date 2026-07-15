"""
tests/minitab/test_generate_boxplot_chart.py

Exploratory script — tests whether ExecuteCommand can regenerate the
Boxplot chart with the current full column list. Command Language
syntax below is a best guess, NOT verified against documentation —
this script's purpose is to find out if it works, not assume it does.

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
    print(f"Found {len(column_names)} columns from "
          f"{config.CHART_DATA_START_YEAR} onward.")

    first_col = column_names[0]
    last_col = column_names[-1]

    chart_config = config.MINITAB_BOXPLOT_CHART
    command_text = (
        f"Boxplot '{first_col}'-'{last_col}';\n"
        f"  Overlay;\n"
        f"  IQRBox;\n"
        f"  Outlier;\n"
        f"  Title \"{chart_config['title']}\";\n"
        f"  AxLabel 1 \"{chart_config['x_axis']}\";\n"
        f"  AxLabel 2 \"{chart_config['y_axis']}\"."
    )
    # AxLabel is documented as a SESSION SUBCOMMAND, not a standalone
    # command — that's why the earlier separate-command attempt failed
    # with "Unknown Minitab command". Testing it nested inside
    # Boxplot's own subcommand block this time.

    print(f"\nAttempting:\n{command_text}\n")

    commands_before = project.Commands.Count
    project.ExecuteCommand(command_text)
    commands_after = project.Commands.Count

    if commands_after > commands_before:
        print(f"Chart created ({commands_before} -> {commands_after}). "
              f"Check visually for both axis labels.")
    else:
        print(f"FAILED: command count unchanged. Check Minitab's "
              f"Session window for the real error.")


if __name__ == "__main__":
    main()