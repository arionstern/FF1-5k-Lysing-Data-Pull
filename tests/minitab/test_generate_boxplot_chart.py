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

    column_names = minitab_utils.get_current_year_boxplot_column_names(boxplot_sheet)
    print(f"Found {len(column_names)} columns for the current year.")

    first_col = column_names[0]
    last_col = column_names[-1]

    command_text = (
        f"Boxplot '{first_col}'-'{last_col}';\n"
        f"  Overlay;\n"
        f"  IQRBox;\n"
        f"  Outlier."
    )

    print(f"\nAttempting real captured syntax:\n{command_text}\n")

    commands_before = project.Commands.Count
    project.ExecuteCommand(command_text)
    commands_after = project.Commands.Count

    if commands_after > commands_before:
        print(f"Command created ({commands_before} -> {commands_after}). "
              f"Check Minitab visually — is it ONE combined chart now, "
              f"not separate tiled ones?")
    else:
        print(f"FAILED: command count unchanged. Check Minitab's "
              f"Session window for the real error.")


if __name__ == "__main__":
    main()