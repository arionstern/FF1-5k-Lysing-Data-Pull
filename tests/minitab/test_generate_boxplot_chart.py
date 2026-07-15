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

    column_names = minitab_utils.get_existing_boxplot_column_names(boxplot_sheet)
    print(f"Found {len(column_names)} columns to include in the boxplot.")

    # Build the command text — best guess at classic Minitab Command
    # Language syntax for a multi-Y simple boxplot with title/axis
    # labels. Column names with special characters (like '/') need
    # single quotes around each one.
    column_list = " ".join(f"'{name}'" for name in column_names)
    chart_config = config.MINITAB_BOXPLOT_CHART

    # XLabel/YLabel are NOT valid Boxplot subcommands (confirmed via
    # real Minitab error) — dropping them. Title alone may still work;
    # axis labels may need a different mechanism entirely (possibly
    # not settable via this command at all, or need "Footnote" etc.)
    command_text = (
        f"Boxplot {column_list};\n"
        f"  Title \"{chart_config['title']}\"."
    )

    print(f"\nAttempting command:\n{command_text[:300]}...\n")

    commands_before = project.Commands.Count
    try:
        project.ExecuteCommand(command_text)
    except Exception as e:
        print(f"Python-level exception: {e}")

    # ExecuteCommand does NOT raise on Minitab-side syntax errors — it
    # just prints the error into Minitab's own Session window. The
    # only real way to check success is whether a new command actually
    # got created.
    commands_after = project.Commands.Count
    if commands_after > commands_before:
        print(f"SUCCESS: new command created "
              f"({commands_before} -> {commands_after}).")
    else:
        print(f"FAILED: command count unchanged ({commands_before}) — "
              f"check Minitab's Session window for the real error "
              f"message, since ExecuteCommand won't raise one here.")


if __name__ == "__main__":
    main()