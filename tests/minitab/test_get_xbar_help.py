"""
tests/minitab/test_get_xbar_help.py

Same technique as test_get_boxplot_help.py: uses Minitab's own
built-in "Help" command to get the real Xbar chart syntax, since we
need to recreate it via ExecuteCommand (to add AxLabel automation)
rather than continuing to rely on the existing auto-updating chart.

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

    commands_before = project.Commands.Count
    project.ExecuteCommand("Help XBARCHART.")
    commands_after = project.Commands.Count

    print(f"Commands before/after: {commands_before}/{commands_after}")

    if commands_after > commands_before:
        help_command = project.Commands.Item(commands_after)
        print(f"New command: {help_command.Name}")
        try:
            output_text = help_command.Outputs.Item(1).Text
            print(f"\n--- Help output ---\n{output_text}")
        except Exception as e:
            print(f"Couldn't read output text: {e}")
    else:
        print("No new command created — try checking Minitab's Help "
              "window directly if it opened separately.")


if __name__ == "__main__":
    main()