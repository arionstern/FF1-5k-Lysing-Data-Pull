"""
tests/minitab/test_explore_existing_charts.py

Exploratory script — not an automated test suite (no assertions).
Lists every recorded Command in the Minitab project, looking for the
existing Xbar and Boxplot charts (per the original instructions, these
already exist manually — step 5 says Xbar "updates automatically",
implying it's a linked/existing chart, not something built fresh each
time). Same approach as everything else: look at what's really there
before writing chart-generation logic.

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

    print(f"Total commands in project: {project.Commands.Count}\n")

    # First, discover the real properties available on a command object
    first_command = project.Commands.Item(1)
    print("--- dir(command) ---")
    print([a for a in dir(first_command) if not a.startswith("_")])
    print()

    for i in range(1, project.Commands.Count + 1):
        command = project.Commands.Item(i)

        # Try several plausible property names for the command text
        cmd_text = None
        for attr_name in ("Text", "CommandLine", "Language", "Title",
                           "Name", "CmdLine"):
            try:
                cmd_text = getattr(command, attr_name)
                if cmd_text:
                    break
            except Exception:
                continue

        try:
            output_count = command.Outputs.Count
        except Exception:
            output_count = "?"

        marker = ""
        text_upper = str(cmd_text).upper()
        if "XBAR" in text_upper or "BOXPLOT" in text_upper:
            marker = "  <-- LOOKS LIKE A CHART COMMAND"

        print(f"[{i}] outputs={output_count}: {cmd_text!r}{marker}")

    print("\nDone. Look for 'LOOKS LIKE A CHART COMMAND' lines above —")
    print("those are the real Xbar/Boxplot command text to build")
    print("chart-generation logic around, instead of guessing syntax.")


if __name__ == "__main__":
    main()