"""
tests/minitab/test_graph_saveas_dimensions.py

Tests whether Graph.SaveAs() accepts optional width/height parameters
— dir() only showed the method name, not its full signature, and COM
methods often hide optional parameters from a bare dir() listing.

Requires: config.USE_TEST_PATHS = True.
"""

import sys
import os

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

    xbar_command = minitab_utils.regenerate_xbar_chart(project, control_chart_sheet)
    graph = xbar_command.Outputs.Item(1).Graph

    output_folder = os.path.join(os.getcwd(), "test_chart_dimensions")
    os.makedirs(output_folder, exist_ok=True)

    # Try a few plausible SaveAs signatures — we don't know the real
    # one, so testing several rather than guessing just one
    attempts = [
        ("default (no size args)", lambda: graph.SaveAs(
            os.path.join(output_folder, "default.png"))),
        ("positional width/height", lambda: graph.SaveAs(
            os.path.join(output_folder, "positional.png"), 1200, 600)),
        ("named Width/Height", lambda: graph.SaveAs(
            os.path.join(output_folder, "named.png"), Width=1200, Height=600)),
    ]

    for label, attempt in attempts:
        print(f"\n--- Trying: {label} ---")
        try:
            attempt()
            print("SUCCESS — check the output file's actual dimensions.")
        except Exception as e:
            print(f"FAILED: {e}")

    print(f"\nDone. Check {output_folder!r} for whichever files were "
          f"created, and compare their actual pixel dimensions/aspect "
          f"ratios (e.g. via image properties or PIL).")


if __name__ == "__main__":
    main()