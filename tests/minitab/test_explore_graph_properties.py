"""
tests/minitab/test_explore_graph_properties.py

Exploratory script — XLabel/YLabel subcommands don't exist for
Boxplot (confirmed twice now, with and without Overlay), yet the
screenshot shows Title/X-axis label/Y-axis label ARE settable via
Minitab's own "Text Annotations" panel. That strongly suggests these
are properties on the Graph object itself, set via COM after chart
creation — not part of the Boxplot command syntax. This creates the
base chart (known-working syntax, no Title/XLabel/YLabel), then
inspects the resulting Graph object's real properties.

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

    command_text = (
        f"Boxplot '{first_col}'-'{last_col}';\n"
        f"  Overlay;\n"
        f"  IQRBox;\n"
        f"  Outlier."
    )
    project.ExecuteCommand(command_text)

    # Get the command we just created and its Graph output
    command = project.Commands.Item(project.Commands.Count)
    print(f"Command created: {command.Name}")

    graph = command.Outputs.Item(1).Graph
    print(f"\n--- dir(graph) ---")
    attrs = [a for a in dir(graph) if not a.startswith("_")]
    print(attrs)

    # Try a handful of plausible title/label-related attribute names
    print(f"\n--- Trying plausible title/label attributes ---")
    for attr_name in ("Title", "MainTitle", "GraphTitle", "XAxis", "YAxis",
                       "XAxisLabel", "YAxisLabel", "Annotations"):
        try:
            value = getattr(graph, attr_name)
            print(f"  {attr_name}: {value!r} (type: {type(value)})")
        except Exception as e:
            print(f"  {attr_name}: not accessible ({e})")


if __name__ == "__main__":
    main()