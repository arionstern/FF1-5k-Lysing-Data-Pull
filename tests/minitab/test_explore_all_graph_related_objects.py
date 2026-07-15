"""
tests/minitab/test_explore_all_graph_related_objects.py

Broader exploration — we only checked command.Outputs.Item(1).Graph
before (very limited: AddRef/CopyToClipboard/SaveAs/etc). This checks
several OTHER object levels that might expose title/axis-label control:
  - project itself (does a GraphWindows-style collection exist?)
  - command.OutputDocument (appeared in dir(command) earlier, never
    actually opened)
  - mtb (application level) for anything graph-window related

Requires: config.USE_TEST_PATHS = True.
"""

import sys

sys.path.append("../..")
import config
import minitab_utils


def safe_dir(obj, label):
    print(f"\n--- dir({label}) ---")
    try:
        attrs = [a for a in dir(obj) if not a.startswith("_")]
        print(attrs)
        return attrs
    except Exception as e:
        print(f"FAILED: {e}")
        return []


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
    command = project.Commands.Item(project.Commands.Count)

    # Level 1: application object
    safe_dir(mtb, "mtb (application)")

    # Level 2: project object — checking specifically for anything
    # graph-window related
    project_attrs = safe_dir(project, "project")
    graph_related = [a for a in project_attrs if "graph" in a.lower()
                      or "window" in a.lower()]
    print(f"\nProject attrs containing 'graph' or 'window': {graph_related}")

    # Level 3: command.OutputDocument — seen in dir(command) earlier,
    # never actually explored
    print(f"\n--- Trying command.OutputDocument ---")
    try:
        output_doc = command.OutputDocument
        safe_dir(output_doc, "command.OutputDocument")
    except Exception as e:
        print(f"FAILED: {e}")

    # Level 4: the Outputs collection itself (not just .Item(1).Graph)
    safe_dir(command.Outputs, "command.Outputs")

    # Level 5: command.Outputs.Item(1) itself, one level above .Graph
    output_item = command.Outputs.Item(1)
    safe_dir(output_item, "command.Outputs.Item(1)")

    # If any GraphWindows-style collection turned up on project, try it
    if graph_related:
        for attr_name in graph_related:
            print(f"\n--- Trying project.{attr_name} ---")
            try:
                collection = getattr(project, attr_name)
                print(f"  Value: {collection}")
                if hasattr(collection, "Count"):
                    print(f"  Count: {collection.Count}")
                    if collection.Count > 0:
                        safe_dir(collection.Item(1), f"project.{attr_name}.Item(1)")
            except Exception as e:
                print(f"  FAILED: {e}")

    print("\nDone. Look through all the dir() dumps above for anything")
    print("Title/Label/Axis/Scale/Annotation-related we haven't tried yet.")


if __name__ == "__main__":
    main()