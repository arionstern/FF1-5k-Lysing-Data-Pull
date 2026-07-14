"""
tests/sap/test_dynamic_routing_lookup.py

Exploratory script — not an automated test suite (no assertions).
Confirms the real root cause found while backfilling 260610C:
navigating directly by WO number lands on whatever operation is
CURRENTLY ACTIVE for that order (e.g. "Final QC Testing" / 0060 for
an older, further-along lot) — not reliably on operation 0030
"In-Process GEM Logs". A fixed row/column selector on the routing grid
was always going to break on lots at a different lifecycle stage.

This script explicitly clicks the Routing tab, then reads every row in
the routing grid and finds the one matching config.ROUTING_STEP_ID by
TEXT, not position — this is the real fix, replacing the hardcoded
item_key/column_key used before.

Requires: SAP GUI open, already navigated to a WO (any operation is
fine — this script explicitly goes to Routing regardless of where you
land).
"""

import sys
import win32com.client

sys.path.append("../..")
import config


def get_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def main():
    session = get_session()
    print(f"Session: {session.Info.SystemName}, "
          f"transaction: {session.Info.Transaction}")

    # Explicitly click the Routing tab (tabpTAB01), rather than relying
    # on it being the default landing tab.
    routing_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB01"
    try:
        session.findById(routing_tab_id).Select()
        print("Selected Routing tab.")
    except Exception as e:
        print(f"FAILED to select Routing tab: {e}")
        return

    # The routing grid's container ID differs from the Documents tab's
    # container (cntlZPPDA_CONT here, vs cntlZPPDR_CONT on Documents) —
    # confirmed from the earlier recording's grid_id.
    routing_grid_id = (
        "wnd[0]/usr/tabsTABS_0200/tabpTAB01/"
        "ssubSUBSCREEN:SAPLZPPDA:0100/cntlZPPDA_CONT/shellcont/"
        "shell/shellcont[1]/shell[1]"
    )
    try:
        routing_grid = session.findById(routing_grid_id)
    except Exception as e:
        print(f"FAILED to find routing grid: {e}")
        return

    # Read every row's text to find the one matching ROUTING_STEP_ID,
    # instead of assuming a fixed row number. GuiShell tree/grid
    # controls expose GetItemText(item, column) — try common column
    # keys since we don't yet know which column holds the op number.
    print(f"\nSearching routing grid for '{config.ROUTING_STEP_ID}' "
          f"or '{config.ROUTING_DESCRIPTION}'...")

    try:
        row_count = routing_grid.RowCount
        print(f"Grid reports {row_count} rows via RowCount.")
    except Exception:
        row_count = None
        print("RowCount not available on this control (tree-type grids "
              "often don't expose it directly) — will try node-based "
              "enumeration instead.")

    # Get the REAL column keys instead of guessing — GuiShell grids
    # expose a ColumnOrder collection with the actual identifiers.
    try:
        column_order = routing_grid.ColumnOrder
        real_columns = [column_order(i) for i in range(column_order.Count)]
        print(f"\nReal column keys: {real_columns}")
    except Exception as e:
        print(f"FAILED to read ColumnOrder: {e}")
        real_columns = []

    found_key = None
    try:
        node_keys = routing_grid.GetAllNodeKeys()
        print(f"\nFound {node_keys.Count} node keys. Dumping each row's "
              f"text across all real columns (including HierarchyHeader, "
              f"debugged explicitly since it came back empty last time)...\n")
        for i in range(node_keys.Count):
            key = node_keys(i)
            row_text = {}

            # Debug HierarchyHeader specifically — don't silently skip
            # empty/failed results this time
            try:
                hierarchy_text = routing_grid.GetItemText(key, "HierarchyHeader")
                print(f"Row (key={key!r}) HierarchyHeader: {hierarchy_text!r}")
                if hierarchy_text:
                    row_text["HierarchyHeader"] = hierarchy_text
            except Exception as e:
                print(f"Row (key={key!r}) HierarchyHeader FAILED: {e}")

            for col_key in real_columns:
                if col_key == "HierarchyHeader":
                    continue
                try:
                    text = routing_grid.GetItemText(key, col_key)
                    if text:
                        row_text[col_key] = text
                except Exception:
                    continue
            print(f"  Full row: {row_text}\n")

            combined = " ".join(row_text.values())
            if (config.ROUTING_STEP_ID in combined
                    or "0030" in combined
                    or config.ROUTING_DESCRIPTION in combined):
                print(f"  ^ MATCH")
                found_key = key
    except Exception as e:
        print(f"GetAllNodeKeys failed: {e}")

    if found_key:
        print(f"\nFound matching row at key: {found_key!r}")
        print("Next: use this key with pressButton/selectItem to open it, "
              "same as before, but with a REAL key instead of a hardcoded "
              "'12'.")
    else:
        print("\nNo automatic match found. Manually note which row number "
              "(top to bottom, 1-indexed) contains operation 0030 in the "
              "Routing tab right now, and we'll try a more targeted "
              "enumeration approach.")


if __name__ == "__main__":
    main()