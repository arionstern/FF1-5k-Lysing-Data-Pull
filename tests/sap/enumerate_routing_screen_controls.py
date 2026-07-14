"""
tests/sap/enumerate_routing_screen_controls.py

Exploratory script — not an automated test suite (no assertions).
SAP GUI Script Recording can't capture clicks inside the embedded
Excel content on the routing/instructions screen (step 4.5), since
that content is native Excel/OLE, not a SAP GUI control. This script
walks the control tree instead, from the routing screen down, and
prints every control's ID/type/name so we can spot the OLE container
(GuiOLEContainer or similar) without needing SAP to record it.

Requires: SAP GUI already open, logged in, and already sitting on the
routing/instructions screen from step 4.3.1 (run
test_zpp_wi_navigation.py first, then run this one without closing
that screen).
"""

import sys
import win32com.client

sys.path.append("../..")
import config  # noqa: F401 (not used directly, kept for consistency)


def get_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def walk_controls(control, depth=0, max_depth=15):
    """Recursively print every child control's Id/Type/SubType/Name."""
    indent = "  " * depth
    try:
        control_id = control.Id
        control_type = control.Type
        control_name = getattr(control, "Name", "")
        # SubType matters most for GuiShell, which is a generic wrapper
        # used for grids, trees, pictures, AND OLE containers alike —
        # Type alone can't tell them apart.
        control_subtype = ""
        try:
            control_subtype = control.SubType
        except Exception:
            pass
    except Exception as e:
        print(f"{indent}[error reading control: {e}]")
        return

    marker = ""
    combined = f"{control_type} {control_subtype}".upper()
    if "OLE" in combined:
        marker = "  <-- LOOKS LIKE AN OLE CONTAINER"

    subtype_display = f" (SubType: {control_subtype})" if control_subtype else ""
    print(f"{indent}{control_type}{subtype_display}: {control_id}{marker}")

    if depth >= max_depth:
        print(f"{indent}  [max depth reached, stopping recursion]")
        return

    try:
        children = control.Children
        for i in range(children.Count):
            walk_controls(children.ElementAt(i), depth + 1, max_depth)
    except Exception:
        # Not all controls have children (e.g. leaf fields) — normal, skip.
        pass


def main():
    session = get_session()
    print(f"Session: {session.Info.SystemName}, "
          f"transaction: {session.Info.Transaction}")

    # First pass found a Documents *list* tab (GuiShell/ALV grid), not
    # an OLE container — the actual spreadsheet content may be a
    # separate top-level window rather than embedded in wnd[0]/usr.
    # Check several window indices, not just wnd[0].
    for wnd_index in range(5):
        wnd_id = f"wnd[{wnd_index}]"
        try:
            wnd = session.findById(wnd_id)
        except Exception:
            if wnd_index == 0:
                print(f"FAILED: couldn't even find {wnd_id}, something's wrong.")
                return
            # No more windows beyond this index — normal, stop looking.
            break

        print(f"\n=== Walking {wnd_id} ===")
        walk_controls(wnd)

    print("\nDone. Look above for any line marked 'LOOKS LIKE AN OLE")
    print("CONTAINER' — that Id is what excel_utils.py will need to")
    print("grab .OleObject from. If nothing was flagged, search the")
    print("printed output manually for anything with 'OLE', 'CONT',")
    print("or 'shell' in its type/id, since naming isn't always exact.")
    print("\nIf still nothing: confirm the actual spreadsheet content")
    print("(cell values, not a list of document entries) was visible")
    print("on screen at the moment this script ran.")


if __name__ == "__main__":
    main()