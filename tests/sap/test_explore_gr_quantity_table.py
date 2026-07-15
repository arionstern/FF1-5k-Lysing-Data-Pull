"""
tests/sap/test_explore_gr_quantity_table.py

Exploratory script — navigates to the GR Quantity results table
(reusing the known F4/"no restriction"/material-number flow, which
turns out to land directly on this table rather than a separate
screen) and inspects its real structure, since the existing recording
only shows clicking one row by position (lbl[37,24]), not reading the
whole table.

Requires: SAP GUI open and logged in.
"""

import sys
import time
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
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_WI"
    session.findById("wnd[0]").sendVKey(0)
    order_field_id = "wnd[0]/usr/ctxtZZWOSCAN-AUFNR"
    session.findById(order_field_id).setFocus()
    session.findById(order_field_id).caretPosition = 0
    session.findById("wnd[0]").sendVKey(4)  # F4, opens the popup
    time.sleep(1)

    checkbox_id = (
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON"
    )
    session.findById(checkbox_id).selected = True

    material_field_id = (
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/"
        "ctxtG_SELFLD_TAB-LOW[0,24]"
    )
    session.findById(material_field_id).text = config.MATERIAL_NUMBER
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print("Landed on results screen. Enumerating wnd[1]/usr children...")
    print(f"Transaction: {session.Info.Transaction}\n")

    try:
        usr_area = session.findById("wnd[1]/usr")
        children = usr_area.Children
        print(f"wnd[1]/usr has {children.Count} direct children:\n")
        for i in range(children.Count):
            child = children.ElementAt(i)
            child_type = child.Type
            child_id = child.Id
            subtype = ""
            try:
                subtype = child.SubType
            except Exception:
                pass
            print(f"  [{i}] {child_type}"
                  f"{f' (SubType: {subtype})' if subtype else ''}: {child_id}")
    except Exception as e:
        print(f"FAILED to enumerate wnd[1]/usr: {e}")

    print("\nDone. Look for a GuiTableControl, GuiShell (ALV grid), or")
    print("similar — that's the real table structure to read Batch/")
    print("Order/GR Quantity from. If it's plain lbl[row,col]/txt[row,col]")
    print("labels instead, we'll need position-based reading like the")
    print("original recording used.")


if __name__ == "__main__":
    main()