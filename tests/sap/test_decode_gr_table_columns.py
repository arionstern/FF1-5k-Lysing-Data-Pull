"""
tests/sap/test_decode_gr_table_columns.py

Reads every label's text across all rows of the GR Quantity table,
printing (column_position, row, text) for each — to empirically map
which column position holds Batch/Order/Item Quantity/GR Quantity/DCI,
by comparing against a known lot (260610C: Order 10556542,
Item Qty 6000, GR Qty 5070).

This is a plain old-style table (lbl[col,row] positions), not a named-
column ALV grid — confirmed via the previous enumeration script.

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
    session.findById("wnd[0]").sendVKey(4)
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

    print("Reading all labels in wnd[1]/usr...\n")

    usr_area = session.findById("wnd[1]/usr")
    children = usr_area.Children

    # Collect all (col, row, text) for GuiLabel elements, parsed from
    # the field ID itself (lbl[col,row])
    rows = {}
    for i in range(children.Count):
        child = children.ElementAt(i)
        if child.Type not in ("GuiLabel", "GuiCheckBox"):
            continue
        field_id = child.Id
        # Extract "col,row" from something like ".../lbl[37,24]"
        try:
            bracket_content = field_id.split("[")[-1].rstrip("]")
            col_str, row_str = bracket_content.split(",")
            col, row = int(col_str), int(row_str)
        except Exception:
            continue

        try:
            text = child.Text if child.Type == "GuiLabel" else child.Selected
        except Exception:
            text = "?"

        rows.setdefault(row, {})[col] = text

    for row in sorted(rows.keys()):
        cols = rows[row]
        col_items = sorted(cols.items())
        print(f"Row {row}: {col_items}")


if __name__ == "__main__":
    main()