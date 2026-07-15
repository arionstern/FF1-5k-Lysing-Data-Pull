"""
tests/sap/test_find_new_lots_backward_scan.py

Scrolls from the BOTTOM of the GR Quantity table upward, stopping as
soon as it reaches a lot already known (in WO Data) — far more
efficient than reading the entire multi-year history (626 rows,
confirmed via the full-scroll test), and gives a natural stopping
condition for "find new lots since last check."

LAST_KNOWN_LOT is hardcoded here for testing — in real use this
should come from reading the actual last row of WO Data column E.

Requires: SAP GUI open and logged in.
"""

import sys
import time
import win32com.client

sys.path.append("../..")
import config

# TODO: read this from the real WO Data sheet instead of hardcoding —
# this is just for testing the backward-scan mechanism itself.
LAST_KNOWN_LOT = "260630C"


def is_valid_lot_name(name):
    """Same validation used in the last project's lot_utils.py:
    YYMMDD + letter(s), e.g. '260610C' or '250121AA'."""
    return len(name) >= 7 and name[:6].isdigit() and name[6:].isalpha()


def get_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def read_visible_rows(usr_area):
    """Read whatever rows are currently rendered, keyed by row number."""
    children = usr_area.Children
    page_rows = {}
    for i in range(children.Count):
        child = children.ElementAt(i)
        if child.Type not in ("GuiLabel", "GuiCheckBox"):
            continue
        field_id = child.Id
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
        page_rows.setdefault(row, {})[col] = text
    return page_rows


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

    usr_area = session.findById("wnd[1]/usr")
    scrollbar = usr_area.verticalScrollbar
    max_position = scrollbar.Maximum
    print(f"Scrollbar maximum: {max_position}")
    print(f"Scanning backward from the bottom, stopping at "
          f"{LAST_KNOWN_LOT!r}...\n")

    new_lots = {}
    position = max_position
    page_size = 20
    found_known_lot = False

    while position >= 0 and not found_known_lot:
        scrollbar.Position = position
        time.sleep(0.3)

        page_rows = read_visible_rows(usr_area)
        for row, cols in sorted(page_rows.items(), reverse=True):
            batch = cols.get(1, "").strip()
            order = cols.get(10, "").strip()
            if not batch or not order or order in ("Order",):
                continue

            if batch == LAST_KNOWN_LOT:
                print(f"Reached known lot {LAST_KNOWN_LOT!r} — stopping scan.")
                found_known_lot = True
                break

            if not is_valid_lot_name(batch):
                continue

            if order not in new_lots:
                new_lots[order] = {
                    "batch": batch,
                    "order": order,
                    "item_qty": cols.get(19, "").strip(),
                    "gr_qty": cols.get(37, "").strip(),
                }

        if position == 0:
            break
        position = max(position - page_size, 0)

    print(f"\nFound {len(new_lots)} lot(s) newer than {LAST_KNOWN_LOT!r}:\n")
    for order, data in sorted(new_lots.items()):
        gr_qty_str = data["gr_qty"].replace(",", "")
        try:
            gr_qty_val = float(gr_qty_str)
        except ValueError:
            gr_qty_val = 0
        ready = "READY" if gr_qty_val != 0 else "NOT READY (GR qty = 0)"
        print(f"  Batch={data['batch']!r} Order={order!r} "
              f"GRQty={data['gr_qty']!r} [{ready}]")


if __name__ == "__main__":
    main()