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

    print("Reading full table by scrolling through it...\n")

    usr_area = session.findById("wnd[1]/usr")
    scrollbar = usr_area.verticalScrollbar
    max_position = scrollbar.Maximum
    print(f"Scrollbar maximum: {max_position}")

    import re
    lysing_lot_pattern = re.compile(r"^\d{6}[A-Z]$")

    all_rows = {}
    seen_any_real_lot = False
    consecutive_non_matches = 0
    position = 0
    page_size = 20
    stopped_early = False

    while True:
        scrollbar.Position = position
        time.sleep(0.3)

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

        for row, cols in sorted(page_rows.items()):
            order = cols.get(10, "").strip()
            if not order or order == "Order":
                continue
            batch = cols.get(1, "").strip()

            if lysing_lot_pattern.match(batch):
                seen_any_real_lot = True
                consecutive_non_matches = 0
            elif seen_any_real_lot:
                # Once we've seen real lots, count how many non-matching
                # rows follow in a row — a real transition into the
                # "odd" section, not just an isolated blank/junk row.
                consecutive_non_matches += 1
                if consecutive_non_matches >= 5:
                    print(f"Stopping early: hit {consecutive_non_matches} "
                          f"consecutive non-Lysing entries after row "
                          f"{row} — crossed into the unrelated section.")
                    stopped_early = True
                    break

            all_rows[order] = {
                "batch": batch,
                "order": order,
                "item_qty": cols.get(19, "").strip(),
                "gr_qty": cols.get(37, "").strip(),
            }

        if stopped_early:
            break
        if position >= max_position:
            break
        position = min(position + page_size, max_position)

    print(f"\nTotal unique orders found: {len(all_rows)}\n")

    # Format alone can't distinguish real recent Lysing lots from old
    # unrelated batches — confirmed both use the same YYMMDD+letter
    # naming convention across a full decade of history. Filtering by
    # the DATE embedded in the batch name instead: only accept batches
    # from the last N days, and not already known.
    from datetime import datetime, timedelta
    RECENCY_WINDOW_DAYS = 90
    cutoff_date = datetime.now() - timedelta(days=RECENCY_WINDOW_DAYS)

    def parse_batch_date(batch):
        """Parse the YYMMDD prefix from a batch name like '260610C'."""
        if not lysing_lot_pattern.match(batch):
            return None
        try:
            yy, mm, dd = int(batch[0:2]), int(batch[2:4]), int(batch[4:6])
            return datetime(2000 + yy, mm, dd)
        except ValueError:
            return None

    new_lots = {}
    for order, data in all_rows.items():
        batch = data["batch"]
        batch_date = parse_batch_date(batch)
        if batch_date is None or batch_date < cutoff_date:
            continue
        if batch in config.KNOWN_LOTS:  # TODO: read from real WO Data
            continue
        new_lots[order] = data

    print(f"Filtered to {len(new_lots)} recent, unknown Lysing lot(s) "
          f"(last {RECENCY_WINDOW_DAYS} days):\n")
    for order, data in sorted(new_lots.items(), key=lambda x: x[1]["batch"]):
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