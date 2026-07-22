"""
sap_utils.py

Reusable functions for navigating SAP (ZPP_WI) and reaching the
embedded GEM log document for a given work order.

APPROACH (confirmed working across 3 real lots, including 2 fully
completed/closed orders that a naive approach failed on):
  1. Navigate directly to a known WO number.
  2. Click the Documents List tab (tabpTAB02) - a FIXED tab ID, always
     the same regardless of lot or order status.
  3. Search its DOKNR column for the document number, which is fully
     DERIVABLE from the WO number alone:
       "0000" + WO_NUMBER + "000006" + "0030" + "-01"
     (000006/0030 are the sub-operation/operation for the GEM log
     document specifically - constant for every lot).
  4. Open that document. SAP writes it to a predictable local path:
       %LOCALAPPDATA%\\SAP\\SAP GUI\\tmp\\{doc_number}.xlsx
  5. Poll for that file and open it directly.

WHY NOT THE ROUTING TAB: an earlier approach tried clicking operation
0030's row directly in the Routing tab (tabpTAB01). That grid is a
custom, non-standard control where:
  - Row numbers are per-lot AND per-click unstable (confirmed: the
    same row shifted between two clicks in one session).
  - selectItem/pressButton/doubleClickItem all failed identically via
    scripting on fully-completed orders, despite working via manual
    click and via read-only enumeration (GetAllNodeKeys) on the same
    control - likely some SAP-side restriction on scripted interaction
    with closed-out orders specifically.
The Documents List approach avoids all of this: it's a standard ALV
grid with real column names (not the custom tree), doesn't depend on
row position, and worked identically on both active and closed orders.
"""

import os
import time
import win32com.client
import config


# ---------------------------------------------------------------------------
# Session / navigation
# ---------------------------------------------------------------------------

def get_sap_session():
    """Attach to the active SAP GUI session. Requires SAP GUI already
    open and logged in, and scripting enabled (Options > Accessibility
    & Scripting > Scripting)."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def navigate_to_wo_directly(session, wo_number):
    """Navigate straight to a known WO number via ZPP_WI, skipping the
    F4 multi-select popup used when searching by material instead.

    Uses "/n" to force a full transaction reset -- without it, SAP
    doesn't reliably return to the initial entry screen if the session
    is already sitting deep inside a previous screen (e.g. left
    mid-navigation from an earlier run).
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_WI"
    session.findById("wnd[0]").sendVKey(0)

    order_field_id = "wnd[0]/usr/ctxtZZWOSCAN-AUFNR"
    session.findById(order_field_id).text = str(wo_number)
    session.findById("wnd[0]").sendVKey(0)  # Enter


# ---------------------------------------------------------------------------
# Document number derivation (the key discovery)
# ---------------------------------------------------------------------------

def derive_gem_log_document_number(wo_number):
    """Derive the GEM log document number directly from the WO number.
    Confirmed against 3 real lots at different completion states, all
    matched. No SAP interaction needed for this step -- pure string
    formatting."""
    return f"0000{wo_number}{config.DOC_NUMBER_SUFFIX}"


def get_expected_temp_file_path(doc_number):
    """Predict the local path SAP will write the opened document to."""
    temp_folder = os.path.expandvars(config.SAP_TEMP_FOLDER_TEMPLATE)
    return os.path.join(temp_folder, f"{doc_number}.xlsx")


# ---------------------------------------------------------------------------
# Documents List navigation (the proven path)
# ---------------------------------------------------------------------------

def find_document_candidates(session, wo_number):
    """Build an ordered list of candidate documents to try, ranked by
    confidence. Does NOT check content — a matched document can still
    be genuinely blank (confirmed on 260610C: the derived-number match
    was a real, correctly-numbered but empty template, not a
    mislabeled autosave). Content must be checked after opening, by
    the caller — see get_gem_log_sheet_for_wo().

    Returns a list of dicts: {"row": int, "doc_number": str, "reason": str}
    """
    docs_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB02"
    session.findById(docs_tab_id).select()
    time.sleep(1)

    grid_id = (
        "wnd[0]/usr/tabsTABS_0200/tabpTAB02/"
        "ssubSUBSCREEN:SAPLZPPDB:0100/cntlZPPDB_CONT/shellcont/shell"
    )
    grid = session.findById(grid_id)
    row_count = grid.RowCount
    print(f"  Documents List has {row_count} rows.")

    candidates = []

    doc_number = derive_gem_log_document_number(wo_number)
    print(f"  Derived document number: {doc_number}")
    for row in range(row_count):
        cell_value = grid.GetCellValue(row, "DOKNR")
        if cell_value and cell_value.strip() == doc_number:
            candidates.append({
                "row": row, "doc_number": cell_value.strip(),
                "reason": "derived number match",
            })
            print(f"  Candidate 1: row {row} (derived number match)")
            break

    keywords = [k.lower().replace("-", " ") for k in config.DOC_LIST_SEARCH_KEYWORDS]
    scored_rows = []
    for row in range(row_count):
        # CONFIRMED real bug: without normalizing hyphens to spaces,
        # the keyword "in-process" never matches a real document
        # description like "ATTACHED IN PROCESS GEM LOGS" (space, not
        # hyphen) -- costing the correct document a point it should
        # get, and tying its score with unrelated master documents
        # (e.g. "FILLING OF GEM 5000 LYSING BAGS...") that happen to
        # match "gem"+"fill" instead. Ties break by row number, which
        # can rank the wrong documents ahead of the genuinely correct
        # one. Normalizing both sides the same way fixes the match.
        description = (grid.GetCellValue(row, "DKTXT") or "").lower().replace("-", " ")
        score = sum(1 for k in keywords if k in description)
        if score >= 2:
            scored_rows.append((score, row))
    scored_rows.sort(reverse=True)  # highest score first

    for score, row in scored_rows:
        row_doc_number = (grid.GetCellValue(row, "DOKNR") or "").strip()
        if any(c["row"] == row for c in candidates):
            continue  # already the derived-number candidate
        candidates.append({
            "row": row, "doc_number": row_doc_number,
            "reason": f"keyword score {score}",
        })
        print(f"  Candidate {len(candidates)}: row {row} "
              f"(keyword score {score})")

    if not candidates:
        raise ValueError(
            f"No candidate documents found for WO {wo_number} — "
            f"neither derived number nor keyword search matched anything."
        )

    return grid, candidates


def navigate_back_to_order_screen(session):
    """Navigate back to the order-level screen (with the tab strip),
    since opening a document navigates one level deeper into a
    "Display Document" screen — needed before trying another
    candidate document."""
    session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Back
    time.sleep(1)


def open_document_row(session, row):
    """Click a specific Documents List row open and return the real
    document number from that row (for temp-file polling).

    Re-selects the tab and sets grid focus first — needed when this
    isn't the first candidate tried in a session, since opening a
    document navigates SAP itself one level deeper (into a "Display
    Document" screen), not just launches Excel — the order-level tab
    strip doesn't exist on that screen, so the caller must navigate
    back first (see navigate_back_to_order_screen).
    """
    docs_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB02"
    session.findById(docs_tab_id).select()
    grid_id = (
        "wnd[0]/usr/tabsTABS_0200/tabpTAB02/"
        "ssubSUBSCREEN:SAPLZPPDB:0100/cntlZPPDB_CONT/shellcont/shell"
    )
    grid = session.findById(grid_id)  # re-fetched fresh, not reused —
                                       # stale after navigating away/back
    grid.SetFocus()
    time.sleep(0.5)

    doc_number = (grid.GetCellValue(row, "DOKNR") or "").strip()
    print(f"  Opening row {row} (doc {doc_number})...")

    grid.setCurrentCell(row, "DOKNR")
    grid.selectedRows = str(row)
    grid.clickCurrentCell()
    time.sleep(1.5)

    file_tree_id = (
        "wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/"
        "ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/"
        "shell/shellcont[1]/shell[1]"
    )
    file_tree = session.findById(file_tree_id)
    file_tree.selectNode("          1")
    file_tree.doubleClickNode("          1")

    return doc_number


def snapshot_temp_folder():
    """Record current files in the SAP temp folder, to detect whatever
    NEW file appears after opening a document — more robust than
    predicting a filename, since different document types use
    different naming conventions (confirmed: 'ATA1...' documents use
    their original uploaded filename, not doc_number.xlsx)."""
    temp_folder = os.path.expandvars(config.SAP_TEMP_FOLDER_TEMPLATE)
    if not os.path.exists(temp_folder):
        return set()
    return set(os.listdir(temp_folder))


def wait_for_new_temp_file(start_time, max_attempts=10, poll_seconds=1):
    """Poll the SAP temp folder for the most recently modified .xlsx
    file, checked against start_time (must be captured BEFORE
    triggering the click — capturing it inside this function was too
    late, since SAP can finish writing during the click's own wait,
    before this function is even called)."""
    temp_folder = os.path.expandvars(config.SAP_TEMP_FOLDER_TEMPLATE)

    for attempt in range(1, max_attempts + 1):
        time.sleep(poll_seconds)
        if not os.path.exists(temp_folder):
            continue
        xlsx_files = [
            f for f in os.listdir(temp_folder)
            if f.lower().endswith(".xlsx") and not f.startswith("~$")
        ]
        for f in xlsx_files:
            full_path = os.path.join(temp_folder, f)
            if os.path.getmtime(full_path) >= start_time:
                return full_path

    raise TimeoutError(
        f"No file modified after {start_time} found in {temp_folder} "
        f"after {max_attempts} attempts."
    )


# ---------------------------------------------------------------------------
# Orchestration helper
# ---------------------------------------------------------------------------

def _document_has_real_content(sheet):
    """Check whether an opened document actually has real data, not
    just a blank template with the right document number (confirmed
    real failure mode on 260610C)."""
    fill_line = sheet.Range(config.SOURCE_FILL_LINE_CELL).Value
    if fill_line in config.VALID_FILL_LINES:
        return True
    return False


def get_gem_log_sheet_for_wo(wo_number):
    """Full pipeline: navigate to a WO, build a ranked list of
    candidate documents, open each in order until one has real content
    (not just a matching document number), and return the
    (workbook, sheet) ready for excel_utils' read functions.

    Caller is responsible for closing the returned workbook when done.
    """
    session = get_sap_session()
    navigate_to_wo_directly(session, wo_number)
    grid, candidates = find_document_candidates(session, wo_number)

    for i, candidate in enumerate(candidates):
        print(f"  Trying candidate: row {candidate['row']} "
              f"({candidate['reason']})")
        if i > 0:
            navigate_back_to_order_screen(session)
        start_time = time.time()
        open_document_row(session, candidate["row"])
        file_path = wait_for_new_temp_file(start_time)
        print(f"  New file: {file_path}")

        excel_app = win32com.client.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(file_path, ReadOnly=True)

        try:
            sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)
        except Exception as e:
            # CONFIRMED real failure case: a candidate matched by
            # keyword score alone (e.g. 'SPM00470088-01') can be a
            # totally different document type -- a Setup Aid/Checklist
            # workbook, not a GEM log at all -- which doesn't even have
            # a sheet named SOURCE_SHEET_NAME. Without this try/except,
            # that raises an uncaught COM exception here, which crashes
            # the whole lot AND leaves this workbook open forever
            # (nothing in scope to close it), leaking an Excel instance
            # every time it happens. Reject and try the next candidate
            # instead, same as the blank-template case below.
            print(f"  REJECTED: candidate doesn't have sheet "
                  f"{config.SOURCE_SHEET_NAME!r} ({e}) -- wrong "
                  f"document type, trying next candidate.")
            workbook.Close(SaveChanges=False)
            continue

        if _document_has_real_content(sheet):
            print(f"  ACCEPTED: real content found.")
            return workbook, sheet

        print(f"  REJECTED: blank/template content, trying next "
              f"candidate.")
        workbook.Close(SaveChanges=False)

    raise ValueError(
        f"No candidate document for WO {wo_number} had real content "
        f"— all {len(candidates)} candidate(s) were blank templates "
        f"or the wrong document type."
    )


# ---------------------------------------------------------------------------
# Finding new lots (GR Quantity table)
# ---------------------------------------------------------------------------

def is_valid_lot_name(name):
    """Same validation used in the last project's lot_utils.py:
    YYMMDD + letter(s), e.g. '260610C' or '250121AA'."""
    return len(name) >= 7 and name[:6].isdigit() and name[6:].isalpha()


def parse_lot_date(batch):
    """Parse the YYMMDD prefix from a batch/lot name into a real
    datetime object, e.g. '260630C' -> datetime(2026, 6, 30). Returns
    None if the batch doesn't match the expected format.

    NOTE: returns datetime, not date — both Excel COM and Minitab COM
    reject a raw datetime.date object ("must be a pywintypes time
    object" / similar errors), confirmed via real testing.
    """
    from datetime import datetime
    if not is_valid_lot_name(batch):
        return None
    try:
        yy, mm, dd = int(batch[0:2]), int(batch[2:4]), int(batch[4:6])
        return datetime(2000 + yy, mm, dd)
    except ValueError:
        return None


def _read_visible_gr_table_rows(usr_area):
    """Read whatever rows of the GR Quantity table are currently
    rendered, keyed by row number. This is an old-style table built
    from individual lbl[col,row] labels, not a named-column grid —
    confirmed via exploration. Only currently-visible rows exist as
    real controls; scrolling is required to see more."""
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


def find_new_lots(session, last_known_lot, page_size=20):
    """Find lots newer than last_known_lot in the GR Quantity table
    for config.MATERIAL_NUMBER, by scrolling from the bottom upward
    and stopping as soon as last_known_lot is reached.

    Column mapping confirmed via exploration:
      col 1 = Batch, col 10 = Order, col 19 = Item Quantity,
      col 37 = GR Quantity, col 55 = DCI (checkbox)

    NOTE: jumping directly to the scrollbar's Maximum position
    produced DIFFERENT (wrong, older) results than reaching the same
    position incrementally during testing — this table's rendering
    appears to depend on scroll path, not just final position. This
    function jumps directly, matching what was actually tested and
    confirmed correct; be cautious about assuming it'll always land
    on the true bottom if SAP's table size changes.

    Returns a dict keyed by Order number: {"batch", "order",
    "item_qty", "gr_qty"} for each lot newer than last_known_lot,
    filtered to real lot-name format only.
    """
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

    new_lots = {}
    position = max_position
    found_known_lot = False

    while position >= 0 and not found_known_lot:
        scrollbar.Position = position
        time.sleep(0.3)

        page_rows = _read_visible_gr_table_rows(usr_area)
        for row, cols in sorted(page_rows.items(), reverse=True):
            batch = cols.get(1, "").strip()
            order = cols.get(10, "").strip()
            if not batch or not order or order in ("Order",):
                continue

            if batch == last_known_lot:
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

    return new_lots


def filter_ready_lots(new_lots):
    """Filter find_new_lots() results down to lots where GR Quantity
    != 0 (per original step 3.3.1 — GR quantity 0 means not ready)."""
    ready = {}
    for order, data in new_lots.items():
        gr_qty_str = data["gr_qty"].replace(",", "")
        try:
            gr_qty_val = float(gr_qty_str)
        except ValueError:
            gr_qty_val = 0
        if gr_qty_val != 0:
            ready[order] = data
    return ready