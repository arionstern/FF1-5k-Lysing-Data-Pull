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

    keywords = [k.lower() for k in config.DOC_LIST_SEARCH_KEYWORDS]
    scored_rows = []
    for row in range(row_count):
        description = (grid.GetCellValue(row, "DKTXT") or "").lower()
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


def open_document_row(session, grid, row):
    """Click a specific Documents List row open and return the real
    document number from that row (for temp-file polling).

    Re-selects the tab and sets grid focus first — needed when this
    isn't the first candidate tried in a session, since the previous
    candidate's Excel window opening/closing can steal SAP's focus.
    """
    docs_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB02"
    session.findById(docs_tab_id).select()
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

    for candidate in candidates:
        print(f"  Trying candidate: row {candidate['row']} "
              f"({candidate['reason']})")
        start_time = time.time()
        open_document_row(session, grid, candidate["row"])
        file_path = wait_for_new_temp_file(start_time)
        print(f"  New file: {file_path}")

        excel_app = win32com.client.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(file_path, ReadOnly=True)
        sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)

        if _document_has_real_content(sheet):
            print(f"  ACCEPTED: real content found.")
            return workbook, sheet

        print(f"  REJECTED: blank/template content, trying next "
              f"candidate.")
        workbook.Close(SaveChanges=False)

    raise ValueError(
        f"No candidate document for WO {wo_number} had real content "
        f"— all {len(candidates)} candidate(s) were blank templates."
    )