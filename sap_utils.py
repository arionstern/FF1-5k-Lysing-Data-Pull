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
       %LOCALAPPDATA%\SAP\SAP GUI\tmp\{doc_number}.xlsx
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

def open_gem_log_via_documents_list(session, wo_number):
    """Open the GEM log document via the Documents List tab, searching
    for the derived document number in the real DOKNR column. Returns
    the document number that was opened (for use in polling the temp
    file path afterward)."""
    doc_number = derive_gem_log_document_number(wo_number)

    docs_tab_id = "wnd[0]/usr/tabsTABS_0200/tabpTAB02"
    session.findById(docs_tab_id).select()
    time.sleep(1)

    grid_id = (
        "wnd[0]/usr/tabsTABS_0200/tabpTAB02/"
        "ssubSUBSCREEN:SAPLZPPDB:0100/cntlZPPDB_CONT/shellcont/shell"
    )
    grid = session.findById(grid_id)

    matched_row = None
    row_count = grid.RowCount
    for row in range(row_count):
        cell_value = grid.GetCellValue(row, "DOKNR")
        if cell_value and cell_value.strip() == doc_number:
            matched_row = row
            break

    if matched_row is None:
        raise ValueError(
            f"Could not find document {doc_number!r} in the Documents "
            f"List grid ({row_count} rows checked). The document "
            f"number formula may not hold for this WO, or the "
            f"document genuinely isn't listed here."
        )

    grid.setCurrentCell(matched_row, "DOKNR")
    grid.selectedRows = str(matched_row)
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


def wait_for_temp_file(doc_number, max_attempts=10, poll_seconds=1):
    """Poll for SAP to finish writing the opened document to its
    predictable local temp path. Returns the path once found, or
    raises TimeoutError."""
    expected_path = get_expected_temp_file_path(doc_number)

    for attempt in range(1, max_attempts + 1):
        time.sleep(poll_seconds)
        if os.path.exists(expected_path):
            return expected_path

    raise TimeoutError(
        f"File not found after {max_attempts} attempts: {expected_path}"
    )


# ---------------------------------------------------------------------------
# Orchestration helper
# ---------------------------------------------------------------------------

def get_gem_log_sheet_for_wo(wo_number):
    """Full pipeline: navigate to a WO, open its GEM log document via
    Documents List, wait for it to be written locally, open it, and
    return the (workbook, sheet) ready for excel_utils' read functions.

    Caller is responsible for closing the returned workbook when done.
    """
    session = get_sap_session()
    navigate_to_wo_directly(session, wo_number)
    doc_number = open_gem_log_via_documents_list(session, wo_number)
    file_path = wait_for_temp_file(doc_number)

    excel_app = win32com.client.Dispatch("Excel.Application")
    workbook = excel_app.Workbooks.Open(file_path)
    sheet = workbook.Sheets(config.SOURCE_SHEET_NAME)

    return workbook, sheet