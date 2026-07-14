"""
tests/sap/test_office_integration_access.py

Exploratory script — not an automated test suite (no assertions).
Tries two different ways to reach the embedded spreadsheet content on
the routing/documents screen (SubType: OfficeIntegration), to figure
out which one actually works before building excel_utils.py around it:

  A) Grab .OleObject directly off the SAP GuiShell control
  B) Attach to a real, already-running Excel.Application process via
     COM directly (bypassing SAP entirely) — SAP's "Office Integration"
     sometimes docks a genuine Excel.exe window rather than truly
     embedding it, in which case this is much simpler and more robust.

Requires: SAP GUI open, logged in, sitting on the routing/documents
screen with the spreadsheet content actually visible (same state as
the last enumerate_routing_screen_controls.py run).
"""

import sys
import win32com.client

sys.path.append("../..")
import config  # noqa: F401


SHELL_ID = (
    "/app/con[0]/ses[0]/wnd[0]/usr/tabsTABS_0200/tabpTAB16/"
    "ssubSUBSCREEN:SAPLZPPDR:0100/cntlZPPDR_CONT/shellcont/shell"
)


def get_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def try_ole_object(session):
    print("--- Approach A: control.OleObject ---")
    try:
        shell = session.findById(SHELL_ID)
        ole_obj = shell.OleObject
        print(f"SUCCESS: got OleObject: {ole_obj}")
        try:
            print(f"Active sheet name: {ole_obj.ActiveSheet.Name}")
        except Exception as e:
            print(f"(got OleObject, but couldn't read ActiveSheet: {e})")
        return ole_obj
    except Exception as e:
        print(f"FAILED: {e}")
        return None


def try_direct_excel_attach():
    print("\n--- Approach B: attach to running Excel.Application ---")
    try:
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        print(f"SUCCESS: attached to Excel. Open workbooks:")
        for wb in excel_app.Workbooks:
            print(f"  - {wb.Name}")
        return excel_app
    except Exception as e:
        print(f"FAILED: {e}")
        return None


def main():
    session = get_session()
    print(f"Session: {session.Info.SystemName}, "
          f"transaction: {session.Info.Transaction}\n")

    ole_result = try_ole_object(session)
    excel_result = try_direct_excel_attach()

    print("\n=== Summary ===")
    print(f"Approach A (OleObject):        {'worked' if ole_result else 'failed'}")
    print(f"Approach B (direct Excel COM): {'worked' if excel_result else 'failed'}")
    print("\nWhichever worked is what excel_utils.py should build on.")
    print("If both failed, we likely need SAP's Office Integration API")
    print("specifically (different from a generic OLE container) — flag")
    print("this back and we'll look into GuiOfficeIntegration-specific")
    print("methods instead.")


if __name__ == "__main__":
    main()