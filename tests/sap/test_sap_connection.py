"""
tests/sap/test_sap_connection.py

Exploratory script — not an automated test suite (no assertions).
Confirms SAP GUI Scripting is enabled and reachable before building
sap_utils.py for real. This is the highest-risk unknown in the whole
project: many companies disable GUI Scripting by policy, so this needs
to succeed before anything else in tests/sap/ is worth writing.

Requires: SAP GUI already open and logged in manually (step 2 in the
original instructions) before running this script.
"""

import sys
import win32com.client

sys.path.append("../..")  # so config.py is importable from tests/sap/
import config


def main():
    # Step 1: can we even reach the scripting engine?
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        print("FAILED to get SAPGUI object. Scripting may not be enabled,")
        print("or SAP GUI isn't running. Check Options > Accessibility &")
        print("Scripting > Scripting in the SAP GUI client, and confirm")
        print("with IT/Basis whether server-side scripting is allowed.")
        print(f"Error: {e}")
        return

    print("Connected to SAPGUI object.")

    application = sap_gui_auto.GetScriptingEngine
    if application is None:
        print("FAILED: GetScriptingEngine returned None.")
        return

    # Step 2: is there an active connection/session to work with?
    if application.Children.Count == 0:
        print("FAILED: No active SAP connections found. Log into SAP")
        print("manually first, then re-run this script.")
        return

    connection = application.Children(0)
    if connection.Children.Count == 0:
        print("FAILED: Connection found but no active session.")
        return

    session = connection.Children(0)
    print(f"Connected to session: {session.Info.SystemName}, "
          f"client {session.Info.Client}, user {session.Info.User}")

    # Step 3: can we navigate at all? (Doesn't run ZPP_WI yet — just
    # confirms we can read/write the transaction field.)
    try:
        current_transaction = session.Info.Transaction
        print(f"Current transaction: {current_transaction}")
    except Exception as e:
        print(f"FAILED to read current transaction: {e}")
        return

    print("\nBasic scripting connection confirmed.")
    print("Next step: try navigating to ZPP_WI and reading the material")
    print(f"filter field, using config.MATERIAL_NUMBER = "
          f"{config.MATERIAL_NUMBER!r}")


if __name__ == "__main__":
    main()