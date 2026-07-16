"""
tests/outlook/test_email_pipeline_staged.py

Staged test for the untested email pipeline (outlook_utils.py +
minitab_utils.export_boxplot_and_xbar). Each stage prints its result
and pauses for confirmation before continuing — nothing here has been
tested live before.

Stage 1: find the real sent email in the chain
Stage 2: export both charts from the currently-open Minitab project
Stage 3: build and DISPLAY (not send) the full reply-all draft

Requires: config.USE_TEST_PATHS = True. Outlook and Minitab both open.
"""

import sys
from datetime import date

sys.path.append("../..")
import config
import minitab_utils
import outlook_utils


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    # --- Stage 1: find the real sent email ---
    print("=== Stage 1: find_latest_sent_email_by_subject ===")
    email = outlook_utils.find_latest_sent_email_by_subject()
    if email is None:
        print(f"FAILED: no email found matching "
              f"{config.VIRAL_EMAIL_SUBJECT_SEARCH!r}")
        return
    print(f"Found: {email.Subject!r}")
    print(f"Sent: {email.SentOn}")
    print(f"To: {email.To}")

    input("\nCheck this is the right email above, then press Enter "
          "to continue to Stage 2 (chart export)...")

    # --- Stage 2: export both charts ---
    print("\n=== Stage 2: export_boxplot_and_xbar ===")
    print("This requires the CURRENT Minitab project already open, "
          "with regenerate_boxplot_chart() already run this session "
          "(so there's a real boxplot command to export).")
    mtb, project = minitab_utils.open_minitab_project()

    try:
        boxplot_command = minitab_utils.regenerate_boxplot_chart(project)
        print(f"Boxplot chart command: {boxplot_command.Name}")
    except Exception as e:
        print(f"FAILED to regenerate boxplot chart: {e}")
        return

    import os
    output_folder = os.path.join(os.getcwd(), "test_chart_exports")
    try:
        chart_paths = minitab_utils.export_boxplot_and_xbar(
            project, boxplot_command, output_folder
        )
        print(f"Exported:")
        print(f"  Xbar: {chart_paths['xbar']}")
        print(f"  Boxplot: {chart_paths['boxplot']}")
    except Exception as e:
        print(f"FAILED to export charts: {e}")
        return

    input(f"\nCheck both PNG files exist and look correct in "
          f"{output_folder!r}, then press Enter to continue to "
          f"Stage 3 (building the draft)...")

    # --- Stage 3: build and display (NOT send) the draft ---
    print("\n=== Stage 3: send_update_reply (display only, not sent) ===")
    test_dates = [date(2026, 7, 1)]  # placeholder — one lot, matching
                                       # today's actual scenario
    try:
        draft = outlook_utils.send_update_reply(
            test_dates, chart_paths, display=True
        )
        print("Draft created and displayed. Review it manually — "
              "nothing has been sent.")
    except Exception as e:
        print(f"FAILED to build draft: {e}")
        return

    print("\nDone. Compare the draft's summary line and chart layout "
          "against the real email examples before trusting this.")


if __name__ == "__main__":
    main()