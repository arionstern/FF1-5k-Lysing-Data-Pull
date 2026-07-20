"""
Run_Lysing_Pull.py

Main orchestrator for the FF1 5k Lysing pull. Ties together:
  sap_utils    -> find new lots, navigate SAP, open the GEM log document
  excel_utils  -> read source data, write to WO Data/Boxplot/Control Chart
  minitab_utils -> write to Boxplot/Control Chart, regenerate the boxplot chart

Auto-detection of new lots IS wired in here (comparing SAP's GR
Quantity table against the last known lot in WO Data), unlike the
earlier draft of this file.

Leaves both destination files open, unsaved, at the end for manual
review before saving -- matches the caution used throughout testing,
and the pattern from the last project's Run_Chart_Prep.py.
"""

import sys
import win32com.client

import config
import sap_utils
import excel_utils
import minitab_utils
import outlook_utils


def lot_exists_in_wo_data(wo_data_sheet, lot_number):
    """Check whether a lot already has a row in WO Data (determines
    whether to append a new row or skip that step)."""
    row = 2
    while wo_data_sheet.Range(f"{config.COL_LOT_NUMBER}{row}").Value not in (None, ""):
        if str(wo_data_sheet.Range(f"{config.COL_LOT_NUMBER}{row}").Value) == lot_number:
            return True
        row += 1
    return False


def process_one_lot(lot_number, wo_number, fill_date,
                     dest_workbook, mtb_project):
    """Runs one lot through the full pipeline: SAP -> Excel -> Minitab."""
    print(f"\n{'=' * 60}")
    print(f"Processing lot {lot_number} (WO {wo_number})")
    print(f"{'=' * 60}")

    # --- SAP: read source data ---
    print("Reading source data from SAP...")
    source_workbook, source_sheet = sap_utils.get_gem_log_sheet_for_wo(wo_number)

    try:
        fill_line = excel_utils.read_fill_line(source_sheet)
        if not excel_utils.validate_fill_line(fill_line):
            print(f"WARNING: fill line '{fill_line}' not in "
                  f"{config.VALID_FILL_LINES} -- proceeding anyway, but "
                  f"this should be reviewed.")

        metadata = excel_utils.read_lot_metadata(source_sheet)
        seal_values = excel_utils.read_seal_strength_values(source_sheet)
        print(f"  Fill line: {fill_line}, {len(seal_values)} readings")
    finally:
        source_workbook.Close(SaveChanges=False)

    # --- Excel: write to destination ---
    print("Writing to Excel destination...")
    wo_data_sheet = dest_workbook.Sheets(config.DEST_WO_SHEET_NAME)
    boxplot_sheet_xl = dest_workbook.Sheets(config.DEST_BOXPLOT_SHEET)
    control_chart_sheet_xl = dest_workbook.Sheets(config.DEST_CONTROL_CHART_SHEET)

    if lot_exists_in_wo_data(wo_data_sheet, lot_number):
        print(f"  Lot {lot_number} already in WO Data -- skipping append, "
              f"writing Boxplot/Control Chart only.")
        excel_utils.write_boxplot_data(
            boxplot_sheet_xl, lot_number, wo_number, fill_date, seal_values
        )
        excel_utils.write_control_chart_data(
            control_chart_sheet_xl, lot_number, wo_number, fill_date, seal_values
        )
    else:
        # NOTE: "Filler" in WO Data is actually the fill line
        # (FF1/FF2), confirmed against real data -- not a separate
        # metadata field.
        result = excel_utils.write_new_lot_to_all_sheets(
            wo_data_sheet, boxplot_sheet_xl, control_chart_sheet_xl,
            fill_date, fill_line, metadata.get("part_number", ""),
            metadata.get("product_name", ""), lot_number, wo_number, seal_values
        )
        print(f"  Wrote WO Data row {result['wo_data_row']}, "
              f"Boxplot col {result['boxplot_column']}, "
              f"Control Chart row {result['control_chart_row']}")

    # --- Minitab: write to destination ---
    print("Writing to Minitab destination...")
    control_chart_sheet_mtb = minitab_utils.get_worksheet(
        mtb_project, config.DEST_CONTROL_CHART_SHEET
    )
    if minitab_utils.lot_exists_in_control_chart(control_chart_sheet_mtb, lot_number):
        print(f"  Lot {lot_number} already in Minitab Control Chart — "
              f"skipping Minitab write entirely to avoid a duplicate.")
    else:
        minitab_result = minitab_utils.write_new_lot_to_minitab(
            mtb_project, lot_number, wo_number, fill_date, seal_values
        )
        print(f"  Wrote Boxplot column {minitab_result['boxplot_column_name']!r}, "
              f"Control Chart row {minitab_result['control_chart_row']}")

    print(f"Lot {lot_number} complete.")


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False -- this will write to "
              "REAL production files. Stopping as a precaution; set "
              "USE_TEST_PATHS = False deliberately once ready.")
        return

    # --- Auto-detect new lots ---
    print(f"Opening Excel to check last known lot: {config.LYSING_WORKBOOK_PATH}")
    dest_excel = win32com.client.Dispatch("Excel.Application")
    dest_excel.Visible = True
    dest_workbook = dest_excel.Workbooks.Open(config.LYSING_WORKBOOK_PATH)
    wo_data_sheet = dest_workbook.Sheets(config.DEST_WO_SHEET_NAME)

    last_known_lot = excel_utils.get_last_known_lot(wo_data_sheet)
    print(f"Last known lot: {last_known_lot!r}")

    print("Checking SAP for new lots...")
    session = sap_utils.get_sap_session()
    new_lots = sap_utils.find_new_lots(session, last_known_lot)
    ready_lots = sap_utils.filter_ready_lots(new_lots)

    print(f"\nFound {len(new_lots)} new lot(s), {len(ready_lots)} ready "
          f"(GR quantity != 0):")
    for order, data in sorted(new_lots.items()):
        status = "READY" if order in ready_lots else "not ready yet"
        print(f"  {data['batch']} (WO {order}): {status}")

    if not ready_lots:
        print("\nNo ready lots to process. Nothing to do.")
        return

    lots_to_process = []
    for order, data in sorted(ready_lots.items()):
        fill_date = sap_utils.parse_lot_date(data["batch"])
        lots_to_process.append((data["batch"], order, fill_date))

    print(f"\nProcessing {len(lots_to_process)} lot(s)...")

    mtb, mtb_project = minitab_utils.open_minitab_project()

    results = []
    successful_fill_dates = []
    for lot_number, wo_number, fill_date in lots_to_process:
        try:
            process_one_lot(lot_number, wo_number, fill_date,
                             dest_workbook, mtb_project)
            results.append((lot_number, "OK"))
            successful_fill_dates.append(fill_date)
        except Exception as e:
            print(f"FAILED on lot {lot_number}: {e}")
            results.append((lot_number, f"FAILED: {e}"))

    print(f"\n{'=' * 60}")
    print("SUMMARY")
    print(f"{'=' * 60}")
    for lot_number, status in results:
        print(f"  {lot_number}: {status}")

    print("\nRegenerating Boxplot chart in Minitab...")
    import time
    time.sleep(2)  # let Minitab fully register the new column's data
                    # first — running the chart command immediately
                    # after writing hit "No data in column C60"
    boxplot_command = None
    try:
        boxplot_command = minitab_utils.regenerate_boxplot_chart(mtb_project)
        print("Chart regenerated.")
    except Exception as e:
        print(f"FAILED to regenerate chart: {e}")

    if boxplot_command is not None and successful_fill_dates:
        print("\nBuilding reply-all email draft...")
        try:
            import os
            output_folder = os.path.join(os.getcwd(), "chart_exports")
            chart_paths = minitab_utils.export_boxplot_and_xbar(
                mtb_project, boxplot_command, output_folder
            )
            outlook_utils.send_update_reply(
                successful_fill_dates, chart_paths, display=True
            )
            print("Draft created and displayed for review — NOT sent "
                  "automatically. Review and send manually.")
        except Exception as e:
            print(f"FAILED to build email draft: {e}")
    else:
        print("\nSkipping email — no successful lots or chart wasn't "
              "regenerated.")

    print("\nBoth destination files left OPEN, NOT SAVED. Review")
    print("everything before saving.")


if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------
# TODO / status notes
# ---------------------------------------------------------------------------
#
# DONE:
# - Auto-detecting new lots (excel_utils.get_last_known_lot +
#   sap_utils.find_new_lots + filter_ready_lots) -- the original TODO
#   from the first version of this file is now resolved and tested.
# - SAP navigation + document lookup (sap_utils.py), including the
#   Documents List / derived-number / keyword-fallback approach and
#   blank-content rejection
# - Excel read (source) and write (WO Data, Boxplot, Control Chart)
# - Minitab write (Boxplot, Control Chart), including automatic
#   duplicate-date "_1" suffix detection
# - Minitab Boxplot chart regeneration, including title AND axis
#   labels (AxLabel session subcommand, nested inside Boxplot's own
#   subcommand block -- NOT a standalone command)
# - Xbar chart: originally thought to need NO code (auto-updates
#   natively). Reversed after testing confirmed any COM-driven write
#   resets its custom axis labels back to defaults, even though a
#   manual keystroke edit doesn't. Now regenerated fresh each run via
#   regenerate_xbar_chart(), same pattern as Boxplot, using the real
#   XBARCHART command (RSUB/STAMP/EXCLUDE/AxLabel/Title). One
#   unverified piece: "TEST 0" to suppress special-cause markers --
#   not found in documentation, just happened to work.
# - The "genuinely new lot" append branch (append_wo_data_row via
#   write_new_lot_to_all_sheets) HAS been tested successfully multiple
#   times now (e.g. 260630C, after manually removing it from the test
#   sheet to simulate a new lot) — including through Excel, Minitab,
#   chart regeneration, and email draft creation end-to-end.
#
# NOT DONE / NOT YET FULLY VALIDATED:
# - Minitab duplicate-prevention (lot_exists_in_control_chart) added
#   and wired in -- mirrors the existing WO Data check. Untested as
#   of this writing (no real duplicate scenario hit yet in testing).
# - Reply-all email (step 7) -- wired in and tested via
#   tests/outlook/test_email_pipeline_staged.py. Draft is displayed
#   for manual review/send, never auto-sent. Date-range phrasing in
#   the summary line is only confirmed against a small number of real
#   examples -- worth double-checking wording on edge cases (e.g. 2
#   lots, not just 1 or 3).
# - Old boxplot AND xbar chart commands both accumulate in Minitab's
#   history each time this runs -- nothing deletes the previous ones.
#   Worth cleaning up / auto-deleting old chart commands before this
#   runs unattended long-term.
# - This script currently does NOT save either destination file -- that
#   stays a manual step for now, given the QA-sensitive nature of this
#   data.
# - find_new_lots() jumps directly to the scrollbar's Maximum position
#   rather than scrolling there incrementally -- testing showed these
#   can produce DIFFERENT results for reasons not fully understood.
#   Worth re-validating if SAP's table size/behavior ever changes.