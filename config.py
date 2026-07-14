"""
config.py
Central place for all file paths, settings, and constants used by the
FF1 5k Lysing pull automation. Change values here, not inline elsewhere.
"""

# ---------------------------------------------------------------------------
# General
# ---------------------------------------------------------------------------

# Set True while developing/testing against copies, False for the real
# production files. Only one path set is active at a time — no relying
# on assignment order or manual commenting-out.
USE_TEST_PATHS = True

VERBOSE = True


# ---------------------------------------------------------------------------
# SAP
# ---------------------------------------------------------------------------

# Material number to filter on in ZPP_WI (step 3.1)
MATERIAL_NUMBER = "000470088"

# "No restriction" checkbox behavior — kept as a named flag in case SAP
# scripting needs to toggle it explicitly rather than relying on defaults
NO_RESTRICTION = True

# Text used to identify the correct routing instruction (step 4.3.1)
ROUTING_STEP_ID = "0030/000006"
ROUTING_DESCRIPTION = "In-Process GEM Logs ( 0030 - 0060 )"

# Fallback search text if the routing step isn't found and we have to
# search Doc List instead (step 4.4)
DOC_LIST_SEARCH_TEXT = "gem log fill lot"


# ---------------------------------------------------------------------------
# Excel — source (GEM log workbook)
# ---------------------------------------------------------------------------

SOURCE_SHEET_NAME = "SOPGEM-45-2"

# Which fill lines are valid for this pull (step 4.5.1)
VALID_FILL_LINES = ["FF1", "FF2"]

# NOTE: the real sheet's column header has a typo — "Sample Seal
# Stength (lbf)" (missing "r", no "Column") — so it's referenced by
# letter below rather than matched by text.
SOURCE_COLUMN_SEAL_STRENGTH = "Sample Seal Strength Column (lbf)"  # display name only, don't match on this

# Real cell addresses confirmed against SOPGEM-45-2 (see
# tests/sap/test_locate_fill_line_and_data.py for how these were found)
SOURCE_PRODUCT_NAME_CELL = "C3"       # e.g. "5K LYSING"
SOURCE_WO_NUMBER_CELL = "G3"
SOURCE_FILL_LINE_CELL = "G6"          # header-level fill line, e.g. "FF1"

# Tensile Tester Use Log table layout
SOURCE_TABLE_HEADER_ROW = 11
SOURCE_TABLE_DATA_START_ROW = 12
SOURCE_TABLE_COLUMNS = {
    "fill_line": "B",
    "date": "C",
    "product_name": "D",
    "lot_number": "E",
    "time": "F",
    "seal_strength": "G",
    "operator_initials": "H",
}


# ---------------------------------------------------------------------------
# Excel — destination (FF1 5k Lysing Tensile data workbook)
# ---------------------------------------------------------------------------

_LYSING_WORKBOOK_PATH_REAL = r"\\obsvr07\Operations\Manufacturing Engineering\19_VPATEL\01_FF1 5K Lysing Tensile Data\FF1 5K Lysing Tensile Data.xlsx"
_LYSING_WORKBOOK_PATH_TEST = r"\\obsvr07\Operations\Manufacturing Engineering\09_INTERNS & CONTRACTORS\Arion Stern\FF1 5K Lysing Tensile Data_copy0.xlsx"

LYSING_WORKBOOK_PATH = _LYSING_WORKBOOK_PATH_TEST if USE_TEST_PATHS else _LYSING_WORKBOOK_PATH_REAL

# Real sheet name confirmed (was previously unnamed/assumed)
DEST_WO_SHEET_NAME = "WO Data"

# Columns where new lot/WO numbers get written (step 4.1) — confirmed
# against real headers "Fill Lot #" (E) and "Fill WO #" (F)
COL_LOT_NUMBER = "E"
COL_WO_NUMBER = "F"

# Destination sheets/regions for the two paste types (step 4.5.2)
DEST_BOXPLOT_SHEET = "Tensile Data (Boxplot)"
DEST_CONTROL_CHART_SHEET = "Tensile Data (Control Chart)"

# Fields to fill on the WO Data sheet (step 4.6) — confirmed real
# header text. Product Name (col D) also exists but wasn't in the
# original instructions; add here if it needs to be filled too.
FF1_SHEET_FIELDS = ["Fill Date", "Filler", "Part Number"]

# Boxplot sheet layout: column A holds row labels ("Fill Lot #",
# "Fill WO #", "Fill Date" in rows 1-3); each lot gets its own column
# to the right, bag values going down starting row 4. New lot's data
# goes in the next empty column (first column with a blank row-1 cell).
DEST_BOXPLOT_HEADER_ROWS = {"lot": 1, "wo": 2, "fill_date": 3}
DEST_BOXPLOT_DATA_START_ROW = 4

# Control Chart sheet layout: header row 1 = "Lot No.", "WO No.",
# "Fill Date", "Bag 1", "Bag 2", ... ; each lot gets its own ROW,
# appended at the bottom (next empty row). Confirms this is a genuine
# transpose of the Boxplot sheet's layout (columns become rows).
DEST_CONTROL_CHART_HEADER_ROW = 1
DEST_CONTROL_CHART_DATA_START_COL = "D"  # Bag 1 starts here (A-C are Lot/WO/Date)

# NOTE: on this destination sheet, lots with fewer bag readings than
# the widest existing lot just get BLANK cells for the unused columns
# — not "N/A", not "*". The asterisk-padding rule only applies later,
# when pasting into Minitab itself (step 4.7.1), not here.

# KNOWN GAP (found while inspecting real data, not yet resolved):
# lots 260610C and 260610D exist in WO Data (as of this writing) but
# are missing from both Boxplot and Control Chart sheets — a real
# backlog case worth testing the eventual "catch up missed lots"
# logic against.

# Duplicate-lot handling (Notes): if 2 lots land on the same day, the
# second one gets a "_1" suffix
DUPLICATE_LOT_SUFFIX = "_1"


# ---------------------------------------------------------------------------
# Minitab
# ---------------------------------------------------------------------------

_MINITAB_PROJECT_PATH_REAL = r"\\obsvr07\Operations\Manufacturing Engineering\19_VPATEL\01_FF1 5K Lysing Tensile Data\FF1 5K Lysing Tensile Data.mpx"
_MINITAB_PROJECT_PATH_TEST = r"\\obsvr07\Operations\Manufacturing Engineering\09_INTERNS & CONTRACTORS\Arion Stern\FF1 5K Lysing Tensile Data_copy0.mpx"

MINITAB_PROJECT_PATH = _MINITAB_PROJECT_PATH_TEST if USE_TEST_PATHS else _MINITAB_PROJECT_PATH_REAL

# NOTE: bag count is NOT fixed per lot — confirmed against real data,
# one lot had 30 real readings, not 38. "bag1-bag38" in the original
# instructions was just whichever lot happened to be widest at the
# time, not a constant. minitab_utils.py needs to determine the actual
# column count per run (read until real data ends, per
# SOURCE_TABLE_DATA_START_ROW logic) and pad every column in the
# destination sheet to match whatever the longest one currently is —
# this is exactly why MINITAB_PAD_CHARACTER exists.
MINITAB_MAX_BAG_COLUMNS_OBSERVED = 38  # highest seen so far, not a hard limit

# Chart titles/axis labels — kept together so both charts stay in sync
# if wording ever changes
MINITAB_XBAR_CHART = {
    "title": "2026 FF1 5K Lysing (Variable Subgroup Size)",
    "x_axis": "Fill Date",
    "y_axis": "Seal Strength (lbf)",
}

MINITAB_BOXPLOT_CHART = {
    "title": "2026 FF1 5K Lysing Boxplot",
    "x_axis": "Fill Date",
    "y_axis": "Seal Strength (lbf)",
}

# Rule for step 4.7.1: Minitab requires all rows/columns to be equal
# length, so any row/column shorter than the longest one gets "*"
# appended (as a placeholder, not a real value) until lengths match.
# The padding character itself lives here; the actual pad-to-longest
# logic belongs in minitab_utils.py, not config.py.
MINITAB_PAD_CHARACTER = "*"


# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------

# Text used to find the correct email thread to reply-all on (step 7)
VIRAL_EMAIL_SUBJECT_SEARCH = "FF1 5K Lysing Bag Seal Strength Trending"