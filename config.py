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

SOURCE_COLUMN_SEAL_STRENGTH = "Sample Seal Strength Column (lbf)"


# ---------------------------------------------------------------------------
# Excel — destination (FF1 5k Lysing Tensile data workbook)
# ---------------------------------------------------------------------------

_LYSING_WORKBOOK_PATH_REAL = r"\\obsvr07\Operations\Manufacturing Engineering\19_VPATEL\01_FF1 5K Lysing Tensile Data\FF1 5K Lysing Tensile Data.xlsx"
_LYSING_WORKBOOK_PATH_TEST = r"\\obsvr07\Operations\Manufacturing Engineering\09_INTERNS & CONTRACTORS\Arion Stern\FF1 5K Lysing Tensile Data_copy0.xlsx"

LYSING_WORKBOOK_PATH = _LYSING_WORKBOOK_PATH_TEST if USE_TEST_PATHS else _LYSING_WORKBOOK_PATH_REAL

# Columns where new lot/WO numbers get written (step 4.1)
COL_LOT_NUMBER = "E"
COL_WO_NUMBER = "F"

# Destination sheets/regions for the two paste types (step 4.5.2)
DEST_BOXPLOT_SHEET = "Tensile Data (Boxplot)"
DEST_CONTROL_CHART_SHEET = "Tensile Data (Control Chart)"

# Fields to fill on the FF1 Lysing sheet (step 4.6)
FF1_SHEET_FIELDS = ["Fill Date", "Filler", "Part Number"]

# Duplicate-lot handling (Notes): if 2 lots land on the same day, the
# second one gets a "_1" suffix
DUPLICATE_LOT_SUFFIX = "_1"


# ---------------------------------------------------------------------------
# Minitab
# ---------------------------------------------------------------------------

_MINITAB_PROJECT_PATH_REAL = r"\\obsvr07\Operations\Manufacturing Engineering\19_VPATEL\01_FF1 5K Lysing Tensile Data\FF1 5K Lysing Tensile Data.mpx"
_MINITAB_PROJECT_PATH_TEST = r"\\obsvr07\Operations\Manufacturing Engineering\09_INTERNS & CONTRACTORS\Arion Stern\FF1 5K Lysing Tensile Data_copy0.mpx"

MINITAB_PROJECT_PATH = _MINITAB_PROJECT_PATH_TEST if USE_TEST_PATHS else _MINITAB_PROJECT_PATH_REAL

# Column range used for the Xbar subgroup chart (step 5.1.1)
MINITAB_BAG_COLUMN_RANGE = ("bag1", "bag38")

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