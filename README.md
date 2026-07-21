# FF1 5k Lysing Tensile Data Pull

Automates the weekly FF1 5K Lysing bag seal-strength tensile data process: it
finds new/ready lots directly in SAP (ZPP_WI), opens each lot's GEM log
document, pulls the real seal-strength readings and fill date, writes the
results into the tracking Excel workbook and a Minitab project, regenerates
the Boxplot and Xbar charts, and drafts the weekly Outlook summary email —
all from one script (`Run_Lysing_Pull.py`).

For a deeper explanation of how the code works internally (function
reference, process flow, SAP/Excel/Minitab COM gotchas, tests folder, known
issues, and areas for improvement), see **[DOCS.md](./DOCS.md)**.

## What it does, in order

1. Opens the destination Excel workbook and reads the last known lot from
   `WO Data` column E.
2. Navigates SAP's `ZPP_WI` transaction, scans the GR Quantity table backward
   from its most recent entries, and stops as soon as it reaches the last
   known lot — this is how new lots are found without scanning the entire
   multi-year history table.
3. Filters those new lots down to ones that are actually ready (GR Quantity
   != 0).
4. For each ready lot:
   - Opens its GEM log document in SAP via the Documents List tab (using a
     derived document number first, falling back to a keyword search),
     verifying the opened document has real content before trusting it.
   - Reads the fill line, product name, part number, WO number, and real
     seal-strength readings from the `SOPGEM-45-2` sheet.
   - Reads the real fill date from the `GEM Fill Logs Header Page` sheet —
     not derived from the lot code text, since that can be wrong for a lot
     that fills overnight past midnight.
   - Flags the lot for a human double-check if either its readings span more
     than one calendar date (overnight fill), the lot-code-implied date
     disagrees with the real fill date, or the lot name itself doesn't match
     the expected format.
   - Writes the new lot into all three destination Excel sheets (`WO Data`,
     `Tensile Data (Boxplot)`, `Tensile Data (Control Chart)`), skipping the
     `WO Data` append if the lot's already there.
   - Writes the new lot into both Minitab worksheets, skipping entirely if
     the lot's already in the Control Chart worksheet (duplicate
     prevention).
5. Regenerates the Boxplot and Xbar charts in Minitab from scratch (writing
   to chart-linked data resets Minitab's custom axis labels, so both charts
   are fully recreated every run rather than relying on auto-update).
6. Exports both charts as PNGs and drafts a Reply-All email on the existing
   weekly email thread, with the charts attached inline and a summary line
   describing the new datapoint(s).

Both destination files (Excel and the Minitab project) are left open and
**not saved** at the end, and the email draft is displayed for manual review
— nothing is saved or sent automatically.

## Requirements

- Windows, with **SAP GUI** (with Scripting enabled and logged in), **Minitab**
  (with the destination `.mpx` project), and **classic Outlook (desktop)**
  installed and signed in — all three are automated via COM, so they need to
  actually be running/available on the machine executing the script.
- **Microsoft Excel** installed (the destination workbook and the
  SAP-hosted source workbook are both driven directly via COM).
- Network access to the destination Excel workbook and Minitab project paths
  referenced in `config.py`.
- Python 3.x with:
  - `pywin32` (`win32com.client`)

Install the Python dependency with:

```bash
pip install pywin32
```

> There's no `requirements.txt` in the project yet, and `pywin32` is the
> only third-party dependency — everything else (`os`, `sys`, `time`,
> `math`, `datetime`) is standard library.

## Setup

1. Clone/copy this project onto a Windows machine that has SAP GUI, Minitab,
   and Outlook installed, with network access to the paths below.
2. Open `config.py` and confirm the paths are correct for your environment
   (destination Excel workbook path, Minitab project path, material number,
   email subject search text).
3. Set `USE_TEST_PATHS` in `config.py`:
   - `True` — points at the personal test copies of the destination
     workbook/project. The script refuses to run against production
     (`main()` exits immediately) unless this is deliberately set `False`.
   - `False` — points at the real production files. Only set this once
     you're ready to run for real.

## How to run

### 1. Navigate to the project folder

The project lives on a network share, and the full path contains both spaces
and an `&` (`09_INTERNS & CONTRACTORS`) — both need the path to be quoted, or
the terminal will misinterpret them.

**PowerShell** (recommended — PowerShell can `cd` straight into a UNC path):

```powershell
cd "\\obsvr07\Operations\Manufacturing Engineering\09_INTERNS & CONTRACTORS\Arion Stern\FF1-5k-Lysing-Data-Pull"
```

Wrapping the *entire* path in double quotes is what makes this work — in
PowerShell, `&` is a reserved "call operator" character, so it's only
treated as plain text when it's inside quotes.

**Command Prompt (cmd.exe)** — unlike PowerShell, `cmd.exe` cannot set its
current directory to a UNC path at all (`\\server\share\...`), even quoted.
If you need to use cmd instead of PowerShell, map a drive letter to the
share first, then `cd` into that:

```cmd
net use Z: "\\obsvr07\Operations\Manufacturing Engineering"
cd /d "Z:\09_INTERNS & CONTRACTORS\Arion Stern\FF1-5k-Lysing-Data-Pull"
```

`net use` only needs to be run once per login session — after that, `Z:`
behaves like a normal local drive letter.

### 2. Make sure SAP GUI is open and logged in

Unlike the destination Excel workbook and Minitab project (which the script
opens itself via COM), SAP GUI needs to already be running and logged in —
`sap_utils.get_sap_session()` attaches to the existing session rather than
launching SAP itself.

### 3. Run the script

```bash
python Run_Lysing_Pull.py
```

The script is NOT interactive — it runs the full pipeline (auto-detect ->
SAP read -> Excel write -> Minitab write -> chart regeneration -> email
draft) end to end and prints progress as it goes, including any overnight-
fill or lot-name-format flags worth a human double-check. Review the printed
`SUMMARY` section, the flagged lots (if any), the destination files, and the
email draft before saving or sending anything.

## Project structure

```
FF1-5k-Lysing-Data-Pull/
├── Run_Lysing_Pull.py      # Main script — orchestrates the whole weekly process
├── config.py                # All file paths, settings, and constants
├── sap_utils.py              # SAP GUI Scripting: navigation, document lookup, new-lot detection
├── excel_utils.py             # Source read (SAP-hosted workbook) + destination Excel read/write
├── minitab_utils.py            # Minitab COM automation (write, duplicate-prevention, charts, export)
├── outlook_utils.py             # Outlook COM automation (find email, build reply draft)
├── chart_exports/            # PNG exports of the Boxplot/Xbar charts from the most recent run
├── sap/                      # Raw SAP GUI Scripting recordings (reference material used while
│                              #   reverse-engineering SAP navigation — see DOCS.md)
└── tests/                    # Manual/exploratory scripts AND real automated unit tests — mixed,
                               #   see DOCS.md's Tests Folder section for which is which
    ├── sap/
    ├── excel/
    ├── minitab/
    ├── outlook/
    └── integration/
```

## Notes

- The `tests/` folder is a mix of exploratory/manual scripts (many require
  SAP/Minitab/Outlook open and manual confirmation) and genuine automated
  unit tests with no COM dependency (the overnight-flag and email-phrasing
  logic tests) — see [DOCS.md](./DOCS.md#9-tests-folder) before relying on
  any of them as regression tests.
- The `sap/` folder holds raw SAP GUI Scripting recordings used as reference
  while figuring out real control IDs and navigation paths — not part of the
  running pipeline.