# FF1 5k Lysing Tensile Data Pull — Developer Documentation


This is the deep-dive companion to [README.md](./README.md). It covers how the code is put together internally: glossary and data structure reference, process flow, a function-by-function walkthrough, general reusable guides to automating SAP GUI Scripting and Minitab via COM, the tests folder, known issues, and areas for improvement.

## 1. Purpose & Scope

`Run_Lysing_Pull.py` automates the weekly FF1 5K Lysing bag tensile seal-strength data process. It scans SAP directly (no network lot folders involved, unlike the sister Minitab Data Pull project) for new, ready lots of material 000470088, opens each lot's GEM log document, reads the real seal-strength readings and fill date, writes the results into a tracking Excel workbook and a Minitab project, regenerates both charts, exports them, and drafts the weekly summary email in Outlook.

This document exists to explain not just what each function does, but how it does it — including the reasoning behind non-obvious choices (why the fill date isn't read from the lot code text, why SAP navigation goes through the Documents List tab instead of the routing tree, why every chart is regenerated from scratch each run instead of relying on Minitab's native auto-update). Much of this project was built by empirically discovering how SAP GUI Scripting and Minitab's COM API actually behave — several early assumptions turned out to be wrong, and the corrections are documented here so they don't get re-discovered (or re-broken) later.

## 2. Glossary

| **Term** | **Meaning** |
| --- | --- |
| Lot | A single production batch of FF1 5K Lysing bags, identified by a name in the format YYMMDD + letter(s) (e.g. `260630C`, or `260630AA` for a second same-day lot). The date prefix means lot names sort chronologically as plain strings. |
| WO (Work Order) | The SAP work order number tied to a lot. Used to navigate directly to a lot in `ZPP_WI` and to derive its GEM log document number. |
| GEM log document | A per-lot Excel workbook embedded in SAP as an "Office Integration" document, reached via the Documents List tab. Contains the lot's header metadata, the Tensile Tester Use Log table (seal-strength readings), and the real fill date. |
| GR Quantity | "Goods Receipt Quantity" — a column in `ZPP_WI`'s results table. A lot is considered "ready" once this is non-zero. |
| `SOPGEM-45-2` | The sheet inside the GEM log document holding the lot's header fields (product name, part number, WO number, fill line) and the Tensile Tester Use Log table (seal-strength readings, per-bag dates/times). |
| `GEM Fill Logs Header Page` | A *different* sheet in the same GEM log document, holding the real, human-recorded fill date (cell `F16`) — not derived from the lot code text, since that can be wrong for an overnight fill. |
| Overnight fill | A fill that runs past midnight, so the lot's readings can genuinely span two calendar dates, and the lot-code-implied date can legitimately disagree with the real fill date. Detected directly from the row-level Date column, not inferred from time-of-day. |
| `_1` suffix | Minitab Boxplot column naming convention for a second lot filling on the same calendar day (e.g. `6/30/2026` and `6/30/2026_1`). |
| `WO Data` sheet | Destination Excel sheet — one row per lot: Fill Date, Filler (actually the fill line, FF1/FF2), Part Number, Product Name, Fill Lot #, Fill WO #. |
| `Tensile Data (Boxplot)` sheet/worksheet | Destination Excel sheet and Minitab worksheet — each lot is its own **column**, seal values going down. |
| `Tensile Data (Control Chart)` sheet/worksheet | Destination Excel sheet and Minitab worksheet — each lot is its own **row** (the transpose of the Boxplot layout), seal values going across. |
| `USE_TEST_PATHS` | `config.py` flag. `True` points at personal test copies of the destination workbook/project; the pipeline refuses to run against production unless this is deliberately set `False`. |

## 3. Data Structure Reference

### 3.1 Source — GEM log document

**`SOPGEM-45-2` sheet:**

| Cell/Range | Meaning |
| --- | --- |
| `C3` | Product Name |
| `C4` | Part Number |
| `G3` | WO Number |
| `G6` | Fill Line (`FF1`/`FF2`) |
| Row 11 | Tensile Tester Use Log table header |
| Rows 12+ | Table data. Columns: B=Fill Line, C=Date, D=Product Name, E=Lot#, F=Time, G=Seal Strength (lbf), H=Operator Initials |

This is a **fixed 138-row template** — unused rows contain literal text `"N/A"`, not blanks, and real data can be interspersed with `N/A` rows (variable subgroup size per lot, historically up to ~38 bags but not a hard limit).

**`GEM Fill Logs Header Page` sheet:**

| Cell | Meaning |
| --- | --- |
| `F16` | The real fill date. `excel_utils.read_fill_date()` raises a clear `ValueError` if this is blank/invalid rather than silently falling back to a lot-code-derived guess. |

### 3.2 Destination — Excel workbook (`FF1 5K Lysing Tensile Data.xlsx`)

| Sheet | Layout |
| --- | --- |
| `WO Data` | One row per lot, appended at the bottom. A=Fill Date, B=Filler (fill line), C=Part Number, D=Product Name, E=Fill Lot #, F=Fill WO #. |
| `Tensile Data (Boxplot)` | Column A = row labels (row 1=Fill Lot#, row 2=Fill WO#, row 3=Fill Date); each lot is its own column to the right, bag values going down from row 4. New lot = next open column. |
| `Tensile Data (Control Chart)` | Row 1 headers = Lot No., WO No., Fill Date, Bag 1...Bag N; each lot is its own row, appended at the bottom. Bag columns beyond a lot's real reading count are left **blank** (not `"N/A"`, not `"*"` — that padding rule only applies in Minitab, not here). |

### 3.3 Destination — Minitab project

| Worksheet | Layout |
| --- | --- |
| `Tensile Data (Boxplot)` | Each lot is a column, and the column **name** is the fill date as a string (e.g. `'6/10/2026'`, with a `_1` suffix for a second same-day lot). Fixed 38 rows; unused cells are Minitab's native missing-value marker. |
| `Tensile Data (Control Chart)` | C1=Lot No. (Text), C2=WO No. (Numeric), C3=Fill Date (Date/Time), C4–C41=Bag 1...Bag 38. |

## 4. Process Flow, In Depth

### 4.1 `main()` — step by step

#### Step 1 — Determine the last known lot

Opens the destination Excel workbook via COM and calls `excel_utils.get_last_known_lot()`, which uses `End(xlUp)` from the bottom of `WO Data` column E rather than looping — fast even once the sheet has hundreds of rows of history.

#### Step 2 — Find new, ready lots in SAP

`sap_utils.get_sap_session()` attaches to the already-open SAP GUI session (SAP is never launched by the script itself). `sap_utils.find_new_lots()` navigates to `ZPP_WI`, filters by material number with "no restriction" checked, and scans the GR Quantity results table **backward from the bottom** (jump to `scrollbar.Maximum`, then scroll up in pages, reading rows bottom-to-top within each page), stopping as soon as `last_known_lot` is encountered. `filter_ready_lots()` then keeps only lots where GR Quantity != 0.

#### Step 3 — Process each ready lot (`process_one_lot()`)

For each lot:

1. **SAP read.** `sap_utils.get_gem_log_sheet_for_wo()` navigates directly to the WO, builds a ranked list of candidate GEM log documents (derived document number first, keyword-score fallback second), and opens each in order until one has real content — not just a matching document number.
2. **Excel read.** `excel_utils.read_fill_line()`, `read_lot_metadata()`, and `read_seal_strength_values()` pull the header fields and real (non-`N/A`) seal-strength readings from `SOPGEM-45-2`.
3. **Real fill date.** `excel_utils.read_fill_date()` reads `GEM Fill Logs Header Page!F16` — deliberately *not* derived from the lot code text, since that can be wrong for an overnight fill.
4. **Overnight/format flags.** `excel_utils.detect_overnight_fill()` checks whether the lot's own readings span more than one calendar date. Combined (OR logic) with a lot-code-vs-header-page date disagreement (via `sap_utils.parse_lot_date()`) through `excel_utils.build_overnight_flag()`. A lot name that doesn't match the expected format gets its own separate, immediate flag via `excel_utils.build_lot_format_flag()`. Both are printed immediately and collected for the end-of-run summary — see section 6.3.
5. **Source workbook closes** (`SaveChanges=False`) once all reads are done.
6. **Excel write.** If the lot isn't already in `WO Data`, `excel_utils.write_new_lot_to_all_sheets()` appends it to all three destination sheets; otherwise only the Boxplot/Control Chart sheets are written (covers a lot that's partially backfilled).
7. **Minitab write.** `minitab_utils.lot_exists_in_control_chart()` checks for a duplicate first; if none, `minitab_utils.write_new_lot_to_minitab()` writes both worksheets, with automatic `_1` suffix detection for a same-day duplicate.

#### Step 4 — Regenerate charts, export, email

After all lots are processed: `minitab_utils.regenerate_boxplot_chart()` and `regenerate_xbar_chart()` fully rebuild both charts (see section 7 for why this is necessary every run, not just on first creation). `minitab_utils.export_boxplot_and_xbar()` saves both as PNGs. `outlook_utils.send_update_reply()` finds the latest sent email in the configured thread, builds a reply-all draft with both charts inline and a summary line (see section 6.4), and **displays** it for manual review — never sends automatically.

#### Step 5 — Leave everything open

Both the destination Excel workbook and the Minitab project are left open and unsaved at the end, for manual review before saving.

## 5. Function Reference

### 5.1 `sap_utils.py` — SAP GUI Scripting

| Function | Purpose |
| --- | --- |
| `get_sap_session()` | Attach to the active SAP GUI session. Requires SAP GUI already open and logged in. |
| `navigate_to_wo_directly(session, wo_number)` | Navigate straight to a known WO via `ZPP_WI`, using `/n` to force a full reset. |
| `derive_gem_log_document_number(wo_number)` | Derive the GEM log document number from the WO number: `"0000" + WO + "000006" + "0030" + "-01"`. |
| `get_expected_temp_file_path(doc_number)` | Predict the local temp path SAP will write the opened document to (only reliable for auto-generated `0000...`-type documents — see 7.4). |
| `find_document_candidates(session, wo_number)` | Build a ranked candidate list: derived-number match first, then a keyword-score fallback (Description text containing "gem"/"fill"/"lot"/"log"/"in-process", score ≥ 2). |
| `navigate_back_to_order_screen(session)` | Navigate back up a level after rejecting a candidate document. |
| `open_document_row(session, row)` | Click a specific Documents List row open and return the real local file path (polling the temp folder, not predicting the filename — see 7.4). |
| `snapshot_temp_folder()` / `wait_for_new_temp_file(start_time, ...)` | Poll-based detection of "the file SAP just wrote," using an mtime comparison against a `start_time` captured **before** the click. |
| `_document_has_real_content(sheet)` | Reject a candidate document if it's a blank autosave placeholder rather than real content. |
| `get_gem_log_sheet_for_wo(wo_number)` | Full pipeline: navigate, build candidates, open each until one has real content. Returns `(workbook, sheet)`; caller must close the workbook. |
| `is_valid_lot_name(name)` | `YYMMDD` + letter(s) format check — same logic as the sister Minitab Data Pull project's `lot_utils.py`. |
| `parse_lot_date(batch)` | Parse the `YYMMDD` prefix into a real `datetime`. Returns `None` if the format doesn't match. No longer used as the real fill-date source (see 4.1 step 3) — now used as the "lot-code-implied date" side of the overnight/disagreement flag. |
| `_read_visible_gr_table_rows(usr_area)` | Read whatever rows of the GR Quantity table are currently rendered (it's a label-based control — only visible rows exist as real controls). |
| `find_new_lots(session, last_known_lot, page_size)` | Scan the GR Quantity table backward from the bottom, stopping at `last_known_lot`. |
| `filter_ready_lots(new_lots)` | Keep only lots where GR Quantity != 0. |

### 5.2 `excel_utils.py` — source read + destination read/write

| Function | Purpose |
| --- | --- |
| `attach_to_source_excel()` | Attach to the Excel process SAP's Office Integration opened. |
| `read_fill_line(sheet)` / `validate_fill_line(fill_line)` | Read and validate the header-level fill line. |
| `read_lot_metadata(sheet)` | Read product name, part number, WO number. |
| `read_fill_date(workbook)` | Read the real fill date from `GEM Fill Logs Header Page!F16`. Always normalizes to a plain naive `datetime` — Excel COM returns `pywintypes.datetime` (tz-aware, but a subclass of `datetime.datetime`), so `isinstance(x, datetime)` alone isn't enough to trust the value as-is. |
| `read_table_dates(sheet)` | Read the row-level Date column for real (non-`N/A`) rows, mirroring `read_seal_strength_values()`'s stopping logic. |
| `detect_overnight_fill(sheet)` | True if `read_table_dates()` returns more than one distinct calendar date. |
| `build_lot_format_flag(lot_number, lot_code_date)` | Pure function — flags a lot name that doesn't match the expected format (`lot_code_date is None`). |
| `build_overnight_flag(lot_number, overnight_detected, fill_date, lot_code_date)` | Pure function — OR logic between overnight evidence and a lot-code/header-page date disagreement. Deliberately COM-free so it's directly unit-testable (see 9.2). |
| `read_seal_strength_values(sheet)` | Read real (non-placeholder) seal-strength values, looping row-by-row (not `End(xlUp)`) since real/`N/A` rows are interspersed, not contiguous. |
| `find_next_open_wo_data_row(sheet)` / `get_last_known_lot(sheet)` | `End(xlUp)`-based next-row lookup and last-lot read. |
| `append_wo_data_row(...)` | Append one row to `WO Data`. |
| `_column_letter(col_index)` | 1-based column index → Excel letter. |
| `find_next_open_boxplot_column(sheet, header_row)` / `write_boxplot_data(...)` | `End(xlToLeft)`-based next-column lookup and write. |
| `find_next_open_control_chart_row(sheet)` / `write_control_chart_data(...)` | `End(xlUp)`-based next-row lookup and write (transpose of the Boxplot layout). |
| `write_new_lot_to_all_sheets(...)` | Orchestrates all three destination writes. Does not handle the `_1` suffix — that's a Minitab-side concern. |

### 5.3 `minitab_utils.py` — Minitab COM automation

| Function | Purpose |
| --- | --- |
| `open_minitab_project(project_path, visible)` | Open the `.mpx` project via COM. Returns `(mtb, project)`. |
| `get_worksheet(project, sheet_name)` | Fetch a worksheet by name — never rely on `ActiveWorksheet`. |
| `lot_exists_in_control_chart(control_chart_sheet, lot_number)` | Duplicate-prevention check before writing. |
| `format_boxplot_column_name(fill_date, is_duplicate)` / `get_existing_boxplot_column_names(worksheet)` / `get_boxplot_column_names_from_year(worksheet, start_year)` / `determine_unique_boxplot_column_name(worksheet, fill_date)` | Build the correct column name for a new lot, including `_1`/`_2` same-day suffix detection, normalizing both 2-digit and 4-digit year formats found in real existing data. |
| `write_boxplot_column(worksheet, column_name, seal_values, max_rows)` | Write one lot's data as a new Boxplot column, reusing a genuinely empty gap column where possible (see 7.3 for how "genuinely empty" is detected). |
| `find_next_open_control_chart_row(worksheet)` / `write_control_chart_row(...)` | Next-row lookup and write for the Control Chart worksheet, converting types to match each column (text/numeric/date). |
| `regenerate_boxplot_chart(project, boxplot_sheet)` | Fully rebuild the combined Boxplot chart (see 7.5 for why "fully rebuild" instead of relying on auto-update). |
| `export_chart(command, output_path, width, height)` | Export any chart command's graph to a PNG. |
| `regenerate_xbar_chart(project, control_chart_sheet)` | Fully rebuild the Xbar chart via the real `XBARCHART` command. |
| `export_boxplot_and_xbar(project, boxplot_command, output_folder, control_chart_sheet)` | Export both charts, matching the real email's image sizing. |
| `write_new_lot_to_minitab(project, lot_number, wo_number, fill_date, seal_values)` | Orchestrates both worksheet writes, including same-day duplicate detection. |

### 5.4 `outlook_utils.py` — Outlook COM automation

| Function | Purpose |
| --- | --- |
| `open_outlook_app()` / `get_default_folder(folder_number)` | Open/attach to Outlook; fetch a default folder (5=Sent Items, 6=Inbox). |
| `find_latest_sent_email_by_subject(subject_text)` | Find the most recent sent email matching `config.VIRAL_EMAIL_SUBJECT_SEARCH`, to reply-all onto the existing thread. |
| `create_reply_all_draft(message, display)` | Build a Reply-All draft. |
| `attach_inline_image(draft, image_path, content_id)` | Attach an image with a Content-ID so it renders inline via `cid:`. |
| `format_update_summary(new_lot_dates, all_processed_ok)` | Build the summary line. Spells out month names consistently for both single- and multi-lot cases (e.g. `"July 1"`, `"April 22 to June 5"`), and collapses a same-day multi-lot batch to a single date rather than a `"June 30 to June 30"` self-range. Confirmed against two real email examples (see 9.2). |
| `build_email_body(summary_line, xbar_content_id, boxplot_content_id)` | Build the HTML body — Xbar chart first, then Boxplot, matching the real email order. |
| `send_update_reply(new_lot_dates, chart_paths, display)` | Full orchestration. `display=True` by default — never sends automatically. |

### 5.5 `Run_Lysing_Pull.py` — orchestrator

| Function | Purpose |
| --- | --- |
| `lot_exists_in_wo_data(wo_data_sheet, lot_number)` | Row-by-row check for whether a lot already has a `WO Data` row. |
| `process_one_lot(lot_number, wo_number, dest_workbook, mtb_project)` | Runs one lot through the full SAP → Excel → Minitab pipeline (see 4.1 step 3). Returns `(fill_date, overnight_flag, format_flag)`. |
| `main()` | Auto-detects new/ready lots, processes each, prints the run summary (including any overnight/format flags), regenerates and exports charts, and drafts the email. Refuses to run if `config.USE_TEST_PATHS` is `False`. |

## 6. Design Decisions Worth Knowing

### 6.1 Why the fill date isn't derived from the lot code

Early code parsed the fill date directly out of the lot code text (`"260630C"` → 6/30/2026). This is wrong whenever a fill genuinely spans midnight — the lot code encodes an assumption, not a fact. `excel_utils.read_fill_date()` reads the real, human-recorded value from `GEM Fill Logs Header Page!F16` instead, and raises a clear error rather than silently falling back to the lot-code guess if that cell is blank/invalid.

### 6.2 Why Excel COM's date values need normalizing

Excel COM returns `pywintypes.datetime` for date cells, not a plain `datetime.datetime` — but `pywintypes.datetime` **is** a subclass of `datetime.datetime` and carries a `tzinfo` (confirmed: `TimeZoneInfo('GMT Standard Time', True)`). An `isinstance(value, datetime)` check alone will pass for both, silently letting a tz-aware value through to code that assumes naive datetimes (Minitab `SetData`, date comparisons, sorting). Both `read_fill_date()` and `read_table_dates()` always rebuild a plain naive `datetime` from the year/month/day components rather than trusting `isinstance()` to mean "already the right type."

### 6.3 The overnight-fill / lot-format flag

Two independent conditions are checked per lot, combined with **OR logic** (either alone is worth a human double-check, not just both together):

1. **Overnight fill** — the lot's own readings (row-level Date column in the Tensile Tester Use Log) span more than one calendar date. This is *direct* evidence, not inferred from time-of-day values.
2. **Lot-code/header-page disagreement** — the date implied by the lot code (`sap_utils.parse_lot_date()`) doesn't match the real fill date read from the header page.

A lot name that doesn't match the expected `YYMMDD`+letter(s) format gets a **separate, immediate** flag (`build_lot_format_flag()`) rather than being folded into the above — almost every real lot follows the naming convention, so one that doesn't is unusual enough to call out on its own.

The decision logic (`build_overnight_flag()`, `build_lot_format_flag()`) is deliberately pure — no SAP/Excel/Minitab objects — specifically so it can be unit-tested with fabricated inputs covering all combinations, rather than waiting for a real lot to happen to hit each case. See section 9.2.

### 6.4 Email phrasing — real confirmed examples

`format_update_summary()`'s wording is confirmed against two real emails:

- Single lot: `"Control chart and boxplot updates with July 1 datapoint.  This update has 1 new datapoint."`
- Multi-lot: `"Control chart and boxplot updates with April 22 to June 5 datapoints.  This update has 10 new datapoints."`

Month names are spelled out consistently regardless of lot count (an earlier numeric `"7/1"` assumption for the single-lot case turned out to be a paraphrase, not a literal quote). A same-day multi-lot batch (e.g. a lot and its `_1` duplicate) collapses to a single date rather than showing a `"June 30 to June 30"` self-range.

## 7. SAP GUI Scripting — Reusable Findings

### 7.1 Reaching the GEM log document

The GEM log document is **not** reliably reached via the routing tree's grid (`cntlZPPDA_CONT`) — that's a fragile old-style control where row positions are per-lot *and* per-click unstable, and scripted `selectItem`/`pressButton`/`doubleClickItem` silently fail on fully-completed/closed orders (works fine on active orders) even though manual clicks and read-only enumeration work fine. The robust path is the **Documents List tab** (`tabpTAB02`, a fixed tab ID) — a standard ALV grid with a real named column `DOKNR`.

### 7.2 Deriving the document number

The document number is derivable from the WO number: `"0000" + WO + "000006" + "0030" + "-01"`. This can point to a genuinely blank autosave placeholder document — a real different document that happens to share the numbering pattern, not a bug. Always verify real content after opening; don't trust a matched document number alone.

### 7.3 Finding the opened document on disk

SAP writes the opened document to `%LOCALAPPDATA%\SAP\SAP GUI\tmp\`, but the filename is **not always** `{doc_number}.xlsx` — that only holds for auto-generated `0000...`-type documents. Real human-uploaded documents (`ATA1...`-type) use their own original filename. The robust approach is to poll the temp folder for the most-recently-modified `.xlsx` file (excluding `~$` lock files), using a `start_time` captured **before** the click (capturing it after is a real bug to avoid — SAP can finish writing before your code even starts timing).

### 7.4 Navigating between candidates

After opening one document and rejecting it (blank), you must explicitly re-select the Documents List tab and re-fetch the grid fresh before trying another candidate — opening a document navigates one level deeper, so `findById` on the tab strip fails otherwise. Navigate back to the order screen between candidates.

### 7.5 Finding new lots

Navigating to `ZPP_WI` with F4 → "no restriction" → material number lands **directly** on the GR Quantity results table — not a separate screen. This table is an old-style **label-based control** (`lbl[col,row]`), not a named-column grid, and only currently-rendered rows exist as real controls — scrolling and re-reading is required to see more.

The table is **not** sorted chronologically throughout its full history (a chronological recent block, then unrelated old batches, then an older chronological block) — jumping straight to the bottom does not reliably find the most recent lots. What does work: jump to `scrollbar.Maximum`, then scroll **up** in pages, reading rows bottom-to-top within each page, stopping at the known `last_known_lot`. This was empirically validated against real data multiple times — jumping directly to a scroll position vs. reaching it incrementally can produce different results for reasons not fully understood; this is flagged as an open item (section 10).

## 8. Minitab COM Automation — Reusable Findings

### 8.1 Column creation and empty-column detection

`worksheet.Columns.Add()` creates a new column — `Columns.Item(index, True)` does **not** accept a "create if necessary" flag despite looking like it should. An untouched column's `.Name` defaults to its own position label (e.g. `'C60'`), not blank, and `.GetData()` **raises an exception** on a truly empty column rather than returning empty data. Real detection: check if `Name == f"C{i}"` (still default), then try `GetData()` — if it raises, the column is genuinely empty and safe to reuse.

### 8.2 Why every chart is fully regenerated

Any **COM-driven write (`SetData`)** to data a chart is linked to silently resets Minitab's auto-update to default axis labels (e.g., "Sample Mean") — but a manual keystroke edit does not cause this reset. There is no way to "gently" auto-update while preserving custom formatting when writing via COM, so both Boxplot and Xbar charts are fully recreated from scratch every run (`regenerate_boxplot_chart()`, `regenerate_xbar_chart()`).

### 8.3 Chart command syntax — get it from real documentation, don't guess

- `Boxplot 'col1'-'col2'; Overlay; IQRBox; Outlier;` produces one combined "Multiple Y's, Simple" boxplot. Without `Overlay`, separate tiled boxplots result (wrong). An explicit **column list** (not a range) is used instead of `'first'-'last'` to avoid sweeping in stray empty gap columns — **confirmed** that list+Overlay together still produces the correct combined chart, not tiled ones (verified against real generated charts in production use).
- `AxLabel` is a **subcommand**, not standalone — it errors with "Unknown Minitab command" if issued alone; must be nested inside the chart command's own subcommand block. `K=1` is X-axis, `K=2` is Y-axis.
- `Scale`/`Label` subcommands exist but overwrite per-category tick labels, not the overall axis title — don't use these for axis titles.
- The real Xbar chart command is `XBARCHART`, not `"Xbar"` (the Commands list shows a cosmetic auto-generated display name, not the real keyword).
- `TEST 0` (confirmed via real Minitab documentation, the `TEST` subcommand's help text): *"Test 0 will perform no tests and override any set control chart options."* This is exactly the intended behavior (suppress special-cause markers) — no longer an unverified guess.
- Range syntax (`'first'-'last'`) spans every column position in between, including stray empty gap columns — breaks with `"No data in column CN"` even if the real named columns are fine.
- `project.ExecuteCommand()` does **not** raise a Python exception on Minitab-side syntax errors — it prints the error into Minitab's own Session window silently from COM's perspective. Always check whether `project.Commands.Count` actually increased to detect real success/failure.
- `Help <command>.` opens a separate window and does not route text back through the COM API — get real documentation text from Minitab's own Help/Session window manually rather than guessing.

### 8.4 Graph export

`Graph.SaveAs()` accepts `Width=`/`Height=` as **named** parameters (positional args fail) — not visible via `dir()`. Chart image export settled on `width=900, height=450` (source PNG) with `600px` CSS width in the email, tuned down from larger initial guesses after review.

### 8.5 General COM gotchas

Long-running scripts can leave the terminal appearing "stuck" after actually completing successfully — a `pywin32`/COM quirk, not a real hang. Check for a final completion message before assuming something's wrong.

## 9. Tests Folder

The `tests/` folder is a mix of **exploratory/manual scripts** (used to discover real SAP/Minitab behavior — most require the relevant system open, print output for a human to read, and have no assertions) and a smaller set of **genuine automated unit tests** (assertion-based, some with no COM dependency at all). Treat these as historical/diagnostic scripts first and regression tests second — many will not run unattended.

### 9.1 `tests/sap/`

| File | Purpose |
| --- | --- |
| `test_sap_connection.py` | Confirms SAP GUI Scripting is enabled and reachable — the highest-risk unknown in the whole project (many companies disable this by policy). |
| `test_zpp_wi_navigation` / `test_zpp_wi_navigation_2` | Navigate to `ZPP_WI`, filter by material, read back the raw lot grid shape before writing real parsing logic. |
| `test_explore_gr_quantity_table.py` | Inspects the real structure of the GR Quantity results table. |
| `test_decode_gr_table_columns.py` | Maps which column position holds Batch/Order/Item Quantity/GR Quantity/DCI by comparing against a known lot. |
| `test_find_new_lots_backward_scan.py` | Validates the backward-scan-from-bottom approach for finding new lots efficiently. |
| `enumerate_routing_screen_controls.py` | Walks the routing screen's control tree to find the embedded Excel OLE container, since SAP GUI Script Recording can't capture clicks inside native Excel/OLE content. |
| `test_office_integration_access.py` | Compares two approaches for reaching the embedded spreadsheet: grabbing `.OleObject` directly vs. attaching to the real running Excel process via COM. |
| `test_dynamic_routing_lookup.py` | Confirms that navigating by WO number lands on whatever operation is *currently active* for that order, not reliably on operation 0030 — root cause for why a fixed routing-grid selector doesn't work across lots at different lifecycle stages. |
| `test_locate_fill_line_and_data.py` | Widens the initial data dump (columns A–I) to find the real cell addresses for header fields and the Tensile Tester Use Log table. |
| `test_read_gem_log_sheet.py` / `test_read_full_seal_strength_column.py` | Confirm the real shape of `SOPGEM-45-2`'s data, including reading the full Seal Strength column to find where real data ends. |
| `test_open_real_document_directly.py` | Isolated test of the snapshot/diff temp-file detection approach, bypassing the derived-number candidate (which was hitting an unrelated file-lock issue). |
| `sap_utils_test.py` | End-to-end test of `get_gem_log_sheet_for_wo()` — navigate, open, read, and properly close the workbook. |

### 9.2 `tests/excel/`

| File | Purpose |
| --- | --- |
| `test_find_next_open_positions.py` | Read-only check of the three `find_next_open_*` functions before trusting them with real writes. |
| `test_find_new_lots_end_to_end.py` | First real test of `get_last_known_lot()` feeding directly into `sap_utils.find_new_lots()`, instead of a hardcoded last lot. |
| `test_backfill_missed_lot.py` | Backfills one specific missed lot (`260610C`) end-to-end through Excel. |
| `test_overnight_flag_logic.py` | **Genuine automated unit test, no COM dependency.** Covers all combinations of the overnight-fill/disagreement flag (6 cases) plus the lot-name-format flag (2 cases), plus real `sap_utils.parse_lot_date()` checks against multi-letter lot names (`260630AA`, `260630AB`) and genuinely malformed names — 8 additional cases. All pure-function, run instantly. |

### 9.3 `tests/minitab/`

Mostly exploratory scripts used to reverse-engineer Minitab's COM API — see section 8 for the findings these produced. Notable ones:

| File | Purpose |
| --- | --- |
| `test_explore_minitab_project.py` | Confirms real worksheet/column names before writing `minitab_utils.py` logic. |
| `test_explore_column_creation.py` | Finds the real way to create a new column (`Columns.Add()`), after `Item(index, True)` turned out not to work. |
| `test_diagnose_empty_column_detection.py` | Diagnoses why the gap-reuse fix wasn't detecting empty columns — led to the `Name == "C{i}"` + `GetData()`-raises detection logic. |
| `test_duplicate_date_suffix.py` | Writes two lots with the same fill date and confirms the second gets a real `_1` suffix. |
| `test_duplicate_prevention.py` | Direct test of `lot_exists_in_control_chart()` against a known-present lot and a known-absent one (control case). |
| `test_explore_existing_charts.py` / `test_explore_graph_properties.py` / `test_explore_all_graph_related_objects.py` | Explore what the Graph/Command COM objects actually expose — led to the finding that `AxLabel` must be a nested subcommand, not a Graph-object property. |
| `test_get_boxplot_help.py` / `test_get_xbar_help.py` | Use Minitab's own Help command to get real chart syntax instead of guessing. |
| `test_xbarchart_command.py` | Tests the real `XBARCHART` command against the Control Chart worksheet. |
| `test_isolate_axis_label_reset.py` | Isolates that Minitab's auto-update mechanism itself (not other project code) resets custom axis labels on a COM-driven write — the root cause behind fully regenerating charts every run. |
| `test_graph_saveas_dimensions.py` | Confirms `Graph.SaveAs()` accepts named `Width=`/`Height=` parameters. |
| `test_write_functions.py` | Tests `write_boxplot_column()`/`write_control_chart_row()` against a lot with already-known-correct values. |

### 9.4 `tests/outlook/`

| File | Purpose |
| --- | --- |
| `test_email_pipeline_staged.py` | Staged, pause-for-confirmation test of the full email pipeline (find email → export charts → build/display draft) — nothing here had been tested live before this script. |
| `test_format_update_summary.py` | **Genuine automated unit test, no COM dependency.** Confirms `format_update_summary()` against both real confirmed email examples (single-lot and 10-lot multi-lot) plus edge cases: same-day multi-lot collapse, unsorted input, empty list. 5 cases. |

### 9.5 `tests/integration/`

| File | Purpose |
| --- | --- |
| `test_read_fill_date.py` | Isolated test of `read_fill_date()` against a real known lot (`260630C`), run before trusting `Run_Lysing_Pull.py`'s fill-date restructuring. |
| `test_find_new_lots_end_to_end.py` *(duplicate name, under `tests/excel/`)* | See 9.2. |
| `test_end_to_end_260610C.py` | Full pipeline integration test — the first time SAP, Excel, and Minitab all ran together in one continuous pass with real data throughout, using the last unconfirmed backfill lot (`260610C`) as the real test case. |

### 9.6 `sap/` (project root, not `tests/sap/`)

Not test scripts — raw SAP GUI Scripting **recordings** (VBScript-format `.vbs`-style output from SAP's own script recorder) captured while reverse-engineering real control IDs and navigation paths (e.g. `click on gr quant`, `instruction_click_doclist_clickondoc`, `routing_operation_0030_reference`). Kept as reference material, not executed by the pipeline.

## 10. Known Issues

- **`find_new_lots()`'s scrollbar-jump behavior is not fully understood.** Jumping directly to `scrollbar.Maximum` vs. reaching the same position incrementally has been observed to produce *different* results in testing, for reasons not fully explained. The current backward-scan-from-max approach has been empirically validated multiple times, but this is worth re-checking if SAP's table size or behavior ever changes materially.
- **Chart history accumulates in the Minitab project.** Every run creates new Boxplot and Xbar chart commands; nothing deletes the old ones. Deliberately deferred to manual cleanup for now (see section 11).
- **Destination files are never saved by the script.** Deliberate QA caution — worth confirming this stays the long-term intention rather than a lingering TODO.

## 11. Areas for Improvement

- **Automated chart-history cleanup.** A candidate approach: once a run's chart PNG export succeeds, delete the *previous* run's Graph window for that chart type (not the Commands history, which is low-cost to leave alone). This needs real investigation into what Minitab's COM API actually exposes for deleting/closing a Graph object — following the same "don't guess COM syntax" discipline used throughout this project (see section 8.3) — before attempting it.
- **`config.KNOWN_LOTS`** is dead weight now that real SAP-based auto-detection exists (`find_new_lots()`/`filter_ready_lots()`) — safe to delete whenever, just not urgent.
- **No `requirements.txt`** — `pywin32` is the only third-party dependency, but it's not currently pinned anywhere.