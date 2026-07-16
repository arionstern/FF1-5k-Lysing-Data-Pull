"""
tests/minitab/test_diagnose_empty_column_detection.py

The gap-reuse fix in write_boxplot_column silently failed — still
appended past the known-empty C60-C69 gap. This prints exactly what
.Name and .GetData() actually return for those columns, since the
detection logic ("if not candidate.Name") was a guess, not confirmed.

Requires: config.USE_TEST_PATHS = True.
"""

import sys

sys.path.append("../..")
import config
import minitab_utils


def main():
    if not config.USE_TEST_PATHS:
        print("WARNING: USE_TEST_PATHS is False. Stopping.")
        return

    mtb, project = minitab_utils.open_minitab_project()
    boxplot_sheet = minitab_utils.get_worksheet(project, config.DEST_BOXPLOT_SHEET)

    print(f"Total columns: {boxplot_sheet.Columns.Count}\n")

    # Check a range covering the known gap and a known-real column
    for i in range(55, min(72, boxplot_sheet.Columns.Count + 1)):
        col = boxplot_sheet.Columns.Item(i)
        name = col.Name
        try:
            synthesized = col.SynthesizedName
        except Exception as e:
            synthesized = f"(error: {e})"
        try:
            data = list(col.GetData())
            data_sample = data[:3]
            has_any_real_value = any(v is not None for v in data)
        except Exception as e:
            data_sample = f"(error: {e})"
            has_any_real_value = "?"

        print(f"Col {i}: Name={name!r} SynthesizedName={synthesized!r} "
              f"has_data={has_any_real_value} sample={data_sample}")


if __name__ == "__main__":
    main()