"""
tests/outlook/test_format_update_summary.py

Unit test for outlook_utils.format_update_summary() -- pure text
formatting, no Outlook/COM dependency, so this runs instantly.

Checks against the two REAL confirmed examples:
  1. Single lot:  "Control chart and boxplot updates with July 1
     datapoint.  This update has 1 new datapoint."
  2. Multi-lot:   "Control chart and boxplot updates with April 22
     to June 5 datapoints.  This update has 10 new datapoints."

Plus edge cases:
  3. Same-day multi-lot (e.g. a lot + its '_1' duplicate) -- should
     collapse to a single date, not a self-range like "June 30 to
     June 30".
  4. Unsorted input -- should still sort correctly before formatting.
  5. Empty list.
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import outlook_utils


def check(label, dates, expected):
    result = outlook_utils.format_update_summary(dates)
    assert result == expected, (
        f"[{label}] expected:\n  {expected!r}\ngot:\n  {result!r}"
    )
    print(f"PASS [{label}]: {result}")


def main():
    # 1. Real confirmed single-lot example
    check(
        "real single-lot example",
        [datetime(2026, 7, 1)],
        "Control chart and boxplot updates with July 1 datapoint.  "
        "This update has 1 new datapoint.",
    )

    # 2. Real confirmed multi-lot example (10 lots, April 22 - June 5)
    multi_dates = (
        [datetime(2026, 4, 22)]
        + [datetime(2026, 5, d) for d in range(1, 9)]  # 8 filler dates in between
        + [datetime(2026, 6, 5)]
    )
    assert len(multi_dates) == 10, "test setup error: expected 10 dates"
    check(
        "real multi-lot example (10 lots)",
        multi_dates,
        "Control chart and boxplot updates with April 22 to June 5 "
        "datapoints.  This update has 10 new datapoints.",
    )

    # 3. Same-day multi-lot -- should NOT show "June 30 to June 30"
    check(
        "same-day multi-lot (2 lots, 1 date)",
        [datetime(2026, 6, 30), datetime(2026, 6, 30)],
        "Control chart and boxplot updates with June 30 datapoints.  "
        "This update has 2 new datapoints.",
    )

    # 4. Unsorted input -- must still sort before building the range
    check(
        "unsorted input",
        [datetime(2026, 7, 1), datetime(2026, 6, 30)],
        "Control chart and boxplot updates with June 30 to July 1 "
        "datapoints.  This update has 2 new datapoints.",
    )

    # 5. Empty list
    check(
        "empty list",
        [],
        "No new datapoints this update.",
    )

    print("\nAll format_update_summary cases passed.")


if __name__ == "__main__":
    main()