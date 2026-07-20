"""
tests/excel/test_overnight_flag_logic.py

Unit test for excel_utils.build_overnight_flag() -- the pure decision
logic behind the overnight-fill / lot-code-disagreement flag added to
Run_Lysing_Pull.py. Deliberately has NO dependency on SAP, Excel, or
Minitab COM -- build_overnight_flag() takes plain Python values
(bool, datetime, datetime-or-None) and returns a string-or-None, so
all 4 True/False combinations can be checked directly and instantly,
without a live run ever needing to happen to hit each case.

Covers:
  1. overnight=False, disagreement=False -> no flag
  2. overnight=True,  disagreement=False -> flag, overnight reason only
  3. overnight=False, disagreement=True  -> flag, disagreement reason only
  4. overnight=True,  disagreement=True  -> flag, both reasons
  + edge case: lot_code_date is None (unparseable lot name) -> treated
    as "no disagreement possible", only overnight_detected matters
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import excel_utils

LOT = "260630C"
SAME_DATE = datetime(2026, 6, 30)
DIFFERENT_DATE = datetime(2026, 7, 1)


def check(label, overnight_detected, fill_date, lot_code_date,
          expect_flag, expect_overnight_reason=False,
          expect_disagreement_reason=False):
    result = excel_utils.build_overnight_flag(
        LOT, overnight_detected, fill_date, lot_code_date
    )

    if not expect_flag:
        assert result is None, f"[{label}] expected no flag, got: {result!r}"
        print(f"PASS [{label}]: no flag, as expected")
        return

    assert result is not None, f"[{label}] expected a flag, got None"

    if expect_overnight_reason:
        assert "overnight fill" in result, (
            f"[{label}] expected overnight reason in message: {result!r}"
        )
    else:
        assert "overnight fill" not in result, (
            f"[{label}] did NOT expect overnight reason in message: {result!r}"
        )

    if expect_disagreement_reason:
        assert "disagrees" in result, (
            f"[{label}] expected disagreement reason in message: {result!r}"
        )
    else:
        assert "disagrees" not in result, (
            f"[{label}] did NOT expect disagreement reason in message: {result!r}"
        )

    print(f"PASS [{label}]: {result}")


def main():
    # 1. Neither condition -> no flag
    check(
        "neither",
        overnight_detected=False,
        fill_date=SAME_DATE,
        lot_code_date=SAME_DATE,
        expect_flag=False,
    )

    # 2. Overnight only
    check(
        "overnight only",
        overnight_detected=True,
        fill_date=SAME_DATE,
        lot_code_date=SAME_DATE,
        expect_flag=True,
        expect_overnight_reason=True,
        expect_disagreement_reason=False,
    )

    # 3. Disagreement only
    check(
        "disagreement only",
        overnight_detected=False,
        fill_date=DIFFERENT_DATE,
        lot_code_date=SAME_DATE,
        expect_flag=True,
        expect_overnight_reason=False,
        expect_disagreement_reason=True,
    )

    # 4. Both
    check(
        "both",
        overnight_detected=True,
        fill_date=DIFFERENT_DATE,
        lot_code_date=SAME_DATE,
        expect_flag=True,
        expect_overnight_reason=True,
        expect_disagreement_reason=True,
    )

    # Edge case: lot_code_date is None (unparseable lot name) --
    # disagreement can't be evaluated, only overnight_detected matters
    check(
        "lot_code_date is None, overnight True",
        overnight_detected=True,
        fill_date=SAME_DATE,
        lot_code_date=None,
        expect_flag=True,
        expect_overnight_reason=True,
        expect_disagreement_reason=False,
    )
    check(
        "lot_code_date is None, overnight False",
        overnight_detected=False,
        fill_date=SAME_DATE,
        lot_code_date=None,
        expect_flag=False,
    )

    print("\nAll overnight-flag logic cases passed.")


if __name__ == "__main__":
    main()