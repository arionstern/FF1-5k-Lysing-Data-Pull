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


def check_format_flag():
    # Valid format -> no flag
    result = excel_utils.build_lot_format_flag(LOT, lot_code_date=SAME_DATE)
    assert result is None, f"expected no format flag, got: {result!r}"
    print("PASS [valid format]: no flag, as expected")

    # Unparseable lot name -> flag. Deliberately uses a clearly-fake
    # placeholder name here (NOT '260630C', which IS valid) to avoid
    # implying the real lot is invalid -- this test only checks that
    # build_lot_format_flag() reacts correctly to a forced None,
    # it isn't asserting anything about a real lot number's validity.
    fake_bad_lot = "BADNAME"
    result = excel_utils.build_lot_format_flag(fake_bad_lot, lot_code_date=None)
    assert result is not None, "expected a format flag, got None"
    assert "does not match" in result, f"unexpected message: {result!r}"
    print(f"PASS [invalid format]: {result}")

    print("\nAll lot-format-flag logic cases passed.")


def check_real_lot_name_parsing():
    """Unlike the cases above (which hand-feed a stand-in lot_code_date
    directly to the pure build_* functions), this calls the REAL
    sap_utils.parse_lot_date() against real-shaped lot names -- so it
    actually proves the parser handles multi-letter suffixes like
    '260630AA' correctly, not just that build_lot_format_flag() reacts
    correctly to whatever value it's handed."""
    import sap_utils

    # Standard single-letter format
    result = sap_utils.parse_lot_date("260630C")
    assert result == datetime(2026, 6, 30), f"got {result!r}"
    print("PASS [real parse '260630C']: ", result)

    # Multi-letter suffix (second/third same-day lot) -- should parse
    # identically to the single-letter case, same date
    result = sap_utils.parse_lot_date("260630AA")
    assert result == datetime(2026, 6, 30), f"got {result!r}"
    print("PASS [real parse '260630AA']: ", result)

    result = sap_utils.parse_lot_date("260630AB")
    assert result == datetime(2026, 6, 30), f"got {result!r}"
    print("PASS [real parse '260630AB']: ", result)

    # Genuinely malformed names -> None, and that None should trigger
    # build_lot_format_flag()
    for bad_name in ["260630", "SAMPLE1", "2A0630B", "abcdefg"]:
        result = sap_utils.parse_lot_date(bad_name)
        assert result is None, f"[{bad_name}] expected None, got {result!r}"
        flag = excel_utils.build_lot_format_flag(bad_name, result)
        assert flag is not None, f"[{bad_name}] expected a format flag"
        print(f"PASS [real parse '{bad_name}' -> None -> flagged]: {flag}")

    print("\nAll real lot-name parsing cases passed.")


if __name__ == "__main__":
    main()
    check_format_flag()
    check_real_lot_name_parsing()