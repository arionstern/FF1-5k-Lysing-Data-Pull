"""
tests/sap/test_open_real_document_directly.py

Isolated test — skips the derived-number candidate entirely (it's
hitting a separate SAP-side file lock/write conflict, unrelated to
the new-file detection logic) and opens the real keyword-matched
document directly, to confirm the snapshot/diff approach works on
its own before re-integrating into sap_utils.py's candidate loop.

Requires: SAP GUI open, already navigated to WO 10556542's Documents
List tab (or run navigate_to_wo_directly + select the tab first).
"""

import sys
import time

sys.path.append("../..")
import config
import sap_utils


def main():
    session = sap_utils.get_sap_session()
    sap_utils.navigate_to_wo_directly(session, "10556542")

    grid, candidates = sap_utils.find_document_candidates(session, "10556542")

    # Skip candidate 1 (derived number match) — known broken right now.
    # Find the first keyword-based candidate instead.
    real_candidate = next(
        (c for c in candidates if "keyword" in c["reason"]), None
    )
    if real_candidate is None:
        print("FAILED: no keyword-based candidate found.")
        return

    print(f"Opening real candidate directly: row {real_candidate['row']} "
          f"({real_candidate['reason']}, doc {real_candidate['doc_number']})")

    import time
    start_time = time.time()
    print(f"Start time: {start_time}")

    sap_utils.open_document_row(session, grid, real_candidate["row"])

    try:
        file_path = sap_utils.wait_for_new_temp_file(start_time)
        print(f"SUCCESS: new file found at {file_path}")
    except TimeoutError as e:
        print(f"FAILED: {e}")
        print("If this also times out, the issue is broader than just "
              "candidate 1's lock conflict.")


if __name__ == "__main__":
    main()