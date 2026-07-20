"""
Sandbox tests for the quote-pipeline entry point -- everything that does NOT need
Outlook: pipeline construction wired to a fake monitor, the dry-run classifier
path (and that it writes nothing), and the disabled-gate safety. The live inbox
scan + minting are exercised by you, at the keyboard, against real Outlook.
Run: python test_quote_pipeline_entry.py
"""
import os
import shutil
import tempfile

import mainbox_quote_pipeline_entry as entry
from inbound_router import InboundMail, Bucket

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


class FakeMonitor:
    """Stands in for OutlookWorkflowMonitor: only vendor_history is read here."""
    def __init__(self):
        self.vendor_history = {"vendors": {"markh@brazill.com": {}}, "customers": {}}


def _isolate_data_dir():
    """Redirect the pipeline's data dir to a temp folder so tests never touch
    %LOCALAPPDATA%\\MaINbox."""
    d = tempfile.mkdtemp(prefix="qp_entry_")
    entry._pipeline_dir = lambda: d
    return d


def test_build_pipeline_caches_and_wires():
    d = _isolate_data_dir()
    try:
        entry._load_triage_rows = lambda: [{"sender_email": "buyer@acme.test", "category": "Quote / Estimate"}]
        mon = FakeMonitor()
        p1 = entry._build_pipeline(mon)
        p2 = entry._build_pipeline(mon)
        check("pipeline builds with all parts", all(k in p1 for k in
              ("store", "inbound", "lookup", "port", "controller", "router")))
        check("pipeline is cached on the monitor (same instance)", p1 is p2)
        check("vendor recognized via monitor.vendor_history",
              p1["lookup"].is_known_vendor("markh@brazill.com", "brazill.com"))
        check("customer recognized via triage data",
              p1["lookup"].is_known_customer("buyer@acme.test", "acme.test"))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_dry_run_classifies_with_no_side_effects():
    d = _isolate_data_dir()
    try:
        entry._load_triage_rows = lambda: [{"sender_email": "buyer@acme.test", "category": "Quote / Estimate"}]
        mon = FakeMonitor()
        pipe = entry._build_pipeline(mon)
        mails = [
            InboundMail(message_id="1", sender="buyer@acme.test", subject="RFQ: 500MCM THHN",
                        body="Please quote price and availability on 500MCM.", entry_id="1"),
            InboundMail(message_id="3", sender="rando@nowhere.test", subject="hello there",
                        body="just saying hi, no business here", entry_id="3"),
        ]
        rows = entry._classify_only(pipe, mails)
        by_sender = {r["sender"]: r["bucket"] for r in rows}
        check("one classification row per mail", len(rows) == 2)
        check("known customer + RFQ -> NEW_CUSTOMER_RFQ",
              by_sender["buyer@acme.test"] is Bucket.NEW_CUSTOMER_RFQ)
        check("unknown + no RFQ content -> IGNORE",
              by_sender["rando@nowhere.test"] is Bucket.IGNORE)
        check("dry-run wrote NOTHING to the queues",
              pipe["inbound"].pending == [] and pipe["inbound"].triage == []
              and pipe["inbound"].handoffs == [] and pipe["inbound"].ledger == {})
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_disabled_gate_touches_nothing():
    # ENABLED=False must return before any Outlook access. A bare object() has no
    # get_my_inbox_folder_safe, so reaching the scan would raise AttributeError.
    saved = entry.QUOTE_PIPELINE_ENABLED
    entry.QUOTE_PIPELINE_ENABLED = False
    try:
        entry.run_quote_scan(object())   # must not raise
        check("OFF gate returns without touching Outlook", True)
    except Exception as e:
        check(f"OFF gate returns without touching Outlook (raised {e!r})", False)
    finally:
        entry.QUOTE_PIPELINE_ENABLED = saved


def test_row_is_confident():
    check("ok status is confident",
          entry._row_is_confident({"notes": "SmartScan | rfq.pdf | ok"}))
    check("confirmed status is confident",
          entry._row_is_confident({"notes": "SmartScan | rfq.pdf | confirmed"}))
    check("corrected status is confident",
          entry._row_is_confident({"notes": "SmartScan | rfq.pdf | corrected"}))
    check("warning is NOT confident",
          not entry._row_is_confident({"notes": "SmartScan | rfq.pdf | warning"}))
    check("suggestion is NOT confident",
          not entry._row_is_confident({"notes": "SmartScan | rfq.pdf | suggestion"}))
    check("no status word is NOT confident",
          not entry._row_is_confident({"notes": "SmartScan | rfq.pdf"}))
    check("empty/missing notes is NOT confident",
          not entry._row_is_confident({}))
    check("substring does not false-match (notebook)",
          not entry._row_is_confident({"notes": "SmartScan | notebook.pdf"}))


def test_prep_rows():
    rows = [
        {"description": "1in EMT", "part_number": "ABC", "qty": "10", "unit": "FT",
         "notes": "SmartScan | rfq.pdf | ok"},
        {"description": "rejected line", "notes": "SmartScan | rfq.pdf | rejected"},
        {"description": "", "notes": "SmartScan | rfq.pdf | ok"},          # empty desc -> dropped
        "not-a-dict",                                                       # junk -> dropped
    ]
    out = entry._prep_rows(rows)
    check("prep_rows drops rejected + empty + non-dict (1 survives)", len(out) == 1)
    check("prep_rows stamps status=green on survivors",
          bool(out) and out[0].get("status") == "green")
    check("prep_rows preserves the real fields",
          bool(out) and out[0]["description"] == "1in EMT" and out[0]["part_number"] == "ABC")
    from mainbox_effects import smartscan_rows_to_items
    items = smartscan_rows_to_items(out)
    check("green survivors map to >=0.85 confidence (engine will draft)",
          bool(items) and items[0].extraction_confidence >= 0.85)
    check("prep_rows on empty/None is safe", entry._prep_rows(None) == [])


def test_match_triage():
    triage = [
        {"message_id": "a", "sender": "orders@harrandelectric.com",
         "subject": "Harrand quote 26106", "received_at": "2026-06-01T10:00:00"},
        {"message_id": "b", "sender": "tom@usis.net",
         "subject": "274870 Quote", "received_at": "2026-06-03T10:00:00"},
        {"message_id": "c", "sender": "x@harrand.com",
         "subject": "follow up", "received_at": "2026-06-05T10:00:00"},
    ]
    # subject substring
    m = entry._match_triage(triage, "274870")
    check("match by subject substring", len(m) == 1 and m[0]["message_id"] == "b")
    # sender substring, case-insensitive, multiple hits, newest first
    m = entry._match_triage(triage, "HARRAND")
    check("match by sender substring (case-insensitive)", len(m) == 2)
    check("newest-first ordering", m[0]["message_id"] == "c")
    # cap
    m = entry._match_triage(triage, "harrand", 1)
    check("limit caps results", len(m) == 1 and m[0]["message_id"] == "c")
    # empty needle / empty list are safe no-ops
    check("empty needle -> no matches", entry._match_triage(triage, "") == [])
    check("none list -> no matches", entry._match_triage(None, "harrand") == [])


def test_review_policy():
    green = {"notes": "SmartScan | rfq.pdf | ok"}
    yellow = {"notes": "SmartScan | rfq.pdf | warning"}
    red = {"notes": "SmartScan | rfq.pdf | error"}
    unknown = {"notes": "SmartScan | rfq.pdf"}
    # banded confidence
    check("green -> ~0.95", entry._row_confidence(green) >= 0.9)
    check("yellow -> ~0.60", 0.5 < entry._row_confidence(yellow) < 0.8)
    check("red -> ~0.30", entry._row_confidence(red) < 0.5)
    check("unknown/unreviewed -> 0.50", entry._row_confidence(unknown) == 0.50)
    # 'always' mode always reviews, regardless of confidence
    check("always mode reviews even all-green", entry._should_review([green, green], "always", 0.9))
    # 'auto' mode: all green clears a 0.9 bar -> no review
    check("auto: all green clears 0.9", not entry._should_review([green, green], "auto", 0.9))
    # 'auto' mode: a yellow trips a 0.9 bar -> review
    check("auto: yellow trips 0.9", entry._should_review([green, yellow], "auto", 0.9))
    # 'auto' mode: lower the bar to 0.5 and yellow passes
    check("auto: yellow clears 0.5", not entry._should_review([green, yellow], "auto", 0.5))
    # no rows -> always review (nothing to auto-draft)
    check("no rows -> review", entry._should_review([], "auto", 0.9))


def test_auto_mint_targets():
    pend = [{"message_id": "a"}, {"message_id": "b"}, {"message_id": "c"}]
    # no confirm needle: newest max_n
    check("auto-mint takes max_n when no confirm", len(entry._auto_mint_targets(pend, "", 1)) == 1)
    check("auto-mint takes 2 when max_n=2", len(entry._auto_mint_targets(pend, "", 2)) == 2)
    # max_n = 0 disables auto-mint
    check("max_n=0 disables auto-mint", entry._auto_mint_targets(pend, "", 0) == [])
    # confirm needle set -> auto-mint stands down entirely, regardless of max_n
    check("confirm needle stands down auto-mint", entry._auto_mint_targets(pend, "274870", 1) == [])
    check("confirm needle beats max_n=5", entry._auto_mint_targets(pend, "harrand", 5) == [])
    # empty/None pending is safe
    check("empty pending is safe", entry._auto_mint_targets(None, "", 1) == [])


def test_saved_sender_storage():
    d = _isolate_data_dir()
    try:
        check("saved customers empty initially", entry.load_saved_senders("customer") == set())
        entry.add_saved_sender("customer", "Buyer@Acme.com")
        check("add lowercases + persists", entry.load_saved_senders("customer") == {"buyer@acme.com"})
        entry.add_saved_sender("customer", "buyer@acme.com")
        check("add is idempotent", entry.load_saved_senders("customer") == {"buyer@acme.com"})
        entry.add_saved_sender("customer", "two@acme.com")
        check("second add accumulates",
              entry.load_saved_senders("customer") == {"buyer@acme.com", "two@acme.com"})
        # vendor list is a separate file, unaffected by customer writes
        check("vendor list independent", entry.load_saved_senders("vendor") == set())
        entry.add_saved_sender("vendor", "rep@supplier.com")
        check("vendor add isolated", entry.load_saved_senders("vendor") == {"rep@supplier.com"})
        check("customer unaffected by vendor add",
              entry.load_saved_senders("customer") == {"buyer@acme.com", "two@acme.com"})
        # remove
        entry.remove_saved_sender("customer", "Buyer@Acme.com")
        check("remove drops one (case-insensitive)", entry.load_saved_senders("customer") == {"two@acme.com"})
        entry.remove_saved_sender("customer", "nope@acme.com")
        check("remove missing is safe", entry.load_saved_senders("customer") == {"two@acme.com"})
        entry.add_saved_sender("customer", "   ")
        check("blank email ignored", entry.load_saved_senders("customer") == {"two@acme.com"})
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_quote_candidates():
    triage = [
        {"message_id": "m1", "sender": "new1@bob.com", "subject": "RFQ 100", "received_at": "2026-06-01T10:00:00"},
        {"message_id": "m2", "sender": "KNOWN@acme.com", "subject": "RFQ 200", "received_at": "2026-06-02T10:00:00"},
        {"message_id": "m3", "sender": "new2@carol.com", "subject": "Need pricing", "received_at": "2026-06-03T10:00:00"},
        {"message_id": "m4", "sender": "new1@bob.com", "subject": "RFQ 100", "received_at": "2026-06-01T09:00:00"},  # dup of m1
    ]
    recognized = {"known@acme.com"}
    out = entry._quote_candidates(triage, recognized)
    senders = [e["sender"] for e in out]
    check("recognized sender filtered out", "KNOWN@acme.com" not in senders)
    check("deduped by (sender, subject)", senders.count("new1@bob.com") == 1)
    check("two unique unrecognized remain", len(out) == 2)
    check("newest first ordering", out[0]["sender"] == "new2@carol.com")
    out2 = entry._quote_candidates(triage, {"NEW1@BOB.COM"})
    check("recognition case-insensitive", "new1@bob.com" not in [e["sender"] for e in out2])
    check("empty triage safe", entry._quote_candidates([], recognized) == [])
    check("None triage safe", entry._quote_candidates(None, recognized) == [])


if __name__ == "__main__":
    test_build_pipeline_caches_and_wires()
    test_dry_run_classifies_with_no_side_effects()
    test_disabled_gate_touches_nothing()
    test_row_is_confident()
    test_prep_rows()
    test_match_triage()
    test_review_policy()
    test_auto_mint_targets()
    test_saved_sender_storage()
    test_quote_candidates()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
