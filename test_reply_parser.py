"""
Vendor-reply parser verification. No Outlook, no API -- a fake parser exercises
the policy gate precisely, the keyword parser is tested on real bodies, and the
whole path (handoff -> parse -> recorded -> engine fires) runs against the
step-3 RecordingPort. Run:  python test_reply_parser.py
"""
import os
import shutil
import tempfile
from datetime import datetime, timedelta, timezone

from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import QuoteJobController
from quote_jobs_model import Customer, Vendor, RequestStatus, StatusSource
from quote_jobs_recording_port import RecordingPort
from inbound_store import InboundStore
from reply_parser import (
    ItemProposal, ParsePolicy, KeywordReplyParser, ReplyParseController,
    extract_learning_pairs,
)

PASS, FAIL = "  ok  ", "  FAIL"
_results = []
T0 = datetime(2026, 6, 5, 12, 0, 0, tzinfo=timezone.utc)


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


class FakeParser:
    """Returns exactly the proposals it's given -- for precise policy-gate tests."""
    def __init__(self, proposals):
        self._p = proposals
    def parse(self, items, body):
        return list(self._p)


def ctx(policy, parser):
    d = tempfile.mkdtemp(prefix="rp_")
    store = QuoteJobStore(os.path.join(d, "quote_jobs.json"))
    ctl = QuoteJobController(store, RecordingPort())
    inbound = InboundStore(os.path.join(d, "inbound.json"))
    rp = ReplyParseController(ctl, inbound, parser, policy)
    return d, store, ctl, inbound, rp


def setup_sent_rfq(store, items=(("LI-1", "500MCM THHN", "thhn-500"),)):
    job = store.add_job(Customer(name="Acme", email="a@x.test"))
    ids = []
    for iid, desc, ckey in items:
        li = store.add_line_item(job.job_id, desc, canonical_key=ckey)
        ids.append(li.item_id)
    rfq = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar", email="q@graybar.test"), ids)
    store.mark_rfq_sent(job.job_id, rfq.rfq_id, "C-1", sent_at=T0.isoformat(),
                        deadline_at=(T0 + timedelta(minutes=30)).isoformat())
    return job.job_id, rfq.rfq_id, ids


def make_handoff(inbound, message_id, job_id, rfq_id):
    inbound.add_handoff({"message_id": message_id, "job_id": job_id, "rfq_id": rfq_id,
                         "sender": "q@graybar.test", "subject": "RE: RFQ"})


# ============================ policy gate ============================
def test_auto_record_high_confidence_silent():
    props = [ItemProposal("LI-1", RequestStatus.PRICED, price=4210.0, confidence=0.95)]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=True, notify_threshold=0.75), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        res = rp.parse_handoff("M1", "body")
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("high-confidence auto-records the fact",
              req.status is RequestStatus.PRICED and req.price == 4210.0)
        check("high-confidence is NOT queued for review", res.review == [] and inbound.parse_review == [])
        check("recording fired the engine (flag_item_quoted)", "flag_item_quoted" in ctl.port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_auto_record_low_confidence_records_and_flags():
    props = [ItemProposal("LI-1", RequestStatus.PRICED, price=99.0, confidence=0.40)]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=True, notify_threshold=0.75), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        res = rp.parse_handoff("M1", "body")
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("low-confidence STILL records (auto-record on)", req.status is RequestStatus.PRICED)
        check("low-confidence ALSO flagged for verification (the slider)",
              res.review == ["LI-1"]
              and any(e["item_id"] == "LI-1" and e["recorded"] for e in inbound.parse_review))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_review_first_records_nothing_until_confirmed():
    props = [ItemProposal("LI-1", RequestStatus.PRICED, price=4210.0, confidence=0.99)]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=False), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        rp.parse_handoff("M1", "body")
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("review-first records NOTHING up front",
              req.status is RequestStatus.AWAITING and "flag_item_quoted" not in ctl.port.names())
        check("review-first queues the proposal (not recorded)",
              any(e["item_id"] == "LI-1" and not e["recorded"] for e in inbound.parse_review))

        rp.confirm_parse_item("M1", "LI-1")
        req2 = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("confirm records the fact and fires the engine",
              req2.status is RequestStatus.PRICED and "flag_item_quoted" in ctl.port.names())
        check("confirm clears the review entry", inbound.parse_review == [])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_correct_preserves_ai_guess_for_learning():
    # AI proposed PARTIAL (low conf) under review-first; user corrects to PRICED
    props = [ItemProposal("LI-1", RequestStatus.PARTIAL, confidence=0.4, excerpt="in stock, no price")]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=False), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        rp.parse_handoff("M1", "body")
        rp.correct_parse_item("M1", "LI-1", RequestStatus.PRICED, price=212.5)
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("correction applied as USER", req.status is RequestStatus.PRICED and req.status_source is StatusSource.USER)
        check("AI's original guess preserved (learning pair)", req.ai_proposed_status is RequestStatus.PARTIAL)

        pairs = extract_learning_pairs(store)
        check("learning pair extracted (predicted vs actual)",
              len(pairs) == 1 and pairs[0]["predicted"] == "partial" and pairs[0]["actual"] == "priced")
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_dismiss_drops_without_recording():
    props = [ItemProposal("LI-1", RequestStatus.QUESTION, confidence=0.3)]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=False), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        rp.parse_handoff("M1", "body")
        rp.dismiss_parse_item("M1", "LI-1")
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("dismiss leaves the request untouched and clears the queue",
              req.status is RequestStatus.AWAITING and inbound.parse_review == [])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_parse_is_idempotent():
    props = [ItemProposal("LI-1", RequestStatus.PRICED, price=10.0, confidence=0.95)]
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=True), FakeParser(props))
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        rp.parse_handoff("M1", "body")
        res2 = rp.parse_handoff("M1", "body")    # handoff already consumed
        check("re-parsing a consumed handoff is a no-op", res2.skipped)
        check("flag_item_quoted fired exactly once", ctl.port.names().count("flag_item_quoted") == 1)
    finally:
        shutil.rmtree(d, ignore_errors=True)


# ============================ keyword parser ============================
def test_keyword_parser_reads_bodies():
    p = KeywordReplyParser()
    items = [{"item_id": "LI-1", "description": "500MCM THHN", "canonical_key": "thhn-500",
              "qty": 300, "uom": "FT"}]
    priced = p.parse(items, "We can do 500MCM at $14.03/ft, in stock.")
    check("price + availability -> PRICED with a number",
          priced and priced[0].status is RequestStatus.PRICED and priced[0].price == 14.03)
    declined = p.parse(items, "That item is discontinued, sorry.")
    check("decline language -> DECLINED", declined and declined[0].status is RequestStatus.DECLINED)
    question = p.parse(items, "Which exact insulation rating do you need?")
    check("a question -> QUESTION", question and question[0].status is RequestStatus.QUESTION)
    partial = p.parse(items, "It's in stock; price to follow.")
    check("availability without a price -> PARTIAL",
          partial and partial[0].status is RequestStatus.PARTIAL)
    silent = p.parse(items, "Thanks for your inquiry, talk soon.")
    check("no statement about the item -> no proposal (stays AWAITING)", silent == [])


def test_end_to_end_keyword_through_engine():
    d, store, ctl, inbound, rp = ctx(ParsePolicy(auto_record=True, notify_threshold=0.75),
                                     KeywordReplyParser())
    try:
        jid, rid, _ = setup_sent_rfq(store)
        make_handoff(inbound, "M1", jid, rid)
        rp.parse_handoff("M1", "500MCM is $14.03/ft and in stock.")
        req = store.get_job(jid).vendor_rfqs[rid].requests["LI-1"]
        check("keyword parser -> recorded PRICED via the full path",
              req.status is RequestStatus.PRICED and req.price == 14.03)
        check("engine flagged the item quoted", "flag_item_quoted" in ctl.port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    test_auto_record_high_confidence_silent()
    test_auto_record_low_confidence_records_and_flags()
    test_review_first_records_nothing_until_confirmed()
    test_correct_preserves_ai_guess_for_learning()
    test_dismiss_drops_without_recording()
    test_parse_is_idempotent()
    test_keyword_parser_reads_bodies()
    test_end_to_end_keyword_through_engine()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
