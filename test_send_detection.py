"""
Send-detection + capture verification, plus the Fork-C signals it lights up.
No Outlook. Run:  python test_send_detection.py
"""
import os
import shutil
import tempfile
from datetime import datetime, timedelta, timezone

from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import QuoteJobController, ControllerConfig
from quote_jobs_model import Customer, Vendor, RequestStatus, RFQState
from quote_jobs_recording_port import RecordingPort
from send_detection import SentMail, match_sent
from inbound_router import InboundMail, resolve_features, KeywordRfqScorer
from inbound_recording import RecordingContactLookup
from quote_jobs_classifier import Bucket, classify

PASS, FAIL = "  ok  ", "  FAIL"
_results = []
T0 = datetime(2026, 6, 5, 12, 0, 0, tzinfo=timezone.utc)


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


def new_ctx(timeout=30):
    d = tempfile.mkdtemp(prefix="snd_")
    store = QuoteJobStore(os.path.join(d, "quote_jobs.json"))
    ctl = QuoteJobController(store, RecordingPort(), ControllerConfig(default_timeout_minutes=timeout))
    return d, store, ctl


def drafted_rfq(ctl, store):
    """Drive a draft via the controller so conversation_id is captured at draft time."""
    job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
    store.add_line_item(job.job_id, "500MCM")
    ctl.select_vendor(job.job_id, Vendor(name="Graybar", email="q@graybar.test"), ["LI-1"])
    j = store.get_job(job.job_id)
    rfq_id = list(j.vendor_rfqs)[-1]
    return job.job_id, rfq_id, j.vendor_rfqs[rfq_id]


def test_conversation_captured_at_draft_time():
    d, store, ctl = new_ctx()
    try:
        jid, rid, rfq = drafted_rfq(ctl, store)
        check("draft has a conversation_id (captured at draft time)", bool(rfq.conversation_id))
        check("RFQ still DRAFT before any send", rfq.state is RFQState.DRAFT)
        check("drafted RFQ is findable by its conversation",
              store.find_rfq_by_conversation(rfq.conversation_id) is not None)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_capture_sent_records_and_snapshots():
    d, store, ctl = new_ctx(timeout=30)
    try:
        jid, rid, rfq = drafted_rfq(ctl, store)
        sm = SentMail(message_id="<rfq-1@apesc>", conversation_id=rfq.conversation_id,
                      subject="Request for quote - 500MCM", sent_on=T0)
        hit = ctl.capture_sent(sm)
        r = store.get_job(jid).vendor_rfqs[rid]
        check("capture_sent matched the drafted RFQ", hit == (jid, rid))
        check("RFQ flipped to SENT", r.state is RFQState.SENT)
        check("Message-ID captured", r.sent_message_id == "<rfq-1@apesc>")
        check("subject normalized + stored", r.subject_norm == "request for quote - 500mcm")
        check("deadline snapshotted = sent_on + 30min",
              r.deadline_at == (T0 + timedelta(minutes=30)).isoformat())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_capture_is_idempotent():
    d, store, ctl = new_ctx()
    try:
        jid, rid, rfq = drafted_rfq(ctl, store)
        sm = SentMail("<m@h>", conversation_id=rfq.conversation_id, subject="x", sent_on=T0)
        first = ctl.capture_sent(sm)
        deadline1 = store.get_job(jid).vendor_rfqs[rid].deadline_at
        second = ctl.capture_sent(SentMail("<m@h>", conversation_id=rfq.conversation_id,
                                           subject="x", sent_on=T0 + timedelta(hours=5)))
        deadline2 = store.get_job(jid).vendor_rfqs[rid].deadline_at
        check("first capture succeeds, re-scan is a no-op", first == (jid, rid) and second is None)
        check("idempotent: deadline not moved by the re-scan", deadline1 == deadline2)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_capture_refuses_to_guess_without_conversation():
    d, store, ctl = new_ctx()
    try:
        drafted_rfq(ctl, store)
        # a sent item with no matching conversation -> no capture, no guessing
        check("unmatched conversation -> capture_sent returns None",
              ctl.capture_sent(SentMail("<x@h>", conversation_id="NOPE", subject="whatever")) is None)
        check("no conversation id at all -> None",
              ctl.capture_sent(SentMail("<x@h>", conversation_id=None, subject="x")) is None)
        check("match_sent is pure-None on a miss",
              match_sent(SentMail("<x@h>", conversation_id="NOPE"), store) is None)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_header_match_lights_up_signal_2():
    d, store, ctl = new_ctx()
    try:
        # an RFQ already sent, with a captured Message-ID
        job = store.add_job(Customer(name="Acme", email="a@x.test"))
        li = store.add_line_item(job.job_id, "wire")
        rfq = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar", email="q@graybar.test"), [li.item_id])
        store.mark_rfq_sent(job.job_id, rfq.rfq_id, "C-OUTLOOK-WEIRD",
                            sent_at=T0.isoformat(), deadline_at=(T0 + timedelta(minutes=30)).isoformat(),
                            sent_message_id="<rfq-7@apesc>", subject="Quote req alpha")
        # vendor reply where Outlook ConversationID does NOT match, but In-Reply-To does
        mail = InboundMail("V9", "q@graybar.test", subject="RE: Quote req alpha",
                           body="$14/ft", conversation_id="SOMETHING-ELSE",
                           in_reply_to="<rfq-7@apesc>")
        feats = resolve_features(mail, store, RecordingContactLookup(vendors=["graybar.test"]),
                                 KeywordRfqScorer())
        r = classify(feats)
        check("header (In-Reply-To) match -> hard vendor match",
              feats.vendor_thread_match == (job.job_id, rfq.rfq_id))
        check("routes as VENDOR_REPLY, auto", r.bucket is Bucket.VENDOR_REPLY and r.auto)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_subject_match_lights_up_signal_3_softly():
    d, store, ctl = new_ctx()
    try:
        job = store.add_job(Customer(name="Acme", email="a@x.test"))
        li = store.add_line_item(job.job_id, "wire")
        rfq = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar", email="q@graybar.test"), [li.item_id])
        store.mark_rfq_sent(job.job_id, rfq.rfq_id, "C-1",
                            sent_at=T0.isoformat(), deadline_at=(T0 + timedelta(minutes=30)).isoformat(),
                            sent_message_id="<rfq-9@apesc>", subject="Quote for project beta")
        # reply with NO conversation match and NO header match -- subject only
        mail = InboundMail("V10", "q@graybar.test", subject="RE: Quote for project beta",
                           body="see attached", conversation_id=None, in_reply_to=None)
        feats = resolve_features(mail, store, RecordingContactLookup(vendors=["graybar.test"]),
                                 KeywordRfqScorer())
        r = classify(feats)
        check("subject match -> SOFT suggestion only",
              feats.subject_suggested_match == (job.job_id, rfq.rfq_id)
              and feats.vendor_thread_match is None)
        check("subject-only routes to NEEDS_TRIAGE, NEVER auto",
              r.bucket is Bucket.NEEDS_TRIAGE and r.auto is False)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_capture_then_reply_matches_by_conversation_end_to_end():
    d, store, ctl = new_ctx()
    try:
        jid, rid, rfq = drafted_rfq(ctl, store)
        ctl.capture_sent(SentMail("<rfq-1@apesc>", conversation_id=rfq.conversation_id,
                                  subject="Request for quote", sent_on=T0))
        # the vendor reply on that same conversation now matches by the primary key
        mail = InboundMail("VR1", "q@graybar.test", subject="RE: Request for quote",
                           body="$x", conversation_id=rfq.conversation_id)
        feats = resolve_features(mail, store, RecordingContactLookup(vendors=["graybar.test"]),
                                 KeywordRfqScorer())
        check("after capture, reply matches by conversation (signal 1)",
              feats.vendor_thread_match == (jid, rid))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_indexes_survive_reload():
    d, store, ctl = new_ctx()
    try:
        jid, rid, rfq = drafted_rfq(ctl, store)
        ctl.capture_sent(SentMail("<rfq-1@apesc>", conversation_id=rfq.conversation_id,
                                  subject="Quote ABC", sent_on=T0))
        reopened = QuoteJobStore(store.path)   # rebuild indexes from disk
        check("message-id index rebuilt on load",
              reopened.find_rfq_by_message_id("<rfq-1@apesc>") is not None)
        check("subject index rebuilt on load",
              len(reopened.find_rfqs_by_subject("quote abc")) == 1)
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    test_conversation_captured_at_draft_time()
    test_capture_sent_records_and_snapshots()
    test_capture_is_idempotent()
    test_capture_refuses_to_guess_without_conversation()
    test_header_match_lights_up_signal_2()
    test_subject_match_lights_up_signal_3_softly()
    test_capture_then_reply_matches_by_conversation_end_to_end()
    test_indexes_survive_reload()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
