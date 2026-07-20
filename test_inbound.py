"""
Inbound verification. Pure routing-table checks (fabricated MailFeatures), then
router behaviour against fakes (RecordingContactLookup + KeywordRfqScorer + a real
QuoteJobStore/QuoteJobController with the step-3 RecordingPort). No Outlook, no API.
Run:  python test_inbound.py
"""
import os
import shutil
import tempfile

from quote_jobs_classifier import (
    Bucket, ClassifierConfig, MailFeatures, classify,
)
from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import QuoteJobController
from quote_jobs_model import Customer, Vendor, RequestStatus
from quote_jobs_recording_port import RecordingPort
from inbound_store import InboundStore
from inbound_router import InboundMail, InboundRouter, KeywordRfqScorer
from inbound_recording import RecordingContactLookup

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


# ============================ pure routing table ============================
def test_routing_table():
    # hard vendor thread match wins, even for an unknown sender with no RFQ signal
    r = classify(MailFeatures(vendor_thread_match=("QJ-1", "RFQ-1")))
    check("vendor thread match -> VENDOR_REPLY, auto",
          r.bucket is Bucket.VENDOR_REPLY and r.auto and r.job_id == "QJ-1" and r.rfq_id == "RFQ-1")

    r = classify(MailFeatures(customer_thread_match="QJ-7"))
    check("customer thread match -> CUSTOMER_FOLLOWUP, auto",
          r.bucket is Bucket.CUSTOMER_FOLLOWUP and r.auto and r.job_id == "QJ-7")

    r = classify(MailFeatures(subject_suggested_match=("QJ-2", "RFQ-9"), rfq_signal=0.99))
    check("subject-only match -> NEEDS_TRIAGE and NEVER auto",
          r.bucket is Bucket.NEEDS_TRIAGE and r.auto is False)

    r = classify(MailFeatures(sender_is_known_customer=True, rfq_signal=0.9))
    check("known customer + strong RFQ -> NEW_CUSTOMER_RFQ, confirm (not auto)",
          r.bucket is Bucket.NEW_CUSTOMER_RFQ and r.auto is False)

    r = classify(MailFeatures(sender_is_known_vendor=True, rfq_signal=0.9))
    check("known vendor + strong RFQ, no thread -> NEEDS_TRIAGE",
          r.bucket is Bucket.NEEDS_TRIAGE and r.auto is False)

    r = classify(MailFeatures(rfq_signal=0.9))
    check("unknown sender + strong RFQ -> NEEDS_TRIAGE (new-customer case)",
          r.bucket is Bucket.NEEDS_TRIAGE and r.auto is False)

    r = classify(MailFeatures(sender_is_known_customer=True, rfq_signal=0.1))
    check("known customer + weak signal -> IGNORE", r.bucket is Bucket.IGNORE)

    r = classify(MailFeatures(rfq_signal=0.1))
    check("unknown + weak signal -> IGNORE", r.bucket is Bucket.IGNORE)

    # threshold is honoured
    cfg = ClassifierConfig(rfq_threshold=0.8)
    r = classify(MailFeatures(sender_is_known_customer=True, rfq_signal=0.7), cfg)
    check("signal below threshold is not 'strong' -> IGNORE", r.bucket is Bucket.IGNORE)


def test_keyword_scorer_separation():
    s = KeywordRfqScorer()
    rfq = InboundMail("m", "x@y.test", subject="Request for quote - 500MCM",
                      body="Please provide price and availability and lead time for 300ft.")
    chat = InboundMail("m", "x@y.test", subject="Lunch?", body="Want to grab lunch tomorrow?")
    check("clear RFQ scores at/above threshold", s.score(rfq) >= 0.6)
    check("chit-chat scores below threshold", s.score(chat) < 0.6)


# ============================ router behaviour ==============================
def new_router(customers=(), vendors=(), extraction=None):
    d = tempfile.mkdtemp(prefix="inb_")
    store = QuoteJobStore(os.path.join(d, "quote_jobs.json"))
    port = RecordingPort(extraction=extraction)
    ctl = QuoteJobController(store, port)
    inbound = InboundStore(os.path.join(d, "inbound.json"))
    router = InboundRouter(ctl, inbound,
                           RecordingContactLookup(customers=customers, vendors=vendors),
                           KeywordRfqScorer())
    return d, store, port, inbound, router


def test_new_customer_confirm_then_followup_no_duplicate():
    d, store, port, inbound, router = new_router(customers=["buyer@acme.test"])
    try:
        m1 = InboundMail("M1", "buyer@acme.test", subject="Request for quote",
                         body="Please quote price and availability for 300ft 500MCM.",
                         conversation_id="C-ACME", entry_id="E1")
        r1 = router.handle(m1)
        check("known customer RFQ -> queued, NOT minted",
              r1.routing.bucket is Bucket.NEW_CUSTOMER_RFQ
              and "draft_ack" not in port.names()
              and any(p["message_id"] == "M1" for p in inbound.pending))

        job, follow = router.confirm_new_customer("M1", name="Acme Co")
        check("confirm mints via controller (draft_ack fired)", "draft_ack" in port.names())
        check("minted job carries the mail's conversation_id",
              store.get_job(job.job_id).customer.conversation_id == "C-ACME")
        check("pending cleared, ledger now has the job id",
              not inbound.pending and inbound.ledger_entry("M1")["job_id"] == job.job_id)

        # a later reply on the SAME thread must attach, not mint a second job
        m2 = InboundMail("M2", "buyer@acme.test", subject="RE: Request for quote",
                         body="Revised the qty to 350ft.", conversation_id="C-ACME", entry_id="E2")
        r2 = router.handle(m2)
        check("same-thread reply -> CUSTOMER_FOLLOWUP on the same job",
              r2.routing.bucket is Bucket.CUSTOMER_FOLLOWUP and r2.routing.job_id == job.job_id)
        check("no duplicate job minted",
              len(store.all_jobs()) == 1 and port.names().count("draft_ack") == 1)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_vendor_reply_handoff_no_parse():
    d, store, port, inbound, router = new_router(vendors=["graybar.test"])
    try:
        # set up a sent RFQ with a conversation_id, directly via the store
        job = store.add_job(Customer(name="Acme", email="a@x.test"))
        li = store.add_line_item(job.job_id, "500MCM")
        rfq = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar", email="q@graybar.test"),
                                   [li.item_id])
        store.mark_rfq_sent(job.job_id, rfq.rfq_id, "C-VEND",
                            sent_at="2026-06-05T12:00:00+00:00",
                            deadline_at="2026-06-05T12:30:00+00:00")

        mail = InboundMail("V1", "q@graybar.test", subject="RE: RFQ",
                           body="$14.03/ft, in stock.", conversation_id="C-VEND", entry_id="EV1")
        r = router.handle(mail)
        check("vendor reply auto-routes to a handoff",
              r.routing.bucket is Bucket.VENDOR_REPLY and r.routing.auto
              and any(h["message_id"] == "V1" for h in inbound.handoffs))
        check("step 4 did NOT parse prices (request still AWAITING)",
              store.get_job(job.job_id).vendor_rfqs[rfq.rfq_id].requests[li.item_id].status
              is RequestStatus.AWAITING)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_dedup_is_a_noop_on_rescan():
    d, store, port, inbound, router = new_router(vendors=["graybar.test"])
    try:
        job = store.add_job(Customer(name="Acme", email="a@x.test"))
        li = store.add_line_item(job.job_id, "x")
        rfq = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar", email="q@graybar.test"), [li.item_id])
        store.mark_rfq_sent(job.job_id, rfq.rfq_id, "C-VEND", deadline_at="2026-06-05T12:30:00+00:00")
        mail = InboundMail("V1", "q@graybar.test", conversation_id="C-VEND", subject="RE: RFQ")

        router.handle(mail)
        r2 = router.handle(mail)   # same message_id again
        check("re-handling the same mail is a no-op", r2.already_processed)
        check("no duplicate handoff", len(inbound.handoffs) == 1)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_ignore_ordinary_traffic():
    d, store, port, inbound, router = new_router()
    try:
        mail = InboundMail("N1", "newsletter@spam.test", subject="Weekly deals",
                           body="Check out our blowout sale!")
        r = router.handle(mail)
        check("ordinary unknown traffic -> IGNORE, nothing queued",
              r.routing.bucket is Bucket.IGNORE
              and not inbound.pending and not inbound.triage and not inbound.handoffs)
        check("but it IS recorded in the ledger (won't be reprocessed)",
              inbound.is_processed("N1"))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_dismiss_a_pending():
    d, store, port, inbound, router = new_router(customers=["buyer@acme.test"])
    try:
        m = InboundMail("M1", "buyer@acme.test", subject="Request for quote",
                        body="please quote pricing and availability", conversation_id="C", entry_id="E")
        router.handle(m)
        router.dismiss("M1")
        check("dismiss clears pending and marks IGNORE",
              not inbound.pending and inbound.ledger_entry("M1")["bucket"] == Bucket.IGNORE.value)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_ledger_persists_across_reload():
    d, store, port, inbound, router = new_router()
    try:
        router.handle(InboundMail("X1", "a@b.test", subject="hi", body="hello"))
        reopened = InboundStore(inbound.path)   # fresh instance, same file
        check("ledger survives reload (dedup is durable)", reopened.is_processed("X1"))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_classifier_is_pure():
    import ast
    import quote_jobs_classifier as c
    tree = ast.parse(open(c.__file__).read())
    mods = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            mods += [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module:
            mods.append(node.module)
    allowed = {"__future__", "dataclasses", "enum", "typing"}
    check("classifier imports stdlib only (no store/controller/api)",
          all(m.split(".")[0] in allowed for m in mods))


if __name__ == "__main__":
    test_routing_table()
    test_keyword_scorer_separation()
    test_new_customer_confirm_then_followup_no_duplicate()
    test_vendor_reply_handoff_no_parse()
    test_dedup_is_a_noop_on_rescan()
    test_ignore_ordinary_traffic()
    test_dismiss_a_pending()
    test_ledger_persists_across_reload()
    test_classifier_is_pure()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
