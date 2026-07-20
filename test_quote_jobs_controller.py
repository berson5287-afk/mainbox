"""
Controller verification. Real QuoteJobStore (temp file) + RecordingPort fake,
no Outlook. Proves the dispatch -> effect -> write-back logic, the Fork-B/C/D
behaviours, before any of it is pointed at live Outlook.
Run:  python test_quote_jobs_controller.py
"""
import os
import shutil
import tempfile
from datetime import datetime, timedelta, timezone

from quote_jobs_model import Customer, RequestStatus, RFQState, Vendor
from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import (
    ControllerConfig, ExtractedItem, ItemReply, QuoteJobController,
)
import quote_jobs_engine as E
from quote_jobs_recording_port import RecordingPort

PASS, FAIL = "  ok  ", "  FAIL"
_results = []
T0 = datetime(2026, 6, 5, 12, 0, 0, tzinfo=timezone.utc)


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


def new_ctx(extraction=None, fail=None, timeout=30):
    d = tempfile.mkdtemp(prefix="qjctl_")
    store = QuoteJobStore(os.path.join(d, "quote_jobs.json"))
    port = RecordingPort(extraction=extraction, fail=fail)
    cfg = ControllerConfig(default_timeout_minutes=timeout)
    return d, store, QuoteJobController(store, port, cfg), port


def test_customer_rfq_drafts_ack_and_extracts():
    extraction = [ExtractedItem("500MCM THHN", extraction_confidence=0.95),
                  ExtractedItem('2" EMT', extraction_confidence=0.40)]
    d, store, ctl, port = new_ctx(extraction=extraction)
    try:
        job, follow = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        check("ack drafted + extraction run", "draft_ack" in port.names() and "run_extraction" in port.names())
        check("ack entry_id persisted (fire-once guard)",
              store.get_job(job.job_id).drafts.ack_entry_id is not None)
        check("extracted items written to store",
              len(store.get_job(job.job_id).line_items) == 2)
        check("RunExtraction hands back ExtractionCompleted (no recursion)",
              len(follow) == 1 and isinstance(follow[0], E.ExtractionCompleted))

        # dispatch the follow-on -> confidence gate splits review vs vendor-selection
        ctl.dispatch(follow[0])
        rev = port.args_for("request_item_review")
        sel = port.args_for("request_vendor_selection")
        check("low-confidence item routed to review", rev and rev[0][1] == ('LI-2',))
        check("high-confidence item routed to vendor selection", sel and sel[0][1] == ('LI-1',))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_vendor_selection_creates_rfq_and_links_draft():
    d, store, ctl, port = new_ctx()
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        store.add_line_item(job.job_id, "item")  # LI-1
        ctl.select_vendor(job.job_id, Vendor(name="Graybar", email="q@g.test"), ["LI-1"])
        j = store.get_job(job.job_id)
        check("an RFQ record was created", len(j.vendor_rfqs) == 1)
        rfq = next(iter(j.vendor_rfqs.values()))
        check("draft entry_id linked to the RFQ", rfq.draft_entry_id is not None)
        check("RFQ starts in draft state", rfq.state is RFQState.DRAFT)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_mark_sent_snapshots_deadline():
    for timeout in (30, 5):
        d, store, ctl, port = new_ctx(timeout=timeout)
        try:
            job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
            store.add_line_item(job.job_id, "item")
            ctl.select_vendor(job.job_id, Vendor(name="Graybar"), ["LI-1"])
            rfq_id = next(iter(store.get_job(job.job_id).vendor_rfqs))
            ctl.mark_rfq_sent(job.job_id, rfq_id, "CONVO-1", sent_at=T0)
            rfq = store.get_job(job.job_id).vendor_rfqs[rfq_id]
            expected = (T0 + timedelta(minutes=timeout)).isoformat()
            check(f"deadline = sent + {timeout}min (Fork D editable knob)",
                  rfq.deadline_at == expected)
            check("state flips to SENT and convo recorded",
                  rfq.state is RFQState.SENT and rfq.conversation_id == "CONVO-1")
        finally:
            shutil.rmtree(d, ignore_errors=True)


def test_reply_priced_flags_quoted():
    d, store, ctl, port = new_ctx()
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        store.add_line_item(job.job_id, "item")
        ctl.select_vendor(job.job_id, Vendor(name="Graybar"), ["LI-1"])
        rfq_id = next(iter(store.get_job(job.job_id).vendor_rfqs))
        ctl.mark_rfq_sent(job.job_id, rfq_id, "CONVO-1", sent_at=T0)
        ctl.ingest_vendor_reply(job.job_id, rfq_id,
                                [ItemReply("LI-1", RequestStatus.PRICED, price=4210.0)])
        req = store.get_job(job.job_id).vendor_rfqs[rfq_id].requests["LI-1"]
        check("reply fact written to store", req.status is RequestStatus.PRICED and req.price == 4210.0)
        check("priced reply -> flag_item_quoted called", "flag_item_quoted" in port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_reply_question_notifies():
    d, store, ctl, port = new_ctx()
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        store.add_line_item(job.job_id, "item")
        ctl.select_vendor(job.job_id, Vendor(name="Graybar"), ["LI-1"])
        rfq_id = next(iter(store.get_job(job.job_id).vendor_rfqs))
        ctl.mark_rfq_sent(job.job_id, rfq_id, "CONVO-1", sent_at=T0)
        ctl.ingest_vendor_reply(job.job_id, rfq_id,
                                [ItemReply("LI-1", RequestStatus.QUESTION)])
        check("question reply -> notify_response_needed called",
              "notify_response_needed" in port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def _make_sent_rfq(ctl, store, job_id, vendor_name, item_id, sent_at):
    store.add_line_item(job_id, item_id, )  # ensure item exists (id auto)
    # find the just-added item id
    iid = list(store.get_job(job_id).line_items)[-1]
    ctl.select_vendor(job_id, Vendor(name=vendor_name), [iid])
    rfq_id = list(store.get_job(job_id).vendor_rfqs)[-1]
    ctl.mark_rfq_sent(job_id, rfq_id, f"CONVO-{rfq_id}", sent_at=sent_at)
    return rfq_id


def test_tick_alerts_once_and_isolates_failures():
    # job A's alert will FAIL; job B's must still fire; A retries next tick
    fail = lambda name, *a: name == "alert_overdue" and a[0] == "QJ-00001"
    d, store, ctl, port = new_ctx(fail=fail)
    try:
        jobA, _ = ctl.ingest_customer_rfq(Customer(name="A", email="a@x.test"))
        jobB, _ = ctl.ingest_customer_rfq(Customer(name="B", email="b@x.test"))
        rfqA = _make_sent_rfq(ctl, store, jobA.job_id, "VA", "ia", T0)
        rfqB = _make_sent_rfq(ctl, store, jobB.job_id, "VB", "ib", T0)

        late = T0 + timedelta(minutes=31)
        ctl.tick(late)
        check("both overdue RFQs had alert attempted",
              port.names().count("alert_overdue") == 2)
        a_alerted = store.get_job(jobA.job_id).vendor_rfqs[rfqA].overdue_alerted
        b_alerted = store.get_job(jobB.job_id).vendor_rfqs[rfqB].overdue_alerted
        check("failed alert (A) did NOT set its guard -> will retry", a_alerted is False)
        check("successful alert (B) set its guard", b_alerted is True)

        # next tick: B is guarded (silent), A retries
        port.calls.clear()
        ctl.tick(late)
        check("guarded B does not re-alert; A retries exactly once",
              port.names() == ["alert_overdue"]
              and port.args_for("alert_overdue")[0][0] == "QJ-00001")
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_partial_reply_stops_the_clock():
    d, store, ctl, port = new_ctx()
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        rfq_id = _make_sent_rfq(ctl, store, job.job_id, "Graybar", "item", T0)
        ctl.ingest_vendor_reply(job.job_id, rfq_id, [ItemReply("LI-1", RequestStatus.PARTIAL)])
        port.calls.clear()
        ctl.tick(T0 + timedelta(minutes=31))
        check("a partial reply means no overdue alert", "alert_overdue" not in port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_draft_failure_is_crash_atomic():
    # draft_vendor_rfq fails -> NO RFQ record should be created (effect before persist)
    fail = lambda name, *a: name == "draft_vendor_rfq"
    d, store, ctl, port = new_ctx(fail=fail)
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        store.add_line_item(job.job_id, "item")
        ctl.select_vendor(job.job_id, Vendor(name="Graybar"), ["LI-1"])
        check("failed draft leaves zero orphan RFQ records",
              len(store.get_job(job.job_id).vendor_rfqs) == 0)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_suggest_customer_quote_when_all_resolved():
    d, store, ctl, port = new_ctx()
    try:
        job, _ = ctl.ingest_customer_rfq(Customer(name="Acme", email="a@x.test"))
        store.add_line_item(job.job_id, "i1")  # LI-1
        store.add_line_item(job.job_id, "i2")  # LI-2
        ctl.mark_item_resolved(job.job_id, "LI-1")
        check("not all resolved -> no suggestion yet", "suggest_customer_quote" not in port.names())
        ctl.mark_item_resolved(job.job_id, "LI-2")
        check("all resolved -> suggest_customer_quote fires once", "suggest_customer_quote" in port.names())
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_controller_does_not_import_mainbox():
    import ast
    import quote_jobs_controller as c
    tree = ast.parse(open(c.__file__).read())
    imported = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported += [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.append(node.module)
    # may import store/engine/model; must never import MaINbox itself
    check("controller imports no mainbox module",
          not any("mainbox" in m.lower() for m in imported))


if __name__ == "__main__":
    test_customer_rfq_drafts_ack_and_extracts()
    test_vendor_selection_creates_rfq_and_links_draft()
    test_mark_sent_snapshots_deadline()
    test_reply_priced_flags_quoted()
    test_reply_question_notifies()
    test_tick_alerts_once_and_isolates_failures()
    test_partial_reply_stops_the_clock()
    test_draft_failure_is_crash_atomic()
    test_suggest_customer_quote_when_all_resolved()
    test_controller_does_not_import_mainbox()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
