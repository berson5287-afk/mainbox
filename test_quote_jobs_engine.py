"""
Standalone verification for the quote-job ENGINE. No store, no Outlook, no disk,
no real clock -- jobs are built in memory and time is handed in. This is the
"test the brain in isolation" property the whole architecture was for.
Run:  python test_quote_jobs_engine.py
"""
from datetime import datetime, timezone

from quote_jobs_model import (
    Customer, Drafts, ItemRequest, LineItem, QuoteJob, RequestStatus, RFQState,
    Vendor, VendorRFQ,
)
import quote_jobs_engine as E

PASS, FAIL = "  ok  ", "  FAIL"
_results = []

T0 = datetime(2026, 6, 5, 12, 0, 0, tzinfo=timezone.utc)        # "send" time
T_LATE = datetime(2026, 6, 5, 12, 31, 0, tzinfo=timezone.utc)   # past a 30m deadline
T_EARLY = datetime(2026, 6, 5, 12, 10, 0, tzinfo=timezone.utc)  # before deadline
DEADLINE = "2026-06-05T12:30:00+00:00"


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


def kinds(intents):
    return [type(i).__name__ for i in intents]


def find(intents, cls):
    return [i for i in intents if isinstance(i, cls)]


# --- small builders ------------------------------------------------------
def li(item_id, conf=0.99, verified=False, resolved=False):
    return LineItem(item_id=item_id, description=item_id,
                    extraction_confidence=conf, user_verified=verified,
                    resolved=resolved)


def rfq(rfq_id, state=RFQState.SENT, deadline=DEADLINE, **item_status):
    """item_status: item_id=RequestStatus, becomes the requests dict."""
    r = VendorRFQ(rfq_id=rfq_id, vendor=Vendor(name=rfq_id), state=state,
                  deadline_at=deadline if state is RFQState.SENT else None,
                  sent_at=T0.isoformat() if state is RFQState.SENT else None)
    for iid, st in item_status.items():
        r.requests[iid] = ItemRequest(status=st)
    return r


def job(items=None, rfqs=None, ack=None, quote=None):
    j = QuoteJob(job_id="QJ-1", customer=Customer(name="Acme"),
                 drafts=Drafts(ack_entry_id=ack, quote_entry_id=quote))
    for x in (items or []):
        j.line_items[x.item_id] = x
    for r in (rfqs or []):
        j.vendor_rfqs[r.rfq_id] = r
    return j


# --- forward pipeline ----------------------------------------------------
def test_customer_rfq_received():
    j = job()
    out = E.plan(j, E.CustomerRfqReceived("QJ-1"))
    check("RFQ in -> draft ack + run extraction",
          kinds(out) == ["DraftAck", "RunExtraction"])

    j2 = job(items=[li("LI-1")], ack="EID-ACK")  # ack already drafted, items exist
    out2 = E.plan(j2, E.CustomerRfqReceived("QJ-1"))
    check("ack not re-drafted, extraction not re-run", out2 == [])


def test_extraction_confidence_gate():
    cfg = E.PlannerConfig(extraction_confidence_threshold=0.85)
    j = job(items=[li("LI-1", conf=0.95), li("LI-2", conf=0.40)])
    out = E.plan(j, E.ExtractionCompleted("QJ-1"), cfg)
    rev = find(out, E.RequestItemReview)
    sel = find(out, E.RequestVendorSelection)
    check("low-confidence item flagged for review", rev and rev[0].item_ids == ["LI-2"])
    check("high-confidence item goes to vendor selection", sel and sel[0].item_ids == ["LI-1"])

    # a user-verified low-confidence item should NOT be sent back to review
    j2 = job(items=[li("LI-2", conf=0.40, verified=True)])
    out2 = E.plan(j2, E.ExtractionCompleted("QJ-1"), cfg)
    check("verified low-conf item skips review",
          not find(out2, E.RequestItemReview) and find(out2, E.RequestVendorSelection))


def test_user_verifies_then_vendor_selection():
    j = job(items=[li("LI-2", conf=0.4, verified=True)])
    out = E.plan(j, E.UserVerifiedItem("QJ-1", "LI-2"))
    sel = find(out, E.RequestVendorSelection)
    check("verifying a murky item prompts vendor selection",
          sel and sel[0].item_ids == ["LI-2"])


def test_vendor_selected_drafts_rfq():
    j = job(items=[li("LI-1"), li("LI-2")])
    v = Vendor(name="Graybar", email="q@g.test")
    out = E.plan(j, E.VendorSelected("QJ-1", v, ["LI-1", "LI-2"]))
    d = find(out, E.DraftVendorRfq)
    check("vendor selection drafts an rfq for those items",
          d and d[0].vendor.name == "Graybar" and d[0].item_ids == ["LI-1", "LI-2"])


def test_sent_is_noop_for_brain():
    j = job(items=[li("LI-1")], rfqs=[rfq("RFQ-1", **{"LI-1": RequestStatus.AWAITING})])
    check("send emits nothing (deadline snapshot does the work)",
          E.plan(j, E.VendorRfqSent("QJ-1", "RFQ-1")) == [])


# --- reply handling ------------------------------------------------------
def test_reply_priced_flags_quoted():
    j = job(items=[li("LI-1")], rfqs=[rfq("RFQ-1", **{"LI-1": RequestStatus.PRICED})])
    out = E.plan(j, E.VendorReplyParsed("QJ-1", "RFQ-1", ["LI-1"]))
    check("priced reply -> flag item quoted", kinds(out) == ["FlagItemQuoted"])


def test_reply_question_notifies():
    j = job(items=[li("LI-1")], rfqs=[rfq("RFQ-1", **{"LI-1": RequestStatus.QUESTION})])
    out = E.plan(j, E.VendorReplyParsed("QJ-1", "RFQ-1", ["LI-1"]))
    n = find(out, E.NotifyResponseNeeded)
    check("question reply -> notify response needed", n and n[0].item_ids == ["LI-1"])


def test_reply_partial_counts_as_needs_response():
    j = job(items=[li("LI-1")], rfqs=[rfq("RFQ-1", **{"LI-1": RequestStatus.PARTIAL})])
    out = E.plan(j, E.VendorReplyParsed("QJ-1", "RFQ-1", ["LI-1"]))
    check("partial reply -> notify response needed (Fork A)",
          find(out, E.NotifyResponseNeeded))


def test_vendor_disagreement_surfaces_all_facets():
    # one item, three vendors disagree: A priced, B asked a question, C silent
    rA = rfq("RFQ-A", **{"LI-1": RequestStatus.PRICED})
    rB = rfq("RFQ-B", **{"LI-1": RequestStatus.QUESTION})
    rC = rfq("RFQ-C", **{"LI-1": RequestStatus.AWAITING})
    j = job(items=[li("LI-1")], rfqs=[rA, rB, rC])
    f = E.derive_item_facets(j, "LI-1")
    check("disagreeing item is quoted AND needs_response AND awaiting at once",
          f.quoted and f.needs_response and f.awaiting and not f.dead)
    check("headline puts the open question first, badge stays separate",
          E.item_headline(j, "LI-1", T_EARLY) == "NEEDS RESPONSE")


def test_dead_item():
    rA = rfq("RFQ-A", **{"LI-1": RequestStatus.DECLINED})
    rB = rfq("RFQ-B", **{"LI-1": RequestStatus.DECLINED})
    j = job(items=[li("LI-1")], rfqs=[rA, rB])
    f = E.derive_item_facets(j, "LI-1")
    check("all-declined item is dead", f.dead and not f.quoted)
    check("dead item headline is NO SUPPLY", E.item_headline(j, "LI-1", T_EARLY) == "NO SUPPLY")


# --- overdue / time ------------------------------------------------------
def test_overdue_rules():
    silent = rfq("RFQ-1", **{"LI-1": RequestStatus.AWAITING})
    check("sent + past deadline + silent = overdue", E.is_rfq_overdue(silent, T_LATE))
    check("not overdue before the deadline", not E.is_rfq_overdue(silent, T_EARLY))

    answered = rfq("RFQ-2", **{"LI-1": RequestStatus.PARTIAL})
    check("a partial reply stops the clock", not E.is_rfq_overdue(answered, T_LATE))

    draft = rfq("RFQ-3", state=RFQState.DRAFT, **{"LI-1": RequestStatus.AWAITING})
    check("a draft (unsent) rfq is never overdue", not E.is_rfq_overdue(draft, T_LATE))

    nodl = rfq("RFQ-4", deadline=None, **{"LI-1": RequestStatus.AWAITING})
    nodl.state = RFQState.SENT
    check("no deadline -> not overdue", not E.is_rfq_overdue(nodl, T_LATE))


def test_tick_alerts_once():
    r = rfq("RFQ-1", **{"LI-1": RequestStatus.AWAITING})
    j = job(items=[li("LI-1")], rfqs=[r])
    out = E.plan(j, E.Tick(T_LATE))
    check("tick past deadline alerts overdue", kinds(out) == ["AlertOverdue"])

    r.overdue_alerted = True                      # adapter set the guard
    out2 = E.plan(j, E.Tick(T_LATE))
    check("already-alerted rfq does not re-alert", out2 == [])

    out3 = E.plan(j, E.Tick(T_EARLY))
    # reset guard to isolate the time check
    r.overdue_alerted = False
    out3 = E.plan(j, E.Tick(T_EARLY))
    check("tick before deadline alerts nothing", out3 == [])


# --- job readiness -------------------------------------------------------
def test_suggest_customer_quote():
    items = [li("LI-1", resolved=True), li("LI-2", resolved=False)]
    j = job(items=items)
    check("not all resolved -> no customer-quote suggestion",
          E.plan(j, E.ItemResolved("QJ-1", "LI-1")) == [])

    j2 = job(items=[li("LI-1", resolved=True), li("LI-2", resolved=True)])
    out = E.plan(j2, E.ItemResolved("QJ-1", "LI-2"))
    check("all resolved + no quote draft -> suggest customer quote",
          kinds(out) == ["SuggestCustomerQuote"])

    j3 = job(items=[li("LI-1", resolved=True)], quote="EID-Q")
    check("existing customer-quote draft -> no re-suggest",
          E.plan(j3, E.ItemResolved("QJ-1", "LI-1")) == [])


def test_job_summary_headline():
    # all priced, none resolved -> "ALL QUOTED" (prices in, not yet picked/priced)
    rA = rfq("RFQ-A", **{"LI-1": RequestStatus.PRICED, "LI-2": RequestStatus.PRICED})
    j = job(items=[li("LI-1"), li("LI-2")], rfqs=[rA])
    s = E.derive_job_summary(j, T_EARLY)
    check("summary counts items + quoted", s.n_items == 2 and s.n_quoted == 2)
    check("all priced, none resolved -> headline ALL QUOTED",
          s.all_quoted and not s.all_resolved and s.headline == "ALL QUOTED")

    # mark both resolved -> READY TO QUOTE
    j.line_items["LI-1"].resolved = True
    j.line_items["LI-2"].resolved = True
    s2 = E.derive_job_summary(j, T_EARLY)
    check("all resolved -> headline READY TO QUOTE",
          s2.all_resolved and s2.headline == "READY TO QUOTE")

    # an open question dominates the headline
    j.vendor_rfqs["RFQ-A"].requests["LI-1"].status = RequestStatus.QUESTION
    s3 = E.derive_job_summary(j, T_EARLY)
    check("open question -> headline NEEDS RESPONSE",
          s3.has_needs_response and s3.headline == "NEEDS RESPONSE")


def test_engine_imports_no_store():
    import quote_jobs_engine as eng
    src = open(eng.__file__).read()
    check("engine never imports the store", "quote_jobs_store" not in src)
    check("engine never reads the wall clock", "datetime.now" not in src)


if __name__ == "__main__":
    test_customer_rfq_received()
    test_extraction_confidence_gate()
    test_user_verifies_then_vendor_selection()
    test_vendor_selected_drafts_rfq()
    test_sent_is_noop_for_brain()
    test_reply_priced_flags_quoted()
    test_reply_question_notifies()
    test_reply_partial_counts_as_needs_response()
    test_vendor_disagreement_surfaces_all_facets()
    test_dead_item()
    test_overdue_rules()
    test_tick_alerts_once()
    test_suggest_customer_quote()
    test_job_summary_headline()
    test_engine_imports_no_store()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
