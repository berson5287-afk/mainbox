"""
MaINbox quote-job ENGINE -- the pure decision layer (the "state machine").

Hard rules that make this testable in isolation:
  * Imports the model ONLY. No store, no Outlook, no file I/O.
  * Never reads the wall clock. `now` arrives inside a Tick event.
  * Never mutates anything. It reads a QuoteJob and RETURNS intentions.
    Step-3 adapters execute those intentions and write resulting facts back
    through the store's mutators. Decision here; effects there.

Two things live here:
  1. Derivations -- pure read-only views the UI renders (item facets, job
     summary, overdue). Status is computed, never stored (that is why step 1
     left status off LineItem and QuoteJob).
  2. plan(job, event) -> [Intention] -- given the job's CURRENT state and the
     event that just happened, what should the system do next.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import List, Optional

from quote_jobs_model import (
    QuoteJob, Vendor, LineItem, VendorRFQ,
    RFQState, RequestStatus,
)


# ==========================================================================
# config
# ==========================================================================
@dataclass
class PlannerConfig:
    # Gates auto-accept vs flag-for-review of an extracted item. This is about
    # how sure SmartScan is of WHAT the item is -- not who to send it to. Tune
    # against your own extraction_confidence distribution.
    extraction_confidence_threshold: float = 0.85


DEFAULT_CONFIG = PlannerConfig()


# ==========================================================================
# events  (what the outside world tells the brain happened)
# ==========================================================================
@dataclass
class CustomerRfqReceived:
    job_id: str


@dataclass
class ExtractionCompleted:
    job_id: str


@dataclass
class UserVerifiedItem:
    job_id: str
    item_id: str


@dataclass
class VendorSelected:
    job_id: str
    vendor: Vendor
    item_ids: List[str]


@dataclass
class VendorRfqSent:
    job_id: str
    rfq_id: str


@dataclass
class VendorReplyParsed:
    job_id: str
    rfq_id: str
    item_ids: List[str]   # the item requests this reply just updated


@dataclass
class ItemResolved:
    job_id: str
    item_id: str


@dataclass
class Tick:
    now: datetime         # the ONLY way the brain learns the time


# ==========================================================================
# intentions  (what the brain says should happen; adapters execute)
# ==========================================================================
@dataclass
class DraftAck:
    job_id: str


@dataclass
class RunExtraction:
    job_id: str


@dataclass
class RequestItemReview:
    job_id: str
    item_ids: List[str]


@dataclass
class RequestVendorSelection:
    job_id: str
    item_ids: List[str]


@dataclass
class DraftVendorRfq:
    job_id: str
    vendor: Vendor
    item_ids: List[str]


@dataclass
class FlagItemQuoted:
    job_id: str
    item_id: str


@dataclass
class NotifyResponseNeeded:
    job_id: str
    rfq_id: str
    item_ids: List[str]


@dataclass
class AlertOverdue:
    # adapter contract: after alerting, record overdue_alerted=True so this
    # fires exactly once (the brain reads that guard each Tick).
    job_id: str
    rfq_id: str


@dataclass
class SuggestCustomerQuote:
    job_id: str


# ==========================================================================
# derived views  (pure, read-only)
# ==========================================================================
@dataclass
class ItemFacets:
    """Orthogonal flags for one line item, folded across every vendor's ask.
    Timeless -- depends only on reply state, never on the clock."""
    quoted: bool = False          # >=1 live request PRICED  -> "flag item quoted"
    needs_response: bool = False  # >=1 live request QUESTION or PARTIAL
    awaiting: bool = False        # >=1 SENT rfq still has an AWAITING ask
    dead: bool = False            # has live requests and ALL of them DECLINED


@dataclass
class JobSummary:
    n_items: int = 0
    n_quoted: int = 0
    n_resolved: int = 0
    all_quoted: bool = False      # every item quoted -> the "all prices in" state
    all_resolved: bool = False    # every item resolved -> ready to quote customer
    has_needs_response: bool = False
    has_overdue: bool = False     # any SENT rfq overdue at `now`
    headline: str = "NEW"


def _parse_iso(s: str) -> datetime:
    dt = datetime.fromisoformat(s)
    return dt if dt.tzinfo is not None else dt.replace(tzinfo=timezone.utc)


def _vendor_response_received(rfq: VendorRFQ) -> bool:
    """True if the vendor has actually said something on any ask. AWAITING and
    our own CANCELLED do not count -- those are silence / our action."""
    answered = (RequestStatus.PRICED, RequestStatus.PARTIAL,
                RequestStatus.QUESTION, RequestStatus.DECLINED)
    return any(req.status in answered for req in rfq.requests.values())


def is_rfq_overdue(rfq: VendorRFQ, now: datetime) -> bool:
    """Per Fork C: overdue only while SENT, past the deadline snapshot, with
    zero vendor response. A partial reply stops the clock."""
    if rfq.state is not RFQState.SENT:
        return False
    if not rfq.deadline_at:
        return False
    if _vendor_response_received(rfq):
        return False
    return now >= _parse_iso(rfq.deadline_at)


def _live_requests_for_item(job: QuoteJob, item_id: str):
    """(rfq, request) pairs for an item, skipping cancelled rfqs and cancelled
    asks -- i.e. the asks that still count toward the item's status."""
    for rfq in job.vendor_rfqs.values():
        if rfq.state is RFQState.CANCELLED:
            continue
        req = rfq.requests.get(item_id)
        if req is not None and req.status is not RequestStatus.CANCELLED:
            yield rfq, req


def derive_item_facets(job: QuoteJob, item_id: str) -> ItemFacets:
    f = ItemFacets()
    live = list(_live_requests_for_item(job, item_id))
    for rfq, req in live:
        if req.status is RequestStatus.PRICED:
            f.quoted = True
        elif req.status in (RequestStatus.QUESTION, RequestStatus.PARTIAL):
            f.needs_response = True
        if rfq.state is RFQState.SENT and req.status is RequestStatus.AWAITING:
            f.awaiting = True
    f.dead = bool(live) and all(req.status is RequestStatus.DECLINED for _, req in live)
    return f


def item_needs_review(item: LineItem, config: PlannerConfig = DEFAULT_CONFIG) -> bool:
    """Derived: a low-confidence, not-yet-verified extraction needs human eyes."""
    return (not item.user_verified
            and item.extraction_confidence < config.extraction_confidence_threshold)


def _item_overdue(job: QuoteJob, item_id: str, now: datetime) -> bool:
    for rfq in job.vendor_rfqs.values():
        req = rfq.requests.get(item_id)
        if req is not None and req.status is RequestStatus.AWAITING and is_rfq_overdue(rfq, now):
            return True
    return False


def item_headline(job: QuoteJob, item_id: str, now: datetime) -> str:
    """Single most-action-first label for a list row. The QUOTED badge should
    still be shown independently (an item can be QUOTED and headlined NEEDS
    RESPONSE at the same time)."""
    f = derive_item_facets(job, item_id)
    if f.needs_response:
        return "NEEDS RESPONSE"
    if _item_overdue(job, item_id, now):
        return "OVERDUE"
    if f.awaiting:
        return "AWAITING"
    if f.quoted:
        return "QUOTED"
    if f.dead:
        return "NO SUPPLY"
    return "PENDING"


def _all_resolved(job: QuoteJob) -> bool:
    return bool(job.line_items) and all(li.resolved for li in job.line_items.values())


def derive_job_summary(job: QuoteJob, now: datetime) -> JobSummary:
    items = job.line_items
    facets = {iid: derive_item_facets(job, iid) for iid in items}
    n = len(items)
    s = JobSummary(
        n_items=n,
        n_quoted=sum(1 for f in facets.values() if f.quoted),
        n_resolved=sum(1 for li in items.values() if li.resolved),
        all_quoted=(n > 0 and all(f.quoted for f in facets.values())),
        all_resolved=_all_resolved(job),
        has_needs_response=any(f.needs_response for f in facets.values()),
        has_overdue=any(is_rfq_overdue(rfq, now) for rfq in job.vendor_rfqs.values()),
    )
    all_dead = n > 0 and all(f.dead for f in facets.values())
    if n == 0:
        s.headline = "NEW"
    elif s.has_needs_response:
        s.headline = "NEEDS RESPONSE"
    elif s.has_overdue:
        s.headline = "OVERDUE"
    elif all_dead:
        s.headline = "NO SUPPLY"
    elif not s.all_quoted:
        s.headline = "AWAITING QUOTES"
    elif not s.all_resolved:
        s.headline = "ALL QUOTED"      # the "all prices in" nudge, as state
    else:
        s.headline = "READY TO QUOTE"
    return s


# ==========================================================================
# the planner
# ==========================================================================
def plan(job: QuoteJob, event, config: PlannerConfig = DEFAULT_CONFIG) -> List:
    """Given the job's current (post-event) state and the event that just
    occurred, return the intentions that should fire. Pure: reads only."""

    if isinstance(event, CustomerRfqReceived):
        out = []
        if job.drafts.ack_entry_id is None:          # don't re-draft the ack
            out.append(DraftAck(job.job_id))
        if not job.line_items:                       # don't re-extract
            out.append(RunExtraction(job.job_id))
        return out

    if isinstance(event, ExtractionCompleted):
        review, ready = [], []
        for iid, li in job.line_items.items():
            (review if item_needs_review(li, config) else ready).append(iid)
        out = []
        if review:
            out.append(RequestItemReview(job.job_id, review))
        if ready:                                    # vendor pick is human in v1
            out.append(RequestVendorSelection(job.job_id, ready))
        return out

    if isinstance(event, UserVerifiedItem):
        # a murky item just cleared review -> now ready for vendor selection
        return [RequestVendorSelection(job.job_id, [event.item_id])]

    if isinstance(event, VendorSelected):
        return [DraftVendorRfq(job.job_id, event.vendor, list(event.item_ids))]

    if isinstance(event, VendorRfqSent):
        return []   # nothing to emit: the deadline_at snapshot does the work

    if isinstance(event, VendorReplyParsed):
        rfq = job.vendor_rfqs.get(event.rfq_id)
        if rfq is None:
            return []
        out = []
        needs = []
        for iid in event.item_ids:
            req = rfq.requests.get(iid)
            if req is None:
                continue
            if req.status is RequestStatus.PRICED:
                out.append(FlagItemQuoted(job.job_id, iid))
            elif req.status in (RequestStatus.QUESTION, RequestStatus.PARTIAL):
                needs.append(iid)
        if needs:
            out.append(NotifyResponseNeeded(job.job_id, event.rfq_id, needs))
        return out

    if isinstance(event, ItemResolved):
        if _all_resolved(job) and job.drafts.quote_entry_id is None:
            return [SuggestCustomerQuote(job.job_id)]
        return []

    if isinstance(event, Tick):
        out = []
        for rid, rfq in job.vendor_rfqs.items():
            if is_rfq_overdue(rfq, event.now) and not rfq.overdue_alerted:
                out.append(AlertOverdue(job.job_id, rid))
        return out

    return []
