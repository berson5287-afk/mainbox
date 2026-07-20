"""
MaINbox quote-job CONTROLLER -- the adapter layer (step 3).

This is the seam: the first place the pure engine and the persistence store
meet, and the ONLY layer that is allowed to touch effects. It still does not
import MaINbox. Instead it depends on an EffectsPort that MaINbox implements
against your real Outlook / SmartScan / notification routines.

Flow per intention (Fork B): run the effect via the port with NO store lock
held, capture its result, then apply that intention's store writes in one short
atomic batch. Effect first means the lock never wraps a COM call and a crash
mid-dispatch can at worst leave a harmless orphan Outlook draft -- never an
orphan record in the store.

Dispatch (Fork C) is single-step: it returns natural follow-on events rather
than recursing, and isolates a failing effect (log it, skip its guard write so
it retries, keep running the other intentions).

SCOPE: manually invoked only -- nothing here fires off the live Outlook scan.
No inbound classifier (you hand it a customer). Customer-quote is a nudge only.
Send-detection / entry_id reverse-index is deferred to the auto-trigger step;
here you pass rfq_id + conversation_id yourself.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Protocol, Tuple

from quote_jobs_model import RequestStatus, StatusSource, Vendor, RFQState
from quote_jobs_store import QuoteJobStore
import quote_jobs_engine as E
from send_detection import SentMail, match_sent

log = logging.getLogger("mainbox.quote_jobs.controller")


@dataclass
class DraftResult:
    """What a drafting effect returns: the created item's EntryID and (for vendor
    RFQs) its Outlook ConversationID, captured so send-detection can match later."""
    entry_id: str
    conversation_id: Optional[str] = None


# ==========================================================================
# config + I/O contracts
# ==========================================================================
@dataclass
class ControllerConfig:
    default_timeout_minutes: int = 30                 # Fork D: the editable knob
    planner: E.PlannerConfig = field(default_factory=E.PlannerConfig)


@dataclass
class ExtractedItem:
    """What run_extraction() hands back for each material line it found."""
    description: str
    canonical_key: Optional[str] = None
    qty: float = 1
    uom: str = "EA"
    extraction_confidence: float = 0.0


@dataclass
class ItemReply:
    """One parsed (item, this-vendor) reply fact, as the step-5 parser will
    eventually produce. In step 3 you build these by hand in the REPL."""
    item_id: str
    status: RequestStatus
    price: Optional[float] = None
    lead_time: Optional[str] = None
    status_confidence: Optional[float] = None
    raw_excerpt: Optional[str] = None
    response_entry_id: Optional[str] = None


# ==========================================================================
# the port MaINbox implements
# ==========================================================================
class EffectsPort(Protocol):
    """Structural interface. MaINbox provides an object with these methods;
    no inheritance or import needed. Methods that create an Outlook item RETURN
    its entry_id so the controller can persist the link."""
    def draft_ack(self, job) -> str: ...                          # -> draft entry_id
    def run_extraction(self, job) -> List[ExtractedItem]: ...
    def request_item_review(self, job, item_ids: List[str]) -> None: ...
    def request_vendor_selection(self, job, item_ids: List[str]) -> None: ...
    def draft_vendor_rfq(self, job, vendor: Vendor, item_ids: List[str]) -> "DraftResult": ...  # entry_id + conversation_id
    def flag_item_quoted(self, job, item_id: str) -> None: ...
    def notify_response_needed(self, job, rfq_id: str, item_ids: List[str]) -> None: ...
    def alert_overdue(self, job, rfq_id: str) -> None: ...
    def suggest_customer_quote(self, job) -> None: ...


class SkeletonPort:
    """Copy this, rename to e.g. MaINboxEffectsPort, and replace each TODO with
    your existing routine. The controller calls these; you map them onto your
    Outlook COM / SmartScan / notification code. Drafting methods must RETURN
    the created item's entry_id."""

    def draft_ack(self, job) -> str:
        # TODO: create the "I will advise" draft to job.customer.email; return its entry_id
        raise NotImplementedError("wire draft_ack to your Outlook draft routine")

    def run_extraction(self, job) -> List[ExtractedItem]:
        # TODO: run SmartScan on the customer RFQ source; return one ExtractedItem per line
        raise NotImplementedError("wire run_extraction to SmartScan")

    def request_item_review(self, job, item_ids) -> None:
        # TODO: surface these low-confidence items for human verification in the UI
        raise NotImplementedError("wire request_item_review to your UI")

    def request_vendor_selection(self, job, item_ids) -> None:
        # TODO: prompt the user to pick/add vendor(s) for these items
        raise NotImplementedError("wire request_vendor_selection to your UI")

    def draft_vendor_rfq(self, job, vendor, item_ids) -> "DraftResult":
        # TODO: draft an RFQ email to vendor for these items; return DraftResult with
        # the draft's EntryID AND its Outlook ConversationID (mail.ConversationID).
        raise NotImplementedError("wire draft_vendor_rfq to your Outlook draft routine")

    def flag_item_quoted(self, job, item_id) -> None:
        # TODO: mark the item quoted in the open-quotes view (idempotent render)
        raise NotImplementedError("wire flag_item_quoted to your UI")

    def notify_response_needed(self, job, rfq_id, item_ids) -> None:
        # TODO: notify the user a vendor reply is incomplete / asked a question
        raise NotImplementedError("wire notify_response_needed to your notifier")

    def alert_overdue(self, job, rfq_id) -> None:
        # TODO: alert the user this RFQ passed its deadline with no reply
        raise NotImplementedError("wire alert_overdue to your notifier")

    def suggest_customer_quote(self, job) -> None:
        # TODO: surface "all items resolved -- ready to quote the customer"
        raise NotImplementedError("wire suggest_customer_quote to your UI")


# ==========================================================================
# the controller
# ==========================================================================
class QuoteJobController:
    def __init__(self, store: QuoteJobStore, port: EffectsPort,
                 config: Optional[ControllerConfig] = None):
        self.store = store
        self.port = port
        self.config = config or ControllerConfig()

    # ---- input side: things you call (REPL / UI / later the parser) --------
    def ingest_customer_rfq(self, customer):
        """Create the job and react. Returns (job, follow_on_events)."""
        job = self.store.add_job(customer)
        return job, self.dispatch(E.CustomerRfqReceived(job.job_id))

    def mark_rfq_sent(self, job_id: str, rfq_id: str, conversation_id: str,
                      sent_at: Optional[datetime] = None,
                      sent_message_id: Optional[str] = None,
                      subject: Optional[str] = None) -> None:
        """Manual override: record a send and snapshot the deadline (Fork D). The
        primary path is capture_sent(); this stays for cases where conversation
        matching can't place a sent item. The adapter may read the clock here --
        only the engine is forbidden to."""
        sent = sent_at or datetime.now(timezone.utc)
        deadline = sent + timedelta(minutes=self.config.default_timeout_minutes)
        self.store.mark_rfq_sent(job_id, rfq_id, conversation_id,
                                 sent_at=sent.isoformat(),
                                 deadline_at=deadline.isoformat(),
                                 sent_message_id=sent_message_id, subject=subject)

    def capture_sent(self, sent_mail: SentMail) -> Optional[Tuple[str, str]]:
        """Detect that a drafted RFQ was sent and capture it: match by
        conversation_id, then record the sent Message-ID + subject and snapshot
        the deadline = sent_on + timeout. Idempotent -- a re-scan of an
        already-captured RFQ is a no-op. Returns (job_id, rfq_id) on capture."""
        m = match_sent(sent_mail, self.store)
        if m is None:
            return None
        job_id, rfq_id = m
        rfq = self.store.get_job(job_id).vendor_rfqs[rfq_id]
        if rfq.state is not RFQState.DRAFT:
            return None                                   # already captured -> idempotent
        sent = sent_mail.sent_on or datetime.now(timezone.utc)
        deadline = sent + timedelta(minutes=self.config.default_timeout_minutes)
        self.store.mark_rfq_sent(job_id, rfq_id, sent_mail.conversation_id,
                                 sent_at=sent.isoformat(),
                                 deadline_at=deadline.isoformat(),
                                 sent_message_id=sent_mail.message_id,
                                 subject=sent_mail.subject)
        return (job_id, rfq_id)

    def ingest_vendor_reply(self, job_id: str, rfq_id: str,
                            replies: List[ItemReply]) -> List:
        """Write the parsed reply facts, then react. (In step 5 the parser calls
        this; in step 3 you pass hand-built ItemReply objects.)"""
        with self.store.batch():
            for r in replies:
                self.store.record_item_response(
                    job_id, rfq_id, r.item_id, r.status,
                    price=r.price, lead_time=r.lead_time,
                    status_confidence=r.status_confidence,
                    status_source=StatusSource.AI,
                    raw_excerpt=r.raw_excerpt,
                    response_entry_id=r.response_entry_id)
        return self.dispatch(
            E.VendorReplyParsed(job_id, rfq_id, [r.item_id for r in replies]))

    def select_vendor(self, job_id: str, vendor: Vendor, item_ids: List[str]) -> List:
        return self.dispatch(E.VendorSelected(job_id, vendor, list(item_ids)))

    def mark_item_verified(self, job_id: str, item_id: str) -> List:
        self.store.set_line_item_verified(job_id, item_id, True)
        return self.dispatch(E.UserVerifiedItem(job_id, item_id))

    def mark_item_resolved(self, job_id: str, item_id: str) -> List:
        self.store.set_line_item_resolved(job_id, item_id, True)
        return self.dispatch(E.ItemResolved(job_id, item_id))

    def tick(self, now: Optional[datetime] = None) -> List:
        return self.dispatch(E.Tick(now or datetime.now(timezone.utc)))

    # ---- the core: plan -> effects -> write-back ---------------------------
    def dispatch(self, event) -> List:
        """Single-step. Plans, executes, returns follow-on events to dispatch
        next (does NOT recurse). Per-intention failures are isolated."""
        if isinstance(event, E.Tick):
            jobs = self.store.open_jobs()              # ticks fan out over all open jobs
        else:
            job = self.store.get_job(getattr(event, "job_id", None))
            jobs = [job] if job is not None else []

        follow_ons: List = []
        for job in jobs:
            for intent in E.plan(job, event, self.config.planner):
                try:
                    follow_ons.extend(self._execute(job, intent))
                except Exception:
                    # log + isolate: guard write was never reached, so it retries;
                    # the remaining intentions still run.
                    log.exception("intention failed, isolated: %r", intent)
        return follow_ons

    def dispatch_until_settled(self, event, max_rounds: int = 50) -> None:
        """Optional convenience: drain follow-ons in one call. Off the main path
        by design -- single-step dispatch() is the predictable default."""
        queue = [event]
        rounds = 0
        while queue and rounds < max_rounds:
            rounds += 1
            queue.extend(self.dispatch(queue.pop(0)))
        if queue:
            log.warning("dispatch_until_settled hit max_rounds=%d; %d events left",
                        max_rounds, len(queue))

    # ---- one executor per intention ----------------------------------------
    def _execute(self, job, intent) -> List:
        jid = job.job_id

        if isinstance(intent, E.DraftAck):
            entry_id = self.port.draft_ack(job)                 # effect (no lock)
            with self.store.batch():                            # then persist
                self.store.set_drafts(jid, ack_entry_id=entry_id)
            return []

        if isinstance(intent, E.RunExtraction):
            items = self.port.run_extraction(job)
            with self.store.batch():
                for it in items:
                    self.store.add_line_item(
                        jid, it.description, canonical_key=it.canonical_key,
                        qty=it.qty, uom=it.uom,
                        extraction_confidence=it.extraction_confidence)
            return [E.ExtractionCompleted(jid)]                 # explicit follow-on

        if isinstance(intent, E.RequestItemReview):
            self.port.request_item_review(job, intent.item_ids)
            return []

        if isinstance(intent, E.RequestVendorSelection):
            self.port.request_vendor_selection(job, intent.item_ids)
            return []

        if isinstance(intent, E.DraftVendorRfq):
            res = self.port.draft_vendor_rfq(job, intent.vendor, intent.item_ids)
            with self.store.batch():                            # record + link together
                rfq = self.store.add_vendor_rfq(jid, intent.vendor, intent.item_ids)
                # capture conversation_id NOW so send-detection can match later
                self.store.set_rfq_draft(jid, rfq.rfq_id, res.entry_id,
                                         conversation_id=res.conversation_id)
            return []

        if isinstance(intent, E.FlagItemQuoted):
            self.port.flag_item_quoted(job, intent.item_id)
            return []

        if isinstance(intent, E.NotifyResponseNeeded):
            self.port.notify_response_needed(job, intent.rfq_id, intent.item_ids)
            return []

        if isinstance(intent, E.AlertOverdue):
            self.port.alert_overdue(job, intent.rfq_id)         # effect first
            with self.store.batch():                            # guard AFTER success
                self.store.set_rfq_overdue_alerted(jid, intent.rfq_id, True)
            return []

        if isinstance(intent, E.SuggestCustomerQuote):
            self.port.suggest_customer_quote(job)
            return []

        log.warning("no executor for intention %r", intent)
        return []
