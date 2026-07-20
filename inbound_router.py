"""
MaINbox inbound ROUTER -- the triage layer (step 4 effects edge).

It resolves a raw InboundMail into features, runs the pure classifier, then either
auto-acts (only for hard-key vendor replies / customer follow-ups) or queues for
your confirmation -- and records every mail in the dedup ledger. Minting always
goes through the step-3 controller; this layer never writes quote-job facts itself
and never touches MaINbox (contacts come through a narrow read-only lookup).

handle() has NO consequential side effects: it reads, queues, and records. The
only thing that mints a job (and thus drafts mail to a person) is your explicit
confirm_new_customer() -- the false-positive-averse spine in code.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Protocol, Tuple

from quote_jobs_classifier import (
    Bucket, ClassifierConfig, MailFeatures, Routing, classify,
)
from quote_jobs_controller import QuoteJobController
from quote_jobs_model import Customer
from quote_jobs_store import normalize_subject
from inbound_store import InboundStore


def _candidate_msgids(*headers) -> list:
    """Extract <message-id> tokens from In-Reply-To / References header values."""
    out = []
    for h in headers:
        if not h:
            continue
        token = ""
        for ch in h:
            if ch == "<":
                token = "<"
            elif ch == ">":
                if token:
                    out.append(token + ">")
                    token = ""
            elif token:
                token += ch
        # also accept a bare id with no angle brackets
        if not out and h.strip():
            out.append(h.strip())
    return out


# ==========================================================================
# inbound value object
# ==========================================================================
@dataclass
class InboundMail:
    message_id: str                       # RFC Message-ID (preferred) or entry_id; the DEDUP key
    sender: str                           # sender email address
    subject: str = ""
    body: str = ""
    conversation_id: Optional[str] = None # Outlook ConversationID -- the primary hard key
    in_reply_to: Optional[str] = None     # RFC header (backup key; resolution deferred)
    references: Optional[str] = None      # RFC header (backup key; resolution deferred)
    received_at: Optional[str] = None
    entry_id: Optional[str] = None        # Outlook EntryID
    has_attachments: bool = False

    @property
    def domain(self) -> str:
        return self.sender.split("@")[-1].lower() if "@" in self.sender else ""


# ==========================================================================
# ports the adapter provides
# ==========================================================================
class ContactLookup(Protocol):
    """Narrow, read-only window onto MaINbox's contacts / categorization."""
    def is_known_customer(self, email: str, domain: str) -> bool: ...
    def is_known_vendor(self, email: str, domain: str) -> bool: ...


class RfqScorer(Protocol):
    """Returns 0..1: how strongly this mail reads as an RFQ."""
    def score(self, mail: InboundMail) -> float: ...


class KeywordRfqScorer:
    """Deterministic scorer for tests and the offline path. Strong RFQ phrasing
    dominates; medium cues and attachments nudge. Tune the cue lists / threshold
    against your real mail."""
    STRONG = ("request for quote", "rfq", "request a quote", "quote request",
              "please quote", "please provide a quote", "requesting a quote",
              "request for quotation")
    MEDIUM = ("price and availability", "pricing", "availability", "lead time",
              "lead-time", "quotation", "quote on", "can you quote", "need a price",
              "need pricing", "quote for", "provide pricing")

    def score(self, mail: InboundMail) -> float:
        text = f"{mail.subject}\n{mail.body}".lower()
        s = 0.0
        if any(k in text for k in self.STRONG):
            s += 0.7
        med = sum(1 for k in self.MEDIUM if k in text)
        s += min(0.3, 0.1 * med)
        if mail.has_attachments:
            s += 0.1
        return min(1.0, s)


class SkeletonClaudeRfqScorer:
    """Copy this for the real path: call your SmartScan / Claude RFQ classifier and
    return a 0..1 confidence. Keeping it OUT of the classifier keeps routing
    deterministic and testable."""
    def score(self, mail: InboundMail) -> float:
        # TODO: send mail.subject + mail.body (and attachment hints) to your Claude
        # RFQ-recognition prompt; map its verdict to a 0..1 score.
        raise NotImplementedError("wire SkeletonClaudeRfqScorer to your Claude/SmartScan classifier")


class SkeletonContactLookup:
    """Copy this for the real path: answer from MaINbox's existing contacts /
    customer-vendor categorization."""
    def is_known_customer(self, email: str, domain: str) -> bool:
        raise NotImplementedError("wire to MaINbox customer contacts")

    def is_known_vendor(self, email: str, domain: str) -> bool:
        raise NotImplementedError("wire to MaINbox vendor contacts")


# ==========================================================================
# feature resolution (effectful: reads the quote store, contacts, scorer)
# ==========================================================================
def resolve_features(mail: InboundMail, quote_store, lookup: ContactLookup,
                     scorer: RfqScorer) -> MailFeatures:
    vendor_match: Optional[Tuple[str, str]] = None
    customer_match: Optional[str] = None

    # Fork C signal 1 -- the primary hard key, fully supported today.
    if mail.conversation_id:
        hit = quote_store.find_rfq_by_conversation(mail.conversation_id)
        if hit is not None:
            job, rfq = hit
            vendor_match = (job.job_id, rfq.rfq_id)
        else:
            job = quote_store.find_job_by_customer_conversation(mail.conversation_id)
            if job is not None:
                customer_match = job.job_id

    # Fork C signal 2 -- RFC headers vs the sent Message-ID (HARD key), used as a
    # backup when Outlook ConversationID threading didn't match above.
    if vendor_match is None and customer_match is None:
        for mid in _candidate_msgids(mail.in_reply_to, mail.references):
            hit = quote_store.find_rfq_by_message_id(mid)
            if hit is not None:
                job, rfq = hit
                vendor_match = (job.job_id, rfq.rfq_id)
                break

    # Fork C signal 3 -- normalized subject (SOFT): only PROPOSE into triage,
    # never auto-route. Suggest only when exactly one RFQ matches the subject.
    subject_suggested: Optional[Tuple[str, str]] = None
    if vendor_match is None and customer_match is None:
        cands = quote_store.find_rfqs_by_subject(normalize_subject(mail.subject))
        if len(cands) == 1:
            job, rfq = cands[0]
            subject_suggested = (job.job_id, rfq.rfq_id)

    return MailFeatures(
        sender_is_known_customer=lookup.is_known_customer(mail.sender, mail.domain),
        sender_is_known_vendor=lookup.is_known_vendor(mail.sender, mail.domain),
        vendor_thread_match=vendor_match,
        customer_thread_match=customer_match,
        subject_suggested_match=subject_suggested,
        rfq_signal=scorer.score(mail),
    )


# ==========================================================================
# router
# ==========================================================================
@dataclass
class RoutingResult:
    routing: Routing
    action: str                  # what handle() actually did
    already_processed: bool = False


class InboundRouter:
    def __init__(self, controller: QuoteJobController, inbound: InboundStore,
                 lookup: ContactLookup, scorer: RfqScorer,
                 config: ClassifierConfig = ClassifierConfig()):
        self.controller = controller
        self.inbound = inbound
        self.lookup = lookup
        self.scorer = scorer
        self.config = config

    # ---- the entry point: route one mail (no minting, no parsing) ----------
    def handle(self, mail: InboundMail) -> RoutingResult:
        # dedup first -- a re-scan must be a no-op.
        if self.inbound.is_processed(mail.message_id):
            prior = self.inbound.ledger_entry(mail.message_id) or {}
            return RoutingResult(
                Routing(Bucket(prior.get("bucket", Bucket.IGNORE.value)),
                        job_id=prior.get("job_id"), rfq_id=prior.get("rfq_id"),
                        reason="already processed"),
                action="skipped (already processed)", already_processed=True)

        features = resolve_features(mail, self.controller.store, self.lookup, self.scorer)
        routing = classify(features, self.config)

        action = "recorded"
        if routing.bucket is Bucket.VENDOR_REPLY:
            # auto, but only a HANDOFF -- step 4 does not read prices.
            self.inbound.add_handoff({
                "message_id": mail.message_id, "job_id": routing.job_id,
                "rfq_id": routing.rfq_id, "sender": mail.sender,
                "subject": mail.subject, "entry_id": mail.entry_id,
                "received_at": mail.received_at})
            action = f"handed off vendor reply -> {routing.job_id}/{routing.rfq_id} (awaiting parse)"

        elif routing.bucket is Bucket.CUSTOMER_FOLLOWUP:
            action = f"noted customer follow-up on {routing.job_id}"

        elif routing.bucket is Bucket.NEW_CUSTOMER_RFQ:
            self.inbound.add_pending({
                "message_id": mail.message_id, "sender": mail.sender,
                "domain": mail.domain, "subject": mail.subject,
                "conversation_id": mail.conversation_id, "entry_id": mail.entry_id,
                "received_at": mail.received_at, "reason": routing.reason})
            action = "queued for confirm-to-mint"

        elif routing.bucket is Bucket.NEEDS_TRIAGE:
            self.inbound.add_triage({
                "message_id": mail.message_id, "sender": mail.sender,
                "subject": mail.subject, "conversation_id": mail.conversation_id,
                "entry_id": mail.entry_id, "received_at": mail.received_at,
                "suggested_job_id": routing.job_id, "suggested_rfq_id": routing.rfq_id,
                "reason": routing.reason})
            action = "queued for triage"

        else:  # IGNORE
            action = "ignored"

        self.inbound.record_ledger(mail.message_id, routing.bucket.value,
                                   job_id=routing.job_id, rfq_id=routing.rfq_id)
        self.inbound.save()
        return RoutingResult(routing, action)

    # ---- human decisions ---------------------------------------------------
    def confirm_new_customer(self, message_id: str, name: Optional[str] = None):
        """Mint the job (THIS is the consequential action -- drafts ack + extracts),
        threading the original mail's conversation_id so future replies attach as
        follow-ups instead of duplicating. Returns (job, follow_on_events)."""
        entry = (self.inbound.pop_from("pending", message_id)
                 or self.inbound.pop_from("triage", message_id))
        if entry is None:
            raise KeyError(f"no pending/triage mail {message_id!r} to confirm")
        customer = Customer(
            name=name or entry.get("sender", ""),
            email=entry.get("sender", ""),
            conversation_id=entry.get("conversation_id"),
            source_entry_id=entry.get("entry_id"))
        job, follow = self.controller.ingest_customer_rfq(customer)
        self.inbound.record_ledger(message_id, Bucket.NEW_CUSTOMER_RFQ.value, job_id=job.job_id)
        self.inbound.save()
        return job, follow

    def assign_vendor_reply(self, message_id: str, job_id: str, rfq_id: str) -> None:
        """Resolve a triage item you've identified as a vendor reply -> handoff."""
        entry = self.inbound.pop_from("triage", message_id) or {}
        self.inbound.add_handoff({
            "message_id": message_id, "job_id": job_id, "rfq_id": rfq_id,
            "sender": entry.get("sender"), "subject": entry.get("subject"),
            "entry_id": entry.get("entry_id"), "received_at": entry.get("received_at")})
        self.inbound.record_ledger(message_id, Bucket.VENDOR_REPLY.value,
                                   job_id=job_id, rfq_id=rfq_id)
        self.inbound.save()

    def dismiss(self, message_id: str) -> None:
        """Reject a pending/triage mail -> mark ignored."""
        self.inbound.pop_from("pending", message_id)
        self.inbound.pop_from("triage", message_id)
        self.inbound.record_ledger(message_id, Bucket.IGNORE.value)
        self.inbound.save()
