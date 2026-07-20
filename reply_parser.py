"""
MaINbox vendor-reply PARSER (step 5) -- turns a matched handoff into per-item facts.

Step 4 hands off "this mail is a reply to job X / RFQ Y". This layer reads the
reply body against that RFQ's items and proposes a status per item (priced /
partial / question / declined), then -- per your ParsePolicy -- either records
the facts automatically or queues them for your review.

Shape mirrors the rest of the system:
  * The LLM read is an injected ReplyParser PORT (SmartScan/Claude in the real
    path); a deterministic KeywordReplyParser drives tests and the offline path.
  * The orchestrator (ReplyParseController) is the effects edge -- it records via
    the step-3 controller (ingest_vendor_reply / correct_item_response) and never
    re-implements engine logic. Recording an RFQ reply still fires the engine
    (flag-quoted / response-needed) through the existing dispatch.
  * Idempotent: a handoff is consumed once (popped); re-parsing is a no-op.
  * Learning pairs reuse step-1's ai_proposed_status -- no new store.

Manually invoked: the adapter fetches the reply body from Outlook by entry_id and
passes it in. Firing this off live mail is the deferred auto-trigger step.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Protocol

from quote_jobs_model import RequestStatus, StatusSource
from quote_jobs_controller import QuoteJobController, ItemReply
from inbound_store import InboundStore
import quote_jobs_engine as E


# ==========================================================================
# proposals + policy
# ==========================================================================
@dataclass
class ItemProposal:
    item_id: str
    status: RequestStatus
    price: Optional[float] = None
    lead_time: Optional[str] = None
    confidence: float = 0.0
    excerpt: str = ""


@dataclass
class ParsePolicy:
    # Your Fork-B settings. auto_record=False -> review-first (queue everything,
    # record nothing until you confirm). auto_record=True -> record all proposals,
    # and also surface any below notify_threshold for verification.
    auto_record: bool = True
    notify_threshold: float = 0.75


@dataclass
class ParseResult:
    job_id: Optional[str] = None
    rfq_id: Optional[str] = None
    recorded: List[str] = field(default_factory=list)   # item_ids auto-recorded
    review: List[str] = field(default_factory=list)     # item_ids queued/flagged
    skipped: bool = False                                # handoff missing / already parsed


# ==========================================================================
# the parser port
# ==========================================================================
class ReplyParser(Protocol):
    """Reads a reply body against the RFQ's items and proposes a status per item.
    `items` is a list of dicts: {item_id, description, canonical_key, qty, uom}."""
    def parse(self, items: List[dict], reply_body: str) -> List[ItemProposal]: ...


class SkeletonClaudeReplyParser:
    """Copy this for the real path: send the reply body + item list (and the
    learning pairs from extract_learning_pairs) to your SmartScan/Claude prompt,
    and map its per-item verdicts to ItemProposal objects with confidences."""
    def parse(self, items, reply_body) -> List[ItemProposal]:
        # TODO: call your Claude reply-parse prompt; return one ItemProposal per
        # item the reply addresses (omit items it doesn't mention -> they stay AWAITING).
        raise NotImplementedError("wire SkeletonClaudeReplyParser to your Claude reply parser")


_PRICE = re.compile(r"\$\s?(\d+(?:[.,]\d+)?)|(\d+(?:\.\d+)?)\s*/\s*(?:ft|ea|each|m|lb|c)\b")
_LEAD = re.compile(r"(in stock|\d+\s*(?:-\s*\d+\s*)?(?:day|week|business day)s?)", re.I)
_DECLINE = ("discontinued", "no longer", "cannot supply", "can't supply",
            "no bid", "unable to", "not available", "we don't carry", "obsolete")


class KeywordReplyParser:
    """Deterministic parser for tests and the offline path. The genuinely hard
    'which of my items is the vendor talking about' matching is the real parser's
    job; this stub attributes the whole body to the lone item of a single-item
    RFQ, and otherwise looks for an item's description words / canonical key."""

    def parse(self, items: List[dict], reply_body: str) -> List[ItemProposal]:
        out: List[ItemProposal] = []
        for it in items:
            chunk = self._chunk_for(it, reply_body, single=(len(items) == 1))
            if chunk is None:
                continue                       # no statement -> stays AWAITING
            prop = self._classify(it["item_id"], chunk)
            if prop is not None:
                out.append(prop)
        return out

    def _chunk_for(self, item, body, single):
        if single:
            return body
        words = [w for w in re.split(r"\W+", (item.get("description") or "").lower()) if len(w) > 3]
        ckey = (item.get("canonical_key") or "").lower()
        for line in re.split(r"[\n\.;]", body):
            low = line.lower()
            if (ckey and ckey in low) or any(w in low for w in words):
                return line
        return None

    def _classify(self, item_id, chunk) -> Optional[ItemProposal]:
        low = chunk.lower()
        m = _PRICE.search(chunk)
        if m:
            raw = (m.group(1) or m.group(2) or "").replace(",", "")
            price = float(raw) if raw else None
            lead = _LEAD.search(chunk)
            return ItemProposal(item_id, RequestStatus.PRICED, price=price,
                                lead_time=(lead.group(1) if lead else None),
                                confidence=0.9, excerpt=chunk.strip())
        if any(k in low for k in _DECLINE):
            return ItemProposal(item_id, RequestStatus.DECLINED, confidence=0.85, excerpt=chunk.strip())
        if "?" in chunk:
            return ItemProposal(item_id, RequestStatus.QUESTION, confidence=0.7, excerpt=chunk.strip())
        if _LEAD.search(chunk):
            lead = _LEAD.search(chunk)
            return ItemProposal(item_id, RequestStatus.PARTIAL, lead_time=lead.group(1),
                                confidence=0.6, excerpt=chunk.strip())
        return None


# ==========================================================================
# the orchestrator (effects edge)
# ==========================================================================
class ReplyParseController:
    def __init__(self, controller: QuoteJobController, inbound: InboundStore,
                 parser: ReplyParser, policy: ParsePolicy = ParsePolicy()):
        self.controller = controller
        self.inbound = inbound
        self.parser = parser
        self.policy = policy

    def parse_handoff(self, message_id: str, reply_body: str) -> ParseResult:
        handoff = self.inbound.pop_handoff(message_id)      # consume once -> idempotent
        if handoff is None:
            return ParseResult(skipped=True)
        job_id, rfq_id = handoff["job_id"], handoff["rfq_id"]
        job = self.controller.store.get_job(job_id)
        rfq = job.vendor_rfqs[rfq_id]
        context = []
        for iid in rfq.requests:
            li = job.line_items.get(iid)
            context.append({"item_id": iid,
                            "description": li.description if li else "",
                            "canonical_key": li.canonical_key if li else None,
                            "qty": li.qty if li else None,
                            "uom": li.uom if li else None})

        proposals = self.parser.parse(context, reply_body)
        records: List[ItemReply] = []
        review: List[str] = []
        for p in proposals:
            if self.policy.auto_record:
                records.append(ItemReply(p.item_id, p.status, price=p.price,
                                         lead_time=p.lead_time, status_confidence=p.confidence,
                                         raw_excerpt=p.excerpt))
                if p.confidence < self.policy.notify_threshold:
                    self.inbound.add_parse_review(self._entry(message_id, job_id, rfq_id, p, recorded=True))
                    review.append(p.item_id)
            else:
                self.inbound.add_parse_review(self._entry(message_id, job_id, rfq_id, p, recorded=False))
                review.append(p.item_id)

        if records:
            # records facts (status_source=AI, ai_proposed_status seeded) AND fires
            # the engine (flag-quoted / response-needed) via the existing dispatch.
            self.controller.ingest_vendor_reply(job_id, rfq_id, records)
        self.inbound.save()
        return ParseResult(job_id=job_id, rfq_id=rfq_id,
                           recorded=[r.item_id for r in records], review=review)

    # ---- human actions on the parse-review queue ---------------------------
    def confirm_parse_item(self, message_id: str, item_id: str) -> bool:
        """Accept the AI's proposal. For a review-first item this records it now
        (and fires the engine); for an already-recorded item it just clears it."""
        entry = self.inbound.pop_parse_review(message_id, item_id)
        if entry is None:
            return False
        if not entry.get("recorded"):
            reply = ItemReply(item_id, RequestStatus(entry["status"]), price=entry.get("price"),
                              lead_time=entry.get("lead_time"),
                              status_confidence=entry.get("confidence"),
                              raw_excerpt=entry.get("excerpt"))
            self.controller.ingest_vendor_reply(entry["job_id"], entry["rfq_id"], [reply])
        self.inbound.save()
        return True

    def correct_parse_item(self, message_id: str, item_id: str, status: RequestStatus,
                           price: Optional[float] = None, lead_time: Optional[str] = None) -> bool:
        """Override with your values. Records as USER while preserving the AI's
        original guess (ai_proposed_status) for the learning DB, then fires the engine."""
        entry = self.inbound.pop_parse_review(message_id, item_id)
        if entry is None:
            return False
        self.controller.store.record_item_response(
            entry["job_id"], entry["rfq_id"], item_id, status, price=price, lead_time=lead_time,
            status_source=StatusSource.USER, ai_proposed_status=RequestStatus(entry["status"]),
            raw_excerpt=entry.get("excerpt"))
        self.controller.dispatch(E.VendorReplyParsed(entry["job_id"], entry["rfq_id"], [item_id]))
        self.inbound.save()
        return True

    def dismiss_parse_item(self, message_id: str, item_id: str) -> bool:
        ok = self.inbound.pop_parse_review(message_id, item_id) is not None
        self.inbound.save()
        return ok

    @staticmethod
    def _entry(message_id, job_id, rfq_id, p: ItemProposal, recorded: bool) -> dict:
        return {"message_id": message_id, "item_id": p.item_id, "job_id": job_id,
                "rfq_id": rfq_id, "status": p.status.value, "price": p.price,
                "lead_time": p.lead_time, "confidence": p.confidence,
                "excerpt": p.excerpt, "recorded": recorded}


# ==========================================================================
# learning pairs (reuses the model -- no separate store)
# ==========================================================================
def extract_learning_pairs(store) -> List[dict]:
    """Every place you corrected the AI: (excerpt, predicted, actual). These feed
    the real Claude parser's prompt. Built by scanning requests where you
    overrode the AI and the corrected status differs from its guess."""
    pairs = []
    for job in store.all_jobs():
        for rfq in job.vendor_rfqs.values():
            for item_id, req in rfq.requests.items():
                if (req.status_source is StatusSource.USER
                        and req.ai_proposed_status is not None
                        and req.ai_proposed_status != req.status):
                    pairs.append({
                        "excerpt": req.raw_excerpt,
                        "predicted": req.ai_proposed_status.value,
                        "actual": req.status.value,
                        "job_id": job.job_id, "rfq_id": rfq.rfq_id, "item_id": item_id,
                    })
    return pairs
