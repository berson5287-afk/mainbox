"""
MaINbox Quote-Job data model (schema v1).

Single source of truth
-----------------------
Every per-(item, vendor) response fact lives in exactly ONE place:
    QuoteJob.vendor_rfqs[rfq_id].requests[item_id]   (an ItemRequest)

Item-level status and job-level status are DERIVED from those facts by the
step-2 state machine and are deliberately NOT stored anywhere. There is no
second copy that can drift -- which is what kept biting the Sent/Waiting
trackers (the v3.8.29 resurrection / subject-normalization class of bug).

Scope of this module
--------------------
Typed records + their JSON (de)serialization. No I/O, no Outlook, no policy.
The store module owns persistence. The state machine owns derivation.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, Optional

SCHEMA_VERSION = 1


# --------------------------------------------------------------------------
# enumerations  (str-valued so they serialize straight to plain JSON strings)
# --------------------------------------------------------------------------
class JobLifecycle(str, Enum):
    OPEN = "open"
    CLOSED = "closed"
    ABANDONED = "abandoned"


class RFQState(str, Enum):
    # "responded" is intentionally NOT a stored state -- it is derived from
    # whether any request under this RFQ has left AWAITING. Only these persist.
    DRAFT = "draft"
    SENT = "sent"
    CANCELLED = "cancelled"


class RequestStatus(str, Enum):
    AWAITING = "awaiting"   # no usable reply yet for this item from this vendor
    PRICED = "priced"       # price + availability present -> drives QUOTED
    PARTIAL = "partial"     # some info, not a complete answer
    QUESTION = "question"   # vendor asked something back
    DECLINED = "declined"   # no-bid / discontinued / can't supply
    CANCELLED = "cancelled" # we withdrew this ask


class StatusSource(str, Enum):
    AI = "ai"
    USER = "user"


# --------------------------------------------------------------------------
# leaf records
# --------------------------------------------------------------------------
@dataclass
class Vendor:
    name: str = ""
    email: str = ""
    vendor_id: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {"name": self.name, "email": self.email, "vendor_id": self.vendor_id}

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Vendor":
        return cls(name=d.get("name", ""), email=d.get("email", ""),
                   vendor_id=d.get("vendor_id"))


@dataclass
class Customer:
    name: str = ""
    email: str = ""
    conversation_id: Optional[str] = None   # source RFQ thread
    source_entry_id: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {"name": self.name, "email": self.email,
                "conversation_id": self.conversation_id,
                "source_entry_id": self.source_entry_id}

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Customer":
        return cls(name=d.get("name", ""), email=d.get("email", ""),
                   conversation_id=d.get("conversation_id"),
                   source_entry_id=d.get("source_entry_id"))


@dataclass
class Drafts:
    ack_entry_id: Optional[str] = None     # the "I will advise" draft
    quote_entry_id: Optional[str] = None   # customer-facing quote (later phase)

    def to_dict(self) -> Dict[str, Any]:
        return {"ack_entry_id": self.ack_entry_id,
                "quote_entry_id": self.quote_entry_id}

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Drafts":
        return cls(ack_entry_id=d.get("ack_entry_id"),
                   quote_entry_id=d.get("quote_entry_id"))


@dataclass
class LineItem:
    item_id: str
    description: str = ""              # extracted material text
    canonical_key: Optional[str] = None   # SmartScan dedup key
    qty: float = 1
    uom: str = "EA"
    extraction_confidence: float = 0.0    # SmartScan: how sure WHAT it is
    user_verified: bool = False
    resolved: bool = False            # user lock: "done shopping this item"
    # NB: no status field -- item status is derived across all RFQ requests.

    def to_dict(self) -> Dict[str, Any]:
        return {
            "item_id": self.item_id,
            "description": self.description,
            "canonical_key": self.canonical_key,
            "qty": self.qty,
            "uom": self.uom,
            "extraction_confidence": self.extraction_confidence,
            "user_verified": self.user_verified,
            "resolved": self.resolved,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "LineItem":
        return cls(
            item_id=d["item_id"],
            description=d.get("description", ""),
            canonical_key=d.get("canonical_key"),
            qty=d.get("qty", 1),
            uom=d.get("uom", "EA"),
            extraction_confidence=d.get("extraction_confidence", 0.0),
            user_verified=d.get("user_verified", False),
            resolved=d.get("resolved", False),
        )


@dataclass
class ItemRequest:
    """The ONE stored response fact for an (item, vendor) pair."""
    status: RequestStatus = RequestStatus.AWAITING
    price: Optional[float] = None
    lead_time: Optional[str] = None          # availability / lead time, free text
    status_confidence: Optional[float] = None  # AI's confidence in the call
    status_source: StatusSource = StatusSource.AI
    ai_proposed_status: Optional[RequestStatus] = None  # pre-correction (learning DB)
    raw_excerpt: Optional[str] = None        # reply text the call was based on
    responded_at: Optional[str] = None
    response_entry_id: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status.value,
            "price": self.price,
            "lead_time": self.lead_time,
            "status_confidence": self.status_confidence,
            "status_source": self.status_source.value,
            "ai_proposed_status": (self.ai_proposed_status.value
                                   if self.ai_proposed_status is not None else None),
            "raw_excerpt": self.raw_excerpt,
            "responded_at": self.responded_at,
            "response_entry_id": self.response_entry_id,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "ItemRequest":
        aps = d.get("ai_proposed_status")
        return cls(
            status=RequestStatus(d.get("status", RequestStatus.AWAITING.value)),
            price=d.get("price"),
            lead_time=d.get("lead_time"),
            status_confidence=d.get("status_confidence"),
            status_source=StatusSource(d.get("status_source", StatusSource.AI.value)),
            ai_proposed_status=RequestStatus(aps) if aps is not None else None,
            raw_excerpt=d.get("raw_excerpt"),
            responded_at=d.get("responded_at"),
            response_entry_id=d.get("response_entry_id"),
        )


@dataclass
class VendorRFQ:
    rfq_id: str
    vendor: Vendor = field(default_factory=Vendor)
    state: RFQState = RFQState.DRAFT
    draft_entry_id: Optional[str] = None
    conversation_id: Optional[str] = None   # captured at DRAFT time -> reply/send-match key
    sent_at: Optional[str] = None
    deadline_at: Optional[str] = None        # snapshot: sent_at + timeout@send
    overdue_alerted: bool = False            # alert-once guard
    sent_message_id: Optional[str] = None    # RFC Message-ID of the sent mail (header matching)
    subject_norm: Optional[str] = None       # normalized sent subject (subject matching)
    # per-item asks inside THIS rfq; an item_id may also appear under other rfqs
    requests: Dict[str, ItemRequest] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "rfq_id": self.rfq_id,
            "vendor": self.vendor.to_dict(),
            "state": self.state.value,
            "draft_entry_id": self.draft_entry_id,
            "conversation_id": self.conversation_id,
            "sent_at": self.sent_at,
            "deadline_at": self.deadline_at,
            "overdue_alerted": self.overdue_alerted,
            "sent_message_id": self.sent_message_id,
            "subject_norm": self.subject_norm,
            "requests": {k: v.to_dict() for k, v in self.requests.items()},
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "VendorRFQ":
        return cls(
            rfq_id=d["rfq_id"],
            vendor=Vendor.from_dict(d.get("vendor", {})),
            state=RFQState(d.get("state", RFQState.DRAFT.value)),
            draft_entry_id=d.get("draft_entry_id"),
            conversation_id=d.get("conversation_id"),
            sent_at=d.get("sent_at"),
            deadline_at=d.get("deadline_at"),
            overdue_alerted=d.get("overdue_alerted", False),
            sent_message_id=d.get("sent_message_id"),
            subject_norm=d.get("subject_norm"),
            requests={k: ItemRequest.from_dict(v)
                      for k, v in d.get("requests", {}).items()},
        )


@dataclass
class QuoteJob:
    job_id: str
    lifecycle: JobLifecycle = JobLifecycle.OPEN   # user-controlled, stored
    created_at: str = ""
    updated_at: str = ""
    customer: Customer = field(default_factory=Customer)
    drafts: Drafts = field(default_factory=Drafts)
    line_items: Dict[str, LineItem] = field(default_factory=dict)
    vendor_rfqs: Dict[str, VendorRFQ] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "job_id": self.job_id,
            "lifecycle": self.lifecycle.value,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "customer": self.customer.to_dict(),
            "drafts": self.drafts.to_dict(),
            "line_items": {k: v.to_dict() for k, v in self.line_items.items()},
            "vendor_rfqs": {k: v.to_dict() for k, v in self.vendor_rfqs.items()},
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "QuoteJob":
        return cls(
            job_id=d["job_id"],
            lifecycle=JobLifecycle(d.get("lifecycle", JobLifecycle.OPEN.value)),
            created_at=d.get("created_at", ""),
            updated_at=d.get("updated_at", ""),
            customer=Customer.from_dict(d.get("customer", {})),
            drafts=Drafts.from_dict(d.get("drafts", {})),
            line_items={k: LineItem.from_dict(v)
                        for k, v in d.get("line_items", {}).items()},
            vendor_rfqs={k: VendorRFQ.from_dict(v)
                         for k, v in d.get("vendor_rfqs", {}).items()},
        )
