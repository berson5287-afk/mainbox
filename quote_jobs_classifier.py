"""
MaINbox inbound CLASSIFIER -- the pure routing brain (step 4).

The whole front door is built around one asymmetry: a false positive (mis-routing
a stranger's email into the pipeline) is far costlier than a false negative
(missing one, which just means you handle it by hand, like today). So this module
auto-acts ONLY on hard mechanical signals and refuses to guess otherwise.

Pure: imports stdlib only. It takes a resolved MailFeatures and returns a Routing.
The effectful feature resolution (querying the quote-job store, contacts, and the
RFQ scorer) lives in the router; classify() never touches any of that, which is
what keeps the routing table unit-testable with fabricated features.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Optional, Tuple


class Bucket(str, Enum):
    NEW_CUSTOMER_RFQ = "new_customer_rfq"     # confirm before minting (drafts to a person)
    CUSTOMER_FOLLOWUP = "customer_followup"   # reply on an existing job's thread; v1: just notify
    VENDOR_REPLY = "vendor_reply"             # reply to a tracked RFQ; hand off for parsing
    IGNORE = "ignore"                         # ordinary traffic; never touches anything
    NEEDS_TRIAGE = "needs_triage"             # relevant but not confidently placeable


@dataclass
class ClassifierConfig:
    rfq_threshold: float = 0.6                # rfq_signal at/above this is "strong RFQ content"


@dataclass
class MailFeatures:
    """Everything classify() is allowed to see, all pre-resolved to plain data."""
    sender_is_known_customer: bool = False
    sender_is_known_vendor: bool = False
    # hard thread matches (mechanical, not a guess) -> may auto-route
    vendor_thread_match: Optional[Tuple[str, str]] = None      # (job_id, rfq_id)
    customer_thread_match: Optional[str] = None                # job_id
    # soft subject-only match -> may only PROPOSE into triage, never auto-route
    subject_suggested_match: Optional[Tuple[str, str]] = None  # (job_id, rfq_id)
    rfq_signal: float = 0.0                                    # 0..1, "is this an RFQ?"


@dataclass
class Routing:
    bucket: Bucket
    job_id: Optional[str] = None
    rfq_id: Optional[str] = None
    auto: bool = False        # may the router act WITHOUT a human confirm?
    reason: str = ""


def classify(features: MailFeatures, config: ClassifierConfig = ClassifierConfig()) -> Routing:
    f = features

    # 1. hard thread match to a tracked RFQ -> vendor reply, auto-route (Fork C signal 1).
    if f.vendor_thread_match is not None:
        jid, rid = f.vendor_thread_match
        return Routing(Bucket.VENDOR_REPLY, job_id=jid, rfq_id=rid, auto=True,
                       reason="matched a tracked RFQ thread (hard key)")

    # 2. hard match to an existing job's customer thread -> follow-up, auto (just notify in v1).
    if f.customer_thread_match is not None:
        return Routing(Bucket.CUSTOMER_FOLLOWUP, job_id=f.customer_thread_match, auto=True,
                       reason="matched an existing job's customer thread (hard key)")

    # 3. subject-only suggestion -> NEVER auto-route; propose into triage (Fork C guarantee).
    if f.subject_suggested_match is not None:
        jid, rid = f.subject_suggested_match
        return Routing(Bucket.NEEDS_TRIAGE, job_id=jid, rfq_id=rid, auto=False,
                       reason="subject-only suggested match; confirm before routing")

    # 4. no thread match -> relevance gate (Fork D). Default is IGNORE; act/confirm only on signal.
    strong = f.rfq_signal >= config.rfq_threshold

    if f.sender_is_known_customer and strong:
        return Routing(Bucket.NEW_CUSTOMER_RFQ, auto=False,
                       reason="known customer + RFQ content; confirm to mint")

    if f.sender_is_known_vendor and strong:
        # vendor sending RFQ-like content we can't attach to a known RFQ -> don't guess.
        return Routing(Bucket.NEEDS_TRIAGE, auto=False,
                       reason="vendor mail with RFQ content but no thread match; confirm")

    if (not f.sender_is_known_customer) and (not f.sender_is_known_vendor) and strong:
        # the new-customer case: relevant, but unknown -> flag, never auto-mint.
        return Routing(Bucket.NEEDS_TRIAGE, auto=False,
                       reason="unknown sender + RFQ content (possible new customer); confirm")

    # known contact with weak signal, or unknown with weak signal -> ordinary traffic.
    return Routing(Bucket.IGNORE, auto=False,
                   reason="no thread match and no strong RFQ signal")
