"""
Send-detection -- linking an actually-sent Outlook item back to the RFQ it came
from, so we can capture its Message-ID / subject and start the deadline clock.

The matcher is pure (reads the store only). It keys on conversation_id, captured
on the RFQ at DRAFT time, which stays stable from draft through send. If that
hard key doesn't match, it returns nothing -- it does NOT guess from recipient /
subject (same refuse-to-guess spine as the rest of the system). When detection
can't place a sent item, the RFQ simply stays draft and you can mark it sent
manually.

The controller's capture_sent() drives this: match -> mark sent with the captured
Message-ID + subject -> snapshot the deadline. Pulling SentMail records out of
Outlook's Sent Items (or an ItemSend event) is the adapter's job and stays
manually invoked for now -- the live hook is the deferred auto-trigger step.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional, Tuple

from quote_jobs_model import RFQState


@dataclass
class SentMail:
    """A sent Outlook item, as the adapter pulls it from Sent Items / ItemSend."""
    message_id: str                       # RFC Message-ID of the sent mail
    conversation_id: Optional[str] = None # Outlook ConversationID (the match key)
    subject: str = ""
    to: str = ""
    sent_on: Optional[datetime] = None    # Outlook SentOn; None -> caller uses 'now'
    entry_id: Optional[str] = None        # EntryID in Sent Items (may differ from the draft's)


def match_sent(sent_mail: SentMail, store) -> Optional[Tuple[str, str]]:
    """Return (job_id, rfq_id) of the DRAFT RFQ this sent item belongs to, or None.

    Hard key only: conversation_id. A SENT match is returned too (the caller treats
    it as 'already captured' and skips, keeping capture idempotent)."""
    if not sent_mail.conversation_id:
        return None
    hit = store.find_rfq_by_conversation(sent_mail.conversation_id)
    if hit is None:
        return None
    job, rfq = hit
    return (job.job_id, rfq.rfq_id)


def is_draft(store, job_id: str, rfq_id: str) -> bool:
    job = store.get_job(job_id)
    if job is None:
        return False
    rfq = job.vendor_rfqs.get(rfq_id)
    return rfq is not None and rfq.state is RFQState.DRAFT
