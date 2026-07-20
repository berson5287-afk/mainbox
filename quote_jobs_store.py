"""
Persistence substrate for MaINbox quote jobs (schema v1).

Does ONLY this:
  * Load / save quote_jobs.json with atomic writes (fsync + os.replace), in its
    OWN store -- not smeared across the existing flat trackers.
  * Depth-counted, exception-safe write-batching (mirrors MaINbox's existing
    _begin/_flush primitives) so a burst of mutations collapses to one write.
  * In-memory reverse indexes for O(1) reply matching, rebuilt on load and
    NEVER persisted (a persisted index is a second copy that can drift).
  * Thread-safe mutators that record FACTS only.

Deliberately NOT here:
  * Status derivation (item/job folds)        -> step-2 state machine.
  * Timeout *policy* (how long, snapshot calc) -> caller passes a precomputed
    deadline_at into mark_rfq_sent(). The store stores; it does not decide.
"""
from __future__ import annotations

import json
import logging
import os
import tempfile
import threading
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Dict, Iterator, List, Optional, Tuple

from quote_jobs_model import (
    SCHEMA_VERSION, JobLifecycle, RFQState, RequestStatus, StatusSource,
    Customer, Drafts, ItemRequest, LineItem, QuoteJob, Vendor, VendorRFQ,
)

log = logging.getLogger("mainbox.quote_jobs")

_UNSET = object()  # sentinel: distinguishes "leave unchanged" from "set to None"


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat()


def _seq_from_ids(ids, prefix: str) -> int:
    """Highest numeric suffix among ids shaped PREFIX<n>, else 0.

    Next ids are derived from existing data rather than a persisted counter, so
    there is no separate counter to fall out of sync. This assumes records are
    closed/abandoned via lifecycle rather than hard-deleted, which keeps ids
    monotonic (a deleted id would otherwise be reusable).
    """
    hi = 0
    plen = len(prefix)
    for k in ids:
        if k.startswith(prefix):
            tail = k[plen:]
            if tail.isdigit():
                hi = max(hi, int(tail))
    return hi


def normalize_subject(subject: Optional[str]) -> str:
    """THE one canonical subject normalizer.

    Per the Fork-3 decision, conversation_id is the primary reply key and
    subject is only ever a fallback. Anything that compares subjects -- here or
    in the step-5 reply matcher -- must call THIS function and no other, so the
    two-different-normalizations split that produced the v3.8.29 placeholder
    bug cannot recur. (The existing MaINbox subject logic should be pointed at
    this too, but that change is out of scope for this module.)
    """
    if not subject:
        return ""
    s = subject.strip()
    low = s.lower()
    # strip any stack of reply/forward prefixes
    while True:
        for p in ("re:", "fw:", "fwd:"):
            if low.startswith(p):
                s = s[len(p):].lstrip()
                low = s.lower()
                break
        else:
            break
    return " ".join(s.split()).lower()


def _norm_msgid(mid: Optional[str]) -> str:
    """Canonicalize an RFC Message-ID for matching: strip angle brackets and
    surrounding whitespace. (Message-IDs are case-sensitive, so case is kept.)"""
    if not mid:
        return ""
    return mid.strip().strip("<>").strip()


class QuoteJobStore:
    # ---- construction / load ------------------------------------------------
    def __init__(self, path: str, autoload: bool = True):
        self.path = path
        self._lock = threading.RLock()       # reentrant: batch + nested mutators
        self._batch_depth = 0
        self._dirty = False
        self.schema_version = SCHEMA_VERSION
        self.jobs: Dict[str, QuoteJob] = {}
        self._next_job_seq = 0
        # reverse indexes -- rebuilt on load, never written to disk
        self._idx_rfq_by_convo: Dict[str, Tuple[str, str]] = {}
        self._idx_job_by_customer_convo: Dict[str, str] = {}
        self._idx_rfq_by_message_id: Dict[str, Tuple[str, str]] = {}
        self._idx_rfq_by_subject: Dict[str, List[Tuple[str, str]]] = {}
        if autoload and os.path.exists(self.path):
            self.load()

    def load(self) -> None:
        """Read from disk. Raises on a corrupt file (never silently overwrites
        your real data) and on a newer-than-supported schema."""
        with self._lock:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            ver = data.get("schema_version", 1)
            if ver > SCHEMA_VERSION:
                raise ValueError(
                    f"quote_jobs.json schema_version {ver} is newer than supported "
                    f"{SCHEMA_VERSION}; refusing to load to avoid data loss."
                )
            self.schema_version = ver
            self.jobs = {jid: QuoteJob.from_dict(jd)
                         for jid, jd in data.get("jobs", {}).items()}
            self._next_job_seq = _seq_from_ids(self.jobs.keys(), "QJ-")
            self._rebuild_indexes()
            self._dirty = False

    # ---- indexes ------------------------------------------------------------
    def _rebuild_indexes(self) -> None:
        convo: Dict[str, Tuple[str, str]] = {}
        cust: Dict[str, str] = {}
        by_mid: Dict[str, Tuple[str, str]] = {}
        by_subj: Dict[str, List[Tuple[str, str]]] = {}
        for jid, job in self.jobs.items():
            cid = job.customer.conversation_id
            if cid:
                if cid in cust and cust[cid] != jid:
                    log.warning("customer convo %s maps to >1 job (%s,%s); last wins",
                                cid, cust[cid], jid)
                cust[cid] = jid
            for rid, rfq in job.vendor_rfqs.items():
                rcid = rfq.conversation_id
                if rcid:
                    if rcid in convo and convo[rcid] != (jid, rid):
                        log.warning("vendor convo %s maps to >1 rfq; last wins", rcid)
                    convo[rcid] = (jid, rid)
                if rfq.sent_message_id:
                    by_mid[_norm_msgid(rfq.sent_message_id)] = (jid, rid)
                if rfq.subject_norm:
                    by_subj.setdefault(rfq.subject_norm, []).append((jid, rid))
        self._idx_rfq_by_convo = convo
        self._idx_job_by_customer_convo = cust
        self._idx_rfq_by_message_id = by_mid
        self._idx_rfq_by_subject = by_subj

    def rebuild_indexes(self) -> None:
        with self._lock:
            self._rebuild_indexes()

    # ---- batching + atomic persistence -------------------------------------
    def _begin_batch(self) -> None:
        with self._lock:
            self._batch_depth += 1

    def _flush_batch(self) -> None:
        with self._lock:
            if self._batch_depth == 0:
                log.warning("flush_batch with depth 0; ignoring")
                return
            self._batch_depth -= 1
            if self._batch_depth == 0 and self._dirty:
                self._write_to_disk()
                self._dirty = False

    @contextmanager
    def batch(self) -> Iterator["QuoteJobStore"]:
        """Collapse a burst of mutations into a single disk write.

        Exception-safe: the batch is always closed on exit, so depth can never
        get stuck. Not transactional -- mutations already applied in memory are
        flushed even if the block raises. Holds the store lock for its whole
        duration, so keep batches short.
        """
        with self._lock:
            self._begin_batch()
            try:
                yield self
            finally:
                self._flush_batch()

    def save(self) -> None:
        """Persist now, or defer to the outermost batch flush if batching."""
        with self._lock:
            if self._batch_depth > 0:
                self._dirty = True
                return
            self._write_to_disk()
            self._dirty = False

    def _write_to_disk(self) -> None:
        data = {
            "schema_version": SCHEMA_VERSION,
            "jobs": {jid: job.to_dict() for jid, job in self.jobs.items()},
        }
        d = os.path.dirname(os.path.abspath(self.path)) or "."
        os.makedirs(d, exist_ok=True)
        # temp file in the SAME dir so os.replace stays on one filesystem
        fd, tmp = tempfile.mkstemp(prefix=".quote_jobs.", suffix=".tmp", dir=d)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.path)  # atomic on POSIX; replace-in-place on Windows
        except BaseException:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)      # leave no temp turd behind on failure
            except OSError:
                pass
            raise
        # best-effort directory fsync for durability (no-op / unsupported on Windows)
        try:
            dfd = os.open(d, os.O_RDONLY)
            try:
                os.fsync(dfd)
            finally:
                os.close(dfd)
        except (OSError, AttributeError, ValueError):
            pass

    # ---- internal helpers ---------------------------------------------------
    def _require_job(self, job_id: str) -> QuoteJob:
        job = self.jobs.get(job_id)
        if job is None:
            raise KeyError(f"no quote job {job_id!r}")
        return job

    def _require_rfq(self, job: QuoteJob, rfq_id: str) -> VendorRFQ:
        rfq = job.vendor_rfqs.get(rfq_id)
        if rfq is None:
            raise KeyError(f"{job.job_id} has no rfq {rfq_id!r}")
        return rfq

    def _mark_and_save(self, job: QuoteJob) -> None:
        job.updated_at = _utcnow()
        self._dirty = True
        self.save()

    def _new_job_id(self) -> str:
        self._next_job_seq += 1
        return f"QJ-{self._next_job_seq:05d}"

    # ---- mutators (record facts only) --------------------------------------
    def add_job(self, customer: Customer) -> QuoteJob:
        with self._lock:
            jid = self._new_job_id()
            now = _utcnow()
            job = QuoteJob(job_id=jid, lifecycle=JobLifecycle.OPEN,
                           created_at=now, updated_at=now,
                           customer=customer, drafts=Drafts())
            self.jobs[jid] = job
            if customer.conversation_id:
                self._idx_job_by_customer_convo[customer.conversation_id] = jid
            self._dirty = True
            self.save()
            return job

    def add_line_item(self, job_id: str, description: str,
                      canonical_key: Optional[str] = None, qty: float = 1,
                      uom: str = "EA", extraction_confidence: float = 0.0) -> LineItem:
        with self._lock:
            job = self._require_job(job_id)
            lid = f"LI-{_seq_from_ids(job.line_items.keys(), 'LI-') + 1}"
            li = LineItem(item_id=lid, description=description,
                          canonical_key=canonical_key, qty=qty, uom=uom,
                          extraction_confidence=extraction_confidence)
            job.line_items[lid] = li
            self._mark_and_save(job)
            return li

    def add_vendor_rfq(self, job_id: str, vendor: Vendor,
                       item_ids: List[str]) -> VendorRFQ:
        with self._lock:
            job = self._require_job(job_id)
            for iid in item_ids:
                if iid not in job.line_items:
                    raise KeyError(f"{job_id} has no line item {iid!r}")
            rid = f"RFQ-{_seq_from_ids(job.vendor_rfqs.keys(), 'RFQ-') + 1}"
            rfq = VendorRFQ(rfq_id=rid, vendor=vendor, state=RFQState.DRAFT)
            for iid in item_ids:
                rfq.requests[iid] = ItemRequest()  # AWAITING by default
            job.vendor_rfqs[rid] = rfq
            self._mark_and_save(job)
            return rfq

    def set_rfq_draft(self, job_id: str, rfq_id: str,
                      draft_entry_id: Optional[str],
                      conversation_id: Optional[str] = None) -> None:
        """Link the created draft. conversation_id is captured HERE (at draft
        time) so send-detection can match the sent item back to this RFQ; it is
        indexed immediately. A reply can't arrive before send, so indexing a
        draft-state RFQ's conversation is safe."""
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            rfq.draft_entry_id = draft_entry_id
            if conversation_id:
                rfq.conversation_id = conversation_id
                self._idx_rfq_by_convo[conversation_id] = (job_id, rfq_id)
            self._mark_and_save(job)

    def mark_rfq_sent(self, job_id: str, rfq_id: str, conversation_id: Optional[str],
                      sent_at: Optional[str] = None,
                      deadline_at: Optional[str] = None,
                      sent_message_id: Optional[str] = None,
                      subject: Optional[str] = None) -> None:
        """Record an RFQ as sent. `deadline_at` is the precomputed snapshot the
        caller derived from the timeout setting in effect at send time -- the
        store does not own that policy. sent_message_id + subject (if captured by
        send-detection) light up header/subject reply-matching; the subject is
        normalized through the one canonical normalizer and both are indexed."""
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            rfq.state = RFQState.SENT
            rfq.sent_at = sent_at or _utcnow()
            rfq.deadline_at = deadline_at
            if conversation_id:
                rfq.conversation_id = conversation_id
                self._idx_rfq_by_convo[conversation_id] = (job_id, rfq_id)
            if sent_message_id:
                rfq.sent_message_id = sent_message_id
                self._idx_rfq_by_message_id[_norm_msgid(sent_message_id)] = (job_id, rfq_id)
            if subject is not None:
                rfq.subject_norm = normalize_subject(subject)
                if rfq.subject_norm:
                    self._idx_rfq_by_subject.setdefault(rfq.subject_norm, [])
                    if (job_id, rfq_id) not in self._idx_rfq_by_subject[rfq.subject_norm]:
                        self._idx_rfq_by_subject[rfq.subject_norm].append((job_id, rfq_id))
            self._mark_and_save(job)

    def cancel_rfq(self, job_id: str, rfq_id: str) -> None:
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            rfq.state = RFQState.CANCELLED
            self._mark_and_save(job)

    def record_item_response(self, job_id: str, rfq_id: str, item_id: str,
                             status: RequestStatus, *,
                             price: Optional[float] = None,
                             lead_time: Optional[str] = None,
                             status_confidence: Optional[float] = None,
                             status_source: StatusSource = StatusSource.AI,
                             ai_proposed_status: Optional[RequestStatus] = None,
                             raw_excerpt: Optional[str] = None,
                             responded_at: Optional[str] = None,
                             response_entry_id: Optional[str] = None) -> None:
        """Full write of an (item, vendor) response -- typically the AI's read of
        a vendor reply. Use correct_item_response() for later user edits so the
        AI's original guess is preserved for the learning DB."""
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            req = rfq.requests.get(item_id)
            if req is None:
                raise KeyError(f"{rfq_id} has no request for item {item_id!r}")
            req.status = status
            req.price = price
            req.lead_time = lead_time
            req.status_confidence = status_confidence
            req.status_source = status_source
            if ai_proposed_status is not None:
                req.ai_proposed_status = ai_proposed_status
            elif status_source is StatusSource.AI and req.ai_proposed_status is None:
                req.ai_proposed_status = status  # remember the first AI guess
            req.raw_excerpt = raw_excerpt
            req.responded_at = responded_at or _utcnow()
            req.response_entry_id = response_entry_id
            self._mark_and_save(job)

    def correct_item_response(self, job_id: str, rfq_id: str, item_id: str, *,
                              status=_UNSET, price=_UNSET, lead_time=_UNSET,
                              status_confidence=_UNSET, raw_excerpt=_UNSET,
                              responded_at=_UNSET, response_entry_id=_UNSET) -> None:
        """User correction: partial update. Only fields you pass are changed
        (passing None clears a field; omitting it leaves it alone). Forces
        status_source=USER and never touches ai_proposed_status, so the AI's
        original call survives as the learning-DB signal."""
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            req = rfq.requests.get(item_id)
            if req is None:
                raise KeyError(f"{rfq_id} has no request for item {item_id!r}")
            if status is not _UNSET:
                req.status = status
            if price is not _UNSET:
                req.price = price
            if lead_time is not _UNSET:
                req.lead_time = lead_time
            if status_confidence is not _UNSET:
                req.status_confidence = status_confidence
            if raw_excerpt is not _UNSET:
                req.raw_excerpt = raw_excerpt
            if responded_at is not _UNSET:
                req.responded_at = responded_at
            if response_entry_id is not _UNSET:
                req.response_entry_id = response_entry_id
            req.status_source = StatusSource.USER
            self._mark_and_save(job)

    def set_rfq_overdue_alerted(self, job_id: str, rfq_id: str,
                                value: bool = True) -> None:
        with self._lock:
            job = self._require_job(job_id)
            rfq = self._require_rfq(job, rfq_id)
            rfq.overdue_alerted = value
            self._mark_and_save(job)

    def set_line_item_resolved(self, job_id: str, item_id: str,
                               value: bool = True) -> None:
        with self._lock:
            job = self._require_job(job_id)
            li = job.line_items.get(item_id)
            if li is None:
                raise KeyError(f"{job_id} has no line item {item_id!r}")
            li.resolved = value
            self._mark_and_save(job)

    def set_line_item_verified(self, job_id: str, item_id: str,
                               value: bool = True) -> None:
        with self._lock:
            job = self._require_job(job_id)
            li = job.line_items.get(item_id)
            if li is None:
                raise KeyError(f"{job_id} has no line item {item_id!r}")
            li.user_verified = value
            self._mark_and_save(job)

    def set_job_lifecycle(self, job_id: str, lifecycle: JobLifecycle) -> None:
        with self._lock:
            job = self._require_job(job_id)
            job.lifecycle = lifecycle
            self._mark_and_save(job)

    def set_drafts(self, job_id: str, ack_entry_id=_UNSET,
                   quote_entry_id=_UNSET) -> None:
        with self._lock:
            job = self._require_job(job_id)
            if ack_entry_id is not _UNSET:
                job.drafts.ack_entry_id = ack_entry_id
            if quote_entry_id is not _UNSET:
                job.drafts.quote_entry_id = quote_entry_id
            self._mark_and_save(job)

    # ---- reads / lookups (reply matching) ----------------------------------
    def get_job(self, job_id: str) -> Optional[QuoteJob]:
        with self._lock:
            return self.jobs.get(job_id)

    def find_rfq_by_conversation(self, conversation_id: str
                                 ) -> Optional[Tuple[QuoteJob, VendorRFQ]]:
        """Primary vendor-reply match: O(1) by conversation_id."""
        with self._lock:
            hit = self._idx_rfq_by_convo.get(conversation_id)
            if not hit:
                return None
            jid, rid = hit
            job = self.jobs.get(jid)
            if job is None:
                return None
            rfq = job.vendor_rfqs.get(rid)
            return (job, rfq) if rfq is not None else None

    def find_job_by_customer_conversation(self, conversation_id: str
                                          ) -> Optional[QuoteJob]:
        with self._lock:
            jid = self._idx_job_by_customer_convo.get(conversation_id)
            return self.jobs.get(jid) if jid else None

    def find_rfq_by_message_id(self, message_id: str
                               ) -> Optional[Tuple[QuoteJob, VendorRFQ]]:
        """Header-match a reply to its RFQ via the sent Message-ID (hard key)."""
        with self._lock:
            hit = self._idx_rfq_by_message_id.get(_norm_msgid(message_id))
            if not hit:
                return None
            jid, rid = hit
            job = self.jobs.get(jid)
            if job is None:
                return None
            rfq = job.vendor_rfqs.get(rid)
            return (job, rfq) if rfq is not None else None

    def find_rfqs_by_subject(self, subject_norm: str
                             ) -> List[Tuple[QuoteJob, VendorRFQ]]:
        """Subject-match candidates (SOFT -- caller must treat as a suggestion,
        never an auto-route). Returns all sent RFQs sharing the normalized subject."""
        out: List[Tuple[QuoteJob, VendorRFQ]] = []
        if not subject_norm:
            return out
        with self._lock:
            for jid, rid in self._idx_rfq_by_subject.get(subject_norm, []):
                job = self.jobs.get(jid)
                if job is None:
                    continue
                rfq = job.vendor_rfqs.get(rid)
                if rfq is not None:
                    out.append((job, rfq))
        return out

    # Subject-fallback matching itself is a step-5 concern (it needs a subject
    # snapshot whose home we haven't decided yet). The ONE canonical normalizer
    # it must use already lives here as module-level normalize_subject(); the
    # matcher will call that and surface candidates for confirmation rather than
    # writing state. Nothing subject-related is auto-applied at this layer.

    def open_jobs(self) -> List[QuoteJob]:
        with self._lock:
            return [j for j in self.jobs.values()
                    if j.lifecycle is JobLifecycle.OPEN]

    def all_jobs(self) -> List[QuoteJob]:
        with self._lock:
            return list(self.jobs.values())
