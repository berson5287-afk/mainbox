"""
Inbound store -- the front door's own small persistence unit (like quote_jobs.json).

Holds:
  * ledger    : processed-mail record keyed by message_id  (DEDUP -- a re-scan can
                never mint a second job or double-route)
  * pending   : NEW_CUSTOMER_RFQ mails awaiting your confirm-to-mint
  * triage    : NEEDS_TRIAGE mails awaiting your manual decision
  * handoffs  : matched VENDOR_REPLY mails ready for step-5 / manual parsing

Same atomic-write discipline as the quote store (temp file -> fsync -> os.replace).
The router mutates in memory and calls save() once per handle(), so a crash before
save just means the mail is re-handled next run -- never half-recorded.
"""
from __future__ import annotations

import json
import os
import tempfile
import threading
from datetime import datetime, timezone
from typing import Dict, List, Optional

SCHEMA_VERSION = 1


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat()


class InboundStore:
    def __init__(self, path: str, autoload: bool = True):
        self.path = path
        self._lock = threading.RLock()
        self.schema_version = SCHEMA_VERSION
        self.ledger: Dict[str, dict] = {}
        self.pending: List[dict] = []
        self.triage: List[dict] = []
        self.handoffs: List[dict] = []
        self.parse_review: List[dict] = []
        if autoload and os.path.exists(self.path):
            self.load()

    # ---- persistence -------------------------------------------------------
    def load(self) -> None:
        with self._lock:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            ver = data.get("schema_version", 1)
            if ver > SCHEMA_VERSION:
                raise ValueError(
                    f"inbound store schema_version {ver} newer than supported "
                    f"{SCHEMA_VERSION}; refusing to load.")
            self.schema_version = ver
            self.ledger = dict(data.get("ledger", {}))
            self.pending = list(data.get("pending", []))
            self.triage = list(data.get("triage", []))
            self.handoffs = list(data.get("handoffs", []))
            self.parse_review = list(data.get("parse_review", []))

    def save(self) -> None:
        with self._lock:
            data = {
                "schema_version": SCHEMA_VERSION,
                "ledger": self.ledger,
                "pending": self.pending,
                "triage": self.triage,
                "handoffs": self.handoffs,
                "parse_review": self.parse_review,
            }
            d = os.path.dirname(os.path.abspath(self.path)) or "."
            os.makedirs(d, exist_ok=True)
            fd, tmp = tempfile.mkstemp(prefix=".inbound.", suffix=".tmp", dir=d)
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, self.path)
            except BaseException:
                try:
                    if os.path.exists(tmp):
                        os.remove(tmp)
                except OSError:
                    pass
                raise
            try:
                dfd = os.open(d, os.O_RDONLY)
                try:
                    os.fsync(dfd)
                finally:
                    os.close(dfd)
            except (OSError, AttributeError, ValueError):
                pass

    # ---- dedup -------------------------------------------------------------
    def is_processed(self, message_id: str) -> bool:
        with self._lock:
            return message_id in self.ledger

    def ledger_entry(self, message_id: str) -> Optional[dict]:
        with self._lock:
            return self.ledger.get(message_id)

    def record_ledger(self, message_id: str, bucket: str,
                      job_id: Optional[str] = None, rfq_id: Optional[str] = None) -> None:
        with self._lock:
            self.ledger[message_id] = {
                "bucket": bucket, "job_id": job_id, "rfq_id": rfq_id, "at": _utcnow()}

    def update_ledger(self, message_id: str, **fields) -> None:
        with self._lock:
            if message_id in self.ledger:
                self.ledger[message_id].update(fields)

    # ---- queues ------------------------------------------------------------
    def add_pending(self, entry: dict) -> None:
        with self._lock:
            self.pending.append(entry)

    def add_triage(self, entry: dict) -> None:
        with self._lock:
            self.triage.append(entry)

    def add_handoff(self, entry: dict) -> None:
        with self._lock:
            self.handoffs.append(entry)

    def pop_handoff(self, message_id: str) -> Optional[dict]:
        with self._lock:
            for i, e in enumerate(self.handoffs):
                if e.get("message_id") == message_id:
                    return self.handoffs.pop(i)
            return None

    def add_parse_review(self, entry: dict) -> None:
        with self._lock:
            self.parse_review.append(entry)

    def pop_parse_review(self, message_id: str, item_id: str) -> Optional[dict]:
        with self._lock:
            for i, e in enumerate(self.parse_review):
                if e.get("message_id") == message_id and e.get("item_id") == item_id:
                    return self.parse_review.pop(i)
            return None

    def pop_from(self, queue_name: str, message_id: str) -> Optional[dict]:
        """Remove and return the entry with this message_id from 'pending' or
        'triage'. Returns None if absent."""
        with self._lock:
            queue = getattr(self, queue_name)
            for i, e in enumerate(queue):
                if e.get("message_id") == message_id:
                    return queue.pop(i)
            return None
