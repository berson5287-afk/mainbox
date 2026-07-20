"""
A recording fake EffectsPort. Two uses:
  * the controller test harness drives the whole back half through it;
  * you can drop it into a REPL to dry-run the pipeline with zero Outlook.

It records every call, returns canned entry_ids / extraction results, and can
be told to fail on specific calls (to exercise failure isolation).
"""
from typing import Callable, List, Optional

from quote_jobs_controller import ExtractedItem, DraftResult


class RecordingPort:
    def __init__(self, extraction: Optional[List[ExtractedItem]] = None,
                 fail: Optional[Callable[..., bool]] = None):
        # extraction: what run_extraction returns
        # fail(method_name, *args) -> bool : raise inside that call if True
        self.calls = []            # list of (method_name, args_tuple)
        self.extraction = extraction or []
        self._fail = fail or (lambda *_: False)
        self._eid = 0
        self._conv = 0

    # --- introspection helpers for tests ---
    def names(self) -> List[str]:
        return [name for name, _ in self.calls]

    def args_for(self, name: str):
        return [args for n, args in self.calls if n == name]

    def _next_eid(self) -> str:
        self._eid += 1
        return f"EID-{self._eid}"

    def _record(self, name: str, *args):
        self.calls.append((name, args))
        if self._fail(name, *args):
            raise RuntimeError(f"simulated failure in {name}{args}")

    # --- the port surface ---
    def draft_ack(self, job) -> str:
        self._record("draft_ack", job.job_id)
        return self._next_eid()

    def run_extraction(self, job) -> List[ExtractedItem]:
        self._record("run_extraction", job.job_id)
        return list(self.extraction)

    def request_item_review(self, job, item_ids) -> None:
        self._record("request_item_review", job.job_id, tuple(item_ids))

    def request_vendor_selection(self, job, item_ids) -> None:
        self._record("request_vendor_selection", job.job_id, tuple(item_ids))

    def draft_vendor_rfq(self, job, vendor, item_ids) -> DraftResult:
        self._record("draft_vendor_rfq", job.job_id, vendor.name, tuple(item_ids))
        self._conv += 1
        return DraftResult(entry_id=self._next_eid(), conversation_id=f"CONV-{self._conv}")

    def flag_item_quoted(self, job, item_id) -> None:
        self._record("flag_item_quoted", job.job_id, item_id)

    def notify_response_needed(self, job, rfq_id, item_ids) -> None:
        self._record("notify_response_needed", job.job_id, rfq_id, tuple(item_ids))

    def alert_overdue(self, job, rfq_id) -> None:
        self._record("alert_overdue", job.job_id, rfq_id)

    def suggest_customer_quote(self, job) -> None:
        self._record("suggest_customer_quote", job.job_id)
