"""
MaINbox production EffectsPort (Path A: headless-but-real).

  * draft_ack / draft_vendor_rfq  -> REAL Outlook drafts (the layer-2 path, proven
    against your Outlook), returning DraftResult(entry_id, conversation_id). The
    ack threads as a reply to the customer's original mail when we have its
    entry_id; the vendor RFQ starts a new thread. NOTHING is sent.
  * the six UI/alert effects        -> append to a JSONL event log (a durable record
    of everything that fired) and print a console line; the genuine interruptions
    (response-needed, overdue, ready-to-quote) also pop a native message box.
    Surfacing these in MaINbox's own UI is Path B -- a later, focused change.
  * run_extraction                  -> an INJECTED hook, because your SmartScan is an
    embedded child app (GUI review + subprocess), not a callable. You provide
    extraction_hook(job) -> [row dicts]; this port maps those rows to ExtractedItem
    via smartscan_rows_to_items(). Without a hook it raises, loudly.

This file talks to Outlook directly (no MaINbox import). At integration you
construct it with your event-log path and your SmartScan hook.
"""
from __future__ import annotations

import json
import re
import threading
from datetime import datetime, timezone
from typing import Callable, List, Optional

from quote_jobs_controller import DraftResult, ExtractedItem

try:
    import pythoncom
    import win32com.client
    _HAVE_WIN32 = True
except Exception:
    _HAVE_WIN32 = False


def _popup(title: str, text: str) -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)
    except Exception:
        pass


def _connect_outlook():
    if not _HAVE_WIN32:
        raise RuntimeError("pywin32 not available -- run inside MaINbox's environment with Outlook open")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook, outlook.GetNamespace("MAPI")


def _to_float(v, default: float = 1.0) -> float:
    m = re.search(r"\d+(?:\.\d+)?", str(v or ""))
    return float(m.group(0)) if m else default


# default SmartScan status -> extraction_confidence. Your per-item review colors
# (green/yellow/red) map to "skip MaINbox review" vs "take another look": green
# clears the default 0.85 gate; yellow/red fall below it and get re-reviewed.
_DEFAULT_STATUS_CONF = {
    "green": 0.95, "ok": 0.95, "good": 0.95, "verified": 0.95,
    "yellow": 0.60, "review": 0.60, "warn": 0.60,
    "red": 0.30, "bad": 0.30, "error": 0.30,
}


def smartscan_rows_to_items(rows, status_confidence: Optional[dict] = None) -> List[ExtractedItem]:
    """Map SmartScan's returned rows (qty/unit/description/part_number/manufacturer/
    status) to ExtractedItem. part_number becomes the canonical_key (a good dedup
    key); status maps to extraction_confidence (unknown status -> 0.8)."""
    conf_map = status_confidence or _DEFAULT_STATUS_CONF
    out: List[ExtractedItem] = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        desc = str(r.get("description", "") or "").strip()
        part = str(r.get("part_number", "") or "").strip()
        if not desc and not part:
            continue
        st = str(r.get("status", "") or "").strip().lower()
        out.append(ExtractedItem(
            description=desc or part,
            canonical_key=(part.lower() or None),
            qty=_to_float(r.get("qty"), 1.0),
            uom=str(r.get("unit") or r.get("uom") or "EA").strip() or "EA",
            extraction_confidence=conf_map.get(st, 0.8)))
    return out


class MaINboxEffectsPort:
    def __init__(self, event_log_path: Optional[str] = None, display_drafts: bool = True,
                 extraction_hook: Optional[Callable[[object], list]] = None,
                 extraction_status_confidence: Optional[dict] = None,
                 enable_popups: bool = True):
        # event_log_path: append-only JSONL record of effects (None -> console only)
        # extraction_hook(job) -> [SmartScan row dicts]; required for run_extraction
        # enable_popups: native modal alert boxes. True for bench use; set False for
        #   headless runs, and the seam for routing alerts to the UI later (a modal
        #   MessageBox from a background thread is unsafe once the auto-trigger fires).
        self.event_log_path = event_log_path
        self.display_drafts = display_drafts
        self.extraction_hook = extraction_hook
        self.extraction_status_confidence = extraction_status_confidence
        self.enable_popups = enable_popups
        self._log_lock = threading.Lock()

    def _alert(self, title: str, text: str) -> None:
        if self.enable_popups:
            _popup(title, text)

    # ---- event log ----------------------------------------------------------
    def _log(self, effect: str, job_id: Optional[str] = None, **details) -> None:
        rec = {"ts": datetime.now(timezone.utc).isoformat(), "effect": effect,
               "job_id": job_id, **details}
        print(f"[quote-jobs] {effect}  {job_id or ''}  {details}")
        if not self.event_log_path:
            return
        line = json.dumps(rec, ensure_ascii=False)
        with self._log_lock:
            try:
                with open(self.event_log_path, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except OSError:
                pass

    # ---- drafting (real Outlook) -------------------------------------------
    def _new_mail(self, outlook, to_addr, subject, html):
        mail = outlook.CreateItem(0)
        mail.To = to_addr or ""
        mail.Subject = subject
        mail.HTMLBody = html
        return mail

    def _finish(self, mail) -> DraftResult:
        mail.Save()
        entry_id = mail.EntryID or ""
        try:
            conv = mail.ConversationID or ""
        except Exception:
            conv = ""
        if self.display_drafts:
            try:
                mail.Display()
            except Exception:
                pass
        return DraftResult(entry_id=entry_id, conversation_id=conv)

    def draft_ack(self, job) -> str:
        outlook, ns = _connect_outlook()
        body = ("<p>Thank you for your request. We have received it and will advise "
                "shortly with pricing and availability.</p>")
        mail = None
        src = getattr(job.customer, "source_entry_id", None)
        if src:                                  # thread the reply to the original mail
            try:
                original = ns.GetItemFromID(src)
                mail = original.Reply()
                mail.HTMLBody = body + (mail.HTMLBody or "")
            except Exception:
                mail = None
        if mail is None:
            mail = self._new_mail(outlook, job.customer.email,
                                  "Re: Your request for quote", body)
        res = self._finish(mail)
        self._log("draft_ack", job.job_id, to=job.customer.email, entry_id=res.entry_id)
        return res.entry_id

    def draft_vendor_rfq(self, job, vendor, item_ids) -> DraftResult:
        outlook, _ = _connect_outlook()
        rows = []
        for iid in item_ids:
            li = job.line_items.get(iid)
            if li is not None:
                rows.append(f"<li>{li.qty} {li.uom} &mdash; {li.description}</li>")
        body = ("<p>Please quote price and availability on the following:</p>"
                f"<ul>{''.join(rows)}</ul>")
        mail = self._new_mail(outlook, vendor.email, "Request for quote", body)
        res = self._finish(mail)
        self._log("draft_vendor_rfq", job.job_id, vendor=vendor.email,
                  items=list(item_ids), entry_id=res.entry_id, conversation_id=res.conversation_id)
        return res

    # ---- extraction (injected SmartScan hook) ------------------------------
    def run_extraction(self, job) -> List[ExtractedItem]:
        if self.extraction_hook is None:
            raise NotImplementedError(
                "set extraction_hook=lambda job: <your SmartScan rows>. SmartScan is "
                "your embedded child app (GUI review + subprocess); fetch the RFQ "
                "attachment(s) for job.customer.source_entry_id, run SmartScan, return "
                "its rows. This port maps them via smartscan_rows_to_items().")
        rows = self.extraction_hook(job)
        items = smartscan_rows_to_items(rows, self.extraction_status_confidence)
        self._log("run_extraction", job.job_id, items=len(items))
        return items

    # ---- UI / alert effects (logged; interruptions also popped) ------------
    def request_item_review(self, job, item_ids) -> None:
        self._log("request_item_review", job.job_id, items=list(item_ids))

    def request_vendor_selection(self, job, item_ids) -> None:
        self._log("request_vendor_selection", job.job_id, items=list(item_ids))

    def flag_item_quoted(self, job, item_id) -> None:
        self._log("flag_item_quoted", job.job_id, item=item_id)

    def notify_response_needed(self, job, rfq_id, item_ids) -> None:
        self._log("notify_response_needed", job.job_id, rfq_id=rfq_id, items=list(item_ids))
        self._alert("MaINbox \u2014 Response Needed",
               f"A vendor reply needs your attention.\nRFQ {rfq_id}, items {list(item_ids)}")

    def alert_overdue(self, job, rfq_id) -> None:
        self._log("alert_overdue", job.job_id, rfq_id=rfq_id)
        self._alert("MaINbox \u2014 Vendor Overdue",
               f"No reply on RFQ {rfq_id} past its deadline.\nFollow up with the vendor.")

    def suggest_customer_quote(self, job) -> None:
        self._log("suggest_customer_quote", job.job_id)
        self._alert("MaINbox \u2014 Ready to Quote",
               f"All items on {job.job_id} are resolved \u2014 ready to quote the customer.")
