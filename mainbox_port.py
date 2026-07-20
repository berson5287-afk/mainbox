"""
Layer-2 TEST port -- drives REAL Outlook drafts via win32com, matching MaINbox's
own connect pattern (dynamic Dispatch, not gencache.EnsureDispatch).

Why a direct-to-Outlook port: it lets you run Layer 2 WITHOUT launching your full
MaINbox GUI, and it imports / changes NONE of your MaINbox code. Your eventual
in-app port will instead call your existing MaINbox draft routines -- this one
just proves the end-to-end path against real Outlook.

  * draft_ack / draft_vendor_rfq -> create a real Outlook draft (Save + Display),
    return its EntryID. Subjects are tagged [MaINbox TEST] so they're trivial to
    spot and delete. NOTHING is ever sent (no .Send() anywhere).
  * run_extraction -> returns two stand-in line items (this is SmartScan's job;
    stubbed so the test runs with no document). One high-confidence, one low.
  * UI / notify effects -> print a clear banner; the user-facing alerts also pop
    a native Windows message box so you see the real "alert the user" behaviour.
"""
from __future__ import annotations

from typing import List, Optional

from quote_jobs_controller import ExtractedItem, DraftResult
from send_detection import SentMail

try:
    import pythoncom
    import win32com.client
    _HAVE_WIN32 = True
except Exception:
    _HAVE_WIN32 = False


def _popup(title: str, text: str) -> None:
    """Native Windows message box, no Tk dependency. No-op if unavailable."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)  # MB_ICONINFORMATION
    except Exception:
        pass


def _connect_outlook():
    """Same dynamic-Dispatch approach MaINbox uses (avoids the stale gencache path)."""
    if not _HAVE_WIN32:
        raise RuntimeError(
            "pywin32 not available. Run this on your Windows box, with the same "
            "Python that runs MaINbox, and with Outlook desktop open.")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    return win32com.client.Dispatch("Outlook.Application")


def _banner(label: str, detail: str = "") -> None:
    line = "-" * 64
    print(f"\n{line}\n  EFFECT  {label}\n          {detail}\n{line}")


class OutlookEffectsPort:
    def __init__(self, display_drafts: bool = True):
        self.display_drafts = display_drafts

    def _make_draft(self, to_addr: str, subject: str, html_body: str):
        outlook = _connect_outlook()
        mail = outlook.CreateItem(0)                 # olMailItem
        mail.To = to_addr or ""
        mail.Subject = "[MaINbox TEST] " + subject
        mail.HTMLBody = html_body
        mail.Save()                                  # commit to Drafts -> stable EntryID + ConversationID
        entry_id = mail.EntryID or ""
        try:
            conversation_id = mail.ConversationID or ""
        except Exception:
            conversation_id = ""
        if self.display_drafts:
            try:
                mail.Display()                       # pop the inspector so you can see it
            except Exception:
                pass
        return entry_id, conversation_id

    # --- drafting effects ---
    def draft_ack(self, job) -> str:
        to = job.customer.email
        body = ("<p>Thank you for your request. We have received it and will "
                "advise shortly with pricing and availability.</p>"
                "<p><i>(MaINbox Layer-2 test draft &mdash; not sent.)</i></p>")
        eid, _conv = self._make_draft(to, "We will advise on your request", body)
        _banner("draft_ack", f"to={to}  entry_id={eid}")
        return eid

    def draft_vendor_rfq(self, job, vendor, item_ids) -> DraftResult:
        rows = []
        for iid in item_ids:
            li = job.line_items.get(iid)
            if li is not None:
                rows.append(f"<li>{li.qty} {li.uom} &mdash; {li.description}</li>")
        body = ("<p>Please quote price and availability on the following:</p>"
                f"<ul>{''.join(rows)}</ul>"
                "<p><i>(MaINbox Layer-2 test draft &mdash; not sent.)</i></p>")
        eid, conv = self._make_draft(vendor.email, "Request for quote", body)
        _banner("draft_vendor_rfq",
                f"vendor={vendor.name} <{vendor.email}>  items={list(item_ids)}  "
                f"entry_id={eid}  conversation_id={conv}")
        return DraftResult(entry_id=eid, conversation_id=conv)

    # --- extraction (SmartScan stand-in) ---
    def run_extraction(self, job) -> List[ExtractedItem]:
        items = [
            ExtractedItem("500MCM THHN CU, black", canonical_key="thhn-500-cu-blk",
                          qty=300, uom="FT", extraction_confidence=0.95),
            ExtractedItem('2" EMT conduit', canonical_key="emt-2in",
                          qty=50, uom="EA", extraction_confidence=0.40),
        ]
        _banner("run_extraction",
                f"returned {len(items)} stand-in items (SmartScan stub: LI-1 high, LI-2 low)")
        return items

    # --- UI / notification effects ---
    def request_item_review(self, job, item_ids) -> None:
        _banner("request_item_review", f"LOW-CONFIDENCE, need your review: {list(item_ids)}")

    def request_vendor_selection(self, job, item_ids) -> None:
        _banner("request_vendor_selection", f"pick / add vendor(s) for: {list(item_ids)}")

    def flag_item_quoted(self, job, item_id) -> None:
        _banner("flag_item_quoted", f"item {item_id} is now QUOTED")

    def notify_response_needed(self, job, rfq_id, item_ids) -> None:
        msg = f"A vendor reply needs your attention.\nRFQ {rfq_id}, items {list(item_ids)}"
        _banner("notify_response_needed", msg.replace("\n", "  "))
        _popup("MaINbox \u2014 Response Needed", msg)

    def alert_overdue(self, job, rfq_id) -> None:
        msg = f"No reply on RFQ {rfq_id} past its deadline.\nFollow up with the vendor."
        _banner("alert_overdue", msg.replace("\n", "  "))
        _popup("MaINbox \u2014 Vendor Overdue", msg)

    def suggest_customer_quote(self, job) -> None:
        msg = f"All items on {job.job_id} are resolved \u2014 ready to quote the customer."
        _banner("suggest_customer_quote", msg)
        _popup("MaINbox \u2014 Ready to Quote", msg)


def scan_sent_items_skeleton(controller, lookback_minutes: int = 120, max_items: int = 50):
    """Reference for the real (manually-invoked) send-detection scan.

    Pull recent Sent Items, build a SentMail per item, and hand each to
    controller.capture_sent(); capture is idempotent, so re-scanning is safe.
    Wiring this to run on a timer / on Outlook's ItemSend event is the deferred
    auto-trigger step -- keep it manual for now.
    """
    outlook = _connect_outlook()
    ns = outlook.GetNamespace("MAPI")
    sent = ns.GetDefaultFolder(5)          # olFolderSentMail
    items = sent.Items
    items.Sort("[SentOn]", True)           # newest first
    captured = []
    seen = 0
    for it in items:
        seen += 1
        if seen > max_items:
            break
        try:
            mail = SentMail(
                message_id=getattr(it, "PropertyAccessor").GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x1035001F"),  # PR_INTERNET_MESSAGE_ID
                conversation_id=getattr(it, "ConversationID", None),
                subject=getattr(it, "Subject", "") or "",
                to=getattr(it, "To", "") or "",
                sent_on=None,              # TODO: convert it.SentOn (a COM time) to a datetime
                entry_id=getattr(it, "EntryID", None),
            )
        except Exception:
            continue
        hit = controller.capture_sent(mail)
        if hit:
            captured.append(hit)
    return captured
