"""
LAYER 2 -- the real-Outlook walkthrough.

Run on Windows with Outlook desktop OPEN, using the same Python that runs MaINbox
(the one with pywin32):

    python layer2_walkthrough.py

It drives the controller through the whole back-half sequence. Real drafts appear
in your Outlook, tagged [MaINbox TEST] and never sent. It pauses so you can look.

Needs these files in the SAME folder (the ones you already ran the tests with):
    quote_jobs_model.py  quote_jobs_store.py  quote_jobs_engine.py
    quote_jobs_controller.py  mainbox_port.py

The store is written to your LOCAL temp dir (NOT OneDrive) on purpose.
"""
import os
import tempfile
from datetime import datetime, timezone

from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import QuoteJobController, ItemReply
from quote_jobs_model import Customer, Vendor, RequestStatus
from mainbox_port import OutlookEffectsPort


def pause(msg: str) -> None:
    input(f"\n>>> {msg}\n    (press Enter to continue) ")


def main() -> None:
    store_path = os.path.join(tempfile.gettempdir(), "qj_layer2_live.json")
    if os.path.exists(store_path):
        os.remove(store_path)
    print(f"Store (local temp, safe from OneDrive sync): {store_path}")

    ctl = QuoteJobController(QuoteJobStore(store_path), OutlookEffectsPort())

    print("\n========== STEP 1: a customer RFQ arrives ==========")
    job, follow = ctl.ingest_customer_rfq(
        Customer(name="Acme Co", email="customer@example.com"))
    print(f"created {job.job_id}; follow-on events: {[type(e).__name__ for e in follow]}")
    pause("CHECK OUTLOOK: an '[MaINbox TEST] We will advise...' draft should be open / in Drafts.")

    print("\n========== STEP 2: extraction completes (confidence gate) ==========")
    for ev in follow:                       # the ExtractionCompleted handed back to us
        ctl.dispatch(ev)
    j = ctl.store.get_job(job.job_id)
    print(f"line items now in store: {list(j.line_items)}")
    print("  (per the banners above: LI-1 high-confidence -> vendor selection; LI-2 low -> review)")

    print("\n========== STEP 3: you pick a vendor for the clear item ==========")
    item_ids = list(j.line_items)           # real ids from extraction
    ctl.select_vendor(job.job_id,
                      Vendor(name="Graybar", email="vendor@example.com"),
                      [item_ids[0]])
    j = ctl.store.get_job(job.job_id)
    rfq_id = list(j.vendor_rfqs)[-1]
    print(f"created {rfq_id}; draft linked: {j.vendor_rfqs[rfq_id].draft_entry_id!r}")
    pause("CHECK OUTLOOK: an '[MaINbox TEST] Request for quote' draft to the vendor should appear.")

    print("\n========== STEP 4: you send it, then tell MaINbox it went ==========")
    ctl.mark_rfq_sent(job.job_id, rfq_id, "CONVO-LAYER2")   # snapshots deadline = now + 30 min
    rfq = ctl.store.get_job(job.job_id).vendor_rfqs[rfq_id]
    print(f"state={rfq.state.value}  deadline_at={rfq.deadline_at}  (Fork D snapshot)")

    print("\n========== STEP 5: the vendor replies with a price ==========")
    ctl.ingest_vendor_reply(
        job.job_id, rfq_id,
        [ItemReply(item_ids[0], RequestStatus.PRICED, price=4210.0, lead_time="in stock")])
    print("  (flag_item_quoted should have fired in the banners above)")

    print("\n========== STEP 6: a tick now -> NO alert (the reply stopped the clock) ==========")
    ctl.tick(datetime.now(timezone.utc))
    print("  (no alert_overdue banner / popup here = correct)")

    print("\nDONE. Delete the two [MaINbox TEST] drafts from Outlook when finished.")
    print(f"You can also delete the store file: {store_path}")


if __name__ == "__main__":
    main()
